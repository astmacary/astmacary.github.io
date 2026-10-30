/* Roster display stays within each ministry's existing SharePoint access.
   Manual Leadership entries are authenticated administrator events in the Serve board. */
const Rosters = {
  open: new Set(), data: {}, errors: {}, tasks: {}, identities: new Map(), photos: new Map(), extraEditors: {}, saving: false, profiles: {}, editingProfile: false, profileDraft: null, profileSaving: false, deleteKey: '', leadershipChanges: new Map(), demoLeadershipEvents: [],
  baseHubAdmins: null, hubAdminTask: null,
  async loadHubAdmins() {
    if (DEMO) return;
    if (this.hubAdminTask) return this.hubAdminTask;
    this.baseHubAdmins ||= [...CHURCH_CAL.editors];
    this.hubAdminTask = (async () => {
      try {
        const base = await Comm.list({ ...CHURCH_CAL, key: 'serve' }, 'Team Board');
        const raw = await API.all(base + '?$expand=fields($select=Title,Body,ParentId)&$top=200');
        const rows = raw.map(r => Comm.normalize(r)).filter(r => r.parent.startsWith('hub-admin:'));
        await Comm.resolveAuthors(rows);
        CHURCH_CAL.editors = this.adminSettings(rows);
      } catch (_) { /* Existing administrators remain available if settings cannot be read. */ }
    })().finally(() => { this.hubAdminTask = null; });
    return this.hubAdminTask;
  },
  adminSettings(events) {
    const base = (this.baseHubAdmins || CHURCH_CAL.editors).map(lc), current = new Set(base);
    const trusted = events.filter(r => base.includes(lc(r.email))).sort((a,b) => a.at.localeCompare(b.at) || Number(a.id) - Number(b.id));
    for (const r of trusted) try {
      const p = JSON.parse(r.text), email = lc(p.email);
      if (!email || r.parent !== 'hub-admin:' + email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) continue;
      if (p.action === 'add') current.add(email); else if (p.action === 'remove' && !base.includes(email)) current.delete(email);
    } catch (_) {}
    return [...current];
  },
  hubAdmins() { return CHURCH_CAL.editors.map(lc); },
  isHubAdmin(p) { return (p.emails || [p.email]).filter(Boolean).some(e => this.hubAdmins().includes(lc(e))); },
  canManage() { return DEMO || !!S.me?.emails.some(e => this.hubAdmins().includes(lc(e))); },
  canManageTeam(key) {
    if (key === 'leadership') return this.canManage();
    const t = TEAMS.find(t => t.key === key);
    return !!t && (DEMO || !!S.me?.emails.some(e => [...t.editors, ...(this.extraEditors[key] || [])].map(lc).includes(lc(e))));
  },
  async ensureTeamAdmin(key) {
    const t = TEAMS.find(t => t.key === key);
    if (!t) throw new Error('Team not found.');
    if (!DEMO) await this.adminEmails(t);
    if (!this.canManageTeam(key)) throw new Error('Only this ministry’s administrators can manage its roster.');
    return t;
  },
  star(p) { return this.isHubAdmin(p) ? '<span class="roster-star" title="RCI Serve administrator" aria-label="RCI Serve administrator">★</span>' : ''; },
  unique(people) {
    const found = new Map();
    for (const p of people) {
      if (!p.name) continue;
      const key = p.userId || lc(p.email) || 'name:' + lc(p.name);
      if (!found.has(key)) found.set(key, { ...p, emails: [...new Set([...(p.emails || []), p.email].filter(Boolean).map(lc))] });
      else { const old = found.get(key); old.emails = [...new Set([...old.emails, ...(p.emails || []), p.email].filter(Boolean).map(lc))]; if (!old.photoRaw && p.photoRaw) Object.assign(old, { photoRaw: p.photoRaw, team: p.team, rowId: p.rowId }); }
    }
    return [...found.values()].sort((a,b) => a.name.localeCompare(b.name));
  },
  async identity(email) {
    email = lc(email);
    if (!this.identities.has(email)) this.identities.set(email, (async () => {
      const quoted = email.replace(/'/g, "''");
      const rows = await API.all('/users?$select=id,displayName,mail,userPrincipalName&$filter=' + encodeURIComponent(`mail eq '${quoted}' or userPrincipalName eq '${quoted}'`));
      if (rows.length !== 1) throw new Error('No single Microsoft account matches this sign-in email.');
      const u = rows[0];
      return { userId: u.id, name: u.displayName, email, emails: [email, u.mail, u.userPrincipalName].filter(Boolean).map(lc) };
    })().catch(e => { this.identities.delete(email); throw e; }));
    return this.identities.get(email);
  },
  async ministry(t) {
    const d = S.teams[t.key];
    if (!d?.accessVerified || d.error || d.noAccess || !myTeams().some(x => x.key === t.key)) throw new Error('This ministry roster is not available for your account.');
    const meta = API.sites[t.key];
    if (!meta) throw new Error('Refresh your teams and try again.');
    const fields = ['Photo', ...(t.key === 'worship' ? ['ActiveStatus'] : [])];
    const rows = await API.all(`/sites/${meta.siteId}/lists/${meta.listId}/items?$expand=fields($select=${fields.join(',')})&$top=500`);
    const byId = new Map(rows.map(r => [String(r.id), r.fields || {}]));
    return this.unique(d.items.filter(p => p.kind === 'profile' && !p.gone && p.fullName && (t.key !== 'worship' || lc(byId.get(p.rowId)?.ActiveStatus) !== 'non active')).map(p => ({
      name: p.fullName, email: p.email, rowId: p.rowId, team: t.key, photoRaw: byId.get(p.rowId)?.Photo || ''
    })));
  },
  async adminEmails(t) {
    let extra = [];
    try {
      let siteId = API.sites[t.key]?.siteId;
      if (!siteId) siteId = (await API.g(`/sites/${CONFIG.siteHost}:${t.sitePath}?$select=id`)).id;
      const drives = await API.g(`/sites/${siteId}/drives?$select=id,name,webUrl`);
      const drive = drives.value?.find(d => /\/SiteAssets$/i.test(d.webUrl || '') || /site assets/i.test(d.name || ''))
        || await API.g(`/sites/${siteId}/lists/${encodeURIComponent('Site Assets')}/drive?$select=id`);
      const file = await API.g(`/drives/${drive.id}/root:/app/admins.json:/content`);
      if (Array.isArray(file)) extra = file.map(lc);
    } catch (_) { /* Built-in admins remain the fallback if the optional file is unavailable. */ }
    this.extraEditors[t.key] = extra;
    return [...t.editors, ...extra].map(lc);
  },
  async manual() {
    const base = await Comm.list({ ...CHURCH_CAL, key: 'serve' }, 'Team Board');
    const raw = await API.all(base + '?$expand=fields($select=Title,Body,ParentId)&$top=200');
    const records = raw.map(r => Comm.normalize(r));
    this.readProfiles(records.filter(r => r.parent.startsWith('hub-profile:')));
    const events = records.filter(r => r.parent.startsWith('leadership:'));
    await Comm.resolveAuthors(events);
    return this.manualPeople(events);
  },
  manualPeople(events) {
    const latest = new Map();
    const trusted = events.filter(r => this.hubAdmins().includes(lc(r.email))).sort((a,b) => a.at.localeCompare(b.at) || Number(a.id) - Number(b.id));
    for (const r of trusted) {
      try {
        const p = JSON.parse(r.text);
        if (!p.userId || r.parent !== 'leadership:' + p.userId || !p.name || !p.email) continue;
        latest.set(p.userId, p);
      } catch (_) {}
    }
    this.leadershipChanges = latest;
    return [...latest.values()].filter(p => p.action === 'add').map(p => ({ ...p, manual: true }));
  },
  async leadership() {
    if (DEMO) { this.manualPeople(this.demoLeadershipEvents); return this.applyProfiles([
      { userId: S.me?.id || 'demo-self', name: S.me?.name || 'Demo Leader', email: S.me?.emails[0] || CHURCH_CAL.editors[0], emails: [...(S.me?.emails || []), CHURCH_CAL.editors[0]] },
      { userId: 'demo-other', name: 'Ministry Leader', email: 'leader@example.org' }
    ].filter(p => this.leadershipChanges.get(p.userId)?.action !== 'remove')); }
    await this.loadHubAdmins();
    const directories = Promise.allSettled(myTeams().filter(t => S.teams[t.key]?.accessVerified && !S.teams[t.key]?.error && !S.teams[t.key]?.noAccess).map(t => this.data[t.key] ? Promise.resolve(this.data[t.key]) : this.ministry(t)));
    const lists = await Promise.all(TEAMS.map(t => this.adminEmails(t)));
    const emails = [...new Set([...this.hubAdmins(), ...lists.flat()])];
    const resolved = await Promise.allSettled(emails.map(e => this.identity(e)));
    let manual = [];
    try { manual = await this.manual(); } catch (_) { this.errors.leadership = 'Additional Leadership members could not be loaded. Refresh to try again.'; }
    const missing = resolved.filter(x => x.status === 'rejected').length;
    if (missing) this.errors.leadership = 'Some administrator names could not be matched to Microsoft accounts. Refresh to try again.';
    const profiles = (await directories).filter(r => r.status === 'fulfilled').flatMap(r => r.value);
    const people = this.unique([...resolved.filter(x => x.status === 'fulfilled').map(x => x.value), ...manual]).filter(p => this.leadershipChanges.get(p.userId)?.action !== 'remove');
    for (const p of people) {
      const profile = profiles.find(profile => profile.email && (p.emails || [p.email]).includes(lc(profile.email)) && profile.photoRaw);
      if (profile) Object.assign(p, { photoRaw: profile.photoRaw, team: profile.team, rowId: profile.rowId });
    }
    // Identical names share one display card; their separate account identities retain all permissions.
    const display = new Map();
    for (const p of people) {
      const key = lc(p.name).trim();
      if (!display.has(key)) display.set(key, { ...p, userIds: [p.userId] });
      else { const old = display.get(key); old.userIds.push(p.userId); old.emails = [...new Set([...(old.emails || [old.email]), ...(p.emails || [p.email])])]; if (!old.photoRaw && p.photoRaw) Object.assign(old, { photoRaw: p.photoRaw, team: p.team, rowId: p.rowId }); }
    }
    return this.applyProfiles([...display.values()]);
  },
  async load(key, refresh = false) {
    if (this.tasks[key]) return this.tasks[key];
    if (!refresh && this.data[key]) return;
    delete this.errors[key];
    const identity = S.me?.id || S.me?.emails.join("|");
    this.tasks[key] = (async () => {
      await Promise.resolve();
      try {
        const t = TEAMS.find(t => t.key === key);
        const people = key === 'leadership' ? await this.leadership() : DEMO
          ? [{ rowId: 'demo-1', name: 'Demo Person', email: S.me.emails[0], team: key }, { rowId: 'demo-2', name: 'Team Member', email: 'member@example.org', team: key }]
          : await this.ministry(t);
        if (identity === (S.me?.id || S.me?.emails.join("|"))) this.data[key] = people;
      } catch (e) { this.errors[key] = e.message || 'Unable to load the roster.'; }
      finally { delete this.tasks[key]; if (S.phase === 'ready' && ['teams','profile'].includes(S.view) && !this.profileSaving) render(); }
    })();
    return this.tasks[key];
  },
  toggle(key) {
    if (this.open.has(key)) this.open.delete(key);
    else { this.open.add(key); void this.load(key); }
    render();
  },
  person(p, index, key) {
    return `<li class="roster-person"><span class="roster-avatar" data-roster-photo="${esc(key + ':' + index)}">${esc(initials(p.name))}</span>${this.star(p)}<b>${esc(p.name)}</b></li>`;
  },
  count(key) {
    if (key === 'leadership') return this.data.leadership?.length;
    if (DEMO) return this.data[key]?.length ?? 2;
    return this.unique((S.teams[key]?.items || []).filter(p => p.kind === 'profile' && !p.gone && p.fullName && (key !== 'worship' || lc(p.active) !== 'non active')).map(p => ({name:p.fullName,email:p.email}))).length;
  },
  icons: {
    leadership: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"><path d="m3 6 4 4 5-6 5 6 4-4-2 14H5L3 6Z"/><path d="M6 16h12"/></svg>',

    add: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg>',
    remove: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18M9 6V3h6v3M5 6l1 15h12l1-15M10 10v7M14 10v7"/></svg>'
  },
  refreshIcon(key,label) {
    return `<button class="iconbtn roster-refresh-small" data-act="roster-refresh" data-k="${key}" aria-label="Refresh ${esc(label)} roster" title="Refresh roster">${ICON.refresh}</button>`;
  },
  controls(key, label) {
    const manage = this.canManageTeam(key), t = TEAMS.find(t => t.key === key);
    return `<div class="roster-actions">${manage ? (key === 'leadership'
      ? `<button class="iconbtn sm" data-act="leadership-add" aria-label="Add RCI Leadership member" title="Add member">${this.icons.add}</button>`
      : `<a class="iconbtn sm" href="${t.path}?roster=add${DEMO ? '&amp;demo' : ''}" aria-label="Add ${esc(label)} member" title="Add member">${this.icons.add}</a>`) + `<button class="iconbtn sm danger" data-act="roster-delete" data-k="${key}" aria-label="Delete ${esc(label)} member" title="Delete member">${this.icons.remove}</button>` : ''}</div>`;
  },
  section(key) {
    const opened = this.open.has(key), people = this.data[key], label = key === 'leadership' ? 'RCI Leadership' : TEAMS.find(t => t.key === key)?.full || key;
    return `<div class="roster-section"><button class="roster-toggle" data-act="roster-toggle" data-k="${key}" aria-label="${esc(label)} team roster" aria-expanded="${opened}" aria-controls="roster-${key}"><span>Team roster · ${this.count(key) ?? '—'}</span><span class="roster-chevron" aria-hidden="true"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="m6 9 6 6 6-6"/></svg></span></button>${opened ? `<div id="roster-${key}" class="roster-content">
      ${this.tasks[key] ? '<p class="hint" role="status">Loading roster…</p>' : ''}
      ${this.errors[key] ? `<p class="hint" role="alert">${esc(this.errors[key])}</p>` : ''}
      ${people?.length ? `<ul class="team-roster">${people.map((p,i) => this.person(p,i,key)).join('')}</ul>` : !this.tasks[key] && !this.errors[key] ? '<p class="hint">No team members yet.</p>' : ''}
    </div>` : ''}</div>`;
  },
  modal() {
    if (this.adding && this.canManage()) return `<div class="roster-overlay"><section class="panel roster-dialog" role="dialog" aria-modal="true" aria-labelledby="roster-dialog-title"><h2 id="roster-dialog-title">Add Leadership member</h2><form id="leadership-form" class="sec"><label for="leadership-email">Sign-in email</label><input id="leadership-email" type="email" name="email" value="${esc(this.draft || '')}" required autocomplete="email"><p class="hint">Adds their name and photo to Leadership. App access stays the same.</p><div class="acts2"><button class="btn primary" ${this.saving ? 'disabled' : ''}>${this.saving ? 'Adding…' : 'Add member'}</button><button class="btn" type="button" data-act="leadership-cancel">Cancel</button></div></form></section></div>`;
    if (!this.deleteKey || !this.canManageTeam(this.deleteKey)) return '';
    const key = this.deleteKey, label = key === 'leadership' ? 'RCI Leadership' : TEAMS.find(t => t.key === key)?.full, people = this.data[key] || [];
    return `<div class="roster-overlay"><section class="panel roster-dialog" role="dialog" aria-modal="true" aria-labelledby="roster-dialog-title"><h2 id="roster-dialog-title">Delete ${esc(label)} member</h2><form id="roster-delete-form" class="sec">
      ${this.tasks[key] ? '<p role="status">Loading roster…</p>' : ''}${this.errors[key] ? `<p role="alert">${esc(this.errors[key])}</p>` : ''}
      <label for="roster-member">Team member</label><select id="roster-member" name="member" required><option value="">Choose a member</option>${people.map(p => `<option value="${esc(key === 'leadership' ? p.userId : p.rowId)}" ${this.deleteSelection === (key === 'leadership' ? p.userId : p.rowId) ? 'selected' : ''}>${esc(p.name)}</option>`).join('')}</select>
      <p class="hint">${key === 'leadership' ? 'Removes this person from Leadership, including automatic administrator entries.' : 'Deletes this person’s directory entry. Existing schedule assignments stay in place.'} Their Microsoft account and app permissions stay the same.</p>
      <div class="acts2"><button class="btn danger" ${this.saving || !people.length ? 'disabled' : ''}>${this.saving ? 'Deleting…' : 'Delete member'}</button><button type="button" class="btn" data-act="roster-delete-cancel">Cancel</button></div></form></section></div>`;
  },
  view() {
    const teams = myTeams();
    if (!this.data.leadership && !this.errors.leadership){
      void this.load('leadership');
      return '<h1 class="page">Your Teams</h1><p class="hint" role="status">Loading teams and roster counts…</p>';
    }
    return `<h1 class="page">Your Teams</h1><div class="team-containers"><section class="panel leadership-panel"><div class="team-heading"><div class="team-title"><span class="t-ic leadership-icon">${this.icons.leadership}</span><b>RCI Leadership</b>${this.refreshIcon('leadership','RCI Leadership')}</div>${this.controls('leadership','RCI Leadership')}</div>${this.section('leadership')}</section>
      ${teams.map(t => `<section class="panel"><div class="team-heading team" style="${tvars(t)}"><div class="team-title"><a class="team-name" href="${t.path}?home${DEMO ? '&amp;demo' : ''}"><span class="t-ic">${ICON[t.icon]}</span><b>${esc(t.full)}</b></a>${this.refreshIcon(t.key,t.full)}</div>${this.controls(t.key,t.full)}</div>${this.section(t.key)}</section>`).join('')}</div>${this.modal()}`;
  },
  async deleteMember(key, person) {
    if (key === 'leadership') { await this.save(person, 'remove'); return; }
    const t = await this.ensureTeamAdmin(key), meta = API.sites[key];
    if (!person?.rowId || (!DEMO && !meta) || !this.data[key]?.some(p => p.rowId === person.rowId)) throw new Error('Refresh the roster before deleting a member.');
    if (!DEMO) {
      const url = `/sites/${meta.siteId}/lists/${meta.listId}/items/${encodeURIComponent(person.rowId)}`;
      const fresh = await API.g(url + '?$expand=fields');
      const normalized = normalize(t, String(fresh.id), fresh.fields || {}, {}, meta.col || {});
      if (!normalized.some(p => p.kind === 'profile' && lc(p.fullName) === lc(person.name))) throw new Error('This directory entry changed. Refresh and try again.');
      const etag = fresh.eTag || fresh['@odata.etag'];
      if (!etag) throw new Error('Unable to verify the current directory entry. Refresh and try again.');
      await API.g(url, { method: 'DELETE', headers: { 'If-Match': etag } });
    }
    S.teams[key].items = S.teams[key].items.filter(p => p.rowId !== person.rowId);
    this.data[key] = this.data[key].filter(p => p.rowId !== person.rowId);
  },
  async photo(p) {
    if (p.hubPhoto !== undefined) return p.hubPhoto || null;
    if (DEMO) return null;
    const key = [p.team, p.rowId, p.photoRaw, p.userId, p.email].join('|');
    if (!this.photos.has(key)) this.photos.set(key, (async () => {
      if (p.photoRaw && p.team) try {
        const f = typeof p.photoRaw === 'string' ? JSON.parse(p.photoRaw) : p.photoRaw;
        const t = TEAMS.find(t => t.key === p.team), meta = API.sites[p.team];
        const q = v => encodeURIComponent(String(v).replace(/'/g, "''"));
        if (f.serverRelativeUrl) return await API.spBlob(t, `GetFileByServerRelativeUrl('${q(f.serverRelativeUrl)}')/$value`);
        if (f.fileName) return await API.spBlob(t, `lists(guid'${meta.listId}')/items(${p.rowId})/AttachmentFiles('${q(f.fileName)}')/$value`);
      } catch (_) {}
      try {
        const person = p.userId ? p : await this.identity(p.email);
        const r = await API.g(`/users/${encodeURIComponent(person.userId)}/photo/$value`, { raw: true });
        return URL.createObjectURL(await r.blob());
      } catch (_) { return null; }
    })());
    return this.photos.get(key);
  },
  loadTeamRoles() {
    if (DEMO || this.rolesLoaded) return;
    this.rolesLoaded = true;
    void Promise.all(myTeams().map(t => this.adminEmails(t))).then(() => {
      if (S.phase === 'ready' && S.view === 'teams' && !this.saving && !document.activeElement?.closest('input,select,textarea')) render();
    });
  },
  hydrate() {
    this.loadTeamRoles();
    document.querySelectorAll('[data-roster-photo]').forEach(async el => {
      const [key, index] = el.dataset.rosterPhoto.split(':');
      const p = this.data[key]?.[Number(index)]; if (!p) return;
      const url = await this.photo(p);
      if (!url || !el.isConnected) return;
      const img = document.createElement('img'); img.src = url; img.alt = ''; el.replaceChildren(img);
    });
  },
  async save(person, action) {
    if (!this.canManage()) throw new Error('Only RCI Serve administrators can change Leadership members.');
    const ids = [...new Set(person.userIds || [person.userId])];
    if (!ids.length || ids.some(id => !id)) throw new Error('Refresh the roster and try again.');
    for (const userId of ids) {
      const entry = { userId, name: person.name, email: person.email, action };
      if (DEMO) this.demoLeadershipEvents.push({ id: String(this.demoLeadershipEvents.length + 1), email: this.hubAdmins()[0], parent: 'leadership:' + userId, text: JSON.stringify(entry), at: new Date().toISOString() });
      else {
        const base = await Comm.list({ ...CHURCH_CAL, key: 'serve' }, 'Team Board');
        await API.g(base, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fields: { Title: 'Leadership roster', ParentId: 'leadership:' + userId, Body: JSON.stringify(entry) } }) });
      }
    }
    await this.load('leadership', true);
  },
  readProfiles(events) {
    const profiles = {};
    for (const r of [...events].sort((a,b) => a.at.localeCompare(b.at) || Number(a.id)-Number(b.id))) try {
      const p = JSON.parse(r.text);
      if (!p.userId || r.authorId !== p.userId || r.parent !== 'hub-profile:' + p.userId || typeof p.name !== 'string' || !p.name.trim() || p.name.length > 120) continue;
      if (p.photo !== undefined && p.photo !== null && (typeof p.photo !== 'string' || p.photo.length > 40000 || !/^data:image\/jpeg;base64,[A-Za-z0-9+/]+=*$/.test(p.photo))) continue;
      profiles[p.userId] = { userId:p.userId, name:p.name.trim(), ...(p.photo !== undefined ? {photo:p.photo} : {}), at:r.at };
    } catch (_) {}
    this.profiles = profiles;
    return profiles;
  },
  applyProfiles(people) {
    return people.map(p => {
      const settings = [...new Set(p.userIds || [p.userId])].map(id => this.profiles[id]).filter(Boolean).sort((a,b) => a.at.localeCompare(b.at)).at(-1);
      return settings ? {...p,name:settings.name,...(settings.photo !== undefined ? {hubPhoto:settings.photo} : {})} : p;
    });
  },
  canEditProfile() {
    return this.canManage() || !!this.data.leadership?.some(p => p.userId === S.me?.id || p.userIds?.includes(S.me?.id) || (p.emails || [p.email]).some(e => S.me?.emails.includes(lc(e))));
  },
  ownProfile() {
    const stored = this.editingProfile ? this.profileDraft : this.profiles[S.me?.id];
    return {name:stored?.name || S.me?.name || '',photo:stored && Object.prototype.hasOwnProperty.call(stored,'photo') ? stored.photo || '' : S.photo};
  },
  profileEditor() {
    if (!this.canEditProfile()) return '<p class="hint">Your ministry profile is edited in that team’s app.</p>';
    if (!this.editingProfile) return '';
    return `<form id="hub-profile-form" class="panel sec hub-profile-editor"><label for="hub-profile-name">Display name</label><input id="hub-profile-name" name="name" maxlength="120" required value="${esc(this.profileDraft?.name || '')}"><p class="hint">Your name and photo appear on the RCI Leadership roster.</p><div class="acts2"><button class="btn" type="button" data-act="hub-profile-photo-remove">Remove photo</button><button class="btn primary" ${this.profileSaving ? 'disabled' : ''}>${this.profileSaving ? 'Saving…' : 'Save profile'}</button><button class="btn" type="button" data-act="hub-profile-cancel">Cancel</button></div></form>`;
  },
  editProfile() {
    if (!this.canEditProfile()) return;
    this.profileDraft = {name:this.ownProfile().name,...(this.profiles[S.me?.id]?.photo !== undefined ? {photo:this.profiles[S.me.id].photo} : {})};
    this.editingProfile = true; render();
  },
  async saveProfile(name) {
    if (!this.canEditProfile() || !S.me?.id) throw new Error('Leadership members can edit only their own profile.');
    name = String(name || '').trim();
    if (!name || name.length > 120) throw new Error('Enter a name up to 120 characters.');
    const p = {userId:S.me.id,name,...(this.profileDraft?.photo !== undefined ? {photo:this.profileDraft.photo} : {})};
    if (p.photo !== undefined && p.photo !== null && (p.photo.length > 40000 || !/^data:image\/jpeg;base64,[A-Za-z0-9+/]+=*$/.test(p.photo))) throw new Error('Choose a smaller profile photo.');
    if (!DEMO){
      const base = await Comm.list({...CHURCH_CAL,key:'serve'},'Team Board');
      await API.g(base,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({fields:{Title:'Leadership profile',ParentId:'hub-profile:'+S.me.id,Body:JSON.stringify(p)}})});
    }
    this.profiles[S.me.id] = {...p,at:new Date().toISOString()};
    if (this.data.leadership) this.data.leadership = this.applyProfiles(this.data.leadership);
    this.editingProfile = false; this.profileDraft = null;
  },
  async choosePhoto(file) {
    if (!this.canEditProfile() || !this.editingProfile || !file) return;
    if (!file.type.startsWith('image/')) throw new Error('Choose an image file.');
    const url = URL.createObjectURL(file);
    try {
      const img = new Image(); img.src = url; await img.decode();
      const canvas = document.createElement('canvas'); canvas.width = 256; canvas.height = 256;
      const size = Math.min(img.naturalWidth,img.naturalHeight); if (!size) throw new Error('Unable to open this photo.');
      canvas.getContext('2d').drawImage(img,(img.naturalWidth-size)/2,(img.naturalHeight-size)/2,size,size,0,0,256,256);
      let photo = canvas.toDataURL('image/jpeg',.8);
      if (photo.length > 40000) photo = canvas.toDataURL('image/jpeg',.5);
      if (photo.length > 40000) throw new Error('Choose a smaller photo.');
      this.profileDraft.photo = photo; render();
    } finally { URL.revokeObjectURL(url); }
  },
  reset() { this.open.clear(); this.data = {}; this.errors = {}; this.identities.clear(); this.photos.clear(); this.adding = false; this.draft = ''; this.deleteKey = ''; this.demoLeadershipEvents = []; this.rolesLoaded = false; this.profiles = {}; this.editingProfile = false; this.profileDraft = null; if (this.baseHubAdmins) CHURCH_CAL.editors = [...this.baseHubAdmins]; }
};
document.addEventListener('submit', async e => {
  if (e.target.id !== 'leadership-form') return;
  e.preventDefault(); if (Rosters.saving || !Rosters.canManage()) return;
  Rosters.draft = String(new FormData(e.target).get('email') || '').trim();
  Rosters.saving = true; render();
  try { const p = await Rosters.identity(Rosters.draft); await Rosters.save(p, 'add'); Rosters.adding = false; Rosters.draft = ''; toast('Leadership member added'); }
  catch (err) { toast(err.message || 'Unable to add this member.', true); }
  Rosters.saving = false; render();
});

document.addEventListener('submit', async e => {
  if (e.target.id !== 'roster-delete-form') return;
  e.preventDefault(); if (Rosters.saving || !Rosters.canManageTeam(Rosters.deleteKey)) return;
  const key = Rosters.deleteKey, id = String(new FormData(e.target).get('member') || '');
  const person = Rosters.data[key]?.find(p => (key === 'leadership' ? p.userId : p.rowId) === id);
  if (!person) return;
  Rosters.deleteSelection = id; Rosters.saving = true; render();
  try { await Rosters.deleteMember(key, person); Rosters.deleteKey = ''; toast('Member deleted from ' + (key === 'leadership' ? 'Leadership' : 'the team roster')); }
  catch (err) { toast(err.message || 'Unable to delete this member.', true); }
  Rosters.saving = false; render();
});

document.addEventListener('input', e => {
  if (e.target.id === 'hub-profile-name' && Rosters.profileDraft) Rosters.profileDraft.name = e.target.value;
});
document.addEventListener('change', async e => {
  if (e.target.id !== 'hub-profile-photo') return;
  try { await Rosters.choosePhoto(e.target.files?.[0]); } catch(err) { toast(err.message,true); }
});
document.addEventListener('submit', async e => {
  if (e.target.id !== 'hub-profile-form') return;
  e.preventDefault(); if (Rosters.profileSaving || !Rosters.canEditProfile()) return;
  const name = String(new FormData(e.target).get('name') || '');
  Rosters.profileSaving = true; render();
  try { await Rosters.saveProfile(name); toast('Profile updated'); } catch(err) { toast(err.message,true); }
  Rosters.profileSaving = false; render();
});
