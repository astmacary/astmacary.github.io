/* Communication records stay in their own ministry's SharePoint site.
   SharePoint permissions, not these UI checks, enforce access. See COMMUNICATIONS.md. */
/* Rich-text formatting for messages: mirrors the Production Notes editor's toolbar/colors/sanitizer
   pattern, but with its own names (COMM_COLORS/commToolbar/commRanges/commPreview) so this file never
   collides with the page-global KEEP/cleanHtml/NOTES_COLORS/notesToolbar some apps already define.
   It still calls the page's own global cleanHtml()/KEEP (present in every app) to sanitize. */
const COMM_COLORS = [
  { label: "Red", hex: "#C0392B" },
  { label: "Orange", hex: "#D98E04" },
  { label: "Green", hex: "#1E8449" },
  { label: "Blue", hex: "#2E5EAA" },
  { label: "Purple", hex: "#6C3483" }
];
function commToolbar(target){
  return `<div class="notes-tools" role="toolbar" aria-label="Formatting">
      <button type="button" class="tool" data-comm="comm-fmt" data-target="${target}" data-cmd="bold" aria-label="Bold"><b>B</b></button>
      <button type="button" class="tool" data-comm="comm-fmt" data-target="${target}" data-cmd="italic" aria-label="Italic"><i>I</i></button>
      <button type="button" class="tool" data-comm="comm-fmt" data-target="${target}" data-cmd="insertUnorderedList" aria-label="Bulleted list">• List</button>
      <button type="button" class="tool" data-comm="comm-fmt" data-target="${target}" data-cmd="insertOrderedList" aria-label="Numbered list">1. List</button>
      <button type="button" class="tool" data-comm="comm-fmt" data-target="${target}" data-cmd="outdent" aria-label="Decrease indent">⇤</button>
      <button type="button" class="tool" data-comm="comm-fmt" data-target="${target}" data-cmd="indent" aria-label="Increase indent">⇥</button>
    </div>
    <div class="notes-tools" role="toolbar" aria-label="Text color">
      <span class="hint" style="margin:0 2px 0 0">Color</span>
      ${COMM_COLORS.map(c => `<button type="button" class="tool swatch" data-comm="comm-fmt" data-target="${target}" data-cmd="foreColor" data-val="${c.hex}" style="background:${c.hex}" aria-label="${c.label} text" title="${c.label}"></button>`).join("")}
      <button type="button" class="tool swatch reset" data-comm="comm-fmt" data-target="${target}" data-cmd="foreColor" data-val="reset" aria-label="Default text color" title="Default">×</button>
    </div>`;
}
/* Remembers the cursor position in each composer so a toolbar click (which steals focus) still applies there. */
const commRanges = {};
document.addEventListener("selectionchange", () => {
  const sel = getSelection(); if (!sel.rangeCount) return;
  ["comm-body","comm-edit-body"].forEach(id => { const ed = document.getElementById(id); if (ed && ed.contains(sel.anchorNode)) commRanges[id] = sel.getRangeAt(0).cloneRange(); });
});
document.addEventListener("mousedown", e => { if (e.target.closest(".tool")) e.preventDefault(); });
/* Plain-text rendering of a sanitized message body, for length checks and the truncated board-preview snippet. */
function commPreview(html){
  const spaced = cleanHtml(html || "")
    .replace(/<(p|div|li|h1|h2|h3|h4|tr|blockquote|br)[^>]*>/gi, " $&")
    .replace(/<\/(p|div|li|h1|h2|h3|h4|tr|blockquote)>/gi, "$& ");
  const d = document.createElement("div"); d.innerHTML = spaced;
  return (d.textContent || "").replace(/\s+/g, " ").trim();
}

const Comm = {
  selected: 'serve', thread: '', mode: 'board', data: {}, meta: {}, drafts: {}, busy: false, edit: null, meId: '', identityKey: '', extraAdmins: {}, showHidden: false, composeOpen: false, generation: 0, moderationReady: true,
  // Verified in SharePoint: author-only reads, member Contribute, designated leader Design.
  privateReady: { serve: true, productions: true, security: true, ushers: true, worship: true },
  lists: {"serve": {"board": "2ec3f7fa-168b-43f4-a1a6-552242ddef87", "contacts": "e507b7a1-8fd9-4495-a75b-14f11fde39ef"}, "productions": {"board": "c2cf8dce-848e-44a5-82f9-22be525b3a27", "contacts": "3483d881-7779-4d60-9c0d-566f453b448f"}, "security": {"board": "e7acd4bb-1b3f-49fd-919f-6f8ee9e7d80d", "contacts": "878e6cc8-5db9-4875-bd56-32566efd84ea"}, "ushers": {"board": "695182cc-9ac4-4897-a4e7-1cab655b8e6b", "contacts": "e0601cbd-55d0-45f3-9e90-3a014155d2fd"}, "worship": {"board": "29e8d328-dadd-4eaf-a4ad-fc484d00adff", "contacts": "23b8b906-50a4-4382-829d-2dd0486ec01c"}},
  spaces() {
    const ministries = myTeams().map(t => t.key === 'ushers' ? {...t,name:'Ushers & Altar Workers'} : t).filter(t => (DEMO || S.teams[t.key]?.accessVerified) && S.teams[t.key]?.items && !S.teams[t.key]?.error && !S.teams[t.key]?.noAccess);
    // The RCI Serve board itself is limited to people on the RCI Leadership roster.
    if (!DEMO && !Rosters.data.leadership && !Rosters.errors.leadership) void Rosters.load('leadership');
    const serveAllowed = Rosters.isLeadershipMember();
    return ministries.length || isEditorOf(CHURCH_CAL)
      ? [...(serveAllowed ? [{ ...CHURCH_CAL, key: 'serve', name: 'RCI Serve', full: 'RCI Serve' }] : []), ...ministries] : [];
  },
  space(key = this.selected) { return this.spaces().find(t => t.key === key); },
  seenKey(key) { return `rci-serve:messages:${DEMO ? 'demo:' : ''}${S.me?.emails?.[0] || ''}:${key}`; },
  seen(key) { try { return new Set(JSON.parse(localStorage.getItem(this.seenKey(key)) || '[]')); } catch (_) { return new Set(); } },
  unread(key) { const seen = this.seen(key); return (this.data[key]?.board || []).filter(r => !this.metadata(r) && !this.hiddenRow(r,this.data[key],key) && !seen.has(r.id) && !this.own(r)).length; },
  totalUnread() { return this.spaces().reduce((n,t) => n + this.unread(t.key), 0); },
  markThread() {
    const d = this.data[this.selected]; if (!d || !this.thread) return;
    const seen = this.seen(this.selected);
    d.board.filter(r => r.id === this.thread || r.parent === this.thread).forEach(r => seen.add(r.id));
    this.saveSeen(this.selected, seen);
  },
  saveSeen(key, seen) { try { localStorage.setItem(this.seenKey(key), JSON.stringify([...seen].slice(-10000))); } catch (_) {} },
  /* Opening the board counts as reading the posts shown on it; no click into the post is needed.
     Posts first seen on this visit keep a "New" tag until the person leaves the board. Replies stay
     unread until the conversation is opened, since only the reply count shows on the board. */
  newOnVisit: new Set(),
  markShown(key, roots) {
    const seen = this.seen(key);
    if (!document.querySelector(`.comm-board[data-space="${CSS.escape(key)}"]`)) this.newOnVisit = new Set();
    let changed = false;
    roots.forEach(r => { if (!seen.has(r.id) && !this.own(r)) { seen.add(r.id); this.newOnVisit.add(r.id); changed = true; } });
    if (changed) this.saveSeen(key, seen);
    return seen;
  },
  normalize(r) { r = { ...r, fields: r.fields || {} }; return { id: String(r.id), parent: String(r.fields.ParentId || ''), title: r.fields.Title || '', text: r.fields.Body || '', name: r.createdBy?.user?.displayName || 'Team member', email: lc(r.createdBy?.user?.email), at: r.createdDateTime || '', authorId: r.createdBy?.user?.id || '', etag: r['@odata.etag'] || r.eTag || '' }; },
  async list(t, name) {
    this.siteIds ||= {};
    let siteId = API.sites?.[t.key]?.siteId || (t.key === 'serve' ? API.sites?.church?.siteId : '') || API.siteId || this.siteIds[t.sitePath];
    if (!siteId) { this.siteRequests ||= {}; this.siteRequests[t.sitePath] ||= API.g(`/sites/${CONFIG.siteHost}:${t.sitePath}?$select=id`).then(r => r.id).finally(() => { delete this.siteRequests[t.sitePath]; }); siteId = await this.siteRequests[t.sitePath]; this.siteIds[t.sitePath] = siteId; }
    const id = this.lists[t.key]?.[name === 'Team Board' ? 'board' : 'contacts'];
    if (!id) throw new Error('Message storage has not been configured.');
    return `/sites/${siteId}/lists/${id}/items`;
  },
  async load() {
    if (this.loadTask) return this.loadTask;
    this.loadTask = this.loadData().finally(() => { this.loadTask = null; });
    return this.loadTask;
  },
  async loadData() {
    const identityKey = S.me?.emails?.join('|') || '';
    if (this.identityKey !== identityKey) { this.meId = ''; this.data = {}; this.meta = {}; this.identityKey = identityKey; }
    const generation = this.generation;
    if (!this.meId && S.me?.id) this.meId = S.me.id;
    if (!DEMO && !this.meId) { try { this.meId = (await API.g('/me?$select=id')).id || ''; } catch (_) {} }

    const allowed = new Set(this.spaces().map(t => t.key));
    for (const k of Object.keys(this.data)) if (!allowed.has(k)) { delete this.data[k]; delete this.meta[k]; }
    await Promise.all(this.spaces().map(async t => {
      const d = { board: [], contacts: [], error: '', contactError: '' };
      if (DEMO) {
        this.data[t.key] ||= { ...d, board: [{ id: 'welcome', parent: '', title: `Welcome to ${t.name}`, text: 'Use this board to stay connected. Open a conversation to read and reply.', name: 'Team leader', email: 'leader@example.org', at: new Date().toISOString() }] };
        return;
      }
      await Promise.all([
        (async () => { try {
          const base = await this.list(t,'Team Board'); (this.meta[t.key] ||= {}).board = base;
          d.board = (await API.all(base + '?$expand=fields($select=Title,Body,ParentId)&$top=200')).map(r => this.normalize(r));
          await this.resolveAuthors(d.board);
        } catch (e) { d.error = e.status === 404 ? 'This board is awaiting setup. Please contact your leader.' : 'Unable to load this board. Refresh to try again.'; } })(),
        (async () => { if (!this.privateReady[t.key]) return; try {
          const base = await this.list(t,'Private Contacts'); (this.meta[t.key] ||= {}).contacts = base;
          d.contacts = (await API.all(base + '?$expand=fields($select=Title,Body,ParentId)&$top=200')).map(r => this.normalize(r));
          await this.resolveAuthors(d.contacts);
        } catch (_) { d.contactError = 'Unable to load private messages. Please try again.'; } })(),
        (async () => { try { const base = await this.list(t,'Team Board'); await this.loadAdminEmails(t,base.split('/')[2]); } catch (_) {} })()
      ]);
      if (generation === this.generation) this.data[t.key] = d;
    }));
    if (!this.space()) { this.selected = this.spaces()[0]?.key || ''; this.thread = ''; }
  },
  draftKey(kind, parent = '') { return `${this.selected}:${kind}:${parent}`; },
  form(kind, parent = '') {
    const draft = this.drafts[this.draftKey(kind, parent)] || {};
    const contact = kind === 'contact';
    return `<form class="panel pad comm-form" data-kind="${kind}" data-parent="${esc(parent)}">
      <h2>${contact ? 'Contact your leaders privately' : parent ? 'Write a reply' : 'Make a post'}</h2>
      ${!parent ? `<div class="field"><label for="comm-title">${contact ? 'Subject' : 'Title'}</label><input id="comm-title" name="title" maxlength="120" value="${esc(draft.title || '')}" required></div>` : ''}
      <div class="field"><label for="comm-body">${parent ? 'Reply' : 'Message'}</label>${commToolbar('comm-body')}<div class="panel notes notes-editor" id="comm-body" contenteditable="true" role="textbox" aria-multiline="true" aria-label="${parent ? 'Reply' : 'Message'}">${cleanHtml(draft.body || '')}</div></div>
      ${contact ? '<p class="hint">Your message is private to you and this space’s leaders. Authorized site administrators retain access.</p>' : '<p class="hint">Visible to members of this board.</p>'}
      <button class="btn primary" type="submit" ${this.busy ? 'disabled' : ''}>${this.busy ? 'Sending…' : contact ? 'Send private message' : parent ? 'Post reply' : 'Post message'}</button>
    </form>`;
  },
  async resolveAuthors(rows) {
    const authors = [...new Set(rows.filter(r => !r.email && r.authorId).map(r => r.authorId))];
    const emails = new Map(await Promise.all(authors.map(async id => { try { const u = await API.g(`/users/${encodeURIComponent(id)}?$select=mail,userPrincipalName`); return [id,lc(u.mail || u.userPrincipalName)]; } catch (_) { return [id,'']; } })));
    rows.forEach(r => { if (!r.email) r.email = emails.get(r.authorId) || ''; });
  },
  own(r) { return !!S.me && (!!this.meId && r.authorId === this.meId || !!r.email && S.me.emails.map(lc).includes(lc(r.email))); },
  adminEmails(t = this.space()) { return [...(t?.editors || []), ...(this.extraAdmins[t?.key] || [])].map(lc); },
  isAdmin() { return DEMO || !!S.me?.emails.some(e => this.adminEmails().includes(lc(e))); },
  async loadAdminEmails(t, siteId) {
    if (t.key === 'serve') return;
    this.extraAdmins[t.key] = [];
    try {
      const drives = await API.g(`/sites/${siteId}/drives?$select=id,name,webUrl`);
      const drive = drives.value?.find(x => /\/SiteAssets$/i.test(x.webUrl || '') || /site assets/i.test(x.name || ''));
      if (!drive) return;
      const list = await API.g(`/drives/${drive.id}/root:/app/admins.json:/content`);
      this.extraAdmins[t.key] = Array.isArray(list) ? list.map(lc) : [];
    } catch (_) { /* Built-in administrators remain available if the optional roster cannot be read. */ }
  },
  metadata(r) { return /^(pin:|hide:|leadership:|hub-admin:|hub-profile:|hub-card:|hub-access:)/.test(r.parent); },
  board(d = this.data[this.selected]) { return (d?.board || []).filter(r => !this.metadata(r)); },
  hidden(id, d = this.data[this.selected], key = this.selected) {
    const admins = this.adminEmails(this.space(key));
    const events = (d?.board || []).filter(r => r.parent === 'hide:' + id && (admins.includes(lc(r.email)) || DEMO && this.own(r)) && ['hide','show'].includes(r.text));
    events.sort((a,b) => b.at.localeCompare(a.at) || String(b.id).localeCompare(String(a.id),undefined,{numeric:true}));
    return events[0]?.text === 'hide';
  },
  hiddenRow(r,d,key = this.selected) { return this.hidden(r.id,d,key) || !!r.parent && this.hidden(r.parent,d,key); },
  visibleBoard(d = this.data[this.selected]) { return this.board(d).filter(r => this.isAdmin() && this.showHidden || !this.hiddenRow(r,d)); },
  canChange(r,kind) { return this.own(r) || kind === 'board' && this.moderationReady && this.isAdmin(); },
  roots(d = this.data[this.selected]) {
    const rows = this.visibleBoard(d), roots = rows.filter(r => !r.parent), ids = new Set(roots.map(r => r.id));
    for (const r of rows) if (r.parent && !ids.has(r.parent)) { ids.add(r.parent); roots.push({id:r.parent,parent:'',title:'Deleted conversation',text:'The original message was deleted.',name:'',email:'',at:r.at,deleted:true}); }
    return roots;
  },
  pinned(id, d = this.data[this.selected]) {
    if (!this.board(d).some(r => r.id === id && !r.parent) || this.hidden(id,d)) return false;
    const admins = this.adminEmails();
    const events = (d?.board || []).filter(r => r.parent === 'pin:' + id && (admins.includes(lc(r.email)) || DEMO && this.own(r)) && ['pin','unpin'].includes(r.text));
    events.sort((a,b) => b.at.localeCompare(a.at) || String(b.id).localeCompare(String(a.id),undefined,{numeric:true}));
    return events[0]?.text === 'pin';
  },
  notesTargets() {
    if (!Rosters.isLeadershipMember()) return [];
    return TEAMS.filter(t => ['worship','productions'].includes(t.key) &&
      (DEMO || S.teams[t.key]?.accessVerified && !S.teams[t.key].error && !S.teams[t.key].noAccess));
  },
  canForward(r, kind) {
    return !S.imp && this.selected === 'serve' && this.mode === 'board' && kind === 'board' &&
      !r.parent && !r.deleted && !this.hiddenRow(r) && this.notesTargets().length > 0;
  },
  forwardForm() {
    const draft = this.forward;
    const r = this.data.serve?.board.find(row => row.id === draft?.id);
    if (!r || !this.canForward(r, 'board')) return '';
    return `<form class="panel pad comm-forward-form"><h2>Forward to notes</h2><p>${esc(r.title)}</p>
      <div class="field"><label for="comm-notes-date">Service date</label><input id="comm-notes-date" name="date" type="date" value="${esc(draft.date)}" required></div>
      <fieldset><legend>Send to</legend>${this.notesTargets().map(t => `<label><input type="checkbox" name="target" value="${t.key}" ${draft.targets.includes(t.key) ? 'checked' : ''}> ${this.notesLabel(t.key)}</label><br>`).join('')}</fieldset>
      <p class="hint">Forwarding again to the same date updates the existing copy.</p>
      <div class="comm-edit-buttons"><button class="btn primary" type="submit" ${this.busy ? 'disabled' : ''}>Forward to notes</button><button class="btn" type="button" data-comm="cancel-forward">Cancel</button></div></form>`;
  },
  notesLabel(key) { return key === 'worship' ? 'Worship Notes' : 'Production Notes'; },
  notesLink(id) { return 'https://astmacary.github.io/?boardPost=' + encodeURIComponent(id); },
  notesBlock(r) {
    // Keep the ministry's SET LIST delimiter out of the forwarded text. The link is kept but hidden —
    // it's only there so a re-forward or a later edit can find and update this same block in place;
    // the note itself shows just the message's own text, with no title and no "RCI Serve Team Board" label.
    const body = cleanHtml(r.text || '').replace(/SET LIST(?=[^<>]*(?:<|$))/g, 'SET<span> </span>LIST');
    return `<blockquote><p style="display:none"><a href="${this.notesLink(r.id)}"></a></p>${body}</blockquote>`;
  },
  hasNotesLink(notes, id) {
    const link = this.notesLink(id);
    return notes.includes(link + '"') || notes.includes(link + "'");
  },
  replaceNotesBlock(notes, r, append = false) {
    const blocks = [];
    const opens = [];
    for (const tag of notes.matchAll(/<\/?blockquote\b[^>]*>/gi)) {
      if (tag[0][1] !== '/') {
        const header = notes.slice(tag.index + tag[0].length).match(/^\s*<p\b[^>]*>\s*<a\b[^>]*>/i);
        opens.push({ start:tag.index, linked:!!header && this.hasNotesLink(header[0], r.id) });
      } else {
        const block = opens.pop();
        if (block?.linked) blocks.push({ start:block.start, end:tag.index + tag[0].length });
      }
    }
    if (!blocks.length) return append ? notes + this.notesBlock(r) : notes;
    let updated = notes;
    for (const block of blocks.reverse()) {
      updated = updated.slice(0, block.start) + this.notesBlock(r) + updated.slice(block.end);
    }
    return updated;
  },
  async writeNotes(key, r, date = '') {
    if (DEMO) {
      this.demoNotes ||= {};
      if (date) {
        const slot = key + ':' + date;
        this.demoNotes[slot] = this.replaceNotesBlock(this.demoNotes[slot] || '', r, true);
      } else {
        for (const slot of Object.keys(this.demoNotes).filter(slot => slot.startsWith(key + ':'))) {
          this.demoNotes[slot] = this.replaceNotesBlock(this.demoNotes[slot], r);
        }
      }
      return;
    }
    const { base, rows } = await API.boardNotesRows(key);
    if (!date) {
      const linked = rows.filter(row => this.hasNotesLink(row.fields?.Notes || '', r.id));
      let failed = false;
      for (const row of linked) {
        const notes = this.replaceNotesBlock(row.fields.Notes, r);
        if (notes === row.fields.Notes) continue;
        try { await API.saveBoardNotes(base, row, notes); } catch (_) { failed = true; }
      }
      if (failed) throw new Error('Some dated copies could not be updated.');
      return;
    }
    const production = key === 'productions';
    const dated = rows.filter(({fields: f}) => spDate(production ? f?.ServiceDates : f?.ServiceDate) === date);
    const row = dated.find(({fields: f}) => production ? f.Title === 'Production Director'
      : lc(f.Title) === 'coordinator' && !f.FullName &&
        ![f.LeaderPositions, f.SingingPositions, f.MusicianPositions].some(value => value?.length));
    if (row) {
      const before = row.fields.Notes || '';
      const notes = this.replaceNotesBlock(before, r, true);
      if (notes !== before) await API.saveBoardNotes(base, row, notes);
      return;
    }
    if (production) throw new Error('Schedule a Production Director for this date first.');
    const legacy = dated.find(({fields: f}) => f.Notes && [f.Title, ...[f.LeaderPositions, f.SingingPositions, f.MusicianPositions].flat()].includes('Worship & Arts Director'));
    const previous = legacy?.fields.Notes || '';
    const marker = previous.indexOf('SET LIST');
    const notes = this.replaceNotesBlock(marker < 0 ? previous : previous.slice(marker), r, true);
    await API.g(base, { method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fields: { Title: 'Coordinator', ServiceDate: date + 'T12:00:00Z', ActiveStatus: 'Non Active', Published: '--', Notes: notes } }) });
  },
  async sendToNotes(form) {
    const draft = this.forward;
    const r = this.data.serve?.board.find(row => row.id === draft?.id);
    if (this.busy || !r || !this.canForward(r, 'board')) return;
    const data = new FormData(form);
    draft.date = String(data.get('date') || '');
    draft.targets = this.notesTargets().map(t => t.key).filter(key => data.getAll('target').includes(key));
    if (!/^\d{4}-\d{2}-\d{2}$/.test(draft.date) || !draft.targets.length) {
      toast('Choose a service date and at least one notes destination.', true);
      return;
    }
    const selectedCount = draft.targets.length;
    this.busy = true; render();
    try {
      const source = DEMO ? r : this.normalize(await API.g(`${this.meta.serve.board}/${encodeURIComponent(r.id)}?$expand=fields($select=Title,Body,ParentId)`));
      if (!this.canForward(source, 'board')) throw new Error('This post is unavailable.');
      const failed = [];
      for (const key of draft.targets) {
        try { await this.writeNotes(key, source, draft.date); }
        catch (e) {
          const reason = e.status === 403 ? 'SharePoint write access is required.'
            : e.status === 412 ? 'The notes changed while saving. Try again.'
            : e.status ? 'Could not save. Try again.' : e.message;
          failed.push({ key, reason });
        }
      }
      if (failed.length) {
        draft.targets = failed.map(result => result.key);
        const saved = selectedCount > failed.length ? 'Other selected notes were saved. ' : '';
        toast(saved + failed.map(({key, reason}) => this.notesLabel(key) + ' — ' + reason).join(' '), true);
      } else {
        this.forward = null;
        toast(DEMO ? 'Forwarded in this demo only.' : 'Post forwarded to notes.');
      }
    } catch (_) { toast('Could not read the Team Board post. Refresh and try again.', true); }
    finally { this.busy = false; render(); }
  },
  icon(name) { const paths = {forward:'M3 20V4h12M7 8h5m-5 4h3m4-1 6 4-6 4m-3-4h9',edit:'M12 20h9M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z',delete:'M3 6h18M9 6V4h6v2M5 6l1 14h12l1-14M10 10v6M14 10v6',pin:'m16 3 5 5-4 1-3 5v3l-3-3-7 7 7-7-4-3h3l5-3Z',hide:'M2 12s4-7 10-7 10 7 10 7-4 7-10 7S2 12 2 12ZM9 12a3 3 0 1 0 6 0 3 3 0 0 0-6 0M3 3l18 18'}; return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="${paths[name]}"/></svg>`; },
  actions(r, kind) {
    if (r.deleted) return '';
    const button = (action,label,pressed) => `<button type="button" class="iconbtn" data-comm="${action}" data-kind="${kind}" data-id="${esc(r.id)}" aria-label="${label}" title="${label}"${pressed == null ? '' : ` aria-pressed="${pressed}"`}${this.busy ? ' disabled' : ''}>${this.icon(action)}</button>`;
    return `<div class="comm-actions">${this.canForward(r,kind) ? button('forward','Forward to notes') : ''}${this.canChange(r,kind) ? button('edit','Edit message') + button('delete','Delete message') : ''}${kind === 'board' && !r.parent && this.isAdmin() ? button('pin',this.pinned(r.id) ? 'Unpin message' : 'Pin message',this.pinned(r.id)) : ''}${kind === 'board' && this.isAdmin() ? button('hide',this.hidden(r.id) ? 'Show message' : 'Hide message',this.hidden(r.id)) : ''}</div>`;
  },
  editForm(r, kind) {
    const edit = this.edit;
    if (!edit || edit.id !== r.id || edit.kind !== kind) return '';
    return `<form class="comm-edit-form"><h3>Edit message</h3>${!r.parent ? `<div class="field"><label for="comm-edit-title">${kind === 'contacts' ? 'Subject' : 'Title'}</label><input id="comm-edit-title" name="title" maxlength="120" value="${esc(edit.title)}" required></div>` : ''}<div class="field"><label for="comm-edit-body">Message</label>${commToolbar('comm-edit-body')}<div class="panel notes notes-editor" id="comm-edit-body" contenteditable="true" role="textbox" aria-multiline="true" aria-label="Message">${cleanHtml(edit.text)}</div></div><div class="comm-edit-buttons"><button class="btn primary" type="submit" ${this.busy ? 'disabled' : ''}>Save changes</button><button class="btn" type="button" data-comm="cancel-edit">Cancel</button></div></form>`;
  },
  async fresh(kind, id) {
    const base = this.meta[this.selected]?.[kind];
    if (!base) throw new Error('Refresh messages and try again.');
    const raw = await API.g(`${base}/${encodeURIComponent(id)}?$expand=fields($select=Title,Body,ParentId)`);
    const row = this.normalize(raw);
    if (!this.canChange(row,kind)) throw new Error('You do not have permission to change this message.');
    return {base,row};
  },
  async act(action, kind, id) {
    if (this.busy || !this.space() || !['board','contacts'].includes(kind)) return;
    const d = this.data[this.selected], r = d?.[kind]?.find(x => x.id === id);
    if (!r || this.metadata(r)) return;
    if (action === 'pin' || action === 'hide') {
      if (kind !== 'board' || action === 'pin' && r.parent || !this.isAdmin()) return;
      this.generation++; this.busy = true; render();
      try {
        const text = action === 'pin' ? this.pinned(id,d) ? 'unpin' : 'pin' : this.hidden(id,d) ? 'show' : 'hide';
        const fields = {Title:action === 'pin' ? 'Board pin' : 'Message visibility',Body:text,ParentId:action+':'+id};
        let event;
        if (DEMO) event = {id:'demo-pin-'+String(this.pinSerial = (this.pinSerial || 0) + 1).padStart(8,'0'),title:fields.Title,text,parent:fields.ParentId,at:new Date().toISOString(),name:S.me.name,email:S.me.emails[0],authorId:this.meId};
        else { const raw = await API.g(this.meta[this.selected].board,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({fields})}); event = {...this.normalize(raw),title:fields.Title,text,parent:fields.ParentId,email:S.me.emails[0],authorId:this.meId}; }
        d.board.push(event); if (text === 'hide' && this.thread === id) this.thread = ''; toast(({pin:'Message pinned.',unpin:'Message unpinned.',hide:'Message hidden from the board.',show:'Message restored.'})[text]);
      } catch (_) { toast('Could not update the message. Please try again.',true); }
      finally { this.busy = false; render(); }
      return;
    }
    if (!this.canChange(r,kind)) return;
    if (action === 'edit') { this.edit = {kind,id,title:r.title,text:r.text,etag:r.etag,syncNotes:this.canForward(r, kind)}; render(); return; }
    if (action !== 'delete' || !confirm('Delete this message? Other people’s replies will remain.')) return;
    this.generation++; this.busy = true; render();
    try {
      if (!DEMO) { const {base,row} = await this.fresh(kind,id); await API.g(`${base}/${encodeURIComponent(id)}`,{method:'DELETE',headers:row.etag ? {'If-Match':row.etag} : {}}); }
      d[kind] = d[kind].filter(x => x.id !== id); this.edit = null;
      if (this.thread === id && !d.board.some(x => x.parent === id)) this.thread = '';
      toast('Message deleted.');
    } catch (_) { toast('Could not delete the message. Refresh and try again.',true); }
    finally { this.busy = false; render(); }
  },
  async saveEdit() {
    const edit = this.edit;
    if (this.busy || !edit || !this.space()) return;
    const d = this.data[this.selected], r = d?.[edit.kind]?.find(x => x.id === edit.id);
    if (!r || !this.canChange(r,edit.kind)) return;
    const text = cleanHtml(edit.text || '').trim(), title = edit.title.trim(), plain = commPreview(text);
    if (!r.parent && (!title || title.length > 120)) { toast(!title ? 'Enter a title.' : 'Title is too long.', true); return; }
    if (!plain || plain.length > 4000 || text.length > 20000) { toast(!plain ? 'Write a message first.' : 'Message is too long.', true); return; }
    this.generation++; this.busy = true; render();
    try {
      if (!DEMO) { const {base,row} = await this.fresh(edit.kind,edit.id); if (edit.etag && row.etag && edit.etag !== row.etag) throw new Error('Message changed elsewhere.'); await API.g(`${base}/${encodeURIComponent(edit.id)}/fields`,{method:'PATCH',headers:{'Content-Type':'application/json',...(row.etag ? {'If-Match':row.etag} : {})},body:JSON.stringify({Title:r.parent ? r.title : title,Body:text})}); }
      r.title = r.parent ? r.title : title; r.text = text; this.edit = null;
      const failures = [];
      if (edit.syncNotes && this.canForward(r,edit.kind)) {
        for (const target of this.notesTargets()) {
          try { await this.writeNotes(target.key, r); }
          catch (_) { failures.push(this.notesLabel(target.key)); }
        }
      }
      if (failures.length) toast(`Message saved. Couldn't sync ${failures.join(' or ')}. Refresh messages, then edit and save again to retry.`, true);
      else toast('Message updated.');
    } catch (_) { toast('Could not save your changes. Your edit is still here. Please try again.',true); }
    finally { this.busy = false; render(); }
  },
  record(r, kind = 'board') { return `<div class="comm-record"><div class="comm-record-head"><div><b>${esc(r.name)}</b><time class="hint">${esc(r.at ? new Date(r.at).toLocaleString() : '')}</time></div>${this.actions(r,kind)}</div>${kind === 'board' && this.hiddenRow(r) ? '<span class="hint comm-hidden-label">Hidden from the board</span>' : ''}${this.edit?.id === r.id && this.edit?.kind === kind ? this.editForm(r,kind) : `<div class="comm-text notes">${cleanHtml(r.text)}</div>`}</div>`; },
  view() {
    const spaces = this.spaces(), t = this.space();
    if (!t) return '<h1 class="page">Messages</h1><div class="empty">Your boards will appear after your ministry access is confirmed.</div>';
    const d = this.data[t.key], privateOn = DEMO || this.privateReady[t.key];
    let content = '';
    if (!d) content = '<div class="empty">Loading messages…</div>';
    else if (this.mode === 'contact') {
      content = !privateOn ? '<div class="panel pad">Private messaging is awaiting secure setup. Please contact your ministry leader directly for now.</div>' : (d.contactError ? `<div class="err">${esc(d.contactError)}</div>` : this.form('contact') + `<h2>Private messages</h2><p class="hint">Your submissions appear here. Leaders see the submissions they are authorized to read.</p>` + (d.contacts.length ? d.contacts.slice().sort((a,b) => b.at.localeCompare(a.at)).map(r => `<article class="panel pad"><h3>${esc(r.title)}</h3>${this.record(r,'contacts')}</article>`).join('') : '<div class="empty">No private messages yet.</div>'));
    } else if (d.error) content = `<div class="err">${esc(d.error)}</div>`;
    else if (this.thread) {
      const root = this.roots(d).find(r => r.id === this.thread);
      content = root ? `<button class="btn" data-comm="back">Back to conversations</button><article class="panel pad"><h2>${this.pinned(root.id,d) ? '📌 ' : ''}${esc(root.title)}</h2>${this.record(root)}${this.visibleBoard(d).filter(r => r.parent === root.id).sort((a,b) => a.at.localeCompare(b.at)).map(r => this.record(r)).join('')}</article>${this.form('reply',root.id)}` : '<div class="empty">This conversation is no longer available.</div>';
    } else {
      const roots = this.roots(d), seen = this.markShown(t.key, roots);
      const activity = r => this.board(d).filter(x => x.id === r.id || x.parent === r.id).reduce((at,x) => x.at > at ? x.at : at,r.at);
      const composer = `<button type="button" class="btn comm-compose-toggle" data-comm="toggle-compose" aria-expanded="${this.composeOpen}">${this.composeOpen ? '− Hide' : '+ Make a post'}</button>` + (this.composeOpen ? this.form('post') : '');
      content = `<div class="comm-board" data-space="${esc(t.key)}" hidden></div>` + composer + (roots.length ? roots.sort((a,b) => Number(this.pinned(b.id,d)) - Number(this.pinned(a.id,d)) || activity(b).localeCompare(activity(a))).map(r => {
        const replies = this.visibleBoard(d).filter(x => x.parent === r.id), fresh = [r,...replies].some(x => (!seen.has(x.id) || this.newOnVisit.has(x.id)) && !this.own(x));
        return `<article class="panel pad"><div class="comm-topic-line"><button class="btn comm-topic" data-comm="thread" data-id="${esc(r.id)}">${this.pinned(r.id,d) ? '📌 ' : ''}${esc(r.title)}${fresh ? ' • New' : ''}</button>${this.actions(r,'board')}</div>${this.hiddenRow(r,d) ? '<span class="hint comm-hidden-label">Hidden from the board</span>' : ''}${this.edit?.id === r.id && this.edit?.kind === 'board' ? this.editForm(r,'board') : ''}<p class="hint">${esc(r.name)} · ${plural(replies.length,'reply','replies')}</p><div class="comm-text notes">${cleanHtml(r.text)}</div></article>`;
      }).join('') : '<div class="empty">No conversations yet. Start the first one.</div>');
    }
    const chips = spaces.map(s => `<button class="chipcheck team team-pill" style="${typeof tvars === 'function' ? tvars(s) : ''}" data-comm="space" data-key="${s.key}" aria-pressed="${s.key === t.key}"><span class="dot" aria-hidden="true"></span>${esc(s.name)}${this.unread(s.key) ? ` <span aria-label="${this.unread(s.key)} unread messages">(${this.unread(s.key)})</span>` : ''}</button>`).join('');
    return `<div class="comm"><div class="comm-head"><h1 class="page">Messages</h1><button class="iconbtn" data-comm="refresh" aria-label="Refresh messages" title="Refresh messages"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M23 4v6h-6M1 20v-6h6M3.5 9a9 9 0 0 1 14.9-5L23 10M1 14l4.6 6A9 9 0 0 0 20.5 15"/></svg></button></div><div class="chiprow" aria-label="Choose a board">${chips}</div><h2>${esc(t.name)}</h2><div class="chiprow"><button class="chipcheck" data-comm="mode" data-mode="board" aria-pressed="${this.mode === 'board'}">Team board</button><button class="chipcheck" data-comm="mode" data-mode="contact" aria-pressed="${this.mode === 'contact'}">Contact leaders privately</button>${this.mode === 'board' && this.isAdmin() ? `<button class="chipcheck" data-comm="toggle-hidden" aria-pressed="${this.showHidden}">${this.showHidden ? 'Hide hidden messages' : 'Show hidden messages'}</button>` : ''}</div>${this.forwardForm()}${content}</div>`;
  },
  async send(form) {
    if (this.busy) return;
    const t = this.space(), kind = form.dataset.kind, parent = form.dataset.parent || '';
    if (!t || !['contact','post','reply'].includes(kind)) return;
    const data = new FormData(form), title = String(data.get('title') || '').trim();
    const edInput = form.querySelector('#comm-body'), body = cleanHtml(edInput ? edInput.innerHTML : '').trim(), plain = commPreview(body);
    if (!parent && (!title || title.length > 120)) { toast(!title ? 'Enter a title.' : 'Title is too long.', true); return; }
    if (!plain || plain.length > 4000 || body.length > 20000) { toast(!plain ? 'Write a message first.' : 'Message is too long.', true); return; }
    const d = this.data[t.key], contact = kind === 'contact';
    if (!d || (contact ? !(DEMO || this.privateReady[t.key]) || d.contactError : d.error)) return;
    if (kind === 'reply' && !this.roots(d).some(r => r.id === parent)) return;
    this.generation++; this.busy = true; const key = this.draftKey(kind,parent); render();
    try {
      const fields = { Title: title || 'Reply', Body: body, ParentId: parent };
      let r;
      if (DEMO) r = { id: 'demo-' + crypto.randomUUID(), parent, title: fields.Title, text: body, name: S.me.name, email: S.me.emails[0], authorId: this.meId, at: new Date().toISOString() };
      else {
        const base = this.meta[t.key]?.[contact ? 'contacts' : 'board']; if (!base) throw new Error('This space is not ready. Refresh and try again.');
        r = this.normalize(await API.g(base, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) }));
        // Some Graph responses omit expanded fields; use the submitted plain text.
        Object.assign(r, { parent, title: fields.Title, text: body, name: S.me.name, email: S.me.emails[0], authorId: this.meId });
      }
      d[contact ? 'contacts' : 'board'].push(r); delete this.drafts[key];
      if (!contact) { this.thread = parent || r.id; this.markThread(); }
      toast(DEMO ? 'Saved in this demo only.' : contact ? 'Private message sent to your leaders.' : 'Message posted.');
    } catch (_) { toast('Your message could not be sent. Your draft is still here; please try again.',true); }
    finally { this.busy = false; render(); }
  }
};
document.addEventListener('input', e => {
  if (e.target.closest('.comm-edit-form') && Comm.edit) { const form = e.target.closest('.comm-edit-form'), title = form.querySelector('#comm-edit-title'), ed = form.querySelector('#comm-edit-body'); Comm.edit.title = String(title ? title.value : ''); Comm.edit.text = ed ? ed.innerHTML : ''; Comm.edit.syncNotes = !!form.elements['sync-notes']?.checked; return; }
  const form = e.target.closest('.comm-form'); if (!form) return;
  const title = form.querySelector('#comm-title'), ed = form.querySelector('#comm-body');
  Comm.drafts[Comm.draftKey(form.dataset.kind,form.dataset.parent)] = { title: String(title ? title.value : ''), body: ed ? ed.innerHTML : '' };
});
document.addEventListener('submit', e => { if (e.target.matches('.comm-forward-form')) { e.preventDefault(); Comm.sendToNotes(e.target); return; } if (e.target.matches('.comm-edit-form')) { e.preventDefault(); Comm.saveEdit(); return; } if (!e.target.matches('.comm-form')) return; e.preventDefault(); Comm.send(e.target); });
document.addEventListener('click', async e => {
  const b = e.target.closest('[data-comm]'); if (!b) return;
  const action = b.dataset.comm;
  if (action === 'comm-fmt') {
    const ed = document.getElementById(b.dataset.target); if (!ed) return;
    ed.focus();
    const range = commRanges[b.dataset.target];
    if (range) { const sel = getSelection(); sel.removeAllRanges(); sel.addRange(range); }
    const cmd = b.dataset.cmd, val = cmd === 'foreColor' ? (b.dataset.val === 'reset' ? getComputedStyle(ed).color : b.dataset.val) : null;
    try { document.execCommand('styleWithCSS', false, cmd === 'foreColor'); document.execCommand(cmd, false, val); } catch (_) {}
    return;
  }
  if (Comm.busy) return;
  if (['edit','delete','pin','hide'].includes(action)) { await Comm.act(action,b.dataset.kind,b.dataset.id); return; }
  if (action === 'forward') {
    const r = Comm.data.serve?.board.find(row => row.id === b.dataset.id);
    if (!r || !Comm.canForward(r, b.dataset.kind)) return;
    Comm.edit = null;
    Comm.forward = { id: r.id, date: todayKey(), targets: [] };
    render();
    document.getElementById('comm-notes-date')?.focus();
    return;
  }
  if (['cancel-forward','space','mode','back','thread','refresh','toggle-hidden'].includes(action)) Comm.forward = null;
  if (action === 'toggle-hidden' && Comm.isAdmin()) { Comm.showHidden = !Comm.showHidden; Comm.thread = ''; Comm.edit = null; }
  if (action === 'cancel-edit') Comm.edit = null;
  if (action === 'space') { if (!Comm.space(b.dataset.key)) return; Comm.selected = b.dataset.key; Comm.thread = ''; Comm.edit = null; Comm.showHidden = false; Comm.composeOpen = false; }
  if (action === 'mode') { Comm.mode = b.dataset.mode; Comm.thread = ''; Comm.edit = null; Comm.showHidden = false; Comm.composeOpen = false; }
  if (action === 'back') { Comm.thread = ''; Comm.composeOpen = false; }
  if (action === 'toggle-compose') Comm.composeOpen = !Comm.composeOpen;
  if (action === 'thread') { Comm.thread = b.dataset.id; Comm.markThread(); }
  if (action === 'refresh') { Comm.edit = null; b.disabled = true; await Comm.load(); }
  render();
});

// Keep badges current while the app is open; do not interrupt an active composer.
let communicationRefreshBusy = false;
async function refreshCommunications(){
  if (DEMO || S.phase !== 'ready' || S.loading || Comm.busy || Comm.forward || communicationRefreshBusy || document.hidden || document.activeElement?.closest('.comm-form,.comm-edit-form')) return;
  communicationRefreshBusy = true;
  try {
    await Comm.load();
    if (S.view === 'messages' && Comm.mode === 'board' && Comm.thread) Comm.markThread();
    if (S.phase === 'ready' && !Comm.busy && !Comm.forward && !document.activeElement?.closest('.comm-form,.comm-edit-form')) render();
  } catch (_) { /* Retry on the next interval; list-specific failures are shown in Messages. */ }
  finally { communicationRefreshBusy = false; }
}
setInterval(refreshCommunications, 60000);
document.addEventListener('visibilitychange', () => { if (!document.hidden) refreshCommunications(); });
