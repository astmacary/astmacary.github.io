/* Communication records stay in their own ministry's SharePoint site.
   SharePoint permissions, not these UI checks, enforce access. See COMMUNICATIONS.md. */
const Comm = {
  selected: 'serve', thread: '', mode: 'board', data: {}, meta: {}, drafts: {}, busy: false, edit: null, meId: '', identityKey: '', extraAdmins: {},
  // Verified in SharePoint: author-only reads, member Contribute, designated leader Design.
  privateReady: { serve: true, productions: true, security: true, ushers: true, worship: true },
  lists: {"serve": {"board": "2ec3f7fa-168b-43f4-a1a6-552242ddef87", "contacts": "e507b7a1-8fd9-4495-a75b-14f11fde39ef"}, "productions": {"board": "c2cf8dce-848e-44a5-82f9-22be525b3a27", "contacts": "3483d881-7779-4d60-9c0d-566f453b448f"}, "security": {"board": "e7acd4bb-1b3f-49fd-919f-6f8ee9e7d80d", "contacts": "878e6cc8-5db9-4875-bd56-32566efd84ea"}, "ushers": {"board": "695182cc-9ac4-4897-a4e7-1cab655b8e6b", "contacts": "e0601cbd-55d0-45f3-9e90-3a014155d2fd"}, "worship": {"board": "29e8d328-dadd-4eaf-a4ad-fc484d00adff", "contacts": "23b8b906-50a4-4382-829d-2dd0486ec01c"}},
  spaces() {
    const ministries = myTeams().filter(t => (DEMO || S.teams[t.key]?.accessVerified) && S.teams[t.key]?.items && !S.teams[t.key]?.error && !S.teams[t.key]?.noAccess);
    return ministries.length || isEditorOf(CHURCH_CAL)
      ? [{ ...CHURCH_CAL, key: 'serve', name: 'RCI Serve', full: 'RCI Serve' }, ...ministries] : [];
  },
  space(key = this.selected) { return this.spaces().find(t => t.key === key); },
  seenKey(key) { return `rci-serve:messages:${DEMO ? 'demo:' : ''}${S.me?.emails?.[0] || ''}:${key}`; },
  seen(key) { try { return new Set(JSON.parse(localStorage.getItem(this.seenKey(key)) || '[]')); } catch (_) { return new Set(); } },
  unread(key) { const seen = this.seen(key); return (this.data[key]?.board || []).filter(r => !r.parent.startsWith('pin:') && !seen.has(r.id) && !this.own(r)).length; },
  totalUnread() { return this.spaces().reduce((n,t) => n + this.unread(t.key), 0); },
  markThread() {
    const d = this.data[this.selected]; if (!d || !this.thread) return;
    const seen = this.seen(this.selected);
    d.board.filter(r => r.id === this.thread || r.parent === this.thread).forEach(r => seen.add(r.id));
    try { localStorage.setItem(this.seenKey(this.selected), JSON.stringify([...seen].slice(-10000))); } catch (_) {}
  },
  normalize(r) { r = { ...r, fields: r.fields || {} }; return { id: String(r.id), parent: String(r.fields.ParentId || ''), title: r.fields.Title || '', text: r.fields.Body || '', name: r.createdBy?.user?.displayName || 'Team member', email: lc(r.createdBy?.user?.email), at: r.createdDateTime || '', authorId: r.createdBy?.user?.id || '', etag: r['@odata.etag'] || r.eTag || '' }; },
  async list(t, name) {
    const site = await API.g(`/sites/${CONFIG.siteHost}:${t.sitePath}?$select=id`);
    const id = this.lists[t.key]?.[name === 'Team Board' ? 'board' : 'contacts'];
    if (!id) throw new Error('Message storage has not been configured.');
    const list = await API.g(`/sites/${site.id}/lists/${id}?$select=id`);
    return `/sites/${site.id}/lists/${list.id}/items`;
  },
  async load() {
    const identityKey = S.me?.emails?.join('|') || '';
    if (this.identityKey !== identityKey) { this.meId = ''; this.data = {}; this.meta = {}; this.identityKey = identityKey; }
    if (!DEMO && !this.meId) { try { this.meId = (await API.g('/me?$select=id')).id || ''; } catch (_) {} }

    const allowed = new Set(this.spaces().map(t => t.key));
    for (const k of Object.keys(this.data)) if (!allowed.has(k)) { delete this.data[k]; delete this.meta[k]; }
    await Promise.all(this.spaces().map(async t => {
      const d = { board: [], contacts: [], error: '', contactError: '' };
      if (DEMO) {
        this.data[t.key] ||= { ...d, board: [{ id: 'welcome', parent: '', title: `Welcome to ${t.name}`, text: 'Use this board to stay connected. Open a conversation to read and reply.', name: 'Team leader', email: 'leader@example.org', at: new Date().toISOString() }] };
        return;
      }
      try {
        const base = await this.list(t, 'Team Board'); this.meta[t.key] = { board: base };
        await this.loadAdminEmails(t, base.split('/')[2]);
        d.board = (await API.all(base + '?$expand=fields($select=Title,Body,ParentId)&$top=200')).map(r => this.normalize(r));
        await this.resolveAuthors(d.board);
      } catch (e) { d.error = e.status === 404 ? 'This board is awaiting setup. Please contact your leader.' : 'Unable to load this board. Refresh to try again.'; }
      if (this.privateReady[t.key]) {
        try {
          const base = await this.list(t, 'Private Contacts'); (this.meta[t.key] ||= {}).contacts = base;
          // Server returns only records the signed-in person is allowed to read.
          d.contacts = (await API.all(base + '?$expand=fields($select=Title,Body,ParentId)&$top=200')).map(r => this.normalize(r));
          await this.resolveAuthors(d.contacts);
        } catch (_) { d.contactError = 'Unable to load private messages. Please try again.'; }
      }
      this.data[t.key] = d;
    }));
    if (!this.space()) { this.selected = this.spaces()[0]?.key || ''; this.thread = ''; }
  },
  draftKey(kind, parent = '') { return `${this.selected}:${kind}:${parent}`; },
  form(kind, parent = '') {
    const draft = this.drafts[this.draftKey(kind, parent)] || {};
    const contact = kind === 'contact';
    return `<form class="panel pad comm-form" data-kind="${kind}" data-parent="${esc(parent)}">
      <h2>${contact ? 'Contact your leaders privately' : parent ? 'Write a reply' : 'Start a conversation'}</h2>
      ${!parent ? `<div class="field"><label for="comm-title">${contact ? 'Subject' : 'Title'}</label><input id="comm-title" name="title" maxlength="120" value="${esc(draft.title || '')}" required></div>` : ''}
      <div class="field"><label for="comm-body">${parent ? 'Reply' : 'Message'}</label><textarea id="comm-body" name="body" maxlength="4000" required>${esc(draft.body || '')}</textarea></div>
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
  board(d = this.data[this.selected]) { return (d?.board || []).filter(r => !r.parent.startsWith('pin:')); },
  roots(d = this.data[this.selected]) {
    const rows = this.board(d), roots = rows.filter(r => !r.parent), ids = new Set(roots.map(r => r.id));
    for (const r of rows) if (r.parent && !ids.has(r.parent)) { ids.add(r.parent); roots.push({id:r.parent,parent:'',title:'Deleted conversation',text:'The original message was deleted.',name:'',email:'',at:r.at,deleted:true}); }
    return roots;
  },
  pinned(id, d = this.data[this.selected]) {
    if (!this.board(d).some(r => r.id === id && !r.parent)) return false;
    const admins = this.adminEmails();
    const events = (d?.board || []).filter(r => r.parent === 'pin:' + id && (admins.includes(lc(r.email)) || DEMO && this.own(r)) && ['pin','unpin'].includes(r.text));
    events.sort((a,b) => b.at.localeCompare(a.at) || String(b.id).localeCompare(String(a.id),undefined,{numeric:true}));
    return events[0]?.text === 'pin';
  },
  icon(name) { const paths = {edit:'M12 20h9M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z',delete:'M3 6h18M9 6V4h6v2M5 6l1 14h12l1-14M10 10v6M14 10v6',pin:'m16 3 5 5-4 1-3 5v3l-3-3-7 7 7-7-4-3h3l5-3Z'}; return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="${paths[name]}"/></svg>`; },
  actions(r, kind) {
    if (r.deleted) return '';
    const button = (action,label,pressed) => `<button type="button" class="iconbtn" data-comm="${action}" data-kind="${kind}" data-id="${esc(r.id)}" aria-label="${label}" title="${label}"${pressed == null ? '' : ` aria-pressed="${pressed}"`}${this.busy ? ' disabled' : ''}>${this.icon(action)}</button>`;
    return `<div class="comm-actions">${this.own(r) ? button('edit','Edit message') + button('delete','Delete message') : ''}${kind === 'board' && !r.parent && this.isAdmin() ? button('pin',this.pinned(r.id) ? 'Unpin message' : 'Pin message',this.pinned(r.id)) : ''}</div>`;
  },
  editForm(r, kind) {
    const edit = this.edit;
    if (!edit || edit.id !== r.id || edit.kind !== kind) return '';
    return `<form class="comm-edit-form"><h3>Edit message</h3>${!r.parent ? `<div class="field"><label for="comm-edit-title">${kind === 'contacts' ? 'Subject' : 'Title'}</label><input id="comm-edit-title" name="title" maxlength="120" value="${esc(edit.title)}" required></div>` : ''}<div class="field"><label for="comm-edit-body">Message</label><textarea id="comm-edit-body" name="body" maxlength="4000" required>${esc(edit.text)}</textarea></div><div class="comm-edit-buttons"><button class="btn primary" type="submit" ${this.busy ? 'disabled' : ''}>Save changes</button><button class="btn" type="button" data-comm="cancel-edit">Cancel</button></div></form>`;
  },
  async fresh(kind, id) {
    const base = this.meta[this.selected]?.[kind];
    if (!base) throw new Error('Refresh messages and try again.');
    const raw = await API.g(`${base}/${encodeURIComponent(id)}?$expand=fields($select=Title,Body,ParentId)`);
    const row = this.normalize(raw);
    if (!this.own(row)) throw new Error('Only the author can change this message.');
    return {base,row};
  },
  async act(action, kind, id) {
    if (this.busy || !this.space() || !['board','contacts'].includes(kind)) return;
    const d = this.data[this.selected], r = d?.[kind]?.find(x => x.id === id);
    if (!r || r.parent.startsWith('pin:')) return;
    if (action === 'pin') {
      if (kind !== 'board' || r.parent || !this.isAdmin()) return;
      this.busy = true; render();
      try {
        const text = this.pinned(id,d) ? 'unpin' : 'pin';
        const fields = {Title:'Board pin',Body:text,ParentId:'pin:'+id};
        let event;
        if (DEMO) event = {id:'demo-pin-'+String(this.pinSerial = (this.pinSerial || 0) + 1).padStart(8,'0'),title:fields.Title,text,parent:fields.ParentId,at:new Date().toISOString(),name:S.me.name,email:S.me.emails[0],authorId:this.meId};
        else { const raw = await API.g(this.meta[this.selected].board,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({fields})}); event = {...this.normalize(raw),title:fields.Title,text,parent:fields.ParentId,email:S.me.emails[0],authorId:this.meId}; }
        d.board.push(event); toast(text === 'pin' ? 'Message pinned.' : 'Message unpinned.');
      } catch (_) { toast('Could not update the pin. Please try again.',true); }
      finally { this.busy = false; render(); }
      return;
    }
    if (!this.own(r)) return;
    if (action === 'edit') { this.edit = {kind,id,title:r.title,text:r.text,etag:r.etag}; render(); return; }
    if (action !== 'delete' || !confirm('Delete your message? This removes its text. Other people’s replies will remain.')) return;
    this.busy = true; render();
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
    if (!r || !this.own(r)) return;
    const text = edit.text.trim(), title = edit.title.trim();
    if (!text || text.length > 4000 || !r.parent && (!title || title.length > 120)) return;
    this.busy = true; render();
    try {
      if (!DEMO) { const {base,row} = await this.fresh(edit.kind,edit.id); if (edit.etag && row.etag && edit.etag !== row.etag) throw new Error('Message changed elsewhere.'); await API.g(`${base}/${encodeURIComponent(edit.id)}/fields`,{method:'PATCH',headers:{'Content-Type':'application/json',...(row.etag ? {'If-Match':row.etag} : {})},body:JSON.stringify({Title:r.parent ? r.title : title,Body:text})}); }
      r.title = r.parent ? r.title : title; r.text = text; this.edit = null; toast('Message updated.');
    } catch (_) { toast('Could not save your changes. Your edit is still here. Please try again.',true); }
    finally { this.busy = false; render(); }
  },
  record(r, kind = 'board') { return `<div class="comm-record"><div class="comm-record-head"><div><b>${esc(r.name)}</b><time class="hint">${esc(r.at ? new Date(r.at).toLocaleString() : '')}</time></div>${this.actions(r,kind)}</div>${this.edit?.id === r.id && this.edit?.kind === kind ? this.editForm(r,kind) : `<p class="comm-text">${esc(r.text)}</p>`}</div>`; },
  view() {
    const spaces = this.spaces(), t = this.space();
    if (!t) return '<h1 class="page">Messages</h1><div class="empty">Your boards will appear after your ministry access is confirmed.</div>';
    const d = this.data[t.key], privateOn = DEMO || this.privateReady[t.key];
    const chips = spaces.map(s => `<button class="chipcheck" data-comm="space" data-key="${s.key}" aria-pressed="${s.key === t.key}">${esc(s.name)}${this.unread(s.key) ? ` <span aria-label="${this.unread(s.key)} unread messages">(${this.unread(s.key)})</span>` : ''}</button>`).join('');
    let content = '';
    if (!d) content = '<div class="empty">Loading messages…</div>';
    else if (this.mode === 'contact') {
      content = !privateOn ? '<div class="panel pad">Private messaging is awaiting secure setup. Please contact your ministry leader directly for now.</div>' : (d.contactError ? `<div class="err">${esc(d.contactError)}</div>` : this.form('contact') + `<h2>Private messages</h2><p class="hint">Your submissions appear here. Leaders see the submissions they are authorized to read.</p>` + (d.contacts.length ? d.contacts.slice().sort((a,b) => b.at.localeCompare(a.at)).map(r => `<article class="panel pad"><h3>${esc(r.title)}</h3>${this.record(r,'contacts')}</article>`).join('') : '<div class="empty">No private messages yet.</div>'));
    } else if (d.error) content = `<div class="err">${esc(d.error)}</div>`;
    else if (this.thread) {
      const root = this.roots(d).find(r => r.id === this.thread);
      content = root ? `<button class="btn" data-comm="back">Back to conversations</button><article class="panel pad"><h2>${this.pinned(root.id,d) ? '📌 ' : ''}${esc(root.title)}</h2>${this.record(root)}${this.board(d).filter(r => r.parent === root.id).sort((a,b) => a.at.localeCompare(b.at)).map(r => this.record(r)).join('')}</article>${this.form('reply',root.id)}` : '<div class="empty">This conversation is no longer available.</div>';
    } else {
      const seen = this.seen(t.key), roots = this.roots(d);
      const activity = r => this.board(d).filter(x => x.id === r.id || x.parent === r.id).reduce((at,x) => x.at > at ? x.at : at,r.at);
      content = this.form('post') + (roots.length ? roots.sort((a,b) => Number(this.pinned(b.id,d)) - Number(this.pinned(a.id,d)) || activity(b).localeCompare(activity(a))).map(r => {
        const replies = this.board(d).filter(x => x.parent === r.id), fresh = [r,...replies].some(x => !seen.has(x.id) && !this.own(x));
        return `<article class="panel pad"><div class="comm-topic-line"><button class="btn comm-topic" data-comm="thread" data-id="${esc(r.id)}">${this.pinned(r.id,d) ? '📌 ' : ''}${esc(r.title)}${fresh ? ' • Unread' : ''}</button>${this.actions(r,'board')}</div>${this.edit?.id === r.id && this.edit?.kind === 'board' ? this.editForm(r,'board') : ''}<p class="hint">${esc(r.name)} · ${plural(replies.length,'reply','replies')}</p><p class="comm-text">${esc(r.text.slice(0,180))}</p></article>`;
      }).join('') : '<div class="empty">No conversations yet. Start the first one.</div>');
    }
    return `<div class="comm"><div class="comm-head"><h1 class="page">Messages</h1><button class="iconbtn" data-comm="refresh" aria-label="Refresh messages" title="Refresh messages"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M23 4v6h-6M1 20v-6h6M3.5 9a9 9 0 0 1 14.9-5L23 10M1 14l4.6 6A9 9 0 0 0 20.5 15"/></svg></button></div><div class="chiprow" aria-label="Choose a board">${chips}</div><h2>${esc(t.name)}</h2><div class="chiprow"><button class="chipcheck" data-comm="mode" data-mode="board" aria-pressed="${this.mode === 'board'}">Team board</button><button class="chipcheck" data-comm="mode" data-mode="contact" aria-pressed="${this.mode === 'contact'}">Contact leaders privately</button></div>${content}</div>`;
  },
  async send(form) {
    if (this.busy) return;
    const t = this.space(), kind = form.dataset.kind, parent = form.dataset.parent || '';
    if (!t || !['contact','post','reply'].includes(kind)) return;
    const data = new FormData(form), body = String(data.get('body') || '').trim(), title = String(data.get('title') || '').trim();
    if (!body || body.length > 4000 || (!parent && (!title || title.length > 120))) return;
    const d = this.data[t.key], contact = kind === 'contact';
    if (!d || (contact ? !(DEMO || this.privateReady[t.key]) || d.contactError : d.error)) return;
    if (kind === 'reply' && !this.roots(d).some(r => r.id === parent)) return;
    this.busy = true; const key = this.draftKey(kind,parent); render();
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
  if (e.target.closest('.comm-edit-form') && Comm.edit) { const data = new FormData(e.target.closest('.comm-edit-form')); Comm.edit.title = String(data.get('title') || ''); Comm.edit.text = String(data.get('body') || ''); return; }
  const form = e.target.closest('.comm-form,.comm-edit-form'); if (!form) return;
  const data = new FormData(form);
  Comm.drafts[Comm.draftKey(form.dataset.kind,form.dataset.parent)] = { title: String(data.get('title') || ''), body: String(data.get('body') || '') };
});
document.addEventListener('submit', e => { if (e.target.matches('.comm-edit-form')) { e.preventDefault(); Comm.saveEdit(); return; } if (!e.target.matches('.comm-form')) return; e.preventDefault(); Comm.send(e.target); });
document.addEventListener('click', async e => {
  const b = e.target.closest('[data-comm]'); if (!b || Comm.busy) return;
  const action = b.dataset.comm;
  if (['edit','delete','pin'].includes(action)) { await Comm.act(action,b.dataset.kind,b.dataset.id); return; }
  if (action === 'cancel-edit') Comm.edit = null;
  if (action === 'space') { if (!Comm.space(b.dataset.key)) return; Comm.selected = b.dataset.key; Comm.thread = ''; Comm.edit = null; }
  if (action === 'mode') { Comm.mode = b.dataset.mode; Comm.thread = ''; Comm.edit = null; }
  if (action === 'back') Comm.thread = '';
  if (action === 'thread') { Comm.thread = b.dataset.id; Comm.markThread(); }
  if (action === 'refresh') { Comm.edit = null; b.disabled = true; await Comm.load(); }
  render();
});

// Keep badges current while the app is open; do not interrupt an active composer.
let communicationRefreshBusy = false;
async function refreshCommunications(){
  if (DEMO || S.phase !== 'ready' || S.loading || Comm.busy || communicationRefreshBusy || document.hidden || document.activeElement?.closest('.comm-form,.comm-edit-form')) return;
  communicationRefreshBusy = true;
  try {
    await Comm.load();
    if (S.view === 'messages' && Comm.mode === 'board' && Comm.thread) Comm.markThread();
    if (S.phase === 'ready' && !Comm.busy && !document.activeElement?.closest('.comm-form,.comm-edit-form')) render();
  } catch (_) { /* Retry on the next interval; list-specific failures are shown in Messages. */ }
  finally { communicationRefreshBusy = false; }
}
setInterval(refreshCommunications, 60000);
document.addEventListener('visibilitychange', () => { if (!document.hidden) refreshCommunications(); });
