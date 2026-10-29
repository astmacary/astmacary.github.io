/* Communication records stay in their own ministry's SharePoint site.
   SharePoint permissions, not these UI checks, enforce access. See COMMUNICATIONS.md. */
const Comm = {
  selected: 'serve', thread: '', mode: 'board', data: {}, meta: {}, drafts: {}, busy: false,
  // Verified in SharePoint: author-only reads, member Contribute, designated leader Design.
  privateReady: { serve: true, productions: true, security: true, ushers: true, worship: true },
  lists: {"serve": {"board": "2ec3f7fa-168b-43f4-a1a6-552242ddef87", "contacts": "e507b7a1-8fd9-4495-a75b-14f11fde39ef"}, "productions": {"board": "c2cf8dce-848e-44a5-82f9-22be525b3a27", "contacts": "3483d881-7779-4d60-9c0d-566f453b448f"}, "security": {"board": "e7acd4bb-1b3f-49fd-919f-6f8ee9e7d80d", "contacts": "878e6cc8-5db9-4875-bd56-32566efd84ea"}, "ushers": {"board": "695182cc-9ac4-4897-a4e7-1cab655b8e6b", "contacts": "e0601cbd-55d0-45f3-9e90-3a014155d2fd"}, "worship": {"board": "29e8d328-dadd-4eaf-a4ad-fc484d00adff", "contacts": "23b8b906-50a4-4382-829d-2dd0486ec01c"}},
  spaces() {
    const ministries = myTeams().filter(t => !S.teams[t.key]?.error && !S.teams[t.key]?.noAccess);
    return ministries.length || isEditorOf(CHURCH_CAL)
      ? [{ ...CHURCH_CAL, key: 'serve', name: 'RCI Serve', full: 'RCI Serve' }, ...ministries] : [];
  },
  space(key = this.selected) { return this.spaces().find(t => t.key === key); },
  seenKey(key) { return `rci-serve:messages:${DEMO ? 'demo:' : ''}${S.me?.emails?.[0] || ''}:${key}`; },
  seen(key) { try { return new Set(JSON.parse(localStorage.getItem(this.seenKey(key)) || '[]')); } catch (_) { return new Set(); } },
  unread(key) { const seen = this.seen(key); return (this.data[key]?.board || []).filter(r => !seen.has(r.id) && !S.me?.emails.includes(r.email)).length; },
  totalUnread() { return this.spaces().reduce((n,t) => n + this.unread(t.key), 0); },
  markThread() {
    const d = this.data[this.selected]; if (!d || !this.thread) return;
    const seen = this.seen(this.selected);
    d.board.filter(r => r.id === this.thread || r.parent === this.thread).forEach(r => seen.add(r.id));
    try { localStorage.setItem(this.seenKey(this.selected), JSON.stringify([...seen].slice(-10000))); } catch (_) {}
  },
  normalize(r) { r = { ...r, fields: r.fields || {} }; return { id: String(r.id), parent: String(r.fields.ParentId || ''), title: r.fields.Title || '', text: r.fields.Body || '', name: r.createdBy?.user?.displayName || 'Team member', email: lc(r.createdBy?.user?.email), at: r.createdDateTime || '' }; },
  async list(t, name) {
    const site = await API.g(`/sites/${CONFIG.siteHost}:${t.sitePath}?$select=id`);
    const id = this.lists[t.key]?.[name === 'Team Board' ? 'board' : 'contacts'];
    if (!id) throw new Error('Message storage has not been configured.');
    const list = await API.g(`/sites/${site.id}/lists/${id}?$select=id`);
    return `/sites/${site.id}/lists/${list.id}/items`;
  },
  async load() {
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
        d.board = (await API.all(base + '?$expand=fields($select=Title,Body,ParentId)&$top=200')).map(r => this.normalize(r));
      } catch (e) { d.error = e.status === 404 ? 'This board is awaiting setup. Please contact your leader.' : 'Unable to load this board. Refresh to try again.'; }
      if (this.privateReady[t.key]) {
        try {
          const base = await this.list(t, 'Private Contacts'); (this.meta[t.key] ||= {}).contacts = base;
          // Server returns only records the signed-in person is allowed to read.
          d.contacts = (await API.all(base + '?$expand=fields($select=Title,Body,ParentId)&$top=200')).map(r => this.normalize(r));
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
  record(r) { return `<div class="comm-record"><b>${esc(r.name)}</b> <time class="hint">${esc(r.at ? new Date(r.at).toLocaleString() : '')}</time><p class="comm-text">${esc(r.text)}</p></div>`; },
  view() {
    const spaces = this.spaces(), t = this.space();
    if (!t) return '<h1 class="page">Messages</h1><div class="empty">Your boards will appear after your ministry access is confirmed.</div>';
    const d = this.data[t.key], privateOn = DEMO || this.privateReady[t.key];
    const chips = spaces.map(s => `<button class="chipcheck" data-comm="space" data-key="${s.key}" aria-pressed="${s.key === t.key}">${esc(s.name)}${this.unread(s.key) ? ` <span aria-label="${this.unread(s.key)} unread messages">(${this.unread(s.key)})</span>` : ''}</button>`).join('');
    let content = '';
    if (!d) content = '<div class="empty">Loading messages…</div>';
    else if (this.mode === 'contact') {
      content = !privateOn ? '<div class="panel pad">Private messaging is awaiting secure setup. Please contact your ministry leader directly for now.</div>' : (d.contactError ? `<div class="err">${esc(d.contactError)}</div>` : this.form('contact') + `<h2>Private messages</h2><p class="hint">Your submissions appear here. Leaders see the submissions they are authorized to read.</p>` + (d.contacts.length ? d.contacts.slice().sort((a,b) => b.at.localeCompare(a.at)).map(r => `<article class="panel pad"><h3>${esc(r.title)}</h3>${this.record(r)}</article>`).join('') : '<div class="empty">No private messages yet.</div>'));
    } else if (d.error) content = `<div class="err">${esc(d.error)}</div>`;
    else if (this.thread) {
      const root = d.board.find(r => r.id === this.thread && !r.parent);
      content = root ? `<button class="btn" data-comm="back">Back to conversations</button><article class="panel pad"><h2>${esc(root.title)}</h2>${this.record(root)}${d.board.filter(r => r.parent === root.id).sort((a,b) => a.at.localeCompare(b.at)).map(r => this.record(r)).join('')}</article>${this.form('reply',root.id)}` : '<div class="empty">This conversation is no longer available.</div>';
    } else {
      const seen = this.seen(t.key), roots = d.board.filter(r => !r.parent);
      const activity = r => d.board.filter(x => x.id === r.id || x.parent === r.id).reduce((at,x) => x.at > at ? x.at : at,r.at);
      content = this.form('post') + (roots.length ? roots.sort((a,b) => activity(b).localeCompare(activity(a))).map(r => {
        const replies = d.board.filter(x => x.parent === r.id), fresh = [r,...replies].some(x => !seen.has(x.id) && !S.me.emails.includes(x.email));
        return `<article class="panel pad"><button class="btn comm-topic" data-comm="thread" data-id="${esc(r.id)}">${esc(r.title)}${fresh ? ' • Unread' : ''}</button><p class="hint">${esc(r.name)} · ${plural(replies.length,'reply','replies')}</p><p class="comm-text">${esc(r.text.slice(0,180))}</p></article>`;
      }).join('') : '<div class="empty">No conversations yet. Start the first one.</div>');
    }
    return `<div class="comm"><h1 class="page">Messages</h1><div class="chiprow" aria-label="Choose a board">${chips}</div><h2>${esc(t.name)}</h2><div class="chiprow"><button class="chipcheck" data-comm="mode" data-mode="board" aria-pressed="${this.mode === 'board'}">Team board</button><button class="chipcheck" data-comm="mode" data-mode="contact" aria-pressed="${this.mode === 'contact'}">Contact leaders privately</button><button class="btn" data-comm="refresh">Refresh messages</button></div>${content}</div>`;
  },
  async send(form) {
    if (this.busy) return;
    const t = this.space(), kind = form.dataset.kind, parent = form.dataset.parent || '';
    if (!t || !['contact','post','reply'].includes(kind)) return;
    const data = new FormData(form), body = String(data.get('body') || '').trim(), title = String(data.get('title') || '').trim();
    if (!body || body.length > 4000 || (!parent && (!title || title.length > 120))) return;
    const d = this.data[t.key], contact = kind === 'contact';
    if (!d || (contact ? !(DEMO || this.privateReady[t.key]) || d.contactError : d.error)) return;
    if (kind === 'reply' && !d.board.some(r => r.id === parent && !r.parent)) return;
    this.busy = true; const key = this.draftKey(kind,parent); render();
    try {
      const fields = { Title: title || 'Reply', Body: body, ParentId: parent };
      let r;
      if (DEMO) r = { id: 'demo-' + crypto.randomUUID(), parent, title: fields.Title, text: body, name: S.me.name, email: S.me.emails[0], at: new Date().toISOString() };
      else {
        const base = this.meta[t.key]?.[contact ? 'contacts' : 'board']; if (!base) throw new Error('This space is not ready. Refresh and try again.');
        r = this.normalize(await API.g(base, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ fields }) }));
        // Some Graph responses omit expanded fields; use the submitted plain text.
        Object.assign(r, { parent, title: fields.Title, text: body, name: S.me.name, email: S.me.emails[0] });
      }
      d[contact ? 'contacts' : 'board'].push(r); delete this.drafts[key];
      if (!contact) { this.thread = parent || r.id; this.markThread(); }
      toast(DEMO ? 'Saved in this demo only.' : contact ? 'Private message sent to your leaders.' : 'Message posted.');
    } catch (_) { toast('Your message could not be sent. Your draft is still here; please try again.',true); }
    finally { this.busy = false; render(); }
  }
};
document.addEventListener('input', e => {
  const form = e.target.closest('.comm-form'); if (!form) return;
  const data = new FormData(form);
  Comm.drafts[Comm.draftKey(form.dataset.kind,form.dataset.parent)] = { title: String(data.get('title') || ''), body: String(data.get('body') || '') };
});
document.addEventListener('submit', e => { if (!e.target.matches('.comm-form')) return; e.preventDefault(); Comm.send(e.target); });
document.addEventListener('click', async e => {
  const b = e.target.closest('[data-comm]'); if (!b || Comm.busy) return;
  const action = b.dataset.comm;
  if (action === 'space') { if (!Comm.space(b.dataset.key)) return; Comm.selected = b.dataset.key; Comm.thread = ''; }
  if (action === 'mode') { Comm.mode = b.dataset.mode; Comm.thread = ''; }
  if (action === 'back') Comm.thread = '';
  if (action === 'thread') { Comm.thread = b.dataset.id; Comm.markThread(); }
  if (action === 'refresh') { b.disabled = true; await Comm.load(); }
  render();
});

// Keep badges current while the app is open; do not interrupt an active composer.
let communicationRefreshBusy = false;
async function refreshCommunications(){
  if (DEMO || S.phase !== 'ready' || S.loading || Comm.busy || communicationRefreshBusy || document.hidden || document.activeElement?.closest('.comm-form')) return;
  communicationRefreshBusy = true;
  try {
    await Comm.load();
    if (S.view === 'messages' && Comm.mode === 'board' && Comm.thread) Comm.markThread();
    if (S.phase === 'ready' && !Comm.busy && !document.activeElement?.closest('.comm-form')) render();
  } catch (_) { /* Retry on the next interval; list-specific failures are shown in Messages. */ }
  finally { communicationRefreshBusy = false; }
}
setInterval(refreshCommunications, 60000);
document.addEventListener('visibilitychange', () => { if (!document.hidden) refreshCommunications(); });
