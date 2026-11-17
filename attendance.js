/* Church attendance counts (shared by the RCI Ushers app and RCI Serve).
   Stored on the RCI Serve site, list "Attendance": Service Date, Service, In the Sanctuary, Notes.
   The host app calls Attendance.init({ g, render, toast, canEdit, services }) once, then:
   - Attendance.view()            → the page body (form + recent entries)
   - Attendance.click(button)     → handles data-act="att-*" buttons; returns true when handled
   - Attendance.submit(form)      → handles the form with id "attForm"; returns true when handled
   Same file in both apps; keep them identical. */
const Attendance = {
  siteHost: "xb3hc3233xdwtuxk4m6un8xymmy.sharepoint.com",
  sitePath: "/sites/RCIServe",
  listId: "06990467-6806-419d-8287-ca7a511eb9d4",
  siteId: "", rows: null, loading: false, error: "", busy: false, editId: null, askDel: null,
  host: { g: null, render: () => {}, toast: () => {}, canEdit: () => false, services: () => [] },
  init(h){ Object.assign(this.host, h); },
  esc(s){ return String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]); },
  pad(n){ return String(n).padStart(2, "0"); },
  key(d){ return `${d.getFullYear()}-${this.pad(d.getMonth() + 1)}-${this.pad(d.getDate())}`; },
  today(){ return this.key(new Date()); },
  // A stored date-only value comes back as an instant; read it in local time (saved at noon UTC, so the day is right).
  dayOf(v){ if (!v) return ""; const d = new Date(v); return isNaN(d) ? String(v).slice(0, 10) : this.key(d); },
  fmt(k){ const [y, m, d] = k.split("-").map(Number); return new Date(y, m - 1, d).toLocaleDateString(undefined, { weekday: "short", month: "short", day: "numeric", year: "numeric" }); },
  isSunday(k){ const [y, m, d] = k.split("-").map(Number); return new Date(y, m - 1, d).getDay() === 0; },
  async base(){
    if (!this.siteId) this.siteId = (await this.host.g(`/sites/${this.siteHost}:${this.sitePath}?$select=id`)).id;
    return `/sites/${this.siteId}/lists/${this.listId}/items`;
  },
  async load(force){
    if (this.loading || (this.rows && !force)) return;
    this.loading = true; this.error = "";
    try {
      const out = []; let next = `${await this.base()}?$expand=fields($select=ServiceDate,ServiceCategory,Sanctuary,Notes)&$top=500`;
      while (next){ const j = await this.host.g(next); out.push(...(j.value || [])); next = j["@odata.nextLink"]; }
      this.rows = out.map(r => ({ id: String(r.id), date: this.dayOf(r.fields.ServiceDate), service: r.fields.ServiceCategory || "", count: r.fields.Sanctuary ?? null,
        notes: r.fields.Notes || "", by: r.lastModifiedBy?.user?.displayName || r.createdBy?.user?.displayName || "", at: r.lastModifiedDateTime || "" }))
        .filter(r => r.date).sort((a, b) => b.date.localeCompare(a.date) || a.service.localeCompare(b.service));
    } catch(e){ this.error = e.message; this.rows = this.rows || []; }
    this.loading = false; this.host.render();
  },
  css(){
    if (document.getElementById("att-css")) return;
    const st = document.createElement("style"); st.id = "att-css";
    st.textContent = `.att{display:flex;flex-direction:column;gap:16px}.att form.panel{display:flex;flex-direction:column;gap:12px;padding:14px}.att .field{display:flex;flex-direction:column;gap:6px}
.att-list{list-style:none;margin:0;padding:4px 14px}.att-list li{padding:10px 0;border-bottom:1px solid var(--line)}.att-list li:last-child{border-bottom:0}
.att-main{display:flex;align-items:center;gap:12px}.att-main>span:first-of-type{flex:1;display:flex;flex-direction:column;min-width:0}.att-n{font-size:1.5rem;min-width:64px;color:var(--accent)}
.att-note{opacity:.8}.att-btns{display:flex;gap:10px;flex:none}.att-del{font-size:13px;margin-top:6px}.att .btn{width:100%;justify-content:center}`;
    document.head.appendChild(st);
  },
  view(){
    this.css();
    if (!this.host.canEdit()) return `<div class="empty">Only admins can enter attendance.</div>`;
    if (!this.rows){ this.load(); return `<div class="empty">Loading attendance…</div>`; }
    const ed = this.editId ? this.rows.find(r => r.id === this.editId) : null;
    const services = [...new Set([...(this.host.services() || []), ...(ed && ed.service ? [ed.service] : [])])];
    const day = ed ? ed.date : this.today();
    const defService = ed ? ed.service : (this.isSunday(day) && services.includes("Sunday Morning") ? "Sunday Morning" : (services[0] || ""));
    const recent = this.rows.slice(0, 30);
    return `<div class="att">${this.error ? `<div class="err">Couldn't read the Attendance list. ${this.esc(this.error)}</div>` : ""}
      <form class="panel stack" id="attForm">
        <h2 class="section-label" style="margin:0">${ed ? "Edit Attendance" : "Enter Attendance"}</h2>
        <div class="field"><label for="att-date">Service date</label><input id="att-date" type="date" required value="${this.esc(day)}" max="${this.today()}"></div>
        <div class="field"><label for="att-service">Service</label><select id="att-service" required>${services.map(s => `<option ${s === defService ? "selected" : ""}>${this.esc(s)}</option>`).join("")}</select></div>
        <div class="field"><label for="att-count">In the Sanctuary</label><input id="att-count" type="number" inputmode="numeric" min="0" step="1" required placeholder="0" value="${ed && ed.count != null ? this.esc(ed.count) : ""}"></div>
        <div class="field"><label for="att-notes">Notes <span class="hint">(optional)</span></label><textarea id="att-notes" rows="2" maxlength="1000">${ed ? this.esc(ed.notes) : ""}</textarea></div>
        <button class="btn primary block" type="submit" ${this.busy ? "disabled" : ""}>${this.busy ? "Saving…" : ed ? "Save Changes" : "Save Attendance"}</button>
        ${ed ? `<button class="btn block" type="button" data-act="att-cancel">Cancel</button>` : ""}
      </form>
      <section><h2 class="section-label">Recent Counts</h2>
        ${recent.length ? `<div class="panel"><ul class="att-list">${recent.map(r => `<li>
          <div class="att-main"><b class="att-n">${r.count == null ? "—" : Number(r.count).toLocaleString()}</b><span><b>${this.esc(this.fmt(r.date))}</b><small>${this.esc(r.service)}${r.by ? ` · ${this.esc(r.by)}` : ""}</small>${r.notes ? `<small class="att-note">${this.esc(r.notes)}</small>` : ""}</span>
            <span class="att-btns"><button class="linkbtn edit" type="button" data-act="att-edit" data-id="${r.id}">Edit</button><button class="linkbtn" type="button" data-act="att-ask" data-id="${r.id}">Remove</button></span></div>
          ${this.askDel === r.id ? `<div class="att-del">Remove this count? <button class="linkbtn danger" type="button" data-act="att-del" data-id="${r.id}">Yes, Remove</button> <button class="linkbtn" type="button" data-act="att-keep">Keep It</button></div>` : ""}
        </li>`).join("")}</ul></div>` : `<div class="empty">No attendance has been entered yet.</div>`}
      </section></div>`;
  },
  async click(b){
    const act = b.dataset.act; if (!act || !act.startsWith("att-")) return false;
    if (act === "att-edit"){ this.editId = b.dataset.id; this.askDel = null; this.host.render(); window.scrollTo(0, 0); }
    else if (act === "att-cancel"){ this.editId = null; this.host.render(); }
    else if (act === "att-ask"){ this.askDel = b.dataset.id; this.host.render(); }
    else if (act === "att-keep"){ this.askDel = null; this.host.render(); }
    else if (act === "att-del" && this.host.canEdit()){
      try { await this.host.g(`${await this.base()}/${b.dataset.id}`, { method: "DELETE" });
        this.rows = this.rows.filter(r => r.id !== b.dataset.id); if (this.editId === b.dataset.id) this.editId = null; this.host.toast("Attendance removed"); }
      catch(e){ this.host.toast("Couldn't remove it. " + e.message, true); }
      this.askDel = null; this.host.render();
    }
    return true;
  },
  async submit(form){
    if (form.id !== "attForm") return false;
    if (!this.host.canEdit() || this.busy) return true;
    const date = form.querySelector("#att-date").value, service = form.querySelector("#att-service").value.trim();
    const count = Number(form.querySelector("#att-count").value), notes = form.querySelector("#att-notes").value.trim();
    if (!date || !service || !Number.isInteger(count) || count < 0){ this.host.toast("Enter the date, service and a whole number.", true); return true; }
    // One count per date and service: saving the same pair again updates it.
    const same = this.rows.find(r => r.date === date && r.service === service && r.id !== this.editId);
    const id = this.editId || (same && same.id);
    const fields = { Title: `${service} ${date}`, ServiceDate: date + "T12:00:00Z", ServiceCategory: service, Sanctuary: count, Notes: notes };
    this.busy = true; this.host.render();
    try {
      if (id) await this.host.g(`${await this.base()}/${id}/fields`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify(fields) });
      else await this.host.g(await this.base(), { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fields }) });
      this.host.toast(id && !this.editId ? `Updated the ${service} count for ${this.fmt(date)}` : "Attendance saved");
      this.editId = null; this.busy = false; await this.load(true);
    } catch(e){ this.busy = false; this.host.toast("Couldn't save attendance. " + e.message, true); this.host.render(); }
    return true;
  }
};
