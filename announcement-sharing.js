/* Copy a church announcement only to destinations the signed-in person administers.
   This module never changes membership or permissions. */
const AnnShare = {
  sourceId:'', snapshot:null, selected:new Set(), sent:new Set(), busy:false, error:'',
  targets(){ return TEAMS.filter(t => isEditorOf(t)); },
  find(id=this.sourceId){ return (S.church.anns || []).find(a => String(a.id) === String(id)); },
  flyerKey(value){ if (!value) return ''; try { const p = typeof value === 'string' ? JSON.parse(value) : value; return p.serverRelativeUrl || p.fileName || JSON.stringify(p); } catch (_) { return String(value); } },
  signature(a){ return JSON.stringify([a.text || '',a.until || '',this.flyerKey(a.flyer)]); },
  open(id){
    const a=this.find(id); if (!a || !this.targets().length) return;
    this.sourceId=String(id); this.snapshot={...a}; this.selected=new Set(); this.sent=new Set(); this.error=''; render();
  },
  close(){ if(this.busy)return; this.sourceId=''; this.snapshot=null; this.error=''; render(); },
  modal(){
    if(!this.sourceId || !this.snapshot)return '';
    return `<div class="roster-overlay"><section class="panel roster-dialog" role="dialog" aria-modal="true" aria-labelledby="share-ann-title"><h2 id="share-ann-title">Share Church Announcement</h2><form id="share-announcement-form" class="sec"><div class="chiprow ann-destinations" role="group" aria-label="Share announcement with">${this.targets().map(t=>annDestinationPill(t,this.selected.has(t.key),'ann-share-target',this.busy || this.sent.has(t.key))).join('')}</div><div class="ann-share-preview">${esc(this.snapshot.text)}</div>${this.snapshot.until?`<p class="hint">Shows until ${esc(fmtShort(this.snapshot.until))}</p>`:''}${this.snapshot.flyer?'<p class="hint">The flyer will be copied too.</p>':''}${this.sent.size?`<p class="hint" role="status">Shared with ${esc([...this.sent].map(k=>teamOf(k).name).join(', '))}.</p>`:''}${this.error?`<p class="err" role="alert">${esc(this.error)}</p>`:''}<div class="acts2"><button class="btn primary" ${this.busy || ![...this.selected].some(k=>!this.sent.has(k))?'disabled':''}>${this.busy?'Sharing…':'Share announcement'}</button><button class="btn" type="button" data-act="ann-share-close" ${this.busy?'disabled':''}>${this.sent.size?'Done':'Cancel'}</button></div></form></section></div>`;
  },
  async freshSource(){
    const local=this.find(); if(!local)throw new Error('This church announcement is no longer available.');
    if(DEMO)return local;
    const {base,listId}=await API.churchMeta();
    const r=await API.g(`${base}/${listId}/items/${encodeURIComponent(this.sourceId)}?$expand=fields`);
    const f=r.fields || {};
    const text=String(f.Announcements || '').trim(); if(!text)throw new Error('This church announcement is no longer available.');
    return {...local,text,until:spDate(f.ShowUntil),flyer:f.Flyer || ''};
  },
  async send(){
    if(this.busy || !this.snapshot)return;
    const keys=[...this.selected].filter(k=>!this.sent.has(k));
    if(!keys.length)return;
    if(keys.some(k=>!this.targets().some(t=>t.key===k))){this.error='Choose a team you administer.';render();return;}
    const identity=S.me?.id || S.me?.emails?.join('|') || '';
    const samePerson=()=>identity===(S.me?.id || S.me?.emails?.join('|') || '');
    this.busy=true;this.error='';render();
    try{
      const source=await this.freshSource();
      if(!samePerson())throw new Error('The signed-in account changed. Close this window and try again.');
      if(this.signature(source)!==this.signature(this.snapshot))throw new Error('This announcement changed. Close this window and share the updated version.');
      let flyerBlob=null;
      if(source.flyer && !DEMO){
        const url=await API.flyer(CHURCH_CAL,source);
        if(!url || !url.startsWith('blob:'))throw new Error('The flyer could not be loaded. Nothing has been shared yet.');
        const r=await fetch(url);if(!r.ok)throw new Error('The flyer could not be loaded. Nothing has been shared yet.');
        const blob=await r.blob();flyerBlob=await resizeImage(new File([blob],'church-announcement.jpg',{type:blob.type}),1600);
      }
      const flyerFailures=[];
      for(const key of keys){
        if(!samePerson())throw new Error('The signed-in account changed. Check the destination before continuing.');
        const t=TEAMS.find(t=>t.key===key);if(!t || !isEditorOf(t))throw new Error('Your access to this destination changed.');
        const opts={text:source.text,until:source.until || '',flyerBlob,flyerName:'church-announcement.jpg'};
        const result=DEMO?{id:'demo-share-'+Date.now()+'-'+key,text:opts.text,until:opts.until,flyer:source.flyer || ''}:await API.announcementPush(t,opts);
        this.sent.add(key);
        if(S.teams[key])S.teams[key].announcements=[...(S.teams[key].announcements || []),result];
        if(result.flyerError)flyerFailures.push(t.name);
      }
      if(flyerFailures.length)this.error='The announcement text was shared, but the flyer could not be attached for '+flyerFailures.join(', ')+'. Add it from that team’s announcement editor.';
      else if([...this.selected].every(k=>this.sent.has(k))){saveCache();this.sourceId='';this.snapshot=null;toast('Announcement shared');}
    }catch(e){this.error=e.message || 'Sharing could not be confirmed. Check the destination before trying again.';}
    this.busy=false;render();
  }
};
document.addEventListener('submit',async e=>{if(e.target.id!=='share-announcement-form')return;e.preventDefault();await AnnShare.send();});
