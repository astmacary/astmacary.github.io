const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
function setup(){
  const c=vm.createContext({document:{addEventListener(){}},CHURCH_CAL:{editors:['admin@example.org']},lc:x=>String(x||'').toLowerCase(),S:{me:{emails:['admin@example.org']}},DEMO:false});
  vm.runInContext(fs.readFileSync(require('node:path').join(__dirname,'../rosters.js'),'utf8')+';globalThis.rosters=Rosters;',c);
  return c.rosters;
}
test('Shanelle displays in both leadership groups without an email or account grant',async()=>{
  const r=setup();
  const list=r.withDirectoryLeaders([]);
  assert.equal(list.length,1);
  assert.equal(list[0].name,'Shanelle Moss');
  assert.equal(list[0].email,undefined);
  assert.equal(r.tagFor(list[0]),'Greeters Leader');
  r.data.leadership=list;
  assert.equal(r.auxiliaryLeaders()[0].p.name,'Shanelle Moss');
  await assert.rejects(r.ensureLeadershipAccess(list[0]),/sign-in email/);
});
test('a later linked account replaces the display-only card without duplication',()=>{
  const r=setup();
  const list=r.withDirectoryLeaders([{userId:'real-account',name:'Shanelle Moss',email:'shanelle@example.org'}]);
  assert.equal(list.length,1);
  assert.equal(list[0].userId,'real-account');
  assert.equal(r.tagFor(list[0]),'Greeters Leader');
});
test('ministry leaders sort ahead of the rest of the Greeters roster',()=>{
  const r=setup();
  const sorted=r.orderPeople('greeters',[{name:'Aaron',ministryLeader:false},{name:'Shanelle Moss',ministryLeader:true}]);
  assert.equal(sorted[0].name,'Shanelle Moss');
});
