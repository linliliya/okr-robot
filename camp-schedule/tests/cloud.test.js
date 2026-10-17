const test = require('node:test');
const assert = require('node:assert/strict');
const { Store, parseShare, readLegacy, normalizeCamp } = require('../cloud.js');
const copy = x => JSON.parse(JSON.stringify(x));
const session = { user: { id: 'member', email: 'member@example.test' } };
const row = { id:'camp', term:1, open_date:'2030-01-06', skips:[], version:1 };
function client(execute = () => ({ camps:[row], tasks:[] })) {
  return {
    rpc: async (name,params) => ({ data:await execute(name,params), error:null }),
    auth: { getSession:async()=>({data:{session}}),onAuthStateChange:()=>({data:{subscription:{unsubscribe(){}}}}),signOut:async()=>({}) },
    channel:()=>({on(){return this},subscribe(){return this}}),removeChannel(){},
  };
}

test('legacy links validate actual dates and never write browser storage', () => {
  assert.equal(parseShare('#1:2026-02-30').length,0);
  assert.equal(parseShare('#%broken').length,0);
  assert.equal(parseShare('#1x:2026-01-01').length,0);
  assert.equal(parseShare('#1:2026-01-01,1:2026-02-01').length,0);
  assert.equal(parseShare('#1:2028-02-29:3.6').length,1);
  const store={getItem:key=>key==='camps'?JSON.stringify([{term:2,openDate:'2026-01-04',skips:[]}]):'{"old-task":1}',setItem(){assert.fail('must not overwrite legacy data')}};
  const old=readLegacy(store,'#1:2026-01-04','camps','done');
  assert.equal(old.local[0].term,2);assert.equal(old.shared[0].term,1);assert.equal(old.done['old-task'],1);
  assert.throws(()=>normalizeCamp({term:1.3,openDate:'2026-01-04',skips:[]}));
});

test('newer snapshots win when requests resolve out of order', async () => {
  const store=new Store(client(),'ai');await store.setSession(session);
  const pending=[];store.client.rpc=()=>new Promise(resolve=>pending.push(resolve));
  const first=store.refresh(),second=store.refresh();
  pending[1]({data:{camps:[{...row,term:2}],tasks:[]}});await second;
  pending[0]({data:{camps:[row],tasks:[]}});await first;
  assert.equal(store.state.camps[0].term,2);store.stop();
});

test('logout discards in-flight snapshots and removes team data', async () => {
  const store=new Store(client(),'ai');await store.setSession(session);
  let resolve;store.client.rpc=()=>new Promise(r=>resolve=r);const pending=store.refresh();
  await store.setSession(null);resolve({data:{camps:[row],tasks:[]}});await pending;
  assert.equal(store.state.phase,'login');assert.deepEqual(store.state.camps,[]);store.stop();
});

test('server rejection never marks a task completed and preserves the server version', async () => {
  const calls=[];
  const store=new Store(client((name,params)=>{
    calls.push([name,params]);
    if(name==='set_schedule_task')throw {code:'P0001',message:'schedule_conflict'};
    return {camps:[row],tasks:[{camp_id:'camp',task_key:'open|all|准备',completed:false,version:3}]};
  }),'ai');await store.setSession(session);
  await assert.rejects(store.setTask('camp|open|all|准备',true),e=>e.code==='conflict');
  assert.equal(store.state.tasks['camp|open|all|准备'].completed,false);
  assert.equal(calls.find(([name])=>name==='set_schedule_task')[1].p_version,3);store.stop();
});

test('connection loss keeps the last snapshot read-only; membership removal clears it', async () => {
  const store=new Store(client(),'crawler');await store.setSession(session);
  store.client.rpc=async()=>{throw new Error('Failed to fetch')};await store.refresh();
  assert.equal(store.state.camps.length,1);assert.equal(store.state.online,false);
  await assert.rejects(store.save({term:2,openDate:'2030-01-06',skips:[]},null));
  store.client.rpc=async()=>({error:{code:'42501',message:'schedule_members_only'}});await store.refresh();
  assert.equal(store.state.phase,'blocked');assert.equal(store.state.camps.length,0);store.stop();
});

test('editing sends the captured version, not a refreshed version', async () => {
  let params;
  const store=new Store(client((name,p)=>{
    if(name==='save_schedule_camp'){params=p;return row;}
    return {camps:[{...row,version:7}],tasks:[]};
  }),'ai');await store.setSession(session);
  await store.save({term:2,openDate:'2030-01-06',skips:[]},{id:'camp',version:4});
  assert.equal(params.p_version,4);assert.equal(params.p_program,'ai');store.stop();
});
