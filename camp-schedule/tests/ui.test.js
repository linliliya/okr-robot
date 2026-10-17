const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const path=require('node:path');
const { JSDOM }=require(process.env.JSDOM_MODULE || 'jsdom');
const root=path.resolve(__dirname,'..');
const tick=()=>new Promise(resolve=>setTimeout(resolve,10));
const copy=x=>JSON.parse(JSON.stringify(x));
function server(){
 let camps=[],tasks=[];let sequence=0;
 return {calls:[], async rpc(name,p){
  this.calls.push([name,copy(p)]);
  const active=()=>camps.filter(c=>c.program===p.p_program);
  if(name==='schedule_snapshot')return {data:copy({camps:active(),tasks:tasks.filter(t=>active().some(c=>c.id===t.camp_id))})};
  const conflict=()=>({error:{code:'P0001',message:'schedule_conflict'}});
  if(name==='save_schedule_camp'){
   if(!p.p_id){const c={id:'camp-'+(++sequence),program:p.p_program,term:p.p_term,open_date:p.p_open_date,skips:p.p_skips,version:1};camps.push(c);return {data:copy(c)};}
   const c=active().find(c=>c.id===p.p_id);if(!c || c.version!==p.p_version)return conflict();
   Object.assign(c,{term:p.p_term,open_date:p.p_open_date,skips:p.p_skips,version:c.version+1});return {data:copy(c)};
  }
  if(name==='set_schedule_task'){
   let task=tasks.find(t=>t.camp_id===p.p_camp_id&&t.task_key===p.p_task_key);
   if((task?.version||0)!==p.p_version)return conflict();
   if(!task){task={camp_id:p.p_camp_id,task_key:p.p_task_key,version:0};tasks.push(task);}
   Object.assign(task,{completed:p.p_completed,version:task.version+1});return {data:copy(task)};
  }
  if(name==='import_schedule_camps'){
   const inserted=[],skipped=[];
   for(const c of p.p_camps){if(active().some(x=>x.term===c.term)){skipped.push(c.term);continue;}
    const saved=(await this.rpc('save_schedule_camp',{...p,p_id:null,p_term:c.term,p_open_date:c.open_date,p_skips:c.skips})).data;
    for(const key of c.completed_tasks)tasks.push({camp_id:saved.id,task_key:key,completed:true,version:1});inserted.push(c.term);
   }
   return {data:{inserted,skipped}};
  }
  throw new Error('Unexpected RPC '+name);
 }};
}
async function windowFor(backend,{loggedIn=true,legacy=false}={}){
 const html=fs.readFileSync(path.join(root,'index.html'),'utf8');
 const dom=new JSDOM(html,{url:'https://internal.aurange.cn/test/#9:2030-02-03',runScripts:'outside-only',pretendToBeVisual:true});
 const w=dom.window;w.alert=()=>{};w.confirm=()=>true;w.HTMLElement.prototype.scrollIntoView=function(){};
 w.eval(fs.readFileSync(path.join(root,'config.js'),'utf8'));
 const prefix=w.SCHEDULE_CONFIG.program==='ai'?'ai':'crawler';
 if(legacy)w.localStorage.setItem(prefix+'-camps-v1',JSON.stringify([{term:8,openDate:'2030-01-06',skips:[]}]));
 const session=loggedIn?{user:{id:'member',email:'member@example.test'}}:null;
 const client={rpc:(name,p)=>backend.rpc(name,p),auth:{getSession:async()=>({data:{session}}),onAuthStateChange:()=>({data:{subscription:{unsubscribe(){}}}})},
  channel:()=>({on(){return this},subscribe(){return this}}),removeChannel(){}};
 w.supabase={createClient:()=>client};w.eval(fs.readFileSync(path.join(root,'cloud.js'),'utf8'));
 const sop=fs.readFileSync(path.join(root,'data.js'),'utf8');
 const inline=[...w.document.scripts].filter(s=>!s.src).map(s=>s.textContent).join('\n');
 w.eval(sop+'\n'+inline+'\nwindow.__test={get cloud(){return cloud}, get legacy(){return legacy}};');await tick();
 return {w,close(){w.__test.cloud.stop();w.close();}};
}

test('two independent pages share camps and completion without URL or local data overwrites',async()=>{
 const backend=server(),first=await windowFor(backend,{legacy:true}),second=await windowFor(backend);
 try{
  const a=first.w,b=second.w;
  assert.equal(a.__test.cloud.state.camps.length,0);
  assert.equal(backend.calls.filter(([name])=>name!=='schedule_snapshot').length,0,'old links never silently write to cloud');
  const key=a.SCHEDULE_CONFIG.program==='ai'?'ai-camps-v1':'crawler-camps-v1';const before=a.localStorage.getItem(key);
  a.openForm(null);a.document.getElementById('f-term').value='1';a.document.getElementById('f-date').value='2030-01-06';await a.saveForm(null);
  await b.__test.cloud.refresh();assert.match(b.document.getElementById('root').textContent,/第 1 期/);
  const firstCheckbox=a.document.querySelector('.ckb');assert.ok(firstCheckbox);await a.toggleDone(Number(firstCheckbox.getAttribute('onclick').match(/\d+/)[0]));await b.__test.cloud.refresh();
  assert.equal(b.document.querySelector('.ckb').getAttribute('aria-pressed'),'true');
  a.openForm(1);b.openForm(1);
  b.document.getElementById('f-date').value='2030-01-13';await b.saveForm(1);await a.__test.cloud.refresh();
  assert.equal(a.document.getElementById('f-date').value,'2030-01-06','background sync preserves the open draft');
  await a.saveForm(1);assert.match(a.document.getElementById('notice').textContent,/其他人/);
  assert.ok(a.document.getElementById('f-date'),'conflict keeps draft visible');
  a.closeForm();a.openImport('local');await a.importLegacy('local');await b.__test.cloud.refresh();
  assert.equal(b.__test.cloud.state.camps.length,2);assert.equal(a.localStorage.getItem(key),before);
  assert.equal(a.__test.cloud.state.camps.some(c=>c.term===9),false,'shared URL still requires explicit import');
  assert.equal(a.__test.cloud.state.camps.find(c=>c.term===1).openDate,'2030-01-13');
 }finally{first.close();second.close();}
});

test('logged out visitors do not fetch or render team data',async()=>{
 const backend=server(),page=await windowFor(backend,{loggedIn:false,legacy:true});
 try{
  assert.ok(page.w.document.getElementById('login-email'));
  assert.equal(backend.calls.length,0);assert.equal(page.w.document.querySelector('.camp'),null);
 }finally{page.close();}
});
