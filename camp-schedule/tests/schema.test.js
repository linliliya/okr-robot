const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const { PGlite }=require(process.env.PGLITE_MODULE || '@electric-sql/pglite');

test('database permissions, program isolation, stale edits, imports and shared task progress',async()=>{
 const pg=new PGlite();
 const member='00000000-0000-0000-0000-000000000001',other='00000000-0000-0000-0000-000000000002';
 try{
  await pg.exec(`create role anon; create role authenticated; create schema auth;
   create table auth.users(id uuid primary key);
   create function auth.uid() returns uuid language sql stable as $$select nullif(current_setting('request.jwt.claim.sub',true),'')::uuid$$;
   create function public.is_camp_review_member() returns boolean language sql stable as $$select auth.uid()='${member}'::uuid$$;
   grant usage on schema auth,public to anon,authenticated;
   insert into auth.users values('${member}'),('${other}');`);
  const migration=fs.readFileSync(require.resolve('../setup.sql'),'utf8');await pg.exec(migration);await pg.exec(migration);
  const identity=async(id,role='authenticated')=>{await pg.exec('reset role');await pg.query("select set_config('request.jwt.claim.sub',$1,false)",[id]);await pg.exec('set role '+role);};
  const rpc=async(name,params,casts)=> (await pg.query(`select public.${name}(${params.map((_,i)=>'$'+(i+1)+(casts?.[i]?'::'+casts[i]:'')).join(',')}) as value`,params)).rows[0].value;
  const save=(program,term,id=null,version=null,date='2030-01-06')=>rpc('save_schedule_camp',[program,id,version,term,date,[]],['text','uuid','integer','integer','date','integer[]']);
  const snapshot=p=>rpc('schedule_snapshot',[p]);
  const setTask=(p,id,key,v,done)=>rpc('set_schedule_task',[p,id,key,v,done],['text','uuid','text','integer','boolean']);
  const importCamps=(p,rows)=>rpc('import_schedule_camps',[p,JSON.stringify(rows)],['text','jsonb']);
  await identity('','anon');await assert.rejects(snapshot('ai'),e=>e.code==='42501');
  await identity(other);await assert.rejects(snapshot('ai'),/schedule_members_only/);
  await identity(member);
  const ai=await save('ai',1),crawler=await save('crawler',1);
  assert.notEqual(ai.id,crawler.id);assert.equal((await snapshot('ai')).camps.length,1);
  await assert.rejects(save('ai',1),e=>e.code==='23505');
  await assert.rejects(pg.query('update public.schedule_camps set term=9'),e=>e.code==='42501');
  const task=await setTask('ai',ai.id,'open|all|测试任务',0,true);
  await setTask('ai',ai.id,'open|all|另一任务',0,true);
  await assert.rejects(setTask('crawler',ai.id,'open|all|测试任务',0,true),/schedule_conflict/);
  await assert.rejects(setTask('ai',ai.id,'open|all|测试任务',0,false),/schedule_conflict/);
  const unchecked=await setTask('ai',ai.id,task.task_key,task.version,false);assert.equal(unchecked.version,2);
  await assert.rejects(setTask('ai',ai.id,task.task_key,task.version,true),/schedule_conflict/);
  assert.equal((await snapshot('ai')).tasks.filter(t=>t.completed).length,1);
  const renamed=await save('ai',2,ai.id,1);assert.equal(renamed.id,ai.id);
  await assert.rejects(save('ai',3,ai.id,1),/schedule_conflict/);
  assert.equal((await snapshot('ai')).tasks.length,2);
  const imported=await importCamps('ai',[
   {term:2,open_date:'2040-01-01',skips:[],completed_tasks:['open|all|旧任务']},
   {term:3,open_date:'2030-02-03',skips:[2],completed_tasks:['rest-2|all|测试导入任务']}
  ]);
  assert.deepEqual(imported.inserted,[3]);assert.deepEqual(imported.skipped,[2]);
  assert.equal((await snapshot('ai')).camps.find(c=>c.id===ai.id).open_date,'2030-01-06');
  await assert.rejects(importCamps('ai',[
   {term:4,open_date:'2030-02-03',skips:[],completed_tasks:[]},
   {term:5,open_date:'2030-02-03',skips:[],completed_tasks:['invalid-key']}
  ]));
  assert.equal((await snapshot('ai')).camps.some(c=>c.term===4),false,'import rolls back entirely');
  await rpc('delete_schedule_camp',['ai',ai.id,renamed.version],['text','uuid','integer']);
  assert.equal((await snapshot('ai')).camps.some(c=>c.id===ai.id),false);
  assert.equal((await snapshot('ai')).tasks.some(t=>t.camp_id===ai.id),false);
  await assert.rejects(setTask('ai',ai.id,'open|all|另一任务',1,false),/schedule_conflict/);
  assert.equal((await snapshot('crawler')).camps[0].id,crawler.id);
  await identity(other);assert.equal((await pg.query('select * from public.schedule_camps')).rows.length,0);
  await assert.rejects(importCamps('ai',[]),/schedule_members_only/);
  await pg.exec('reset role');assert.equal((await pg.query('select count(*)::int as n from public.schedule_tasks where camp_id=$1',[ai.id])).rows[0].n,2,'soft deletion preserves recovery data');
 }finally{await pg.close();}
});
