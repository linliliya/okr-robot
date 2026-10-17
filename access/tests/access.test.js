const test=require('node:test');const assert=require('node:assert/strict');const fs=require('node:fs');const path=require('node:path');
const {PGlite}=require('@electric-sql/pglite');const {JSDOM}=require('jsdom');const common=require('../common.js');
const read=f=>fs.readFileSync(path.join(__dirname,'..',f),'utf8');
test('safe return URLs and access decisions',()=>{
 assert.equal(common.safeNext('https://evil.example/x','https://internal.example'),'/');
 assert.equal(common.safeNext('//evil.example','https://internal.example'),'/');
 assert.equal(common.safeNext('/ai-camp-schedule/#1:2030-01-01','https://internal.example'),'/ai-camp-schedule/#1:2030-01-01');
 assert.equal(common.safeNext('/access/login.html','https://internal.example'),'/');
 assert.equal(common.safeNext('/access/admin.html','https://internal.example'),'/access/admin.html');
 assert.equal(common.permits({allowed:true,role:'staff',apps:['ai']},'ai'),true);
 assert.equal(common.permits({allowed:true,role:'analyst',apps:['ai']},'admin'),false);
 assert.equal(common.permits({allowed:false,role:'admin',apps:['ai']},'ai'),false);
});
test('page gate hides unauthorized systems and never unlocks on network failure',async()=>{
 for(const mode of ['staff','admin','denied','offline']){
  const d=new JSDOM('<html data-access="pending"><body><main>Private data</main><a data-app-link="admin">Admin</a></body></html>',{url:'https://internal.example/',runScripts:'outside-only'});const w=d.window;
  Object.defineProperty(w.document,'currentScript',{value:{dataset:{app:mode==='denied'?'sales-script':'portal'}}});
  w.INTERNAL_ACCESS_CONFIG={url:'x',key:'y'};w.AccessCommon=common;
  w.supabase={createClient:()=>({auth:{getSession:async()=>({data:{session:{user:{id:'x'}}}}),onAuthStateChange:()=>{},signOut:async()=>{}},rpc:async()=>mode==='offline'?{error:{message:'offline'}}:{data:{allowed:true,role:mode,apps:['ai']}}})};
  w.eval(read('gate.js'));const ok=await w.InternalAccess.ready;
  assert.equal(ok,['staff','admin'].includes(mode));assert.equal(w.document.documentElement.dataset.access,ok?'ready':'blocked');
  if(ok)assert.equal(w.document.querySelector('a').hidden,mode!=='admin');
  else assert.ok(w.document.getElementById('access-message'));
  w.close();
 }
});
test('database allowlist, role isolation, session revocation and email identity protections',async()=>{
 const pg=new PGlite();const ids=['00000000-0000-0000-0000-000000000001','00000000-0000-0000-0000-000000000002','00000000-0000-0000-0000-000000000003','00000000-0000-0000-0000-000000000004'];
 try{
 await pg.exec(`create role anon;create role authenticated;create schema auth;create schema storage;
 create table auth.users(id uuid primary key,email text,email_confirmed_at timestamptz,banned_until timestamptz,is_anonymous boolean default false);
 create table auth.refresh_tokens(user_id varchar,revoked boolean default false);create table auth.sessions(user_id uuid,not_after timestamptz);
 create function auth.uid() returns uuid language sql stable as $$select nullif(current_setting('request.jwt.claim.sub',true),'')::uuid$$;
 grant usage on schema auth,public to anon,authenticated;
 create table public.profiles(id uuid);create table public.skills(id bigint,downloads int,source_type text);create table public.versions(id bigint);
 create table public.camp_review_members(email text);create table public.sales_members(email text);
 create table public.camp_cohorts(id int);create table public.sales_scripts(id int);
 create table storage.objects(bucket_id text);
 create function public.is_camp_review_member() returns boolean language sql as $$select false$$;
 create function public.is_sales_member() returns boolean language sql as $$select false$$;
 insert into auth.users(id,email,email_confirmed_at) values('${ids[0]}','owner@example.test',now()),('${ids[1]}','analyst@example.test',now()),('${ids[2]}','staff@example.test',now()),('${ids[3]}','old@example.test',now());
 insert into auth.sessions(user_id)values('${ids[2]}');insert into auth.refresh_tokens(user_id)values('${ids[2]}');
 grant select,insert,update,delete on all tables in schema public to authenticated;
 alter table public.skills enable row level security;create policy existing_open on public.skills for all to authenticated using(true)with check(true);
 alter table public.camp_cohorts enable row level security;create policy existing_member on public.camp_cohorts for all to authenticated using(public.is_camp_review_member()) with check(public.is_camp_review_member());
 alter table public.sales_scripts enable row level security;create policy existing_member on public.sales_scripts for all to authenticated using(public.is_sales_member()) with check(public.is_sales_member());
 insert into public.skills values(1,0,'upload');insert into public.camp_cohorts values(1);insert into public.sales_scripts values(1);`);
 await pg.exec(fs.readFileSync(path.join(__dirname,'../../camp-schedule/setup.sql'),'utf8'));
 await pg.exec(read('setup.sql'));await pg.exec(read('setup.sql'));
 await pg.exec("insert into app_private.staff_access(email,role)values('owner@example.test','admin'),('analyst@example.test','analyst'),('staff@example.test','staff'),('new@example.test','staff')");
 const as=async(id,role='authenticated')=>{await pg.exec('reset role');await pg.query("select set_config('request.jwt.claim.sub',$1,false)",[id]);await pg.exec('set role '+role);};
 const rpc=async(n,args=[],casts=[])=> (await pg.query(`select public.${n}(${args.map((_,i)=>'$'+(i+1)+(casts[i]?'::'+casts[i]:'')).join(',')}) as value`,args)).rows[0].value;
 await as('','anon');await assert.rejects(rpc('my_app_access'),e=>e.code==='42501');assert.equal(await rpc('can_register_email',[' NEW@EXAMPLE.TEST ']),true);assert.equal(await rpc('can_register_email',['stranger@example.test']),false);
 await as(ids[3]);assert.equal((await rpc('my_app_access')).allowed,false);assert.equal((await pg.query('select * from skills')).rows.length,0);await assert.rejects(rpc('increment_downloads',[1]),e=>e.code==='42501');
 await as(ids[2]);assert.equal((await rpc('my_app_access')).role,'staff');assert.equal((await pg.query('select * from skills')).rows.length,1);
 assert.equal((await pg.query('select * from camp_cohorts')).rows.length,0);assert.equal((await pg.query('select * from sales_scripts')).rows.length,0);
 assert.deepEqual((await rpc('schedule_snapshot',['ai'])).camps,[]);await assert.rejects(rpc('list_staff_access'),e=>e.code==='42501');await assert.rejects(rpc('set_staff_access',['new@example.test','analyst',true]),e=>e.code==='42501');
 await as(ids[1]);assert.equal((await pg.query('select * from camp_cohorts')).rows.length,1);assert.equal((await pg.query('select * from sales_scripts')).rows.length,1);
 await assert.rejects(pg.query("insert into camp_review_members values('stranger@example.test')"),e=>e.code==='42501');await assert.rejects(rpc('add_sales_member',['stranger@example.test']),e=>e.code==='42501');
 await assert.rejects(rpc('set_staff_access',['new@example.test','analyst',true]),e=>e.code==='42501');
 await as(ids[0]);assert.equal((await rpc('list_staff_access')).length,4);await assert.rejects(rpc('set_staff_access',['owner@example.test','staff',false]),e=>e.code==='42501');
 await rpc('set_staff_access',['staff@example.test','staff',false]);
 await as(ids[2]);assert.equal((await rpc('my_app_access')).allowed,false);await assert.rejects(rpc('schedule_snapshot',['ai']),/schedule_members_only/);
 await pg.exec('reset role');assert.equal((await pg.query('select revoked from auth.refresh_tokens')).rows[0].revoked,true);assert.ok((await pg.query('select not_after from auth.sessions')).rows[0].not_after);
 await assert.rejects(pg.exec("insert into auth.users(id,email) values(gen_random_uuid(),'stranger@example.test')"),/名单/);
 await assert.rejects(pg.exec("insert into auth.users(id,email) values(gen_random_uuid(),'staff@example.test')"),/名单/);
 await pg.exec("insert into auth.users(id,email,email_confirmed_at)values(gen_random_uuid(),'new@example.test',now())");
 await assert.rejects(pg.exec("update auth.users set email='owner@example.test' where email='new@example.test'"),/不能自行更改/);
 }finally{await pg.close();}
});
