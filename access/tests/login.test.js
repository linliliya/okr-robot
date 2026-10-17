const test=require('node:test');const assert=require('node:assert/strict');const fs=require('node:fs');const path=require('node:path');const {JSDOM,VirtualConsole}=require('jsdom');
const common=require('../common.js');const html=fs.readFileSync(path.join(__dirname,'../login.html'),'utf8'),source=fs.readFileSync(path.join(__dirname,'../login.js'),'utf8');
const tick=()=>new Promise(resolve=>setImmediate(resolve));
test('login form cannot submit credentials while scripts are unavailable',async()=>{
 const d=new JSDOM(html,{url:'https://internal.example/access/login.html',runScripts:'outside-only'});d.window.eval(source);
 assert.equal(d.window.document.getElementById('submit').disabled,true);assert.equal(d.window.document.getElementById('form').method,'post');
 assert.match(d.window.document.getElementById('message').textContent,/未加载/);d.window.close();
});
test('approved registration creates a session immediately; unlisted emails never reach Auth',async()=>{
 for(const allowed of [true,false]){
  const d=new JSDOM(html,{url:'https://internal.example/access/login.html',runScripts:'outside-only',virtualConsole:new VirtualConsole()});const w=d.window;const calls=[];
  w.AccessCommon=common;w.INTERNAL_ACCESS_CONFIG={url:'x',key:'y'};w.supabase={createClient:()=>({
   auth:{getSession:async()=>({data:{session:null}}),signUp:async input=>{calls.push('signup');assert.equal(input.email,'employee@example.test');return {data:{session:{user:{id:'1'}}}};},signOut:async()=>{}},
   rpc:async name=>{calls.push(name);return {data:name==='can_register_email'?allowed:{allowed:true,role:'staff',apps:['ai']}};}
  })};
  w.eval(source);await tick();w.document.getElementById('switch').click();
  const f=w.document.getElementById('form');f.elements.email.value='employee@example.test';f.elements.password.value='synthetic-test-password';f.elements.display_name.value='Test employee';
  f.dispatchEvent(new w.Event('submit',{cancelable:true}));await tick();await tick();
  assert.deepEqual(calls,allowed?['can_register_email','signup','my_app_access']:['can_register_email']);
  if(!allowed)assert.match(w.document.getElementById('message').textContent,/名单/);
  w.close();
 }
});
