(function(){
 'use strict';
 const app=document.currentScript.dataset.app;
 const C=window.INTERNAL_ACCESS_CONFIG, A=window.AccessCommon;
 const client=window.supabase && C ? window.supabase.createClient(C.url,C.key,{auth:{detectSessionInUrl:false}}) : null;
 const api=window.InternalAccess={client,access:null};
 const domReady=()=>document.readyState!=='loading'?Promise.resolve():new Promise(resolve=>document.addEventListener('DOMContentLoaded',resolve,{once:true}));
 const next=()=>location.pathname+location.search+location.hash;
 let generation=0, initialized=false;
 function login(){location.replace('/access/login.html?next='+encodeURIComponent(next()));}
 async function hide(message){
   await domReady();document.documentElement.dataset.access='blocked';
   let gate=document.getElementById('access-message');
   if(!gate){gate=document.createElement('section');gate.id='access-message';document.body.append(gate);}
   gate.innerHTML='<h1>暂时无法打开</h1><p></p><a href="/">返回工具首页</a> <button type="button">重试</button>';
   gate.querySelector('p').textContent=message;gate.querySelector('button').onclick=()=>location.reload();
 }
 async function check(){
  const own=++generation;
  try{
   if(!client || !A)throw new Error('Page resources unavailable');
   const {data,error}=await client.auth.getSession();if(error)throw error;
   if(own!==generation)return false;
   if(!data.session){login();return false;}
   const result=await client.rpc('my_app_access');if(result.error)throw result.error;
   if(own!==generation)return false;
   api.access=result.data;
   if(!A.permits(result.data,app)){
    if(!result.data?.allowed){await client.auth.signOut({scope:'local'});login();return false;}
    await hide('这个账号没有此系统的访问权限，请联系管理员。');return false;
   }
   await domReady();if(own!==generation)return false;
   document.getElementById('access-message')?.remove();
   document.querySelectorAll('[data-app-link]').forEach(a=>{a.hidden=!A.permits(result.data,a.dataset.appLink);});
   // Replace legacy membership editors with the single administrator-owned roster.
   document.querySelectorAll('a[href*="#/members"]').forEach(a=>{a.href='/access/admin.html';a.dataset.appLink='admin';a.textContent='账号权限';a.hidden=result.data.role!=='admin';});
   document.documentElement.dataset.access='ready';
   initialized=true;return true;
  }catch(_){if(own===generation)await hide('权限校验失败，请检查网络后重试。');return false;}
 }
 api.ready=check();
 client?.auth.onAuthStateChange(event=>{
   if(event==='SIGNED_OUT'){generation++;document.documentElement.dataset.access='pending';login();}
   else if(event==='USER_UPDATED' || (event==='SIGNED_IN' && initialized))setTimeout(check,0);
 });
 window.addEventListener('focus',()=>{if(initialized)check();});
 setInterval(()=>{if(initialized)check();},30000);
})();
