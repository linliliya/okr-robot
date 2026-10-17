(async function(){
 'use strict';
 const client=supabase.createClient(INTERNAL_ACCESS_CONFIG.url,INTERNAL_ACCESS_CONFIG.key,{auth:{detectSessionInUrl:false}});
 const next=AccessCommon.safeNext(new URLSearchParams(location.search).get('next'),location.origin);
 let signup=false;
 const form=document.getElementById('form'),message=document.getElementById('message'),button=document.getElementById('submit');
 document.getElementById('switch').onclick=()=>{
   signup=!signup;document.getElementById('signup-fields').hidden=!signup;
   form.elements.display_name.required=signup;form.elements.password.autocomplete=signup?'new-password':'current-password';
   document.getElementById('title').textContent=signup?'注册账号':'登录';button.textContent=signup?'注册并进入':'登录';
   document.getElementById('switch').textContent=signup?'已有账号？直接登录':'第一次使用？注册账号';message.textContent='';
 };
 async function enter(){const {data,error}=await client.rpc('my_app_access');if(error)throw error;if(!data?.allowed){await client.auth.signOut({scope:'local'});throw new Error('邮箱不在可用名单中');}location.replace(next);}
 form.onsubmit=async e=>{
  e.preventDefault();button.disabled=true;message.textContent='';
  try{
   const email=form.elements.email.value.trim().toLowerCase(),password=form.elements.password.value;
   const allowed=await client.rpc('can_register_email',{p_email:email});if(allowed.error)throw allowed.error;
   if(!allowed.data)throw new Error('邮箱不在可用名单中');
   const result=signup?await client.auth.signUp({email,password,options:{data:{display_name:form.elements.display_name.value.trim(),department:form.elements.department.value}}}):await client.auth.signInWithPassword({email,password});
   if(result.error)throw result.error;
   if(!result.data.session)throw new Error('already registered');
   await enter();
  }catch(error){message.textContent=AccessCommon.message(error);}finally{button.disabled=false;}
 };
 const {data}=await client.auth.getSession();if(data.session){try{await enter();}catch(error){message.textContent=AccessCommon.message(error);}}
})();
