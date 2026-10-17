(async function(){
 if(!await InternalAccess.ready)return;
 const client=InternalAccess.client,rows=document.getElementById('rows'),msg=document.getElementById('message');
 const esc=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 async function load(){
  const {data,error}=await client.rpc('list_staff_access');if(error)throw error;
  rows.innerHTML=data.map((u,i)=>`<tr><td>${esc(u.email)}</td><td>${u.role==='admin'?'管理员':`<select data-role="${i}"><option value="staff" ${u.role==='staff'?'selected':''}>通用工具和排期</option><option value="analyst" ${u.role==='analyst'?'selected':''}>全部系统</option></select>`}</td><td>${u.active?(u.registered?'已注册':'待注册'):'已停用'}</td><td>${u.role==='admin'?'管理员自身权限受保护':`<button data-save="${i}">保存权限</button> <button data-toggle="${i}">${u.active?'停用':'启用'}</button>`}</td></tr>`).join('');
  rows.querySelectorAll('button').forEach(b=>b.onclick=()=>run(async()=>{
   const i=Number(b.dataset.save??b.dataset.toggle),u=data[i],active=b.hasAttribute('data-toggle')?!u.active:u.active;
   if(!active && !confirm(`停用 ${u.email}？该账号将无法继续登录或访问数据。`))return;
   const role=rows.querySelector(`[data-role="${i}"]`).value;
   const {error}=await client.rpc('set_staff_access',{p_email:u.email,p_role:role,p_active:active});if(error)throw error;
   await load();msg.textContent='权限已保存。';
  }));
 }
 async function run(fn){msg.textContent='';document.querySelectorAll('button').forEach(b=>b.disabled=true);try{await fn();}catch(e){msg.textContent=e.message||'操作失败，请重试。';}finally{document.querySelectorAll('button').forEach(b=>b.disabled=false);}}
 document.getElementById('add').onsubmit=e=>{e.preventDefault();const f=e.target;run(async()=>{const {error}=await client.rpc('set_staff_access',{p_email:f.elements.email.value,p_role:f.elements.role.value,p_active:true});if(error)throw error;f.elements.email.value='';await load();msg.textContent='邮箱已添加，可以注册。';});};
 await run(load);
})();
