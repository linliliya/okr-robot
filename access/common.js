(function(root){
  'use strict';
  function safeNext(value, origin) {
    try { const u=new URL(value || '/',origin); return u.origin===origin && (!u.pathname.startsWith('/access/') || u.pathname==='/access/admin.html') ? u.pathname+u.search+u.hash : '/'; }
    catch(_){ return '/'; }
  }
  function permits(access, app){return !!(access && access.allowed && (app==='portal' || (app==='admin' ? access.role==='admin' : access.apps.includes(app))));}
  function message(error){
    const s=String(error?.message||error||'');
    if(/Invalid login/i.test(s))return '邮箱或密码不正确。';
    if(/banned|not.allowed|allowlist|saving new user|未获授权|名单/i.test(s))return '这个邮箱不在可用名单中，请联系管理员。';
    if(/already registered/i.test(s))return '这个邮箱已经注册，请直接登录。';
    if(/confirm/i.test(s))return '账号暂时无法登录，请联系管理员。';
    return '操作未完成，请检查网络后重试。';
  }
  const api={safeNext,permits,message};
  if(typeof module==='object'&&module.exports)module.exports=api;
  else root.AccessCommon=api;
})(typeof window==='object'?window:globalThis);
