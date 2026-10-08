// Supabase 项目配置：与 Skill 市集（skill-hub/config.js）共用同一个项目和账号体系
// Publishable key 设计上就是公开放在网页里的，数据安全靠 setup.sql 里的成员白名单和行级权限保证
// 两个值都留空（或打开页面时带 ?demo=1）会进入演示模式，数据只存在当前浏览器
window.CAMP_REVIEW_CONFIG = {
  SUPABASE_URL: "https://hxlsquvryfiixvjpuxfr.supabase.co",
  SUPABASE_ANON_KEY: "sb_publishable_eTMK5odEP3pubKfD76qx1w_hBq61jj7",
};
