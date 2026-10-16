// 和 skill-hub 用同一个 Supabase 项目（同一套账号），值与 skill-hub/config.js 相同
// AI_FUNCTION：Supabase 里 AI 中转函数的名字（见 SETUP.md）
// AI_REGION：选填。指定 Supabase 函数运行区域；留空使用平台默认区域。
window.SALES_SCRIPT_CONFIG = {
  SUPABASE_URL: "https://hxlsquvryfiixvjpuxfr.supabase.co",
  SUPABASE_ANON_KEY: "sb_publishable_eTMK5odEP3pubKfD76qx1w_hBq61jj7",
  AI_FUNCTION: "sales-ai",
  AI_REGION: "",
};
