const CFG = window.SKILLHUB_CONFIG || {};
const PREVIEW = !!window.SKILLHUB_PREVIEW_CLIENT;
const CONFIGURED = PREVIEW || (CFG.SUPABASE_URL && !CFG.SUPABASE_URL.includes("YOUR-PROJECT"));
const sb = window.SKILLHUB_PREVIEW_CLIENT || window.InternalAccess?.client || (CONFIGURED ? supabase.createClient(CFG.SUPABASE_URL, CFG.SUPABASE_ANON_KEY) : null);
const H = window.SkillHub;

const CATEGORIES = ["文案生成", "数据处理", "课程制作", "招聘筛选", "资讯采集", "效率工具", "其他"];
const DEPARTMENTS = ["内容", "运营", "交付", "产品", "其他"];

const $app = document.getElementById("app");
let me = null;          // 当前登录用户 { id, email, profile }
let listState = { q: "", category: "", role: "", source: "all", sort: "new" };
let skillCache = [];
const readmeRequests = new Map();

// ---------- 小工具 ----------
const esc = s => String(s ?? "").replace(/[&<>"']/g,
  c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

function flash(msg, ms = 2600) {
  const el = document.getElementById("flash");
  el.textContent = msg; el.style.display = "block";
  clearTimeout(el._t); el._t = setTimeout(() => el.style.display = "none", ms);
}

function fmtDate(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}

// 从 frontmatter 里取字段值；支持多行写法（description: > 或 | 后接缩进续行）
function fmValue(fm, key) {
  const lines = fm.split(/\r?\n/);
  const idx = lines.findIndex(l => l.startsWith(key + ":"));
  if (idx === -1) return "";
  let first = lines[idx].slice(key.length + 1).trim();
  if (/^[>|][+-]?$/.test(first)) first = "";
  const cont = [];
  for (let i = idx + 1; i < lines.length; i++) {
    if (/^\s+\S/.test(lines[i])) cont.push(lines[i].trim());
    else break;
  }
  const parts = (first ? [first] : []).concat(cont);
  return parts.join(" ").replace(/^['"]|['"]$/g, "").trim();
}

function slugify(name) {
  return String(name).toLowerCase().trim()
    .replace(/[^a-z0-9一-鿿-]+/g, "-").replace(/^-+|-+$/g, "") || "skill";
}

function newSlug(title) {
  return slugify(title).slice(0, 70) + "-" + crypto.randomUUID().slice(0, 8);
}

function errorMessage(error) {
  if (["42703", "PGRST204"].includes(error?.code)) return "这项功能还未启用，请联系管理员完成市集更新。";
  return error?.message || "暂时没能完成，请稍后重试。";
}

function showError(error) {
  cancelReadmeReads();
  $app.innerHTML = `<a class="backlink" href="#/">← 返回市集</a><div class="empty" role="alert">${esc(errorMessage(error))}</div>`;
}

function cancelReadmeReads() {
  for (const cancel of readmeRequests.values()) cancel();
  readmeRequests.clear();
}

function readmeHTML(skill, open = false) {
  return `<details class="readme-section" data-readme-slug="${esc(skill.slug)}" ${open ? "open" : ""}>
    <summary><span>README 项目说明</span><span class="readme-expand">查看 README ↓</span><span class="readme-collapse">收起 ↑</span></summary>
    <div class="readme-content" aria-live="polite"><p class="readme-status">展开后读取原项目的说明。</p></div>
  </details>`;
}

function readmeErrorMessage(error) {
  const messages = {
    UNSUPPORTED_SOURCE: "这个来源暂时无法直接读取 README，可以前往原项目查看。",
    NOT_FOUND: "暂时没有找到可读取的项目说明，可以前往原项目查看。",
    NO_README: "这个项目暂时没有 README 或 SKILL.md，可以前往原项目查看。",
    RATE_LIMIT: "GitHub 暂时限制了读取请求，请稍后重试，或直接查看原项目。",
    TIMEOUT: "读取项目说明超时，请重试，或直接查看原项目。",
    NETWORK_ERROR: "暂时连不上 GitHub，请重试，或直接查看原项目。",
    FILE_TOO_LARGE: "这份说明比较长，请前往原项目查看完整内容。",
    RESPONSE_TOO_LARGE: "项目返回的内容过大，请前往原项目查看。"
  };
  return messages[error?.code] || "暂时无法读取项目说明，请重试，或直接查看原项目。";
}

function bindReadmes(skills) {
  const bySlug = new Map(skills.map(skill => [skill.slug, skill]));
  $app.querySelectorAll("[data-readme-slug]").forEach(panel => {
    const skill = bySlug.get(panel.dataset.readmeSlug);
    if (!skill || H.sourceType(skill) !== "external") return;
    const content = panel.querySelector(".readme-content");
    let loaded = false, controller = null, generation = 0;
    const stop = () => { generation++; controller?.abort(); controller = null; readmeRequests.delete(panel); };
    const active = token => generation === token && panel.open && $app.contains(panel);
    async function load() {
      if (loaded || controller || !panel.open) return;
      controller = new AbortController();
      const token = ++generation;
      readmeRequests.set(panel, stop);
      content.innerHTML = `<p class="readme-status" role="status">正在读取项目说明…</p>`;
      try {
        if (!window.SkillHubImporter?.readReadme || !window.SkillHubReadme) throw new Error("README reader unavailable");
        const result = await window.SkillHubImporter.readReadme(skill, { signal: controller.signal });
        if (!active(token)) return;
        const rendered = window.SkillHubReadme.render(result, { marked: window.marked, DOMPurify: window.DOMPurify, document });
        const original = H.safeUrl(result.url);
        const note = result.kind === "skill" ? "此项目未提供 README，以下展示 SKILL.md。"
          : result.scope === "repository" ? "此 Skill 目录未提供 README，以下为整个仓库的说明。" : "";
        content.innerHTML = `<div class="readme-meta"><span>${esc(result.path)}</span>${original ? `<a href="${esc(original)}" target="_blank" rel="noopener noreferrer">查看原文 ↗</a>` : ""}</div>${note ? `<p class="readme-note">${esc(note)}</p>` : ""}<div class="doc readme-document" tabindex="0" role="region" aria-label="项目说明正文">${rendered}</div>`;
        loaded = true;
      } catch (error) {
        if (!active(token) || error?.name === "AbortError") return;
        const source = H.safeUrl(skill.source_url);
        content.innerHTML = `<p class="readme-status" role="status">${esc(readmeErrorMessage(error))}</p><div class="readme-actions"><button class="btn small ghost" type="button" data-readme-retry>重新读取</button>${source ? `<a href="${esc(source)}" target="_blank" rel="noopener noreferrer">查看原项目 ↗</a>` : ""}</div>`;
        content.querySelector("[data-readme-retry]").onclick = () => load();
      } finally {
        if (generation === token) { controller = null; readmeRequests.delete(panel); }
      }
    }
    panel.ontoggle = () => { if (panel.open) return load(); if (!loaded) stop(); };
    if (panel.open) load();
  });
}

async function copyText(text, message = "已复制") {
  try {
    await navigator.clipboard.writeText(text);
    flash(message);
  } catch {
    flash("无法自动复制，请选中文字后手动复制。", 5000);
  }
}

function displayTitleValue(skill) {
  return skill?.display_title || H.suggestDisplayTitle(skill || {});
}

function validateDisplayTitle(value) {
  const title = String(value || "").trim();
  if (!title) throw new Error("请写一个功能标题，让同事知道这个 Skill 能做什么。");
  if ([...title].length > 80 || /[\u0000-\u001f\u007f-\u009f\u2028\u2029]/.test(title)) {
    throw new Error("功能标题请控制在 80 个字以内，并写在一行。");
  }
  return title;
}

function skillHeadingHTML(skill) {
  const canEdit = me && (H.sourceType(skill) === "upload" || me.id === skill.author_id);
  return `<div class="skill-heading-row"><div><h1>${esc(H.getDisplayTitle(skill))}</h1><p class="skill-original-name"><span>Skill 名称</span> <code>${esc(skill.title)}</code></p></div>${canEdit ? `<button class="text-button" id="edit-skill-title">修改功能标题</button>` : ""}</div>`;
}

function bindTitleEditor(skill) {
  const edit = document.getElementById("edit-skill-title");
  if (!edit || !me || (H.sourceType(skill) === "external" && me.id !== skill.author_id)) return;
  edit.onclick = () => {
    const dialog = document.createElement("dialog");
    dialog.className = "confirm-dialog title-dialog";
    dialog.setAttribute("aria-labelledby", "title-edit-heading");
    dialog.innerHTML = `<h2 id="title-edit-heading">修改功能标题</h2><p class="hint">标题最多 80 字，原始 Skill 名称会保留在副标题中。</p><form id="title-edit-form"><div class="field"><label for="edit-display-title">功能标题</label><input id="edit-display-title" name="display_title" required value="${esc(displayTitleValue(skill))}" placeholder="例如 去AI味写作"></div><p class="skill-original-name"><span>Skill 名称</span> <code>${esc(skill.title)}</code></p><div id="title-edit-error" class="form-error" role="alert"></div><div class="actions"><button type="button" class="btn ghost" id="cancel-title-edit">取消</button><button class="btn" id="save-title-edit">保存标题</button></div></form>`;
    document.body.appendChild(dialog);
    let saving = false;
    const close = () => { dialog.close(); dialog.remove(); };
    document.getElementById("cancel-title-edit").onclick = () => { if (!saving) close(); };
    dialog.oncancel = event => { event.preventDefault(); if (!saving) close(); };
    document.getElementById("title-edit-form").onsubmit = async event => {
      event.preventDefault();
      if (saving) return;
      saving = true;
      const submit = document.getElementById("save-title-edit"), cancel = document.getElementById("cancel-title-edit"), errorBox = document.getElementById("title-edit-error");
      submit.disabled = true; cancel.disabled = true; submit.textContent = "正在保存…"; errorBox.textContent = "";
      try {
        const display_title = validateDisplayTitle(document.getElementById("edit-display-title").value);
        const { data, error } = await sb.from("skills").update({ display_title, updated_at: new Date().toISOString() }).eq("id", skill.id).select("id").single();
        if (error) throw error;
        if (!data) throw new Error("没有保存成功，请确认你仍有编辑权限。");
        close(); flash("功能标题已更新");
        if (location.hash === "#/skill/" + encodeURIComponent(skill.slug)) await viewDetail(skill.slug).catch(showError);
      } catch (error) { errorBox.textContent = errorMessage(error); }
      finally { saving = false; submit.disabled = false; cancel.disabled = false; submit.textContent = "保存标题"; }
    };
    dialog.showModal();
    document.getElementById("edit-display-title").focus();
  };
}

function roleTagsHTML(roles) {
  const selected = H.normalizeRoles(roles);
  return selected.length ? selected.map(role => `<span class="role-tag">${esc(role === "通用" ? "所有岗位通用" : role)}</span>`).join("")
    : `<span class="role-unset">暂未标注岗位</span>`;
}

function roleFieldsHTML(roles, id = "skill-role-fields") {
  const selected = H.normalizeRoles(roles);
  return `<fieldset class="role-field" id="${id}"><legend>适用岗位 <span>选填，可多选</span></legend>
    <div class="role-choices">${H.ROLE_OPTIONS.map(role => `<label class="role-choice"><input type="checkbox" name="roles" value="${esc(role)}" ${selected.includes(role) ? "checked" : ""}><span>${esc(role)}</span></label>`).join("")}</div>
    <p class="hint">适合所有岗位可以选「通用」，暂不确定可以留空。</p></fieldset>`;
}

function bindRoleFields(id = "skill-role-fields") {
  const group = document.getElementById(id);
  group.onchange = event => {
    const changed = event.target;
    if (changed.name !== "roles" || !changed.checked) return;
    group.querySelectorAll('input[name="roles"]').forEach(input => {
      if (input !== changed && (changed.value === "通用" || input.value === "通用")) input.checked = false;
    });
  };
}

function detailRolesHTML(skill) {
  const canEdit = me && (H.sourceType(skill) === "upload" || me.id === skill.author_id);
  return `<div class="detail-roles"><span class="role-caption">适用岗位</span>${roleTagsHTML(skill.roles)}${canEdit ? `<button class="text-button" id="edit-skill-roles">${H.normalizeRoles(skill.roles).length ? "修改岗位" : "补充岗位"}</button>` : ""}</div>`;
}

function bindRoleEditor(skill) {
  const edit = document.getElementById("edit-skill-roles");
  if (!edit) return;
  edit.onclick = () => {
    const dialog = document.createElement("dialog");
    dialog.className = "confirm-dialog roles-dialog";
    dialog.setAttribute("aria-labelledby", "roles-title");
    dialog.innerHTML = `<h2 id="roles-title">为 ${esc(H.getDisplayTitle(skill))} 设置岗位</h2><form id="role-edit-form">${roleFieldsHTML(skill.roles, "edit-role-fields")}<div id="role-edit-error" class="form-error" role="alert"></div><div class="actions"><button type="button" class="btn ghost" id="cancel-role-edit">取消</button><button class="btn" id="save-role-edit">保存岗位</button></div></form>`;
    document.body.appendChild(dialog);
    bindRoleFields("edit-role-fields");
    let saving = false;
    const close = () => { dialog.close(); dialog.remove(); };
    document.getElementById("cancel-role-edit").onclick = () => { if (!saving) close(); };
    dialog.oncancel = event => { event.preventDefault(); if (!saving) close(); };
    document.getElementById("role-edit-form").onsubmit = async event => {
      event.preventDefault();
      if (saving) return;
      saving = true;
      const submit = document.getElementById("save-role-edit"), cancel = document.getElementById("cancel-role-edit"), errorBox = document.getElementById("role-edit-error");
      submit.disabled = true; cancel.disabled = true; submit.textContent = "正在保存…"; errorBox.textContent = "";
      try {
        const roles = H.normalizeRoles(new FormData(event.target).getAll("roles"));
        const { data, error } = await sb.from("skills").update({ roles, updated_at: new Date().toISOString() }).eq("id", skill.id).select("id").single();
        if (error) throw error;
        if (!data) throw new Error("没有保存成功，请确认你仍有编辑权限。");
        close(); flash("适用岗位已更新");
        if (location.hash === "#/skill/" + encodeURIComponent(skill.slug)) await viewDetail(skill.slug).catch(showError);
      } catch (error) { errorBox.textContent = errorMessage(error); }
      finally { saving = false; submit.disabled = false; cancel.disabled = false; submit.textContent = "保存岗位"; }
    };
    dialog.showModal();
  };
}

function confirmAction(message, action = "确认") {
  return new Promise(resolve => {
    const dialog = document.createElement("dialog");
    dialog.className = "confirm-dialog";
    dialog.setAttribute("aria-labelledby", "confirm-title");
    dialog.innerHTML = `<h2 id="confirm-title">确认操作</h2><p>${esc(message)}</p><div class="actions"><button class="btn ghost" data-cancel autofocus>取消</button><button class="btn" data-confirm>${esc(action)}</button></div>`;
    const finish = result => { dialog.close(); dialog.remove(); resolve(result); };
    dialog.querySelector("[data-cancel]").onclick = () => finish(false);
    dialog.querySelector("[data-confirm]").onclick = () => finish(true);
    dialog.oncancel = event => { event.preventDefault(); finish(false); };
    document.body.appendChild(dialog);
    dialog.showModal();
  });
}

// ---------- 认证 ----------
async function refreshUser() {
  if (!sb) return;
  const { data: { session } } = await sb.auth.getSession();
  if (!session) { me = null; renderNav(); return; }
  const { data: profile } = await sb.from("profiles")
    .select("*").eq("id", session.user.id).maybeSingle();
  me = { id: session.user.id, email: session.user.email, profile };
  renderNav();
}

function renderNav() {
  document.getElementById("nav-user").textContent =
    me ? `${me.profile?.display_name || me.email} · ${me.profile?.department || ""}` : "";
  document.getElementById("nav-upload").style.display = me ? "" : "none";
  document.getElementById("nav-logout").style.display = me ? "" : "none";
  if (PREVIEW) document.getElementById("nav-logout").style.display = "none";
}

document.getElementById("nav-logout").onclick = async () => {
  await sb.auth.signOut(); me = null; renderNav(); location.hash = "#/login";
};

// ---------- 路由 ----------
async function route() {
  if(window.InternalAccess && !await InternalAccess.ready) return;
  cancelReadmeReads();
  try {
  if (!CONFIGURED) return viewSetupNote();
  await refreshUser();
  const hash = location.hash || "#/";
  if (!me && hash !== "#/register") return viewAuth(hash === "#/login" ? "login" : "login");
  if (hash === "#/register") return viewAuth("register");
  if (hash === "#/login") return viewAuth("login");
  if (hash === "#/share") return viewShare();
  if (hash === "#/recommend") return await viewRecommend();
  const edit = hash.match(/^#\/recommend\/(.+)$/);
  if (edit) return await viewRecommend(decodeURIComponent(edit[1]));
  if (hash === "#/upload") return await viewUpload();
  const upload = hash.match(/^#\/upload\/(.+)$/);
  if (upload) return await viewUpload(decodeURIComponent(upload[1]));
  const m = hash.match(/^#\/skill\/(.+)$/);
  if (m) return await viewDetail(decodeURIComponent(m[1]));
  return await viewList();
  } catch (error) { showError(error); }
}
window.addEventListener("hashchange", route);

// ---------- 视图：未配置提示 ----------
function viewSetupNote() {
  $app.innerHTML = `
    <div class="setup-note">
      <h2>⚙️ 平台还没连接数据库</h2>
      <p>这个页面已经部署好了，但还差最后一步：按照仓库里 <code>skill-hub/SETUP.md</code>
      的指引注册一个免费的 Supabase 项目（约 5 分钟），然后把项目地址和密钥填进
      <code>skill-hub/config.js</code>，重新发布即可。</p>
    </div>`;
}

// ---------- 视图：登录 / 注册 ----------
function viewAuth(mode) {
  const isLogin = mode === "login";
  $app.innerHTML = `
    <div class="panel">
      <h2>${isLogin ? "登录" : "注册账号"}</h2>
      <form id="auth-form">
        ${isLogin ? "" : `
        <div class="field"><label>姓名（同事们看到的名字，填大家熟悉的称呼就行）</label>
          <input name="display_name" required maxlength="20" placeholder="例如：小王 / Lily"></div>
        <div class="field"><label>部门</label>
          <select name="department">${DEPARTMENTS.map(d => `<option>${d}</option>`).join("")}</select></div>`}
        <div class="field"><label>邮箱</label>
          <input name="email" type="email" required placeholder="you@example.com"></div>
        <div class="field"><label>密码${isLogin ? "" : "（至少 6 位）"}</label>
          <input name="password" type="password" required minlength="6"></div>
        <button class="btn" style="width:100%">${isLogin ? "登录" : "注册并进入"}</button>
      </form>
      <div class="authswitch">
        ${isLogin
          ? `第一次来？<a href="#/register">注册一个账号</a>`
          : `已有账号？<a href="#/login">直接登录</a>`}
      </div>
    </div>`;
  document.getElementById("auth-form").onsubmit = async e => {
    e.preventDefault();
    const fd = new FormData(e.target);
    const email = fd.get("email"), password = fd.get("password");
    try {
      if (isLogin) {
        const { error } = await sb.auth.signInWithPassword({ email, password });
        if (error) throw error;
      } else {
        const { error } = await sb.auth.signUp({
          email, password,
          options: { data: {
            display_name: fd.get("display_name"),
            department: fd.get("department"),
          } },
        });
        if (error) throw error;
      }
      location.hash = "#/"; route();
    } catch (err) {
      const msg = String(err.message || err);
      if (msg.includes("Invalid login")) flash("邮箱或密码不对");
      else if (msg.includes("already registered")) flash("这个邮箱已经注册过了，直接登录吧");
      else if (msg.includes("confirm")) flash("账号暂时无法登录，请联系管理员", 5000);
      else flash("出错了：" + msg, 4000);
    }
  };
}

// ---------- 视图：列表 ----------
async function viewList() {
  $app.innerHTML = `<div class="empty">加载中…</div>`;
  const { data, error } = await sb.from("skills")
    .select("*, author:profiles(display_name, department)")
    .order("updated_at", { ascending: false });
  if (error) { $app.innerHTML = `<div class="empty">加载失败：${esc(error.message)}</div>`; return; }
  skillCache = data || [];
  renderList();
}

function renderList() {
  cancelReadmeReads();
  const { q, category, role, source, sort } = listState;
  const items = H.filterSkills(skillCache, listState);
  const sources = [["all", "全部来源"], ["upload", "内部上传"], ["external", "外部推荐"]];
  $app.innerHTML = `
    <section class="hero">
      <div><div class="eyebrow">TEAM SKILL HUB</div><h1>把好用的 Skill 分享给同事</h1>
      <p>团队做的、网上发现的，都在这里。找到适合手头工作的那一个。</p></div>
    </section>
    <div class="source-tabs" aria-label="按来源筛选">
      ${sources.map(([key, label]) => `<button class="source-tab ${source === key ? "active" : ""}" data-source="${key}" aria-pressed="${source === key}">${label}<span>${key === "all" ? skillCache.length : skillCache.filter(s => H.sourceType(s) === key).length}</span></button>`).join("")}
    </div>
    <div class="toolbar">
      <input type="search" id="search" aria-label="搜索 Skill" placeholder="搜索功能、Skill 名称或简介" value="${esc(q)}">
      <button class="chip ${sort === "new" ? "active" : ""}" data-sort="new" aria-pressed="${sort === "new"}">最近更新</button>
      ${source === "upload" ? `<button class="chip ${sort === "hot" ? "active" : ""}" data-sort="hot" aria-pressed="${sort === "hot"}">下载最多</button>` : ""}
    </div>
    <div class="role-filter" aria-label="按岗位筛选">
      <span class="filter-caption">适用岗位</span><div class="chips">
        <button class="chip ${!role ? "active" : ""}" data-role="" aria-pressed="${!role}">全部岗位</button>
        ${H.ROLE_OPTIONS.map(r => `<button class="chip ${role === r ? "active" : ""}" data-role="${r}" aria-pressed="${role === r}">${r}</button>`).join("")}
      </div>
    </div>
    <div class="chips" aria-label="按用途筛选">
      <button class="chip ${!category ? "active" : ""}" data-cat="" aria-pressed="${!category}">全部用途</button>
      ${CATEGORIES.map(c => `<button class="chip ${category === c ? "active" : ""}" data-cat="${c}" aria-pressed="${category === c}">${c}</button>`).join("")}
    </div>
    <div class="list-status" aria-live="polite">${items.length} 个 Skill${role ? ` · ${esc(role === "通用" ? "所有岗位通用" : role + "（含通用）")}` : ""}${q ? ` · 搜索「${esc(q)}」` : ""}</div>
    ${items.length ? `<div class="grid">${items.map(cardHTML).join("")}</div>` : `<div class="empty">${skillCache.length ? "没有找到匹配的 Skill，试试其他岗位、关键词或用途。" : "还没有人分享，来分享第一个好用的 Skill。"}<br><a href="#/share">分享一个 Skill →</a></div>`}
    <div class="discover"><span>去外面找找好 Skill</span><a href="https://skills.sh" target="_blank" rel="noopener noreferrer">浏览 skills.sh ↗</a><span>发现适合团队的，再来这里推荐给同事。</span></div>`;

  document.getElementById("search").oninput = e => {
    listState.q = e.target.value; renderList();
    const inp = document.getElementById("search");
    inp.focus();
  };
  $app.querySelectorAll("[data-source]").forEach(b => b.onclick = () => {
    listState.source = b.dataset.source; listState.sort = "new"; renderList();
  });
  $app.querySelectorAll("[data-sort]").forEach(b => b.onclick = () => { listState.sort = b.dataset.sort; renderList(); });
  $app.querySelectorAll("[data-cat]").forEach(b => b.onclick = () => { listState.category = b.dataset.cat; renderList(); });
  $app.querySelectorAll("[data-role]").forEach(b => b.onclick = () => { listState.role = b.dataset.role; renderList(); });
  bindReadmes(items);
}

function cardHTML(s) {
  const external = H.sourceType(s) === "external";
  const tags = s.tags ? s.tags.split(/[,，\s]+/).filter(Boolean) : [];
  return `<article class="skill-card"><a class="card" href="#/skill/${encodeURIComponent(s.slug)}">
    <div class="card-top"><span class="source-badge ${external ? "external" : ""}">${external ? "外部推荐" : "内部上传"}</span>
      <span class="trial-status">${external ? s.tested ? "分享者已试用" : "尚未试用" : `${s.downloads || 0} 次下载`}</span></div>
    <div class="card-heading"><h3>${esc(H.getDisplayTitle(s))}</h3><div class="skill-original-name">${esc(s.title)}</div></div>
    <div class="desc">${esc(s.description) || "查看介绍和使用方法"}</div>
    <div class="card-roles">${roleTagsHTML(s.roles)}</div>
    <div class="meta"><span class="tag">${esc(s.category)}</span>${tags.slice(0, 2).map(t => `<span>${esc(t)}</span>`).join("")}</div>
    <div class="card-bottom"><div class="meta"><span>${esc(s.author?.display_name || "同事")} ${external ? "推荐" : "分享"}</span><span>${fmtDate(s.updated_at)}</span></div><span class="card-action">${external ? "了解与安装" : "查看详情"} →</span></div>
  </a>${external ? readmeHTML(s) : ""}</article>`;
}

function viewShare() {
  $app.innerHTML = `<a class="backlink" href="#/">← 返回市集</a>
    <div class="share-intro"><h1>分享一个好用的 Skill</h1><p>可以上传你整理的文件，也可以把网上发现的好资源推荐给同事。</p></div>
    <div class="share-options">
      <a class="share-option" href="#/upload"><span class="option-symbol" aria-hidden="true">↑</span><h2>上传 Skill 文件</h2><p>上传完整 ZIP 压缩包，自动读取名称和介绍，再用功能标题告诉同事它能做什么。</p><span class="btn ghost">选择文件 →</span></a>
      <a class="share-option" href="#/recommend"><span class="option-symbol" aria-hidden="true">↗</span><h2>分享网上的 Skill</h2><p>粘贴链接自动读取介绍，确认标题和简介后即可分享。同事可以把安装请求发给 AI 工具。</p><span class="btn ghost">粘贴链接 →</span></a>
    </div>`;
}

// ---------- 视图：详情 ----------
async function viewDetail(slug) {
  cancelReadmeReads();
  $app.innerHTML = `<div class="empty">加载中…</div>`;
  const { data: skill, error } = await sb.from("skills")
    .select("*, author:profiles(display_name, department)").eq("slug", slug).maybeSingle();
  if (error) throw error;
  if (!skill) { showError(new Error("没有找到这个 Skill，可能已被分享者移除。")); return; }
  if (H.sourceType(skill) === "external") return viewExternalDetail(skill);
  const { data: versions, error: versionError } = await sb.from("versions")
    .select("*, uploader:profiles(display_name)")
    .eq("skill_id", skill.id).order("version", { ascending: false });
  if (versionError) throw versionError;
  const latest = versions?.[0];
  const docBody = (latest?.skill_md || "").replace(/^---\s*\n[\s\S]*?\n---\s*\n?/, "");
  const docHTML = docBody ? DOMPurify.sanitize(marked.parse(docBody)) : "";
  const tags = skill.tags ? skill.tags.split(/[,，\s]+/).filter(Boolean) : [];

  $app.innerHTML = `
    <a class="backlink" href="#/">← 返回列表</a>
    <div class="detail-head">
      ${skillHeadingHTML(skill)}
      <div class="meta">
        <span class="source-badge">内部上传</span>
        <span class="tag">${esc(skill.category)}</span>
        ${tags.map(t => `<span class="tag">${esc(t)}</span>`).join("")}
        <span>👤 ${esc(skill.author?.display_name || "?")} · ${esc(skill.author?.department || "")}</span>
        <span>⬇️ 已下载 ${skill.downloads} 次</span>
        <span>更新于 ${fmtDate(skill.updated_at)}</span>
      </div>
      <div class="desc">${esc(skill.description)}</div>
      ${detailRolesHTML(skill)}
      <div class="actions">
        ${latest ? `<button class="btn" id="dl-btn">⬇️ 下载最新版 v${latest.version}</button>` : ""}
        <a class="btn ghost" href="#/upload/${encodeURIComponent(skill.slug)}">上传新版本</a>
      </div>
      <div class="install-guides">
        <div class="ig-title">📦 下载后怎么用？按你用的工具选一个：</div>
        <details open>
          <summary>Claude Code（终端 / 桌面版）</summary>
          <div class="g-body">
            解压 zip，把<b>整个文件夹</b>放进 <code>~/.claude/skills/</code> 目录
            （访达按 <code>Cmd+Shift+G</code> 输入这个路径即可打开；Windows 在
            <code>C:\\Users\\你的用户名\\.claude\\skills\\</code>）。<br>
            放好后新开一个对话，说出 skill 描述里的触发词就会自动生效，
            也可以直接说「用 xx skill 帮我…」。
          </div>
        </details>
        <details>
          <summary>Claude 网页版 / App（claude.ai）</summary>
          <div class="g-body">
            不用解压。打开 claude.ai → 设置（Settings）→ 功能（Capabilities）→
            <b>Skills</b> → 上传这个 zip 文件即可（需要先把设置里的
            「Code execution / 代码执行」开关打开，skill 依赖它运行）。<br>
            上传后可在列表里开关这个 skill，对话中提到相关需求会自动触发。
            skill 只对自己的账号生效，同事需要各自上传。
          </div>
        </details>
        <details>
          <summary>Codex（OpenAI 编程助手）</summary>
          <div class="g-body">
            Codex 原生支持 skill，格式和 Claude 通用，同一个 zip 直接用：<br>
            解压后把<b>整个文件夹</b>放进 <code>~/.agents/skills/</code> 目录
            （没有就新建一个）。对话里输入 <code>$技能名</code> 显式调用，
            输入 <code>/skills</code> 可查看已安装的 skill，匹配需求时也会自动触发。<br>
            提示：旧版 Codex 用的是 <code>~/.codex/skills/</code>，如果不生效，
            先升级 Codex 到最新版，或把文件夹放到这个旧路径试试。
          </div>
        </details>
        <details>
          <summary>ChatGPT / 豆包 / 其他 AI 聊天助手</summary>
          <div class="g-body">
            新版 ChatGPT 桌面版已支持 skill（对话里输入 <code>@</code> 查看）。
            其他聊天助手用通用方法：解压后打开里面的 <code>SKILL.md</code>
            （右键 → 打开方式 → 文本编辑），把全文复制，作为对话的<b>第一条消息</b>发给
            AI，然后再说你的需求。<br>
            注意：如果这个 skill 带了脚本文件（比如 scripts 文件夹），纯聊天助手无法运行脚本，
            只能用到里面的方法和模板部分。
          </div>
        </details>
      </div>
    </div>
    ${docHTML ? `<div class="doc">${docHTML}</div>` : ""}
    <div class="vlist">
      ${(versions || []).map(v => `
        <div class="vitem">
          <span class="vnum">v${v.version}</span>
          <span class="vlog">${esc(v.changelog) || "（没写更新说明）"}</span>
          <span class="vmeta">${esc(v.uploader?.display_name || "?")} · ${fmtDate(v.uploaded_at)}</span>
          <button class="btn small ghost" data-ver="${v.version}" data-path="${esc(v.file_path)}">下载</button>
          ${me && me.id === v.uploader_id
            ? `<button class="btn small danger" data-del="${v.id}" data-delver="${v.version}" data-delpath="${esc(v.file_path)}">删除</button>`
            : ""}
        </div>`).join("")}
    </div>`;

  bindRoleEditor(skill);
  bindTitleEditor(skill);
  async function download(path, version) {
    flash("正在下载…");
    const { data, error } = await sb.storage.from("skills").download(path);
    if (error) { flash("下载失败：" + error.message, 4000); return; }
    const url = URL.createObjectURL(data);
    const a = document.createElement("a");
    a.href = url; a.download = `${skill.slug}-v${version}.zip`; a.click();
    URL.revokeObjectURL(url);
    await sb.rpc("increment_downloads", { p_skill_id: skill.id });
  }
  const dlBtn = document.getElementById("dl-btn");
  if (dlBtn) dlBtn.onclick = () => download(latest.file_path, latest.version);
  $app.querySelectorAll("[data-ver]").forEach(b =>
    b.onclick = () => download(b.dataset.path, b.dataset.ver));

  $app.querySelectorAll("[data-del]").forEach(b => b.onclick = async () => {
    const isLast = versions.length === 1;
    if (!await confirmAction(`确定删除 v${b.dataset.delver} 吗？删除后无法恢复。${isLast ? "\n这是唯一的版本，删除后整个 Skill 会一并移除。" : ""}`, "删除版本")) return;
    const { data: deleted, error } = await sb.from("versions")
      .delete().eq("id", b.dataset.del).select("id");
    if (error || !deleted?.length) {
      flash("删除失败：" + (error?.message || "没有删除权限（管理员需在 Supabase 运行 update-delete-permission.sql）"), 6000);
      return;
    }
    await sb.storage.from("skills").remove([b.dataset.delpath]);
    if (isLast) {
      await sb.from("skills").delete().eq("id", skill.id);
      flash("已删除，这是最后一个版本，skill 已一并移除");
      location.hash = "#/";
    } else {
      flash("版本已删除");
      viewDetail(slug);
    }
  });
}

// ---------- 视图：上传 ----------
async function viewUpload(targetSlug = "") {
  let target = null;
  if (targetSlug) {
    const { data, error } = await sb.from("skills").select("*").eq("slug", targetSlug).maybeSingle();
    if (error) throw error;
    if (!data || H.sourceType(data) !== "upload") throw new Error("这个条目不支持上传版本，请回到市集重新选择。");
    target = data;
  }
  $app.innerHTML = `
    <div class="panel wide">
      <a class="backlink" href="${target ? "#/skill/" + encodeURIComponent(target.slug) : "#/share"}">← ${target ? "返回详情" : "选择分享方式"}</a>
      <h2>${target ? "上传新版本" : "上传 Skill 文件"}</h2>
      ${target ? `<div class="notice">正在更新 <strong>${esc(H.getDisplayTitle(target))}</strong>。这次上传会保留旧版本。</div>` : `<p class="form-intro">上传一个完整的 Skill 文件夹压缩包，同事可以下载后安装。</p>`}
      <form id="up-form">
        <div class="field">
          <div class="dropzone" id="dz">
            📦 点击选择（或拖入）skill 的 <b>.zip 压缩包</b><br>
            <span class="hint">在电脑上找到 skill 文件夹，右键 →「压缩」就能得到 zip</span>
            <div id="parse-result"></div>
          </div>
          <input type="file" id="file" accept=".zip" hidden>
        </div>
        <div class="field"><label for="f-display-title">功能标题 *</label>
          <input name="display_title" id="f-display-title" required value="${esc(displayTitleValue(target))}" placeholder="例如 去AI味写作、生成 AI 资讯朋友圈"><p class="hint">最多 80 字，显示在卡片和详情页顶部，让同事一眼看懂它能做什么。</p></div>
        <div class="field"><label for="f-title">原始 Skill 名称 *</label>
          <input name="title" id="f-title" required maxlength="200" value="${esc(target?.title || "")}" placeholder="从 SKILL.md 自动读取，例如 de-ai-copy"><p class="hint">作为副标题展示，方便识别和使用原 Skill。</p></div>
        <div class="field"><label>一句话描述（这个 skill 是干嘛的、什么时候用）</label>
          <textarea name="description" id="f-desc" maxlength="4000" placeholder="上传 zip 后自动识别 SKILL.md 里的描述">${esc(target?.description || "")}</textarea></div>
        <div class="field"><label>分类</label>
          <select name="category">${CATEGORIES.map(c => `<option ${target?.category === c ? "selected" : ""}>${c}</option>`).join("")}</select></div>
        <div class="field"><label>标签（用逗号分隔，方便同事搜索）</label>
          <input name="tags" maxlength="500" value="${esc(target?.tags || "")}" placeholder="朋友圈, 文案, 小红书"></div>
        ${roleFieldsHTML(target?.roles)}
        <div class="field"><label>更新说明（这一版改了什么；首次上传可不填）</label>
          <input name="changelog" placeholder="例如：优化了触发词，修复了日期格式问题"></div>
        <button class="btn" style="width:100%" id="up-btn">发布</button>
        <div class="hint" style="margin-top:10px">如需更新已有 Skill，请从它的详情页选择「上传新版本」。这里新分享的文件会独立保存，单个文件最大 50MB。</div>
      </form>
    </div>`;

  bindRoleFields();
  const fileInput = document.getElementById("file");
  const dz = document.getElementById("dz");
  let picked = null, parsedMd = "", parsing = false, submitting = false, parseSequence = 0;
  const metadataInputs = {
    title: document.getElementById("f-title"),
    description: document.getElementById("f-desc"),
    display_title: document.getElementById("f-display-title")
  };
  const automaticValues = { title: "", description: "", display_title: "" };
  const editedFields = new Set(target ? Object.keys(metadataInputs) : []);
  for (const [field, input] of Object.entries(metadataInputs)) {
    input.oninput = () => editedFields.add(field);
  }
  function suggestMetadata(values) {
    for (const [field, value] of Object.entries(values)) {
      const input = metadataInputs[field];
      if (!editedFields.has(field) && (!input.value || input.value === automaticValues[field])) {
        input.value = value;
        automaticValues[field] = value;
      }
    }
  }

  dz.onclick = () => fileInput.click();
  dz.ondragover = e => { e.preventDefault(); dz.classList.add("drag"); };
  dz.ondragleave = () => dz.classList.remove("drag");
  dz.ondrop = e => {
    e.preventDefault(); dz.classList.remove("drag");
    if (e.dataTransfer.files[0]) { fileInput.files = e.dataTransfer.files; handleFile(); }
  };
  fileInput.onchange = handleFile;

  async function handleFile() {
    if (submitting) return;
    const sequence = ++parseSequence;
    parsedMd = ""; parsing = false;
    const publishButton = document.getElementById("up-btn");
    publishButton.disabled = false; publishButton.textContent = "发布";
    picked = fileInput.files[0];
    suggestMetadata({ title: "", description: "", display_title: "" });
    if (!picked) return;
    if (!picked.name.toLowerCase().endsWith(".zip")) { flash("请选择 .zip 压缩包"); picked = null; return; }
    if (picked.size > 50 * 1024 * 1024) { flash("文件超过 50MB，请精简后再上传。"); picked = null; return; }
    parsing = true;
    publishButton.disabled = true; publishButton.textContent = "正在读取文件…";
    const result = document.getElementById("parse-result");
    result.innerHTML = `<div class="parse-ok">已选择：${esc(picked.name)}（${(picked.size/1024/1024).toFixed(1)}MB）</div>`;
    try {
      const zip = await JSZip.loadAsync(picked);
      if (sequence !== parseSequence) return;
      const names = Object.keys(zip.files)
        .filter(n => !n.includes("__MACOSX") && !zip.files[n].dir);
      const cands = names.filter(n => n.toLowerCase().endsWith("skill.md"));
      if (!cands.length) {
        const shown = names.slice(0, 6).map(esc).join("、");
        result.innerHTML += `<div class="hint">⚠️ 压缩包里没找到 SKILL.md，仍可上传，请在下方手动填写名称和描述。<br>
          包内文件：${shown || "（空）"}${names.length > 6 ? ` 等 ${names.length} 个` : ""}</div>`;
        parsedMd = ""; return;
      }
      const path = cands.sort((a, b) => a.split("/").length - b.split("/").length)[0];
      const markdown = await zip.files[path].async("string");
      if (sequence !== parseSequence) return;
      parsedMd = markdown;
      const fm = parsedMd.match(/^---\s*\n([\s\S]*?)\n---/);
      let name = "", desc = "";
      if (fm) {
        name = fmValue(fm[1], "name");
        desc = fmValue(fm[1], "description");
      }
      if (!name) name = path.split("/").length >= 2 ? path.split("/").slice(-2)[0] : picked.name.replace(".zip", "");
      suggestMetadata({
        title: name,
        description: desc,
        display_title: H.suggestDisplayTitle({ title: name, description: desc })
      });
      result.innerHTML += `<div class="parse-ok">✅ 已自动识别 SKILL.md${name ? `：${esc(name)}` : ""}</div>`;
    } catch (err) {
      if (sequence !== parseSequence) return;
      picked = null; parsedMd = "";
      result.innerHTML += `<div class="form-error" role="alert">无法读取压缩包，请重新选择有效的 ZIP 文件。</div>`;
    } finally {
      if (sequence === parseSequence) {
        parsing = false; publishButton.disabled = false; publishButton.textContent = "发布";
      }
    }
  }

  document.getElementById("up-form").onsubmit = async e => {
    e.preventDefault();
    if (submitting) return;
    if (parsing) { flash("文件还在读取，请稍等。"); return; }
    if (!picked) { flash("请先选择 zip 压缩包"); return; }
    const btn = document.getElementById("up-btn");
    submitting = true;
    btn.disabled = true; btn.textContent = "发布中…";
    let uploadedPath = "", newSkillId = null, versionSaved = false;
    try {
      const fd = new FormData(e.target);
      const title = fd.get("title").trim();
      if (!title) throw new Error("请填写 Skill 名称。");
      const display_title = validateDisplayTitle(fd.get("display_title"));
      const slug = target?.slug || newSlug(title);
      const existing = target;
      let skillId, version = 1;
      if (existing) {
        skillId = existing.id;
        const { data: lastVer, error: lastError } = await sb.from("versions")
          .select("version").eq("skill_id", skillId)
          .order("version", { ascending: false }).limit(1).maybeSingle();
        if (lastError) throw lastError;
        version = (lastVer?.version || 0) + 1;
      }

      const filePath = `${slug}/v${version}-${crypto.randomUUID()}.zip`;
      const { error: upErr } = await sb.storage.from("skills").upload(filePath, picked);
      if (upErr) throw upErr;
      uploadedPath = filePath;

      const meta = {
        title, display_title, description: fd.get("description").trim(),
        category: fd.get("category"), tags: fd.get("tags").trim(),
        roles: H.normalizeRoles(fd.getAll("roles")),
        updated_at: new Date().toISOString(),
      };
      if (!existing) {
        const { data, error } = await sb.from("skills")
          .insert({ ...meta, slug, author_id: me.id }).select("id").single();
        if (error) throw error;
        skillId = data.id;
        newSkillId = skillId;
      }
      const { error: verErr } = await sb.from("versions").insert({
        skill_id: skillId, version, changelog: fd.get("changelog").trim(),
        file_path: filePath, skill_md: parsedMd, uploader_id: me.id,
      });
      if (verErr) throw verErr;
      versionSaved = true;
      if (existing) {
        const { error } = await sb.from("skills").update(meta).eq("id", skillId);
        if (error) {
          location.hash = "#/skill/" + encodeURIComponent(slug);
          flash("文件版本已保存，但介绍更新失败，请稍后重试。", 6000);
          return;
        }
      }

      flash(`发布成功！当前为 v${version} 🎉`);
      location.hash = "#/skill/" + encodeURIComponent(slug);
    } catch (err) {
      let cleanupFailed = false;
      if (!versionSaved) {
        if (newSkillId) {
          try { const result = await sb.from("skills").delete().eq("id", newSkillId).select("id"); cleanupFailed = !!result.error || !result.data?.length; }
          catch { cleanupFailed = true; }
        }
        if (uploadedPath) {
          try { const result = await sb.storage.from("skills").remove([uploadedPath]); cleanupFailed ||= !!result.error; }
          catch { cleanupFailed = true; }
        }
      }
      flash("发布失败，" + errorMessage(err) + (cleanupFailed ? " 部分临时内容未能清理，请联系管理员。" : ""), 6000);
    } finally {
      submitting = false;
      btn.disabled = false; btn.textContent = "发布";
    }
  };
}

// ---------- 外部推荐 ----------
async function viewRecommend(slug = "") {
  let editing = null;
  if (slug) {
    const { data, error } = await sb.from("skills").select("*").eq("slug", slug).maybeSingle();
    if (error) throw error;
    if (!data || H.sourceType(data) !== "external" || data.author_id !== me.id) throw new Error("只有分享者可以编辑这条推荐。");
    editing = data;
  }
  const value = field => esc(editing?.[field] || "");
  $app.innerHTML = `<div class="panel wide recommend-panel">
    <a class="backlink" href="${editing ? "#/skill/" + encodeURIComponent(editing.slug) : "#/share"}">← ${editing ? "返回详情" : "选择分享方式"}</a>
    <h2>${editing ? "编辑推荐" : "推荐网上的 Skill"}</h2>
    <p class="form-intro">${editing ? "可以更新功能标题、简介和使用信息，让同事更容易上手。" : "粘贴链接自动读取信息，确认功能标题和简介后即可分享。"}</p>
    <form id="recommend-form">
      <div class="field"><label for="r-url">Skill 链接 *</label>
        <div class="import-input"><input id="r-url" name="source_url" type="url" required maxlength="2048" value="${value("source_url")}" ${editing ? "readonly" : ""} placeholder="粘贴 GitHub 或 skills.sh 链接">
        ${editing ? "" : `<button class="btn" id="import-button" type="button">自动读取</button>`}</div>
        <p class="hint">${editing ? "来源链接已固定。" : "支持具体 Skill 或整个 GitHub 仓库。其他网站可以手动补充信息。"}</p></div>
      <div id="import-status" class="import-status" role="status" aria-live="polite"></div>
      <div class="field skill-picker" id="skill-choice-field" hidden><label for="skill-choice">这个仓库有多个 Skill，选一个推荐</label><select id="skill-choice"><option value="">请选择 Skill</option></select></div>
      <div id="imported-preview" class="imported-preview" hidden></div>
      <div class="field"><label for="r-display-title">功能标题 *</label><input id="r-display-title" name="display_title" required value="${esc(displayTitleValue(editing))}" placeholder="例如 生成读书感悟朋友圈"><p class="hint">用一句短话说明它能做什么，最多 80 字。原始名称会显示在副标题中。</p></div>
      <details id="metadata-fields" class="form-details">
        <summary>查看或修改基本信息 <span>名称、介绍、作者等</span></summary>
        <div class="details-content">
          <div class="form-grid">
            <div class="field"><label for="r-title">原始 Skill 名称 *</label><input id="r-title" name="title" maxlength="200" value="${value("title")}" placeholder="例如 book-insights"></div>
            <div class="field"><label for="r-author">原作者</label><input id="r-author" name="original_author" maxlength="200" value="${value("original_author")}" placeholder="作者或组织名称"></div>
          </div>
          <div class="field"><label for="r-description">用途介绍 *</label><textarea id="r-description" name="description" maxlength="4000" placeholder="这个 Skill 能做什么，什么时候用">${value("description")}</textarea></div>
          <div class="form-grid">
            <div class="field"><label for="r-category">分类</label><select id="r-category" name="category">${CATEGORIES.map(c => `<option ${(editing?.category || "其他") === c ? "selected" : ""}>${c}</option>`).join("")}</select></div>
            <div class="field"><label for="r-tags">标签</label><input id="r-tags" name="tags" maxlength="500" value="${value("tags")}" placeholder="选填，用逗号隔开"></div>
          </div>
          <div class="field"><label for="r-path">仓库内的 Skill 目录</label><input id="r-path" name="source_path" maxlength="1024" value="${value("source_path")}" ${editing ? "readonly" : ""} placeholder="自动识别，或手动填写，例如 skills/pdf"><p class="hint">仅 GitHub 需要。手动推荐仓库首页时要填具体目录；根目录填一个英文句点。</p></div>
        </div>
      </details>
      ${roleFieldsHTML(editing?.roles)}
      <details id="extra-fields" class="form-details" ${editing?.tested || editing?.install_notes ? "open" : ""}>
        <summary>补充使用经验 <span>选填</span></summary>
        <div class="details-content">
          <div class="field"><label for="r-tested">你用过了吗</label><select id="r-tested" name="tested"><option value="false" ${!editing?.tested ? "selected" : ""}>还没试过，先分享给大家</option><option value="true" ${editing?.tested ? "selected" : ""}>我用过，可以参考我的经验</option></select></div>
          <div class="form-grid" id="trial-fields" ${editing?.tested ? "" : "hidden"}>
            <div class="field"><label for="r-tool">试用工具 *</label><input id="r-tool" name="tested_with" maxlength="200" value="${value("tested_with")}" placeholder="例如 Claude Code"></div>
            <div class="field"><label for="r-version">试用版本或日期</label><input id="r-version" name="tested_version" maxlength="200" value="${value("tested_version")}" placeholder="填版本、commit 或试用日期"></div>
          </div>
          <div class="field"><label for="r-install">安装与使用补充</label><textarea id="r-install" name="install_notes" maxlength="8000" placeholder="装好后可以怎么用？有什么需要提前准备？">${value("install_notes")}</textarea></div>
        </div>
      </details>
      <div id="recommend-error" class="form-error" role="alert"></div>
      <div class="actions"><button class="btn" id="recommend-submit">${editing ? "保存修改" : "发布推荐"}</button><a class="btn ghost" href="${editing ? "#/skill/" + encodeURIComponent(editing.slug) : "#/"}">取消</a></div>
    </form>
  </div>`;

  const el = id => document.getElementById(id);
  const form = el("recommend-form");
  bindRoleFields();
  const urlInput = el("r-url"), pathInput = el("r-path");
  const errorBox = el("recommend-error"), status = el("import-status"), preview = el("imported-preview");
  const metadataFields = el("metadata-fields"), extraFields = el("extra-fields");
  const button = el("recommend-submit"), importButton = el("import-button");
  const picker = el("skill-choice"), pickerField = el("skill-choice-field");
  const fields = { title: el("r-title"), display_title: el("r-display-title"), description: el("r-description"), original_author: el("r-author"), category: el("r-category"), tags: el("r-tags") };
  let request = 0, controller = null, loading = false, submitting = false;
  let customDisplayTitle = false;
  let imported = null, choices = [], lastAttempt = "", pasteTimer;
  const active = () => document.getElementById("recommend-form") === form;
  function updateButtons() {
    button.disabled = loading || submitting;
    button.textContent = submitting ? "正在保存…" : editing ? "保存修改" : "发布推荐";
    if (importButton) { importButton.disabled = loading || submitting; importButton.textContent = loading ? "读取中…" : "自动读取"; }
    picker.disabled = loading || submitting;
  }
  function trialState() { el("trial-fields").hidden = el("r-tested").value !== "true"; }
  el("r-tested").onchange = trialState;
  trialState();
  function renderSummary() {
    const title = fields.title.value.trim(), description = fields.description.value.trim();
    preview.hidden = !title;
    const displayTitle = H.getDisplayTitle({ title, display_title: fields.display_title.value });
    preview.innerHTML = title ? `<div class="imported-heading"><span class="imported-icon" aria-hidden="true">↗</span><div><strong>${esc(displayTitle)}</strong><p class="skill-original-name">${esc(title)}</p><p>${esc(fields.original_author.value || "原作者未填写")}${fields.category.value ? " · " + esc(fields.category.value) : ""}</p></div><span class="source-badge external">${imported ? "已读取" : "基本信息"}</span></div><p class="imported-description">${esc(description || "请补充用途介绍")}</p>` : "";
  }
  Object.values(fields).forEach(field => { field.oninput = renderSummary; });
  fields.display_title.oninput = () => { customDisplayTitle = true; renderSummary(); };
  const suggestTitle = () => {
    if (!customDisplayTitle && !editing) fields.display_title.value = H.suggestDisplayTitle({ title: fields.title.value, description: fields.description.value });
    renderSummary();
  };
  fields.title.oninput = suggestTitle;
  fields.description.oninput = suggestTitle;
  fields.category.onchange = renderSummary;
  renderSummary();
  function cancelImport() {
    request++; controller?.abort(); controller = null; loading = false; updateButtons();
  }
  function resetSource() {
    cancelImport();
    if (imported) {
      Object.entries(fields).forEach(([key, field]) => { field.value = key === "category" ? "其他" : ""; });
      customDisplayTitle = false;
    }
    imported = null; choices = []; pathInput.value = ""; lastAttempt = "";
    pickerField.hidden = true; status.textContent = ""; errorBox.textContent = "";
    renderSummary();
  }
  async function readSource(choice = null) {
    if (editing || submitting || !active()) return;
    const rawUrl = urlInput.value.trim();
    if (!rawUrl) { status.textContent = "先粘贴一个 Skill 链接。"; urlInput.focus(); return; }
    cancelImport();
    const token = request;
    controller = new AbortController(); loading = true; lastAttempt = rawUrl;
    errorBox.textContent = ""; status.textContent = "正在读取项目里的 Skill 信息…"; updateButtons();
    try {
      if (!window.SkillHubImporter) throw new Error("自动读取暂时不可用，请手动补充基本信息。");
      const result = choice
        ? await window.SkillHubImporter.readChoice(choice, { signal: controller.signal })
        : await window.SkillHubImporter.inspect(rawUrl, { signal: controller.signal });
      if (token !== request || !active()) return;
      if (result.type === "choices") {
        choices = result.choices;
        picker.innerHTML = `<option value="">请选择 Skill</option>` + choices.map((item, i) => `<option value="${i}">${esc(item.label || item.path)}</option>`).join("");
        picker.value = ""; pickerField.hidden = false;
        status.textContent = result.notice || `找到了 ${choices.length} 个 Skill，选中后会自动填好信息。`;
        return;
      }
      if (result.type !== "skill" || !result.skill?.source) throw new Error("没有读到完整的 Skill 信息，请手动补充。");
      imported = result.skill;
      // The importer verifies the directory and ref before this URL is saved.
      const source = H.normalizeSource(imported.source.url, imported.source.path);
      urlInput.value = source.url; pathInput.value = source.path; lastAttempt = source.url;
      imported.source = source;
      for (const [key, field] of Object.entries(fields)) {
        if (key === "display_title" && customDisplayTitle) continue;
        const limit = key === "description" ? 4000 : key === "tags" ? 500 : key === "display_title" ? 80 : 200;
        field.value = key === "category" ? (CATEGORIES.includes(imported.category) ? imported.category : "其他")
          : Array.from(String(Array.isArray(imported[key]) ? imported[key].join(", ") : imported[key] || "")).slice(0, limit).join("");
      }
      if (!customDisplayTitle && !fields.display_title.value) fields.display_title.value = H.suggestDisplayTitle(imported);
      choices = []; pickerField.hidden = true; metadataFields.open = !fields.title.value || !fields.description.value;
      status.textContent = result.notice || "已读取，确认功能标题和简介后即可分享。";
      renderSummary();
    } catch (error) {
      if (token !== request || !active()) return;
      status.textContent = `${errorMessage(error)}${error?.message?.includes("手动") ? "" : " 可以展开基本信息，继续手动填写。"}`;
      metadataFields.open = true;
    } finally {
      if (token === request && active()) { loading = false; controller = null; updateButtons(); }
    }
  }
  if (!editing) {
    importButton.onclick = () => readSource();
    urlInput.oninput = resetSource;
    urlInput.onblur = () => {
      if (!loading && urlInput.value.trim() && urlInput.value.trim() !== lastAttempt) return readSource();
    };
    urlInput.onpaste = () => { clearTimeout(pasteTimer); pasteTimer = setTimeout(() => { if (active()) readSource(); }, 0); };
    picker.onchange = () => picker.value !== "" && choices[Number(picker.value)] ? readSource(choices[Number(picker.value)]) : undefined;
    pathInput.oninput = () => { cancelImport(); imported = null; choices = []; pickerField.hidden = true; status.textContent = "已改为手动填写，请核对来源目录。"; renderSummary(); };
  }
  form.onsubmit = async event => {
    event.preventDefault();
    if (loading || submitting) return;
    errorBox.textContent = ""; submitting = true; updateButtons();
    try {
      if (choices.length) { picker.focus(); throw new Error("请先选择要推荐的 Skill。"); }
      const fd = new FormData(form);
      const get = key => String(fd.get(key) || "").trim();
      let source;
      try { source = H.normalizeSource(get("source_url"), get("source_path")); }
      catch (error) { metadataFields.open = true; throw error; }
      const metadata = {
        title: get("title"), description: get("description"), category: get("category"), tags: get("tags"),
        roles: H.normalizeRoles(fd.getAll("roles")),
        original_author: get("original_author"),
        tested: get("tested") === "true", tested_with: get("tested") === "true" ? get("tested_with") : "",
        tested_version: get("tested") === "true" ? get("tested_version") : "", install_notes: get("install_notes"), updated_at: new Date().toISOString(),
      };
      if (!metadata.title || !metadata.description) {
        metadataFields.open = true; (!metadata.title ? fields.title : fields.description).focus();
        throw new Error("还缺少名称或用途介绍，请补充后发布。");
      }
      try { metadata.display_title = validateDisplayTitle(get("display_title")); }
      catch (error) { fields.display_title.focus(); throw error; }
      if (metadata.tested && !metadata.tested_with) { extraFields.open = true; el("r-tool").focus(); throw new Error("请补充试用工具，方便同事参考。"); }
      let saved;
      if (editing) {
        const { data, error } = await sb.from("skills").update(metadata).eq("id", editing.id).eq("author_id", me.id).select("slug").single();
        if (error) throw error;
        saved = data;
      } else {
        const findDuplicate = async () => {
          const keys = [...new Set([source.key, ...(imported?.source.key === source.key ? imported.aliasKeys || [] : [])])];
          for (const key of keys) {
            const result = await sb.from("skills").select("slug").eq("source_type", "external").eq("source_key", key).maybeSingle();
            if (result.error) throw result.error;
            if (result.data) return result.data;
          }
          return null;
        };
        const existing = await findDuplicate();
        if (existing) {
          location.hash = "#/skill/" + encodeURIComponent(existing.slug);
          flash("这项资源已经有人推荐，已为你打开现有条目。", 5000);
          return;
        }
        const { data, error } = await sb.from("skills").insert({ ...metadata, slug: newSlug(metadata.title), author_id: me.id,
          source_type: "external", source_url: source.url, source_key: source.key, source_repo: source.repo, source_path: source.path,
        }).select("slug").single();
        if (error?.code === "23505") {
          const duplicate = await findDuplicate();
          if (duplicate) { location.hash = "#/skill/" + encodeURIComponent(duplicate.slug); flash("这项资源刚刚被推荐，已打开现有条目。"); return; }
        }
        if (error) throw error;
        saved = data;
      }
      if (!saved?.slug) throw new Error("保存没有成功，请稍后重试。");
      location.hash = "#/skill/" + encodeURIComponent(saved.slug);
      flash(editing ? "推荐已更新" : "已推荐给同事");
    } catch (error) {
      errorBox.textContent = errorMessage(error);
    } finally {
      submitting = false; updateButtons();
    }
  };
}

function viewExternalDetail(skill) {
  const sourceUrl = H.safeUrl(skill.source_url);
  const installPrompt = H.buildInstallPrompt(skill);
  let source = null;
  try { source = H.normalizeSource(skill.source_url, skill.source_path); } catch { /* Keep original instructions available. */ }
  const owner = me?.id === skill.author_id;
  $app.innerHTML = `<a class="backlink" href="#/">← 返回市集</a>
    <div class="detail-head">
      <div class="meta"><span class="source-badge external">外部推荐</span><span class="tag">${esc(skill.category)}</span><span>${skill.tested ? "分享者已试用" : "尚未试用"}</span></div>
      ${skillHeadingHTML(skill)}<p class="desc">${esc(skill.description)}</p>
      <div class="meta"><span>${esc(skill.author?.display_name || "同事")} · ${esc(skill.author?.department || "")}</span><span>更新于 ${fmtDate(skill.updated_at)}</span></div>
      ${detailRolesHTML(skill)}
      <div class="actions"><a class="btn" href="#installation" id="jump-install">让 AI 帮我安装 ↓</a>${sourceUrl ? `<a class="btn ghost" href="${esc(sourceUrl)}" target="_blank" rel="noopener noreferrer">查看原项目 ↗</a>` : ""}
        ${owner ? `<a class="btn ghost" href="#/recommend/${encodeURIComponent(skill.slug)}">编辑推荐</a>` : ""}</div>
    </div>
    <div class="detail-layout">
      <div class="detail-main">
        ${readmeHTML(skill, true)}
        <section class="detail-panel" id="installation"><h2>安装与使用</h2>
          ${installPrompt ? `<div class="install-step"><span class="step-number">1</span><div><h3>把安装请求发给 AI</h3><p>复制下面这段话，发到你常用的 AI 工具里，让它帮你安装。</p></div></div>
          <div class="install-prompt" id="install-prompt">${esc(installPrompt)}</div>
          <button class="btn" id="copy-install">复制安装请求</button>
          <div class="install-step next-step"><span class="step-number">2</span><div><h3>告诉它你要做什么</h3><p>安装完成后，说一句「用这个 Skill 帮我……」，接着描述任务，并附上相关资料。也可以让 AI 先给你演示一个用法。</p></div></div>
          <p class="install-note">能否直接安装取决于当前工具。若不支持，AI 会根据这段请求说明原因和其他用法。</p>` : `<p>这条推荐的来源链接暂不可用，请联系分享者补充，或参考下方的使用说明。</p>`}
          ${skill.install_notes ? `<hr class="section-divider"><h2>分享者补充</h2><p class="prose">${esc(skill.install_notes)}</p>` : ""}
        </section>
      </div>
      <aside>
        <section class="detail-panel"><h2>来源信息</h2><dl class="identity-list">
          <dt>原作者</dt><dd>${esc(skill.original_author || source?.repo?.split("/")[0] || "未填写")}</dd>
          <dt>公司分享者</dt><dd>${esc(skill.author?.display_name || "同事")}</dd>
          <dt>来源</dt><dd>${sourceUrl ? `<a href="${esc(sourceUrl)}" target="_blank" rel="noopener noreferrer">${esc(source?.repo || new URL(sourceUrl).hostname)} ↗</a>` : "原链接暂不可用"}</dd>
          ${source?.path ? `<dt>Skill 目录</dt><dd>${esc(source.path)}</dd>` : ""}
          <dt>分享时间</dt><dd>${fmtDate(skill.created_at)}</dd>
        </dl></section>
        <section class="detail-panel"><h2>试用记录</h2>
          ${skill.tested ? `<dl class="identity-list"><dt>试用者</dt><dd>${esc(skill.author?.display_name || "分享者")}</dd><dt>使用工具</dt><dd>${esc(skill.tested_with || "未填写")}</dd><dt>版本或日期</dt><dd>${esc(skill.tested_version || "未填写")}</dd></dl><p class="install-note">这条记录由分享者填写，安装时请留意来源是否已更新。</p>` : `<p>分享者还没有试用。</p><p class="install-note">可以先看看原项目的说明，了解它是否适合你的工作。</p>`}
        </section>
        ${owner ? `<button class="btn small danger" id="remove-recommendation">移除这条推荐</button>` : ""}
      </aside>
    </div>`;
  document.getElementById("jump-install").onclick = event => {
    event.preventDefault(); document.getElementById("installation").scrollIntoView({ behavior: "auto", block: "start" });
  };
  const copy = document.getElementById("copy-install");
  if (copy) copy.onclick = () => copyText(installPrompt, "已复制，发给 AI 就可以开始安装");
  bindRoleEditor(skill);
  bindTitleEditor(skill);
  bindReadmes([skill]);
  const remove = document.getElementById("remove-recommendation");
  if (remove) remove.onclick = async () => {
    if (!await confirmAction("确定移除这条推荐吗？原项目和已安装的 Skill 不受影响。", "移除推荐")) return;
    remove.disabled = true;
    try {
      const { data, error } = await sb.from("skills").delete().eq("id", skill.id).eq("author_id", me.id).select("id");
      if (error) throw error;
      if (!data?.length) throw new Error("没有移除成功，请确认你仍有权限。");
      location.hash = "#/"; flash("推荐已移除");
    } catch (error) { flash(errorMessage(error), 5000); remove.disabled = false; }
  };
}

// ---------- 启动 ----------
if (PREVIEW) {
  const banner = document.getElementById("preview-banner");
  banner.hidden = false;
  banner.innerHTML = `<span>${esc(window.SKILLHUB_PREVIEW_NOTICE || "本地预览 · 示例内容与公司数据分开，修改保存在当前浏览器会话")}</span><button id="reset-preview">重置示例</button>`;
  document.getElementById("reset-preview").onclick = async () => { if (await confirmAction("重置本次预览的所有修改？", "重置预览")) window.resetSkillHubPreview(); };
}
route();
