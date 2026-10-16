// UI：路由、页面渲染、事件
(function () {
  'use strict';
  const L = window.SalesLogic, DB = window.SalesDB, AI = window.SalesAI, W = window.SalesWorkflows;
  const CFG = window.SALES_SCRIPT_CONFIG || {};
  const AIStatus = window.SalesAIStatus;
  const DEPARTMENTS = ['销售', '运营', '内容', '交付', '产品', '其他'];
  const BG_KEY = 'ai_background';
  const CONTENT_FIELDS = ['stage', 'scene', 'customer_says', 'reply', 'tips', 'channel'];
  const KNOWLEDGE_KINDS = { detail: '课程详情页', live: '销转直播', faq: '常见问题', other: '其他资料' };

  // ---------- 小工具 ----------
  const $ = (sel, el) => (el || document).querySelector(sel);
  const $$ = (sel, el) => [...(el || document).querySelectorAll(sel)];
  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const app = () => $('#app');
  const clone = x => JSON.parse(JSON.stringify(x));
  const uniq = arr => [...new Set(arr.filter(Boolean))];

  function toast(msg, isErr) {
    const el = document.createElement('div');
    el.className = 'toast' + (isErr ? ' err' : '');
    el.textContent = msg;
    $('#toast').appendChild(el);
    setTimeout(() => el.remove(), isErr ? 5000 : 2200);
  }
  function showErr(e) {
    console.error(e);
    toast((e && e.message) || String(e), true);
    if (e && e.code === 'auth' && DB.mode === 'supabase' && S.started) setTimeout(() => location.reload(), 1500);
  }

  async function withBusy(btn, label, fn) {
    const old = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = `<span class="spin"></span> ${esc(label)}`;
    try { return await fn(); } finally {
      if (btn.isConnected) { btn.disabled = false; btn.innerHTML = old; }
    }
  }

  function openModal(html, opts) {
    const mask = document.createElement('div');
    mask.className = 'modal-mask';
    mask.innerHTML = `<div class="modal ${opts && opts.small ? 'sm' : ''}">${html}</div>`;
    $('#modal-root').appendChild(mask);
    const close = () => mask.remove();
    $$('[data-close]', mask).forEach(b => b.addEventListener('click', () => (opts && opts.onCancel ? opts.onCancel() : close())));
    return { el: mask, close };
  }
  function choiceBox(title, text, choices) {
    return new Promise(resolve => {
      const m = openModal(`<h2>${esc(title)}</h2>${text ? `<p class="muted pre">${esc(text)}</p>` : ''}
        <div class="modal-actions">${choices.map((c, i) => `<button class="btn ${c.cls || ''}" data-i="${i}">${esc(c.label)}</button>`).join('')}</div>`, { small: true });
      $$('[data-i]', m.el).forEach(b => b.addEventListener('click', () => { m.close(); resolve(choices[+b.dataset.i].value); }));
    });
  }
  const confirmBox = (title, text, okLabel, danger) =>
    choiceBox(title, text, [{ label: '取消', value: false, cls: 'plain' }, { label: okLabel || '确定', value: true, cls: danger ? 'danger' : '' }]);

  function autosize(ta) { ta.style.height = 'auto'; ta.style.height = (ta.scrollHeight + 2) + 'px'; }

  async function copyText(t, successMessage) {
    let copied = false;
    try { await navigator.clipboard.writeText(t); copied = true; } catch (e) {
      const active = document.activeElement;
      const ta = document.createElement('textarea');
      ta.value = t; document.body.appendChild(ta); ta.select();
      try { copied = document.execCommand('copy'); } catch (_) { /* 提示用户手动复制 */ }
      ta.remove();
      if (active && active.isConnected) active.focus({ preventScroll: true });
    }
    toast(copied ? (successMessage || '已复制') : '复制失败，请选中文字手动复制', !copied);
    return copied;
  }

  function replyMessagesHtml(text, opts) {
    opts = opts || {};
    const messages = L.splitReplyMessages(text);
    if (!messages.length) return '<div class="message-empty">填写发送内容后，这里会按消息预览</div>';
    const shown = opts.compact ? messages.slice(0, 2) : messages;
    return `<div class="reply-messages ${opts.compact ? 'compact' : ''}">
      <div class="message-summary">${messages.length} 条消息${opts.compact ? '' : ' · 按顺序发送'}</div>
      <div class="message-list" role="list" aria-label="发送消息">${shown.map((message, i) => `<div class="script-message" role="listitem">
        <div class="message-meta"><span>消息 ${i + 1}</span>${opts.copy ? `<button type="button" class="linkbtn" data-copy-message="${i + 1}" aria-label="复制第 ${i + 1} 条消息">复制这条</button>` : ''}</div>
        <div class="message-text">${esc(message)}</div>
      </div>`).join('')}</div>
      ${shown.length < messages.length ? `<div class="message-more">还有 ${messages.length - shown.length} 条，查看完整话术</div>` : ''}
    </div>`;
  }

  // 委托处理，导入和编辑中的实时预览更新后也能使用。
  document.addEventListener('click', event => {
    const button = event.target.closest && event.target.closest('[data-copy-message]');
    if (!button || button.disabled) return;
    event.preventDefault();
    event.stopPropagation();
    const text = $('.message-text', button.closest('.script-message')).textContent;
    copyText(text, `已复制第 ${button.dataset.copyMessage} 条消息`);
  });

  // 截图压缩：宽度缩到 1080，太长的长截图切成几段（模型看得清字）
  async function prepareImages(files) {
    const out = [];
    for (const f of files) {
      if (!/^image\//.test(f.type)) continue;
      const url = URL.createObjectURL(f);
      try {
        const img = new Image();
        img.src = url;
        await img.decode();
        const scale = Math.min(1, 1080 / img.naturalWidth);
        const w = Math.round(img.naturalWidth * scale), h = Math.round(img.naturalHeight * scale);
        const SEG = 2400, OVERLAP = 120;
        for (let y = 0; y < h; y += SEG - OVERLAP) {
          const sh = Math.min(SEG, h - y);
          const c = document.createElement('canvas');
          c.width = w; c.height = sh;
          const ctx = c.getContext('2d');
          ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, w, sh);
          ctx.drawImage(img, 0, y / scale, img.naturalWidth, sh / scale, 0, 0, w, sh);
          out.push(c.toDataURL('image/jpeg', 0.85));
          if (y + sh >= h) break;
        }
      } finally { URL.revokeObjectURL(url); }
    }
    if (out.length > 40) throw DB.makeError('invalid', '截图太多了，一次最多 40 张（长截图会被切成几张）');
    return out;
  }
  const imagesFromClipboard = e => [...((e.clipboardData && e.clipboardData.files) || [])].filter(f => /^image\//.test(f.type));

  const chTag = ch => `<span class="tag ch-${esc(ch)}">${esc((L.CHANNELS[ch] || L.CHANNELS.all).label)}</span>`;
  const resultBadge = r => `<span class="badge ${esc(r)}">${esc(L.RESULTS[r] || '')}</span>`;
  const statusBadge = s => s === 'discussed' ? '<span class="badge discussed">已讨论</span>'
    : s === 'skipped' ? '<span class="badge discussed">不讨论</span>' : '<span class="badge pending">待复盘</span>';
  const channelOptions = (keys, cur) => keys.map(k => `<option value="${k}" ${cur === k ? 'selected' : ''}>${esc(L.CHANNELS[k].label)}</option>`).join('');

  // ---------- 状态 ----------
  const S = {
    session: null, me: '', email: '', userId: '',
    scripts: [], cases: [], reviews: [], bg: '', knowledgeProject: '',
    dirty: new Set(), lastHash: '', skipHash: false, started: false, seq: 0,
  };
  // 每次渲染页面拿一个令牌；异步数据回来时页面已经换了，就不再往页面上写
  function viewToken() { const t = ++S.seq; return () => t === S.seq; }
  async function loadAll() {
    const [scripts, cases, reviews, bg] = await Promise.all([
      DB.listScripts(), DB.listCases(), DB.listReviews(), DB.getSetting(BG_KEY),
    ]);
    Object.assign(S, { scripts, cases, reviews, bg });
  }
  const scriptById = id => S.scripts.find(s => s.id === Number(id));
  const caseById = id => S.cases.find(c => c.id === Number(id));
  const reviewByWeek = w => S.reviews.find(r => r.week === w);
  const reviewById = id => S.reviews.find(r => r.id === Number(id));
  function upsert(list, row) {
    const i = list.findIndex(x => x.id === row.id);
    if (i >= 0) list[i] = row; else list.unshift(row);
  }
  const allStages = () => L.sortStages(L.DEFAULT_STAGES.concat(S.scripts.map(s => s.stage)));
  const allScenes = () => uniq(S.scripts.filter(s => s.status !== 'archived').map(s => s.scene)
    .concat(...S.cases.map(c => c.scenes || [])));
  const markDirty = k => S.dirty.add(k);
  const clearDirty = k => S.dirty.delete(k);

  function scriptRef(id, version) {
    const s = scriptById(id);
    if (!s) return `<span class="faint">话术#${esc(id)}（已删除）</span>`;
    const v = arguments.length < 2 ? s.version : Number.isInteger(Number(version)) && Number(version) > 0 ? Number(version) : null;
    return `<a href="#/script/${s.id}">「${esc(s.scene)}」${v ? `v${v}` : '（版本未知）'}</a>${v && v < s.version ? `<span class="faint">（现在是 v${s.version}）</span>` : ''}`;
  }

  // ---------- 路由 ----------
  const ROUTES = [
    [/^\/?$/, 'assistant', () => viewAssistant()],
    [/^\/library$/, 'library', () => viewLibrary()],
    [/^\/script\/(\d+)$/, 'library', m => viewScript(+m[1])],
    [/^\/import$/, 'library', () => viewImport()],
    [/^\/knowledge$/, 'knowledge', () => viewKnowledge()],
    [/^\/knowledge\/new$/, 'knowledge', () => viewKnowledgeEdit(null)],
    [/^\/knowledge\/(\d+)$/, 'knowledge', m => viewKnowledgeEdit(+m[1])],
    [/^\/background$/, 'library', () => viewBackground()],
    [/^\/cases$/, 'cases', () => viewCases()],
    [/^\/case\/new$/, 'cases', () => viewCaseEdit(null)],
    [/^\/case\/(\d+)$/, 'cases', m => viewCase(+m[1])],
    [/^\/case\/(\d+)\/edit$/, 'cases', m => viewCaseEdit(+m[1])],
    [/^\/review(?:\/(\d{4}-\d{2}-\d{2}))?$/, 'review', m => viewReview(m[1])],
    [/^\/members$/, '', () => viewMembers()],
  ];

  async function route() {
    if (S.skipHash) { S.skipHash = false; return; }
    const modalOpen = $('#modal-root').children.length > 0;
    if ((S.dirty.size || modalOpen) && location.hash !== S.lastHash) {
      if (!confirm('有修改还没保存，确定离开吗？')) { S.skipHash = true; location.hash = S.lastHash; return; }
    }
    S.dirty.clear();
    if (S.assistantAbort) { S.assistantAbort.abort(); S.assistantAbort = null; }
    $('#modal-root').innerHTML = '';
    viewToken();
    S.lastHash = location.hash;
    if (document.body.classList.contains('present') && !/^#\/review/.test(location.hash)) exitPresent();
    const path = (location.hash || '#/').slice(1) || '/';
    for (const [re, nav, fn] of ROUTES) {
      const m = re.exec(path);
      if (!m) continue;
      $$('#nav a').forEach(a => a.classList.toggle('on', a.dataset.nav === nav));
      window.scrollTo(0, 0);
      try { await fn(m); } catch (e) { renderError(e); }
      return;
    }
    location.hash = '#/';
  }

  function renderError(e) {
    console.error(e);
    app().innerHTML = `<div class="card empty"><div class="note err" style="display:inline-block">${esc((e && e.message) || e)}</div><br><button class="btn plain" onclick="location.reload()">刷新重试</button></div>`;
  }
  function notFound(msg) {
    app().innerHTML = `<div class="card empty">${esc(msg)}<br><a href="#/">回到首页</a></div>`;
  }

  // ---------- 启动 / 登录 ----------
  function renderAIStatus(state) {
    const box = $('#ai-status-content');
    const button = $('#ai-check');
    if (!box || !button) return;
    button.hidden = state.demo;
    button.disabled = state.checking || state.pending > 0;
    button.textContent = state.checking ? '检查中…' : '检查连接';
    if (state.demo) {
      box.innerHTML = '<strong>AI 演示模式</strong><p class="hint">使用本地模拟结果，尚未连接 ChatGPT 或 DeepSeek。</p>';
      return;
    }
    const time = value => value ? new Date(value).toLocaleString('zh-CN', { hour12: false }) : '';
    const names = { openai: 'ChatGPT（OpenAI）', deepseek: 'DeepSeek' };
    const last = state.lastCall;
    let summary = '优先使用 ChatGPT，失败时自动切换 DeepSeek';
    if (state.pending) summary = `正在处理 ${state.pending} 个 AI 请求，优先尝试 ChatGPT`;
    else if (last && !last.ok) summary = '最近一次 AI 调用未完成';
    else if (last) summary = last.provider ? `最近使用 ${names[last.provider]}${last.model ? ' · ' + last.model : ''}` : '最近调用已完成，服务商待确认';
    const rows = Object.entries(names).map(([id, label]) => {
      const p = state.providers[id];
      const text = p.state === 'available' ? (p.reason ? '连接正常，请求未完成' : p.source === 'probe' ? '连接检查通过' : '最近调用成功')
        : p.state === 'not_configured' ? '尚未配置'
          : p.state === 'unavailable' ? '暂不可用' : '尚未检查';
      return `<div class="ai-provider ${esc(p.state)}"><div><span class="ai-dot" aria-hidden="true"></span><b>${label}</b><span>${text}</span></div>
        ${p.message && (p.state !== 'available' || p.reason) && p.state !== 'unknown' ? `<p>${esc(p.message)}</p>` : ''}
        <small>${p.checked_at ? '更新于 ' + esc(time(p.checked_at)) : p.state === 'not_configured' ? '配置后点击检查连接' : '等待连接检查或实际调用'}</small></div>`;
    }).join('');
    box.innerHTML = `<strong>${esc(summary)}</strong>
      ${last && last.ok && last.fallback ? `<p class="ai-fallback">上次因 ${esc(last.reason || 'ChatGPT 暂不可用')} 改用 DeepSeek。下次仍优先尝试 ChatGPT。</p>` : ''}
      <div class="ai-providers">${rows}</div>
      ${state.issue ? `<p class="ai-issue" role="alert">${esc(state.issue)}</p>` : ''}
      <p class="hint">检查的是服务器到 AI 服务的连接，浏览器无需直连 ChatGPT。连接检查通过后，实际生成仍可能遇到额度或模型权限问题。</p>`;
  }

  function initAIStatus() {
    if (!AIStatus) return;
    $('#ai-status').hidden = false;
    AIStatus.subscribe(renderAIStatus);
    $('#ai-check').addEventListener('click', () => AIStatus.check().catch(() => {}));
    // 状态区会保留错误和重试入口，连接检查不阻塞业务页面。
    AIStatus.check().catch(() => {});
  }

  async function boot() {
    const { mode } = DB.init(CFG);
    $('#demo-bar').hidden = mode !== 'demo';
    if (mode === 'demo') {
      $('#logout').hidden = true;
      $('#demo-fill').addEventListener('click', fillDemo);
      $('#demo-reset').addEventListener('click', async () => {
        if (!(await confirmBox('清空演示数据', '会删掉这台电脑上演示模式里的所有话术、案例和复盘记录。', '清空', true))) return;
        DB.resetDemo(); await loadAll(); route();
      });
    } else {
      DB.onAuthChange(ev => { if (ev === 'SIGNED_OUT' && S.started) location.reload(); });
      $('#logout').addEventListener('click', async () => { await DB.signOut(); location.reload(); });
    }
    window.addEventListener('beforeunload', e => { if (S.dirty.size) { e.preventDefault(); e.returnValue = ''; } });
    document.addEventListener('fullscreenchange', () => { if (!document.fullscreenElement) document.body.classList.remove('present'); });
    try { await start(); } catch (e) { renderError(e); }
  }

  async function start() {
    S.session = await DB.getSession();
    if (!S.session) return renderLogin();
    S.me = S.session.profile.display_name || S.session.user.email;
    S.email = S.session.user.email;
    S.userId = S.session.user.id;
    $('#me').textContent = S.me;
    const acc = await DB.checkAccess();
    if (!acc.ok) return renderNoAccess(acc);
    await loadAll();
    $('#top').hidden = false;
    initAIStatus();
    S.started = true;
    S.lastHash = location.hash;
    window.addEventListener('hashchange', route);
    route();
  }

  function renderLogin(isSignUp) {
    $('#top').hidden = true;
    app().innerHTML = `<div class="auth card">
      <h1>话术复盘</h1>
      <p class="sub">和 Skill 市集共用账号，已经注册过的直接登录</p>
      <form id="auth-form">
        ${isSignUp ? `<div class="field"><label>姓名</label><input type="text" name="name" placeholder="你的名字" required></div>
          <div class="field"><label>部门</label><select name="dept">${DEPARTMENTS.map(d => `<option>${d}</option>`).join('')}</select></div>` : ''}
        <div class="field"><label>邮箱</label><input type="email" name="email" required autocomplete="email"></div>
        <div class="field"><label>密码</label><input type="password" name="password" required minlength="6" autocomplete="${isSignUp ? 'new-password' : 'current-password'}"></div>
        <button class="btn" style="width:100%;justify-content:center" type="submit">${isSignUp ? '注册' : '登录'}</button>
      </form>
      <p class="hint" style="text-align:center;margin-top:14px">${isSignUp ? '已经有账号了？' : '还没有账号？'}<a id="auth-switch">${isSignUp ? '去登录' : '注册一个'}</a></p>
    </div>`;
    $('#auth-switch').addEventListener('click', () => renderLogin(!isSignUp));
    $('#auth-form').addEventListener('submit', async e => {
      e.preventDefault();
      const f = e.target;
      await withBusy($('button[type=submit]', f), isSignUp ? '注册中…' : '登录中…', async () => {
        try {
          if (isSignUp) await DB.signUp(f.email.value, f.password.value, f.name.value, f.dept.value);
          else await DB.signIn(f.email.value, f.password.value);
          await start();
        } catch (err) { showErr(err); }
      });
    });
  }

  function renderNoAccess(acc) {
    $('#top').hidden = true;
    const msg = acc.reason === 'not_member'
      ? `你的账号（${S.email}）还没开通话术复盘系统，请让已开通的同事在「成员」页添加你的邮箱。`
      : acc.reason === 'tables_missing' ? '数据库还没初始化，请按 sales-script/SETUP.md 执行 setup.sql。' : acc.message;
    app().innerHTML = `<div class="auth card" style="text-align:center"><h1>还不能使用</h1><p class="note warn" style="margin:16px 0;text-align:left">${esc(msg)}</p>
      <div class="row" style="justify-content:center"><button class="btn plain" onclick="location.reload()">刷新</button><button class="btn plain" id="na-out">换个账号登录</button></div></div>`;
    $('#na-out').addEventListener('click', async () => { await DB.signOut(); location.reload(); });
  }

  // ======================================================================
  // 话术助手
  // ======================================================================
  function viewAssistant() {
    const alive = viewToken();
    const empty = !S.scripts.some(s => s.status !== 'archived');
    app().innerHTML = `
      <div class="page-head"><div><h1>话术助手</h1><p class="sub">把客户的话贴进来，优先匹配团队原话，需要时再用 AI 调整</p></div></div>
      ${empty ? '<div class="note warn" style="margin-bottom:14px">话术库还是空的，推荐会不太准。先去 <a href="#/import">导入话术文档</a> 吧。</div>' : ''}
      <div class="grid2">
        <div class="card">
          <div class="field"><label>客户说了什么 <b>*</b></label>
            <textarea id="as-msg" rows="3" placeholder="例：太贵了，我再考虑考虑"></textarea></div>
          <div class="field"><label>前面的聊天（选填，贴上会更准）</label>
            <textarea id="as-ctx" rows="5" placeholder="可以粘贴文字，也可以在这里 ⌘V / Ctrl+V 直接粘贴聊天截图"></textarea>
            <div class="row" style="margin-top:6px"><button class="btn plain sm" id="as-shot">上传截图识别</button><span class="hint" id="as-shot-st"></span></div>
            <input type="file" id="as-file" accept="image/*" multiple hidden></div>
          <div class="row">
            <span class="muted">渠道</span>
            <select id="as-ch" style="width:auto"><option value="all">不限</option>${channelOptions(['xiaoe', 'bilibili'])}</select>
            <span class="spacer"></span>
            <button class="btn plain" id="as-stop" hidden>停止生成</button>
            <button class="btn" id="as-go">生成推荐回答</button>
          </div>
          ${knowledgeScopeHtml('as-knowledge')}
          <p class="hint" id="as-progress" role="status" aria-live="polite"></p>
        </div>
        <div id="as-result"><div class="card empty" style="padding:48px 20px">推荐回答会显示在这里<br><span class="faint">⌘ / Ctrl + Enter 快速生成</span></div></div>
      </div>
      <h2>我最近的提问 <small>点一条查看推荐，也可以单独删除</small></h2>
      <div id="as-logs-error" class="note warn" style="margin-bottom:10px" hidden></div>
      <div class="card" id="as-logs"><span class="muted">加载中…</span></div>`;

    bindKnowledgeScope('as-knowledge');
    let logs = [], pending = false, requestId = 0, logsLoadId = 0, activeLogId = null;
    const deletingLogs = new Set(), deletedLogs = new Set();
    const go = async (forceGenerate = false) => {
      if (pending) return;
      const message = $('#as-msg').value.trim();
      if (!message) return toast('先填客户说了什么', true);
      const context = $('#as-ctx').value.trim(), channel = $('#as-ch').value;
      const knowledgeProject = $('#as-knowledge').value;
      const controller = new AbortController(), id = ++requestId;
      activeLogId = null;
      const current = () => alive() && id === requestId;
      S.assistantAbort = controller;
      pending = true;
      const button = $('#as-go'), stop = $('#as-stop'), progress = $('#as-progress'), resultBox = $('#as-result');
      const fields = $$('#as-msg, #as-ctx, #as-ch, #as-knowledge, #as-shot');
      fields.forEach(el => { el.disabled = true; });
      button.disabled = true; button.innerHTML = '<span class="spin"></span> 处理中…'; stop.hidden = false;
      progress.textContent = '正在匹配话术库…';
      resultBox.innerHTML = '<div class="card empty">正在查找适合这条消息的回复…</div>';
      let legacyNote = '';
      try {
        const result = await AI.suggest({ message, context, channel, scripts: S.scripts, cases: S.cases, bg: S.bg,
          knowledgeProject, forceGenerate, signal: controller.signal, onProgress(event) {
            if (!current() || controller.signal.aborted) return;
            if (event.type === 'stage') progress.textContent = event.message;
            if (event.type === 'status') progress.textContent = `正在使用 ${event.provider === 'deepseek' ? 'DeepSeek' : 'ChatGPT'} 生成…`;
            if (event.type === 'legacy') legacyNote = event.message;
            if (event.type === 'reset') {
              resultBox.innerHTML = '<div class="card empty">正在切换服务，重新生成回复…</div>';
              progress.textContent = event.message || '正在切换到 DeepSeek…';
            }
            if (event.type === 'preview' && event.replies) {
              resultBox.innerHTML = assistantPreviewHtml(event.replies);
            }
          } });
        if (!current()) return;
        if (controller.signal.aborted) throw Object.assign(new Error('已停止生成'), { name: 'AbortError' });
        renderSuggestion(resultBox, result, null, message, loadLogs, () => go(true));
        progress.textContent = legacyNote || (result.source === 'library' ? '已匹配话术库原话 · 未调用 AI' : '回复已完成');
        // 先展示完整结果，提问记录的写入不再增加用户等待。
        DB.createAiLog({ message: L.maskSensitive(message), context: L.maskSensitive(context), channel, result })
          .then(log => {
            if (!current()) return;
            if (deletedLogs.has(String(log.id))) {
              activeLogId = null;
              resultBox.innerHTML = '<div class="card empty">这条提问记录已删除</div>';
              progress.textContent = '';
              if ($('#as-msg').value === message) $('#as-msg').value = '';
              if ($('#as-ctx').value === context) $('#as-ctx').value = '';
              return;
            }
            activeLogId = log.id;
            renderSuggestionFeedback($('[data-feedback-slot]', resultBox), log, loadLogs);
            loadLogs();
          }).catch(e => {
            console.warn('记录提问失败', e);
            if (current()) progress.textContent += ' · 提问记录未保存，仍可复制本次回复';
          });
      } catch (e) {
        if (!current()) return;
        const cancelled = controller.signal.aborted || e.name === 'AbortError' || e.code === 'cancelled';
        progress.textContent = cancelled ? '已停止生成，可以修改问题后重试' : '生成未完成，请重试';
        resultBox.innerHTML = `<div class="card empty">${cancelled ? '已停止，本次草稿未保存' : esc(e.message || '生成中断，请重试')}</div>`;
        if (!cancelled) showErr(e);
      } finally {
        if (S.assistantAbort === controller) S.assistantAbort = null;
        pending = false;
        if (current()) {
          fields.forEach(el => { el.disabled = false; });
          button.disabled = false; button.textContent = '生成推荐回答'; stop.hidden = true;
        }
      }
    };
    $('#as-go').addEventListener('click', () => go());
    $('#as-stop').addEventListener('click', () => { if (S.assistantAbort) S.assistantAbort.abort(); });
    $$('#as-msg, #as-ctx').forEach(t => t.addEventListener('keydown', e => {
      if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) { e.preventDefault(); go(); }
    }));

    async function handleImages(files) {
      if (!files.length || pending) return;
      const st = $('#as-shot-st');
      st.innerHTML = '<span class="spin"></span> 正在识别截图…';
      try {
        const dlg = await AI.transcribeImages(await prepareImages(files), S.me);
        if (!alive() || pending) return;
        const ta = $('#as-ctx');
        ta.value = (ta.value.trim() ? ta.value.trim() + '\n' : '') + L.dialogueToText(dlg);
        // 最后一条客户消息自动填进「客户说了什么」
        const last = dlg.slice().reverse().find(x => x.role === 'customer');
        if (last && !$('#as-msg').value.trim()) $('#as-msg').value = last.text;
        st.textContent = `识别出 ${dlg.length} 条消息，检查一下没问题就可以生成了`;
      } catch (e) { if (alive()) { st.textContent = ''; showErr(e); } }
    }
    $('#as-shot').addEventListener('click', () => $('#as-file').click());
    $('#as-file').addEventListener('change', e => { handleImages([...e.target.files]); e.target.value = ''; });
    $('#as-ctx').addEventListener('paste', e => {
      const imgs = imagesFromClipboard(e);
      if (imgs.length) { e.preventDefault(); handleImages(imgs); }
    });

    async function loadLogs() {
      const id = ++logsLoadId;
      let rows;
      try {
        rows = await DB.listAiLogs({ createdBy: S.userId, limit: 15 });
      } catch (e) {
        if (alive() && id === logsLoadId) {
          const error = $('#as-logs-error');
          error.textContent = '提问记录读取失败，请刷新页面重试'; error.dataset.kind = 'load'; error.hidden = false;
        }
        return;
      }
      if (!alive() || id !== logsLoadId) return;
      logs = rows.filter(row => !deletedLogs.has(String(row.id)));
      const error = $('#as-logs-error');
      if (error.dataset.kind === 'load') error.hidden = true;
      drawLogs();
    }

    function drawLogs() {
      const box = $('#as-logs');
      if (!box || !alive()) return;
      box.innerHTML = logs.length ? logs.map((l, i) => `<div class="log-item assistant-log" data-log-id="${esc(l.id)}">
          <button type="button" class="log-open" data-log-open="${i}" ${deletingLogs.has(String(l.id)) ? 'disabled' : ''}>
            <span class="log-msg">${esc(l.message)}</span>
            <span class="faint">${esc(L.fmtDateTime(l.created_at))}${l.result && l.result.scene ? ' · ' + esc(l.result.scene) : ''}${l.feedback === 'good' ? ' · 👍 好用' : l.feedback === 'bad' ? ' · 👎 不好用' : ''}</span>
          </button>
          <button type="button" class="linkbtn danger log-delete" data-log-delete="${i}" aria-label="删除提问 ${esc(l.message)}" ${deletingLogs.has(String(l.id)) ? 'disabled' : ''}>${deletingLogs.has(String(l.id)) ? '处理中…' : '删除'}</button>
        </div>`).join('') : '<span class="muted">还没有提问记录</span>';
      $$('[data-log-open]', box).forEach(el => el.addEventListener('click', () => {
        if (pending) return;
        ++requestId;
        const l = logs[+el.dataset.logOpen];
        activeLogId = l.id;
        $('#as-progress').textContent = '';
        $('#as-ch').value = l.channel || 'all';
        const project = l.result && l.result.context_used && l.result.context_used.knowledge_project || '';
        const scope = $('#as-knowledge');
        scope.dataset.historyProject = project;
        if (![...scope.options].some(o => o.value === project)) scope.add(new Option(project + '（历史范围）', project));
        scope.value = project; S.knowledgeProject = project;
        $('#as-msg').value = l.message;
        $('#as-ctx').value = l.context || '';
        if (l.result) renderSuggestion($('#as-result'), l.result, l, l.message, loadLogs, () => go(true));
        else $('#as-result').innerHTML = '<div class="card empty">这条提问没有保存回复</div>';
        window.scrollTo({ top: 0, behavior: 'smooth' });
      }));
      $$('[data-log-delete]', box).forEach(button => button.addEventListener('click', event => {
        event.stopPropagation();
        deleteLog(logs[+button.dataset.logDelete]);
      }));
    }

    async function deleteLog(log) {
      const key = String(log.id);
      if (deletingLogs.has(key)) return;
      deletingLogs.add(key); drawLogs();
      try {
        const excerpt = Array.from(log.message || '').slice(0, 100).join('');
        const confirmed = await confirmBox('删除这条提问？', `「${excerpt}」\n删除后，这条提问、推荐回复和反馈都无法恢复。已保存的话术、案例和复盘内容会保留。`, '删除提问', true);
        if (!confirmed || !alive()) return;
        const error = $('#as-logs-error'); error.hidden = true;
        await DB.deleteAiLog(log.id);
        deletedLogs.add(key);
        if (!alive()) return;
        logs = logs.filter(row => String(row.id) !== key);
        if (String(activeLogId) === key) {
          activeLogId = null; ++requestId;
          $('#as-result').innerHTML = '<div class="card empty">这条提问记录已删除</div>';
          $('#as-progress').textContent = '';
          if ($('#as-msg').value === log.message) $('#as-msg').value = '';
          if ($('#as-ctx').value === (log.context || '')) $('#as-ctx').value = '';
        }
        drawLogs(); toast('提问记录已删除');
        await loadLogs();
      } catch (e) {
        if (!alive()) return;
        if (e.code === 'ai_log_delete_missing') {
          const error = $('#as-logs-error');
          error.innerHTML = '删除功能还需要更新数据库权限。在 Supabase 的 SQL Editor 执行 <a href="ai-log-delete-setup.sql" target="_blank" rel="noopener">ai-log-delete-setup.sql</a> 后重试。<a href="SETUP.md" target="_blank" rel="noopener">查看步骤</a>';
          error.dataset.kind = 'migration'; error.hidden = false;
        } else showErr(e);
      } finally {
        deletingLogs.delete(key);
        if (alive()) drawLogs();
      }
    }
    loadLogs();
  }

  function prepareCustomerReply(text, notes) {
    const helper = window.SalesReplyContent;
    if (!helper) return { text: '', internal_notes: ['回复检查组件未加载，请刷新页面重试。'], withheld_messages: [String(text || '')] };
    return helper.prepare(text, notes);
  }

  function prepareAiScriptPrefill(script, defaults = {}) {
    const prefill = Object.assign({}, defaults, pickContent(script || {}));
    const prepared = prepareCustomerReply(prefill.reply || '', script && script.internal_notes);
    if (prepared.withheld_messages.length) return null;
    if (prepared.internal_notes.length) prefill.tips = [prefill.tips, ...prepared.internal_notes].filter(Boolean).join('\n');
    return prefill;
  }

  function salesInternalHtml(notes, extra = '') {
    const clean = prepareCustomerReply('', notes).internal_notes;
    if (!clean.length && !extra) return '';
    return `<div class="sales-internal"><b>销售内部提示 <span class="faint">· 不会复制</span></b>
      ${clean.length ? `<ul>${clean.map(note => `<li>${esc(note)}</li>`).join('')}</ul>` : ''}${extra}</div>`;
  }

  function assistantPreviewHtml(replies) {
    const visible = (replies || []).flatMap(reply => {
      const raw = String(reply.text || '');
      // 一条消息写完后再显示，避免半截的来源说明或占位备注进入客户气泡。
      const end = Math.max(raw.lastIndexOf('\n'), raw.lastIndexOf('\r'));
      const complete = reply.text_complete === true ? raw : raw.slice(0, end + 1);
      if (!complete.trim() || prepareCustomerReply(raw).withheld_messages.length) return [];
      return [`<div class="reply-card"><b>${esc(reply.label || '回复草稿')}</b>${replyMessagesHtml(complete)}</div>`];
    });
    return `<div class="card streaming-draft"><p class="hint"><span class="spin"></span> 正在生成，完成后可复制和反馈</p>
      ${visible.length ? visible.join('') : '<p class="muted">正在整理下一条客户消息…</p>'}</div>`;
  }

  function renderSuggestion(el, r, log, message, onFeedback, onAdapt) {
    const replies = (r.replies || []).map(reply => ({ ...reply, prepared: prepareCustomerReply(reply.text, reply.internal_notes) }));
    const blocked = replies.some(reply => reply.prepared.withheld_messages.length);
    const refs = ids => ids.map(id => {
      const snapshot = ((r.context_used && r.context_used.style_scripts) || []).find(s => Number(s.id) === Number(id));
      return scriptRef(id, snapshot && snapshot.version);
    }).join('、');
    el.innerHTML = `<div class="card">
      ${r.demo ? '<div class="note warn" style="margin-bottom:10px">演示模式，当前显示本地话术，未调用 AI。</div>' : ''}
      ${r.source === 'library' ? `<div class="note library-match">已匹配话术库原话 · 未调用 AI${r.matched_script ? `<div>${scriptRef(r.matched_script.id, r.matched_script.version)}</div>` : ''}${onAdapt ? '<button class="btn plain sm" data-adapt style="margin-top:8px">结合当前聊天调整</button>' : ''}</div>` : ''}
      ${contextUsedHtml(r)}
      <div class="row" style="margin-bottom:12px"><span class="tag">场景：${esc(r.scene || '未识别')}</span><span class="muted">${esc(r.intent)}</span></div>
      ${salesInternalHtml(r.internal_notes)}
      ${replies.length ? '<p class="hint" style="margin-bottom:10px">客户消息可以逐条复制，内部提示供销售参考。</p>' : ''}
      ${replies.map((x, i) => {
        const withheld = x.prepared.withheld_messages.length > 0;
        const explanation = `${x.why || x.based_on && x.based_on.length ? `<div class="why">${esc(x.why)}${x.based_on && x.based_on.length ? ` · 依据：${refs(x.based_on)}` : ''}</div>` : ''}${knowledgeCitationsHtml(x, r.context_used)}`;
        return `<div class="reply-card ${withheld ? 'reply-needs-revision' : ''}">
          <div class="reply-head"><b>${esc(x.label || '回复 ' + (i + 1))}</b>${withheld ? '' : `<button type="button" class="btn plain sm" data-copy="${i}">复制全部</button>`}</div>
          ${withheld ? '<div class="note warn">这份回复混入了内部说明或待核实内容，请重新生成后再发送。</div>' : `<div class="customer-reply">${replyMessagesHtml(x.prepared.text, { copy: true })}</div>`}
          ${salesInternalHtml(x.prepared.internal_notes, explanation + (withheld ? `<details class="original-reply"><summary>查看需要整理的原稿</summary><div class="pre">${esc(x.text)}</div></details>` : ''))}
        </div>`;
      }).join('')}
      ${blocked && onAdapt ? '<button type="button" class="btn sm" data-regenerate style="margin-top:12px">重新生成客户话术</button>' : ''}
      ${r.gap ? `<div class="note warn" style="margin-top:12px">${esc(r.gap_note || '话术库里还没有针对这个问题的话术。')}
          <div style="margin-top:8px"><button class="btn sm" id="sg-add">把这个问题补进话术库</button></div></div>` : ''}
      <div data-feedback-slot></div>
    </div>`;
    $$('[data-copy]', el).forEach(b => b.addEventListener('click', () => {
      const prepared = replies[+b.dataset.copy].prepared;
      if (!prepared.withheld_messages.length) copyText(prepared.text);
    }));
    const add = $('#sg-add', el);
    if (add) add.addEventListener('click', () => {
      const prefill = prepareAiScriptPrefill(r.new_script, { customer_says: message || '' });
      if (!prefill) return toast('这份草稿含有内部说明，请先重新生成客户话术', true);
      openScriptEditor(null, { prefill, change_note: '话术助手发现话术库里缺这个场景' });
    });
    $$('[data-adapt], [data-regenerate]', el).forEach(button => button.addEventListener('click', onAdapt));
    renderSuggestionFeedback($('[data-feedback-slot]', el), log, onFeedback);
  }

  function renderSuggestionFeedback(el, log, onFeedback) {
    if (!el || !log || !log.id) return;
    el.innerHTML = `<div style="margin-top:14px;border-top:1px solid var(--line);padding-top:12px">
      <div class="row"><span class="muted">这次推荐好用吗？</span>
        <button class="chip ${log.feedback === 'good' ? 'on' : ''}" data-fb="good">👍 好用</button>
        <button class="chip ${log.feedback === 'bad' ? 'on' : ''}" data-fb="bad">👎 不好用</button></div>
      <div class="row" id="fb-note" style="margin-top:8px" ${log.feedback === 'bad' ? '' : 'hidden'}>
        <input type="text" id="fb-text" placeholder="哪里不好用？复盘会上会看到（选填）" value="${esc(log.feedback_note || '')}" style="flex:1">
        <button class="btn sm" id="fb-save">提交</button></div></div>`;
    let savingFeedback = false;
    const setFb = async (fb, note) => {
      if (savingFeedback) return;
      savingFeedback = true;
      const controls = $$('button, input', el);
      controls.forEach(control => { control.disabled = true; });
      try {
        await DB.setAiFeedback(log.id, fb, note);
        log.feedback = fb; log.feedback_note = note || '';
        $$('[data-fb]', el).forEach(b => b.classList.toggle('on', b.dataset.fb === fb));
        toast(fb === 'good' ? '收到，谢谢反馈' : '已记录，复盘会上会看到');
        if (onFeedback) onFeedback();
      } catch (e) { showErr(e); } finally {
        savingFeedback = false;
        controls.forEach(control => { control.disabled = false; });
      }
    };
    $$('[data-fb]', el).forEach(b => b.addEventListener('click', () => {
      const fb = b.dataset.fb;
      $('#fb-note', el).hidden = fb !== 'bad';
      setFb(fb, fb === 'bad' ? $('#fb-text', el).value.trim() : '');
    }));
    const fbSave = $('#fb-save', el);
    if (fbSave) fbSave.addEventListener('click', () => setFb('bad', $('#fb-text', el).value.trim()));
  }

  const pickContent = obj => { const o = {}; for (const k of CONTENT_FIELDS) if (obj && obj[k] != null) o[k] = obj[k]; return o; };

  // ======================================================================
  // 话术库
  // ======================================================================
  function viewLibrary() {
    const f = S.libFilter || (S.libFilter = { q: '', ch: '', arch: false });
    app().innerHTML = `
      <div class="page-head"><div><h1>话术库</h1><p class="sub">按阶段和场景整理，每次修改都会留下版本记录</p></div>
        <div class="row"><a class="btn plain" href="#/background">给 AI 的指引</a><a class="btn plain" href="#/knowledge">训练营知识库</a><a class="btn ghost" href="#/import">导入话术文档</a><button class="btn" id="lib-new">+ 新增话术</button></div></div>
      <div class="toolbar">
        <input type="search" id="lib-q" placeholder="搜索场景、客户说法、话术内容" value="${esc(f.q)}">
        <div class="chips" id="lib-ch">${[['', '全部渠道'], ['xiaoe', '小鹅通'], ['bilibili', 'B站']].map(([k, t]) => `<button class="chip ${f.ch === k ? 'on' : ''}" data-ch="${k}">${t}</button>`).join('')}</div>
        <label class="muted"><input type="checkbox" id="lib-arch" ${f.arch ? 'checked' : ''}> 显示已停用</label>
      </div>
      <div id="lib-body"></div>`;

    const draw = () => {
      const body = $('#lib-body');
      if (!S.scripts.length) {
        body.innerHTML = `<div class="card empty">话术库还是空的<br>可以把现有的话术文档直接粘贴进来，AI 会帮你拆成一条条话术<br><a class="btn" href="#/import" style="margin-top:8px">导入话术文档</a></div>`;
        return;
      }
      let list = S.scripts.filter(s => f.arch || s.status !== 'archived');
      if (f.ch) list = list.filter(s => L.scriptChannelMatch(s, f.ch));
      list = L.searchScripts(list, f.q);
      if (!list.length) { body.innerHTML = '<div class="empty">没有找到匹配的话术</div>'; return; }
      const usage = L.scriptUsage(S.cases);
      const groups = L.groupScriptsByStage(list);
      body.innerHTML = (groups.length > 1 ? `<div class="stage-nav">${groups.map((g, i) => `<button class="chip" data-jump="${i}">${esc(g.stage)} · ${g.scripts.length}</button>`).join('')}</div>` : '')
        + groups.map((g, i) => `<h2 id="stage-${i}">${esc(g.stage)} <small>${g.scripts.length} 条</small></h2>
          <div class="script-grid">${g.scripts.map(s => scriptCard(s, usage.get(s.id))).join('')}</div>`).join('');
      $$('[data-jump]', body).forEach(b => b.addEventListener('click', () => $('#stage-' + b.dataset.jump).scrollIntoView({ behavior: 'smooth' })));
    };
    draw();
    $('#lib-q').addEventListener('input', e => { f.q = e.target.value; draw(); });
    $$('#lib-ch [data-ch]').forEach(b => b.addEventListener('click', () => {
      f.ch = b.dataset.ch;
      $$('#lib-ch .chip').forEach(x => x.classList.toggle('on', x === b));
      draw();
    }));
    $('#lib-arch').addEventListener('change', e => { f.arch = e.target.checked; draw(); });
    $('#lib-new').addEventListener('click', async () => {
      const saved = await openScriptEditor(null, {});
      if (saved) draw();
    });
  }

  function scriptCard(s, u) {
    const recent = s.version > 1 && Date.now() - new Date(s.updated_at).getTime() < 7 * 864e5;
    return `<a class="script-card ${s.status === 'archived' ? 'archived' : ''}" href="#/script/${s.id}">
      <div class="sc-head"><b>${esc(s.scene)}</b>${s.channel !== 'all' ? chTag(s.channel) : ''}<span class="ver">v${s.version}</span>${recent ? '<span class="badge new">最近更新</span>' : ''}${s.status === 'archived' ? '<span class="tag gray">已停用</span>' : ''}</div>
      ${s.customer_says ? `<div class="sc-says">客户：${esc(s.customer_says.split('\n')[0])}</div>` : ''}
      ${replyMessagesHtml(s.reply, { compact: true })}
      <div class="sc-meta">${esc(L.fmtDateTime(s.updated_at))} ${esc(s.updated_by_name)}${s.version > 1 ? ' 改过' : ' 添加'}${u ? ' · ' + L.usageText(u) : ''}</div>
    </a>`;
  }

  function usageTable(u) {
    if (!u) return '<div class="faint">还没有案例用到这条话术。提交案例时标上「用到的话术」，这里就能看到每个版本的成交情况。</div>';
    const rows = Object.keys(u.byVersion).map(Number).sort((a, b) => b - a)
      .map(v => `<tr><td>${v ? 'v' + v : '未记录版本'}</td><td>${u.byVersion[v].won + u.byVersion[v].lost}</td><td style="color:var(--won)">${u.byVersion[v].won}</td><td style="color:var(--lost)">${u.byVersion[v].lost}</td></tr>`).join('');
    return `<table class="mini"><tr><th>版本</th><th>用到</th><th>成交</th><th>流失</th></tr>${rows}</table>
      <div class="faint" style="margin-top:4px">案例数少的时候只能参考，别只看比例</div>`;
  }

  async function viewScript(id) {
    const alive = viewToken();
    const s = scriptById(id);
    if (!s) return notFound('这条话术不存在');
    app().innerHTML = `
      <a class="back" href="#/library">← 话术库</a>
      <div class="grid-side">
        <div class="card">
          <div class="row between" style="align-items:flex-start">
            <div><div class="faint">${esc(s.stage || '其他')}</div>
              <h1 class="row" style="gap:10px">${esc(s.scene)} <span class="ver" style="font-size:15px">v${s.version}</span>${chTag(s.channel)}${s.status === 'archived' ? '<span class="tag gray">已停用</span>' : ''}</h1></div>
            <div class="row"><button class="btn plain sm" id="sc-arch">${s.status === 'archived' ? '恢复启用' : '停用'}</button><button class="btn sm" id="sc-edit">修改话术</button></div>
          </div>
          <div class="detail-block" style="margin-top:14px"><div class="label">客户常说</div>${s.customer_says ? `<div class="pre">${esc(s.customer_says)}</div>` : '<span class="faint">没填</span>'}</div>
          <div class="detail-block"><div class="row between"><div class="label">发送消息</div><button type="button" class="linkbtn" id="sc-copy">复制全部</button></div>${replyMessagesHtml(s.reply, { copy: true })}</div>
          ${s.tips ? `<div class="detail-block"><div class="label">要点 / 禁忌</div><div class="pre">${esc(s.tips)}</div></div>` : ''}
          <div class="detail-block"><div class="label">实际效果（按案例里记录的话术版本统计）</div>${usageTable(L.scriptUsage(S.cases).get(s.id))}</div>
        </div>
        <div class="card"><h3>修改记录</h3><div id="sc-history"><span class="muted">加载中…</span></div></div>
      </div>`;
    $('#sc-copy').addEventListener('click', () => copyText(s.reply));
    $('#sc-edit').addEventListener('click', async () => { if (await openScriptEditor(s, {})) viewScript(id); });
    $('#sc-arch').addEventListener('click', async () => {
      const archiving = s.status !== 'archived';
      if (archiving && !(await confirmBox('停用这条话术？', '停用后话术助手不会再推荐它，历史记录都保留，随时可以恢复。', '停用'))) return;
      try {
        const saved = await DB.saveScript(Object.assign(clone(s), {
          status: archiving ? 'archived' : 'active', change_note: archiving ? '停用' : '恢复启用', change_review_id: null, change_case_ids: [],
        }));
        upsert(S.scripts, saved);
        viewScript(id);
      } catch (e) { e.code === 'conflict' ? (toast('刚被别人改过，已刷新', true), await reloadScripts(), viewScript(id)) : showErr(e); }
    });

    let versions;
    try { versions = await DB.listScriptVersions(s.id); } catch (e) {
      if (alive()) $('#sc-history').innerHTML = `<div class="note err">${esc(e.message)}</div>`;
      return;
    }
    const box = $('#sc-history');
    if (!box || !alive()) return;
    if (!versions.length) { box.innerHTML = '<span class="faint">没有记录</span>'; return; }
    box.innerHTML = `<ul class="timeline">${versions.map((v, i) => {
      const r = v.review_id ? reviewById(v.review_id) : null;
      const links = [
        r ? `<a href="#/review/${r.week}">复盘会 ${esc(L.weekLabel(r.week))}</a>` : '',
        ...(v.case_ids || []).map(cid => `<a href="#/case/${cid}">案例#${cid}</a>`),
      ].filter(Boolean).join(' · ');
      return `<li>
        <div class="tl-head"><b>v${v.version}</b><span>${esc(L.fmtDateTime(v.changed_at))}</span><span>${esc(v.changed_by_name)}</span>${v.action === 'insert' ? '<span class="tag gray">新建</span>' : ''}</div>
        ${v.change_note ? `<div class="tl-note">${esc(v.change_note)}</div>` : ''}
        ${links ? `<div class="faint">${links}</div>` : ''}
        <div class="tl-actions"><button class="linkbtn" data-diff="${i}">${i < versions.length - 1 ? '看改了什么' : '看这一版'}</button>${i > 0 ? `<button class="linkbtn" data-restore="${i}">恢复到这一版</button>` : ''}</div>
        <div class="tl-diff" hidden></div>
      </li>`;
    }).join('')}</ul>`;
    $$('[data-diff]', box).forEach(b => b.addEventListener('click', () => {
      const i = +b.dataset.diff;
      const pane = b.closest('li').querySelector('.tl-diff');
      if (!pane.hidden) { pane.hidden = true; return; }
      pane.innerHTML = versionDiffHtml(versions[i + 1] && versions[i + 1].snapshot, versions[i].snapshot);
      pane.hidden = false;
    }));
    $$('[data-restore]', box).forEach(b => b.addEventListener('click', async () => {
      const v = versions[+b.dataset.restore];
      const saved = await openScriptEditor(s, { prefill: pickContent(v.snapshot), change_note: `恢复到 v${v.version}` });
      if (saved) viewScript(id);
    }));
  }

  function diffHtml(a, b) {
    return L.diffText(a, b).map(p => p.type === 'same' ? esc(p.text) : p.type === 'add' ? `<ins>${esc(p.text)}</ins>` : `<del>${esc(p.text)}</del>`).join('');
  }
  function versionDiffHtml(prev, cur) {
    const FIELDS = [['scene', '场景'], ['stage', '阶段'], ['customer_says', '客户常说'], ['reply', '发送消息'], ['tips', '要点 / 禁忌'], ['channel', '适用渠道'], ['status', '状态']];
    const show = (k, v) => k === 'channel' ? (L.CHANNELS[v] || L.CHANNELS.all).label : k === 'status' ? (v === 'archived' ? '已停用' : '启用') : (v || '');
    if (!prev) {
      return FIELDS.filter(([k]) => cur[k]).map(([k, t]) => `<div class="diff-label">${t}</div>${k === 'reply' ? replyMessagesHtml(cur[k]) : `<div class="diff">${esc(show(k, cur[k]))}</div>`}`).join('');
    }
    const changed = FIELDS.filter(([k]) => String(prev[k] || '') !== String(cur[k] || ''));
    if (!changed.length) return '<div class="faint" style="margin-top:6px">内容没变</div>';
    return changed.map(([k, t]) => `<div class="diff-label">${t}</div>${k === 'reply' ? replyMessagesHtml(cur[k]) + replyChangesHtml(prev[k], cur[k]) : `<div class="diff">${['customer_says', 'tips'].includes(k) ? diffHtml(prev[k], cur[k]) : `<del>${esc(show(k, prev[k]))}</del> → <ins>${esc(show(k, cur[k]))}</ins>`}</div>`}`).join('');
  }

  function replyChangesHtml(before, after) {
    return `<details class="message-changes"><summary>查看文字改动 · ${L.splitReplyMessages(before).length} 条消息 → ${L.splitReplyMessages(after).length} 条消息</summary><div class="diff">${diffHtml(before, after)}</div></details>`;
  }

  async function reloadScripts() { S.scripts = await DB.listScripts(); }

  // 话术编辑弹窗：新增 / 修改 / 恢复 / 复盘会上采纳建议，全走这里
  // opts: { prefill, change_note, review_id, case_ids, hint, allowUnchanged }
  function openScriptEditor(script, opts) {
    opts = opts || {};
    const d = Object.assign(script ? clone(script) : L.emptyScript(), pickContent(opts.prefill || {}));
    d.change_note = opts.change_note || '';
    d.change_review_id = opts.review_id || null;
    d.change_case_ids = uniq((opts.case_ids || []).map(Number));
    const start = JSON.stringify(pickContent(d));
    const caseLabel = c => `#${c.id} ${L.RESULTS[c.result]} · ${c.seller}${c.scenes && c.scenes.length ? ' · ' + c.scenes.join('、') : ''}`;

    return new Promise(resolve => {
      const m = openModal(`
        <h2>${script ? `修改话术 <small>「${esc(script.scene)}」当前 v${script.version}</small>` : '新增话术'}</h2>
        ${opts.hint ? `<div class="note info pre" style="margin-bottom:14px">${esc(opts.hint)}</div>` : ''}
        <form id="se-form">
          <div class="inline-fields">
            <div class="field"><label>阶段</label><input type="text" name="stage" list="se-stages" value="${esc(d.stage)}" placeholder="如：直播后追单"></div>
            <div class="field"><label>场景 / 客户异议 <b>*</b></label><input type="text" name="scene" list="se-scenes" value="${esc(d.scene)}" placeholder="如：嫌贵"></div>
            <div class="field"><label>适用渠道</label><select name="channel">${channelOptions(L.SCRIPT_CHANNELS, d.channel)}</select></div>
          </div>
          <div class="field"><label>客户常说（一行一句）</label><textarea name="customer_says" rows="2">${esc(d.customer_says)}</textarea></div>
          <div class="field"><label for="se-reply">发送消息 <b>*</b></label><textarea id="se-reply" name="reply" rows="5" aria-describedby="se-reply-hint">${esc(d.reply)}</textarea>
            <div class="hint" id="se-reply-hint">每行是一条消息，按回车开始下一条。空行会忽略，自动折行仍算同一条。</div>
            <div class="message-preview"><div class="label">发送预览</div><div id="se-message-preview">${replyMessagesHtml(d.reply)}</div></div><div id="se-diff"></div></div>
          <div class="field"><label>要点 / 禁忌</label><textarea name="tips" rows="2">${esc(d.tips)}</textarea></div>
          <div class="card" style="background:var(--bg);border:none">
            <div class="field"><label>${script ? '这次为什么改' : '来源 / 说明'}（会记在修改记录里）</label>
              <textarea name="change_note" rows="2" placeholder="${script ? '例：复盘会上发现客户嫌贵时直接给优惠，容易被当成催单，改成先问清楚顾虑' : '例：复盘会上新增'}">${esc(d.change_note)}</textarea></div>
            <div class="field"><label>在哪次复盘会上定的</label><select name="change_review_id"><option value="">（不关联）</option>
              ${S.reviews.map(r => `<option value="${r.id}" ${Number(d.change_review_id) === r.id ? 'selected' : ''}>复盘会 ${esc(L.weekLabel(r.week))}</option>`).join('')}</select></div>
            <div class="field" style="margin-bottom:0"><label>依据的案例</label><div class="chips" id="se-cases"></div>
              <select id="se-case-add" style="margin-top:6px"><option value="">+ 关联一个案例</option>${S.cases.slice(0, 80).map(c => `<option value="${c.id}">${esc(caseLabel(c))}</option>`).join('')}</select></div>
          </div>
          <div id="se-err" style="margin-top:10px"></div>
          <div class="modal-actions"><button type="button" class="btn plain" data-close>取消</button><button type="submit" class="btn" id="se-save">${opts.allowUnchanged ? '确认采纳' : script ? `保存为 v${script.version + 1}` : '保存'}</button></div>
        </form>
        <datalist id="se-stages">${allStages().map(s => `<option value="${esc(s)}">`).join('')}</datalist>
        <datalist id="se-scenes">${allScenes().map(s => `<option value="${esc(s)}">`).join('')}</datalist>`, {
        onCancel: async () => {
          collect();
          if (JSON.stringify(pickContent(d)) !== start && !(await confirmBox('放弃这次修改？', '', '放弃', true))) return;
          m.close(); resolve(null);
        },
      });
      const form = $('#se-form', m.el);
      // textarea 会规范化换行；未修改回复时保留原文，避免凭空生成新版本。
      const originalReply = d.reply;
      form.reply.value = originalReply;
      const initialReplyInput = form.reply.value;
      const readReply = () => form.reply.value === initialReplyInput ? originalReply : form.reply.value;
      const collect = () => {
        for (const k of CONTENT_FIELDS) d[k] = k === 'reply' ? readReply() : form[k].value;
        d.change_note = form.change_note.value.trim();
        d.change_review_id = form.change_review_id.value ? Number(form.change_review_id.value) : null;
      };
      const drawCases = () => {
        $('#se-cases', m.el).innerHTML = d.change_case_ids.length
          ? d.change_case_ids.map(id => { const c = caseById(id); return `<span class="chip-x">${esc(c ? caseLabel(c) : '#' + id)}<button type="button" data-rm="${id}">×</button></span>`; }).join('')
          : '<span class="faint">没有关联案例</span>';
        $$('[data-rm]', m.el).forEach(b => b.addEventListener('click', () => { d.change_case_ids = d.change_case_ids.filter(x => x !== +b.dataset.rm); drawCases(); }));
      };
      drawCases();
      $('#se-case-add', m.el).addEventListener('change', e => {
        const id = Number(e.target.value);
        if (id && !d.change_case_ids.includes(id)) d.change_case_ids.push(id);
        e.target.value = '';
        drawCases();
      });
      const drawDiff = () => {
        const reply = readReply();
        $('#se-message-preview', m.el).innerHTML = replyMessagesHtml(reply);
        const box = $('#se-diff', m.el);
        if (!script || reply === script.reply) { box.innerHTML = ''; return; }
        box.innerHTML = `<div class="diff-label">与当前 v${script.version} 对比</div>${replyChangesHtml(script.reply, reply)}`;
      };
      form.reply.addEventListener('input', drawDiff);
      drawDiff();
      $$('textarea', m.el).forEach(t => { autosize(t); t.addEventListener('input', () => autosize(t)); });

      form.addEventListener('submit', async e => {
        e.preventDefault();
        collect();
        const errs = L.validateScript(d);
        if (errs.length) { $('#se-err', m.el).innerHTML = `<div class="note err">${errs.map(esc).join('<br>')}</div>`; return; }
        if (script && JSON.stringify(pickContent(d)) === JSON.stringify(pickContent(script))) {
          if (opts.allowUnchanged) { m.close(); resolve(script); return; }
          $('#se-err', m.el).innerHTML = '<div class="note warn">内容没有变化</div>'; return;
        }
        await withBusy($('#se-save', m.el), '保存中…', async () => {
          try {
            const saved = await DB.saveScript(d);
            upsert(S.scripts, saved);
            toast(script ? `已保存为 v${saved.version}` : '已添加到话术库');
            m.close(); resolve(saved);
          } catch (err) {
            if (err.code !== 'conflict') { $('#se-err', m.el).innerHTML = `<div class="note err">${esc(err.message)}</div>`; return; }
            const pick = await choiceBox('这条话术刚被别人改过', '可以先看看别人改了什么，也可以直接用你的版本覆盖（别人的改动在修改记录里还能找到）。',
              [{ label: '看看别人改了什么', value: 'reload', cls: 'plain' }, { label: '用我的版本覆盖', value: 'force', cls: 'danger' }]);
            await reloadScripts();
            if (pick === 'force') {
              const latest = scriptById(d.id);
              d.version = latest ? latest.version : d.version;
              try {
                const saved = await DB.saveScript(d);
                upsert(S.scripts, saved);
                toast(`已保存为 v${saved.version}`);
                m.close(); resolve(saved);
              } catch (err2) { $('#se-err', m.el).innerHTML = `<div class="note err">${esc(err2.message)}</div>`; }
            } else {
              m.close(); resolve(null);
              location.hash = '#/script/' + d.id;
            }
          }
        });
      });
    });
  }

  // ---------- 导入话术文档 ----------
  function viewImport() {
    const alive = viewToken();
    let items = [];
    let importing = false;
    let importNotice = '';
    app().innerHTML = `
      <a class="back" href="#/library">← 话术库</a>
      <div class="page-head"><div><h1>导入话术文档</h1><p class="sub">把现有的话术文档整篇复制粘贴进来，AI 会按场景拆成一条条话术，导入前可以逐条检查和修改</p></div></div>
      <div class="card">
        <textarea id="im-text" rows="12" placeholder="直接粘贴整篇文档（Word、飞书、语雀里全选复制即可）"></textarea>
        <div class="row" style="margin-top:10px"><button class="btn" id="im-go">AI 拆分</button><span class="hint" id="im-st">文档长的话会分段处理，每段大约要半分钟</span></div>
      </div>
      <div id="im-result"></div>`;
    $('#im-text').addEventListener('input', () => markDirty('import'));

    $('#im-go').addEventListener('click', () => withBusy($('#im-go'), '拆分中…', async () => {
      const text = $('#im-text').value.trim();
      if (!text) return toast('先粘贴文档内容', true);
      // 也支持直接粘贴整理好的 JSON
      try {
        const j = JSON.parse(text);
        const arr = Array.isArray(j) ? j : j && j.items;
        if (Array.isArray(arr)) { items = arr.map(L.normImportedScript).filter(s => s.scene && s.reply); return drawItems(); }
      } catch (e) { /* 不是 JSON，交给 AI */ }
      const chunks = L.splitDoc(text, 3000);
      const results = new Array(chunks.length);
      const failed = [];
      let done = 0, next = 0;
      const st = $('#im-st');
      st.textContent = `共 ${chunks.length} 段，处理中…`;
      const worker = async () => {
        while (next < chunks.length) {
          const i = next++;
          try { results[i] = await AI.importChunk(chunks[i], allStages()); } catch (e) { failed.push({ i, msg: e.message }); }
          done++;
          st.textContent = `已处理 ${done} / ${chunks.length} 段`;
        }
      };
      await Promise.all([worker(), worker(), worker()]);
      items = results.filter(Boolean).flat();
      st.textContent = failed.length ? `有 ${failed.length} 段失败：${failed[0].msg}（可以把失败的部分单独再粘贴一次）` : `拆分完成，共 ${items.length} 条`;
      drawItems();
    }));

    function drawItems() {
      if (!alive()) return;
      const box = $('#im-result');
      if (!items.length) { box.innerHTML = '<div class="card empty">没有拆出话术，检查一下粘贴的内容？</div>'; return; }
      const exists = new Set(S.scripts.map(s => (s.stage + '|' + s.scene).trim()));
      items.forEach(it => { it._dup = exists.has((it.stage + '|' + it.scene).trim()); if (it._on === undefined) it._on = !it._dup; });
      box.innerHTML = `<div class="card" style="margin-top:14px">
        ${importNotice ? `<div class="note warn" style="margin-bottom:12px">${esc(importNotice)}</div>` : ''}
        <div class="row between"><h3 style="margin:0">拆出 ${items.length} 条话术 <small>检查一下阶段和场景名，话术内容可以直接改</small></h3>
          <label class="muted"><input type="checkbox" id="im-all"> 全选</label></div>
        ${items.map((it, i) => `<div style="border-top:1px solid var(--line);margin-top:12px;padding-top:12px" data-i="${i}">
          <div class="row">
            <input type="checkbox" data-k="_on" aria-label="导入这条话术" ${it._on ? 'checked' : ''}>
            <label class="import-field"><span class="label">销售阶段</span><input type="text" data-k="stage" value="${esc(it.stage)}" list="im-stages" style="width:150px" placeholder="阶段"></label>
            <label class="import-field"><span class="label">场景 / 客户异议</span><input type="text" data-k="scene" value="${esc(it.scene)}" style="width:200px" placeholder="场景"></label>
            <label class="import-field"><span class="label">适用渠道</span><select data-k="channel" style="width:auto">${channelOptions(L.SCRIPT_CHANNELS, it.channel)}</select></label>
            ${it._dup ? '<span class="tag gray">话术库里已有同名场景，默认不导入</span>' : ''}
          </div>
          <div class="grid2" style="margin-top:8px">
            <div><label class="label" for="im-reply-${i}">发送消息</label><textarea id="im-reply-${i}" data-k="reply" rows="4" aria-describedby="im-reply-hint-${i}">${esc(it.reply)}</textarea>
              <div class="hint" id="im-reply-hint-${i}">每行是一条消息，按回车开始下一条。空行会忽略，自动折行仍算同一条。</div>
              <div class="message-preview"><div class="label">发送预览</div><div data-message-preview>${replyMessagesHtml(it.reply)}</div></div></div>
            <div><div class="label">客户常说</div><textarea data-k="customer_says" rows="2">${esc(it.customer_says)}</textarea>
              <div class="label" style="margin-top:6px">要点 / 禁忌</div><textarea data-k="tips" rows="1">${esc(it.tips)}</textarea></div>
          </div>
        </div>`).join('')}
        <datalist id="im-stages">${allStages().map(s => `<option value="${esc(s)}">`).join('')}</datalist>
        <div class="row" style="margin-top:16px"><span class="spacer"></span><button class="btn" id="im-save"></button></div>
      </div>`;
      // 通过 value 赋值，避免 HTML 解析吞掉回复开头的第一个换行。
      $$('textarea[data-k="reply"]', box).forEach(el => {
        el.value = items[+el.closest('[data-i]').dataset.i].reply;
      });
      const count = () => { const n = items.filter(x => x._on).length; $('#im-save').textContent = `导入选中的 ${n} 条`; $('#im-save').disabled = !n; $('#im-all').checked = n === items.length; };
      count();
      $$('[data-k]', box).forEach(el => el.addEventListener(el.type === 'checkbox' ? 'change' : 'input', () => {
        const it = items[+el.closest('[data-i]').dataset.i];
        it[el.dataset.k] = el.type === 'checkbox' ? el.checked : el.value;
        if (el.dataset.k === 'reply') {
          $('[data-message-preview]', el.closest('[data-i]')).innerHTML = replyMessagesHtml(el.value);
          autosize(el);
        }
        markDirty('import');
        count();
      }));
      $$('textarea', box).forEach(autosize);
      $('#im-all').addEventListener('change', e => { items.forEach(x => { x._on = e.target.checked; }); markDirty('import'); drawItems(); });
      $('#im-save').addEventListener('click', () => withBusy($('#im-save'), '导入中…', async () => {
        if (importing) return;
        const todo = items.filter(x => x._on);
        const bad = todo.filter(x => L.validateScript(x).length);
        if (bad.length) return toast(`有 ${bad.length} 条缺场景名或回复内容`, true);
        importing = true;
        markDirty('import');
        const controls = $$('input, textarea, select, button', box).concat($('#im-go'), $('#im-text'));
        const disabled = controls.map(el => el.disabled);
        controls.forEach(el => { el.disabled = true; });
        try {
          const result = await W.importScripts(todo, it => DB.saveScript(Object.assign(L.emptyScript(), pickContent(it), { change_note: '从话术文档导入' })));
          result.saved.forEach(saved => upsert(S.scripts, saved));
          items = items.filter(it => !it._saved);
          if (!alive()) return;
          if (result.error) {
            importNotice = `已导入 ${result.saved.length} 条，剩余 ${items.filter(it => it._on).length} 条未导入。${result.error.message}。未导入内容及修改已保留，可继续重试。`;
            drawItems();
          } else {
            toast(`已导入 ${result.saved.length} 条`);
            clearDirty('import');
            location.hash = '#/library';
          }
        } finally {
          importing = false;
          controls.forEach((el, i) => { if (el.isConnected) el.disabled = disabled[i]; });
        }
      }));
    }
  }

  // ---------- 训练营知识库 ----------
  function knowledgeScopeHtml(id) {
    return `<div class="field knowledge-scope"><label for="${id}">参考哪个训练营的资料</label>
      <select id="${id}"><option value="${esc(S.knowledgeProject)}">${esc(S.knowledgeProject || '仅通用资料')}</option></select>
      <p class="hint" id="${id}-note">选择训练营后，会同时检索该训练营和通用资料。<a href="#/knowledge">管理知识库</a></p></div>`;
  }

  async function bindKnowledgeScope(id) {
    if (typeof DB.listKnowledgeDocuments !== 'function') return;
    const select = $('#' + id), note = $('#' + id + '-note');
    if (!select) return;
    select.addEventListener('change', () => { delete select.dataset.historyProject; S.knowledgeProject = select.value; });
    try {
      const docs = await DB.listKnowledgeDocuments();
      if (!select.isConnected) return;
      const projects = uniq(docs.map(d => d.project)).sort((a, b) => a.localeCompare(b, 'zh-CN'));
      const selected = select.value;
      const historical = selected && select.dataset.historyProject === selected && !projects.includes(selected);
      select.innerHTML = '<option value="">仅通用资料</option>' + projects.map(p => `<option value="${esc(p)}">${esc(p)}</option>`).join('');
      if (historical) select.add(new Option(selected + '（历史范围）', selected));
      select.value = projects.includes(selected) || historical ? selected : '';
      if (selected && !projects.includes(selected) && !historical) S.knowledgeProject = '';
      note.innerHTML = (historical ? '当前列表没有历史记录中的训练营资料，调整前请确认资料范围。' : '选择训练营后，会同时检索该训练营和通用资料。未选具体渠道时，只参考通用渠道资料。') + '<a href="#/knowledge">管理知识库</a>';
    } catch (e) {
      if (!select.isConnected) return;
      note.textContent = e.message + '。仍可使用话术库生成，结果中会说明资料是否读取成功。';
    }
  }

  function knowledgeSourcesHtml(chunks) {
    return chunks.map(c => `<div class="knowledge-source">
      <div><a href="#/knowledge/${Number(c.document_id)}">${esc(c.title)}</a> <span class="faint">${esc(c.project || '通用资料')} · v${esc(c.version || 1)} · 片段 ${Number(c.ordinal)}</span></div>
      ${c.source_note ? `<div class="hint">${esc(c.source_note)}</div>` : ''}
      <div class="pre">${esc(c.content)}</div></div>`).join('');
  }

  function contextUsedHtml(result) {
    const c = result && result.context_used;
    if (!c) return '';
    const styles = c.style_scripts || [], chunks = c.knowledge || [];
    const direct = c.sent_to_ai === false && result.source === 'library';
    return `${c.knowledge_status === 'unavailable' ? `<div class="note warn knowledge-notice">${esc(c.knowledge_note || '本次未能读取知识库，请核实涉及课程的事实。')}</div>` : ''}
      <details class="knowledge-reference"><summary>${direct ? '本次匹配的原话' : `本次提供的参考 · ${styles.length} 条风格示例 · ${chunks.length} 段资料`}</summary>
        <p class="hint">资料范围为${esc(c.knowledge_project || '通用资料')}。这里保留本次参考的版本；打开资料可查看当前内容。</p>
        <p style="margin:8px 0">${styles.length ? '风格示例 ' + styles.map(s => `<a href="#/script/${Number(s.id)}">${esc(s.scene)} v${esc(s.version || 1)}</a>`).join('、') : '没有可用的团队话术，暂时无法对齐团队风格。'}</p>
        ${styles.filter(s => s.reply).map(s => `<details style="margin:8px 0"><summary>查看「${esc(s.scene)}」当时的原话</summary>${replyMessagesHtml(s.reply)}</details>`).join('')}
        ${direct ? '<p class="hint">本次直接使用已保存话术，未检索知识库，未核对课程资料是否更新。</p>' : chunks.length ? knowledgeSourcesHtml(chunks) : `<p class="hint">${esc(c.knowledge_note || '没有检索到相关资料，涉及课程事实时需要补充或核实。')}</p>`}
        <p class="hint">${direct ? '这里保留匹配时的原话和版本，未发送给 AI。' : result.demo || c.demo ? '演示模式只展示本地检索和已有话术，未调用 AI 生成。' : '以上片段已提供给 AI 参考；每个方案标注的资料仍需结合原文核对。'}</p>
      </details>`;
  }

  function knowledgeCitationsHtml(item, context) {
    const refs = new Set(item.knowledge_refs || []);
    const sources = ((context && context.knowledge) || []).filter(c => refs.has(c.ref));
    if (!sources.length) return '';
    return `<div class="hint knowledge-citations">AI 标注的资料 ${sources.map(c => `<a href="#/knowledge/${Number(c.document_id)}">${esc(c.title)} · 片段 ${Number(c.ordinal)}</a>`).join('、')}</div>`;
  }

  function knowledgeErrorHtml(e) {
    return `<div class="note warn">${esc(e.message)}${e.code === 'knowledge_missing' ? '<p style="margin-top:8px">在 Supabase 的 SQL Editor 执行 <a href="knowledge-setup.sql" target="_blank" rel="noopener">knowledge-setup.sql</a>，再刷新页面。已有话术和案例会保留。<a href="SETUP.md" target="_blank" rel="noopener">查看步骤</a></p>' : ''}</div>`;
  }

  async function viewKnowledge() {
    const alive = viewToken();
    app().innerHTML = `<div class="page-head"><div><h1>训练营知识库</h1><p class="sub">保存课程详情、直播逐字稿和常见问题。生成时只取相关片段，话术风格沿用团队话术库。</p></div><div class="row"><a class="btn plain" href="#/background">给 AI 的指引</a><a class="btn" href="#/knowledge/new">+ 添加资料</a></div></div>
      <div class="card" style="margin-bottom:16px"><h3>试试能否找到资料</h3>
        <p class="hint">输入客户常问的问题，检查相关段落是否能被找到。检索按中文关键词匹配，不调用 AI。</p>
        ${knowledgeScopeHtml('kb-search-project')}
        <form id="kb-search" class="row" style="margin-top:10px"><input id="kb-query" type="search" aria-label="检索问题" placeholder="例如 直播错过以后可以看回放吗" style="flex:1;min-width:160px" required>
          <select id="kb-search-channel" aria-label="检索渠道" style="width:auto">${channelOptions(L.SCRIPT_CHANNELS, 'all')}</select><button class="btn plain" type="submit">检索资料</button></form><div id="kb-hits"></div></div>
      <div class="toolbar"><input id="kb-filter" type="search" aria-label="筛选资料标题" placeholder="搜索标题或来源说明"><select id="kb-project" aria-label="筛选训练营" style="width:auto"><option value="*">全部训练营</option></select><label class="muted"><input type="checkbox" id="kb-archived"> 显示已停用</label></div>
      <div id="kb-documents"><div class="empty">加载资料列表…</div></div>`;
    bindKnowledgeScope('kb-search-project');
    $('#kb-search').addEventListener('submit', e => {
      e.preventDefault();
      const query = $('#kb-query').value.trim(), project = $('#kb-search-project').value, channel = $('#kb-search-channel').value;
      if (!query) return;
      withBusy($('button', e.target), '检索中…', async () => {
        try {
          const chunks = await DB.searchKnowledge(query, { project, channel, limit: 6 });
          if (!alive()) return;
          $('#kb-hits').innerHTML = chunks.length ? `<div class="knowledge-reference"><p class="hint">找到 ${chunks.length} 段相关资料，生成时最多提供这些片段。</p>${knowledgeSourcesHtml(chunks)}</div>` : '<div class="note warn" style="margin-top:12px">没有找到相关资料。试试资料中的关键词，或补充包含这个问题的常见问答。</div>';
        } catch (err) { if (alive()) $('#kb-hits').innerHTML = knowledgeErrorHtml(err); }
      });
    });
    try {
      const docs = await DB.listKnowledgeDocuments();
      if (!alive()) return;
      $('#kb-project').innerHTML = '<option value="*">全部训练营</option><option value="">通用资料</option>' + uniq(docs.map(d => d.project)).sort().map(p => `<option value="${esc(p)}">${esc(p)}</option>`).join('');
      const draw = () => {
        const query = $('#kb-filter').value.trim().toLowerCase(), project = $('#kb-project').value, archived = $('#kb-archived').checked;
        const rows = docs.filter(d => (archived || d.status !== 'archived') && (project === '*' || d.project === project) && (!query || (d.title + ' ' + d.source_note).toLowerCase().includes(query)));
        $('#kb-documents').innerHTML = rows.length ? `<div class="knowledge-grid">${rows.map(d => `<div class="card knowledge-card ${d.status === 'archived' ? 'archived' : ''}"><div class="row"><span class="tag">${esc(d.project || '通用资料')}</span><span class="tag gray">${esc(KNOWLEDGE_KINDS[d.kind] || '其他资料')}</span>${chTag(d.channel)}${d.status === 'archived' ? '<span class="tag gray">已停用</span>' : ''}</div><h3 style="margin-top:12px"><a href="#/knowledge/${Number(d.id)}">${esc(d.title)}</a></h3><p class="hint">${esc(d.source_note || '未填写来源说明')}</p><p class="faint">${Number(d.char_count || 0).toLocaleString()} 字 · ${Number(d.chunk_count || 0)} 段 · v${esc(d.version)}</p><div class="row between"><span class="faint">${esc(L.fmtDateTime(d.updated_at))}</span><a href="#/knowledge/${Number(d.id)}">查看 / 编辑</a></div></div>`).join('')}</div>` : '<div class="card empty">还没有符合条件的资料<br>把详情页正文或直播逐字稿粘贴进来，就可以按问题查找了。<br><a class="btn" href="#/knowledge/new" style="margin-top:10px">添加第一份资料</a></div>';
      };
      $('#kb-filter').addEventListener('input', draw);
      $('#kb-project').addEventListener('change', draw);
      $('#kb-archived').addEventListener('change', draw);
      draw();
    } catch (e) { if (alive()) $('#kb-documents').innerHTML = knowledgeErrorHtml(e); }
  }

  async function viewKnowledgeEdit(id) {
    const alive = viewToken();
    app().innerHTML = '<div class="empty">加载资料…</div>';
    let doc;
    try { doc = id ? await DB.getKnowledgeDocument(id) : { title: '', project: S.knowledgeProject, kind: 'detail', channel: 'all', content: '', source_note: '', status: 'active' }; }
    catch (e) { if (alive()) app().innerHTML = `<a class="back" href="#/knowledge">← 知识库</a>${knowledgeErrorHtml(e)}`; return; }
    if (!alive()) return;
    if (!doc) return notFound('这份资料不存在');
    app().innerHTML = `<a class="back" href="#/knowledge">← 知识库</a><div class="page-head"><div><h1>${id ? '编辑资料' : '添加资料'}</h1><p class="sub">每份资料保留完整正文，保存时自动分段。可用于多个训练营的资料，训练营名称留空即可。</p></div></div>
      <form class="card" id="kd-form"><div class="field"><label for="kd-title">资料标题 <b>*</b></label><input id="kd-title" type="text" maxlength="200" required placeholder="例如 Python 训练营课程详情和答疑服务" value="${esc(doc.title)}"></div>
        <div class="grid2"><div class="field"><label for="kd-project">训练营名称</label><input id="kd-project" type="text" maxlength="100" placeholder="留空表示通用资料" list="kd-projects" value="${esc(doc.project)}"><datalist id="kd-projects"></datalist><p class="hint">同一个训练营使用相同名称；不同期次政策有差异时，可把期次写进名称。</p></div>
          <div class="field"><label for="kd-kind">资料类型</label><select id="kd-kind">${Object.entries(KNOWLEDGE_KINDS).map(([k, v]) => `<option value="${k}" ${doc.kind === k ? 'selected' : ''}>${v}</option>`).join('')}</select></div></div>
        <div class="grid2"><div class="field"><label for="kd-channel">适用渠道</label><select id="kd-channel">${channelOptions(L.SCRIPT_CHANNELS, doc.channel)}</select></div><div class="field"><label for="kd-status">使用状态</label><select id="kd-status"><option value="active" ${doc.status === 'active' ? 'selected' : ''}>启用，生成时可参考</option><option value="archived" ${doc.status === 'archived' ? 'selected' : ''}>停用，不再用于新生成</option></select></div></div>
        <div class="field"><label for="kd-source">来源 / 有效期说明</label><input id="kd-source" type="text" maxlength="1000" value="${esc(doc.source_note)}" placeholder="例如 10 月课程详情页，优惠仅适用于本期"></div>
        <div class="field"><label for="kd-content">资料正文 <b>*</b></label><textarea id="kd-content" class="knowledge-text" rows="16" required maxlength="200000" placeholder="在这里粘贴详情页正文、直播逐字稿、课程安排或常见问答。"></textarea><div class="hint" id="kd-count" aria-live="polite"></div><p class="hint">每份最多 20 万字。资料用于提供课程事实，生成时的表达习惯仍参考话术库。过期优惠请及时修改或停用，来源说明中的日期不会自动触发停用。</p></div>
        <div id="kd-error"></div><div class="row"><span class="spacer"></span><a class="btn plain" href="#/knowledge">取消</a><button type="submit" class="btn" id="kd-save">保存资料</button></div></form>`;
    const form = $('#kd-form'), content = $('#kd-content');
    content.value = doc.content;
    const read = () => ({ ...doc, title: $('#kd-title').value.trim(), project: $('#kd-project').value.trim(), kind: $('#kd-kind').value, channel: $('#kd-channel').value, status: $('#kd-status').value, source_note: $('#kd-source').value.trim(), content: content.value });
    const initial = JSON.stringify(read());
    const count = () => { $('#kd-count').textContent = `${content.value.length.toLocaleString()} / 200,000 字`; };
    count();
    form.addEventListener('input', () => { markDirty('knowledge'); count(); });
    form.addEventListener('change', () => markDirty('knowledge'));
    DB.listKnowledgeDocuments().then(docs => {
      if (alive()) $('#kd-projects').innerHTML = uniq(docs.map(d => d.project)).map(p => `<option value="${esc(p)}">`).join('');
    }).catch(() => {});
    form.addEventListener('submit', e => {
      e.preventDefault();
      const data = read();
      if (!data.title || !data.content.trim()) { $('#kd-error').innerHTML = '<div class="note warn">请填写资料标题和正文。</div>'; return; }
      if (id && JSON.stringify(data) === initial) { clearDirty('knowledge'); toast('内容没有变化'); return; }
      withBusy($('#kd-save'), '保存并更新检索片段…', async () => {
        const controls = $$('input, textarea, select', form);
        controls.forEach(el => { el.disabled = true; });
        try {
          const saved = await DB.saveKnowledgeDocument(data);
          if (!alive()) return;
          S.knowledgeProject = saved.project || '';
          clearDirty('knowledge');
          toast(saved.status === 'archived' ? '已停用，后续生成不再参考这份资料' : '资料已保存，可以检索了');
          location.hash = '#/knowledge';
        } catch (err) {
          if (!alive()) return;
          $('#kd-error').innerHTML = err.code === 'conflict' ? '<div class="note warn">这份资料刚被同事修改，你的编辑仍保留在这里。请先复制需要保留的正文，再重新载入最新资料。<button type="button" class="linkbtn" id="kd-reload">重新载入最新资料</button></div>' : knowledgeErrorHtml(err);
          const reload = $('#kd-reload');
          if (reload) reload.addEventListener('click', async () => {
            if (!(await confirmBox('重新载入最新资料？', '当前未保存的编辑会被替换，请确认已复制需要保留的内容。', '重新载入'))) return;
            clearDirty('knowledge'); viewKnowledgeEdit(id);
          });
        } finally { controls.forEach(el => { el.disabled = false; }); }
      });
    });
  }

  // ---------- 给 AI 的背景资料 ----------
  function viewBackground() {
    const PLACEHOLDER = `填写团队统一指引，例如
先问清客户的学习目标，再介绍合适的课程。
不承诺就业，不承诺收入。
课程价格、服务安排和直播逐字稿可以分别放进知识库。`;
    app().innerHTML = `
      <a class="back" href="#/library">← 话术库</a>
      <div class="page-head"><div><h1>给 AI 的指引</h1><p class="sub">填写团队统一的销售原则和禁忌。这段内容每次都会提供给 AI，建议保持简短；课程详情和直播长文请放进 <a href="#/knowledge">知识库</a>，按问题选取片段。</p></div></div>
      <div class="card">
        <textarea id="bg" rows="18" placeholder="${esc(PLACEHOLDER)}">${esc(S.bg)}</textarea>
        <div class="row" style="margin-top:10px"><span class="hint">只保存在数据库里，代码仓库里不会有这些内容</span><span class="spacer"></span><button class="btn" id="bg-save">保存</button></div>
      </div>`;
    $('#bg').addEventListener('input', () => markDirty('bg'));
    $('#bg-save').addEventListener('click', () => withBusy($('#bg-save'), '保存中…', async () => {
      try {
        const v = $('#bg').value;
        await DB.setSetting(BG_KEY, v);
        S.bg = v;
        clearDirty('bg');
        toast('已保存');
      } catch (e) { showErr(e); }
    }));
  }

  // ---------- 成员 ----------
  async function viewMembers() {
    const alive = viewToken();
    app().innerHTML = `
      <div class="page-head"><div><h1>成员</h1><p class="sub">只有名单里的邮箱能使用话术复盘系统</p></div></div>
      <div class="card">
        <div class="row"><input type="email" id="mb-email" placeholder="同事的邮箱" style="max-width:320px"><button class="btn" id="mb-add">添加</button></div>
        <p class="hint">对方要先用这个邮箱在 Skill 市集或本页注册账号，注册之后才能加进来；加好后对方刷新页面就能用了</p>
        <div id="mb-list" style="margin-top:14px"><span class="muted">加载中…</span></div>
      </div>`;
    const draw = async () => {
      try {
        const list = await DB.listMembers();
        if (!alive()) return;
        $('#mb-list').innerHTML = `<table class="mini" style="width:100%"><tr><th>邮箱</th><th>添加时间</th><th>添加人</th><th></th></tr>
          ${list.map(m => `<tr><td>${esc(m.email)}</td><td>${esc(L.fmtDateTime(m.added_at))}</td><td>${esc(m.added_by_name)}</td>
            <td>${m.email === String(S.email).toLowerCase() ? '<span class="faint">你自己</span>' : `<button class="linkbtn danger" data-rm="${esc(m.email)}">移除</button>`}</td></tr>`).join('')}</table>`;
        $$('[data-rm]').forEach(b => b.addEventListener('click', async () => {
          if (!(await confirmBox('移除成员', `移除后 ${b.dataset.rm} 就看不到话术和案例了。`, '移除', true))) return;
          try { await DB.removeMember(b.dataset.rm); draw(); } catch (e) { showErr(e); }
        }));
      } catch (e) { if (alive()) $('#mb-list').innerHTML = `<div class="note err">${esc(e.message)}</div>`; }
    };
    $('#mb-add').addEventListener('click', () => withBusy($('#mb-add'), '添加中…', async () => {
      try { await DB.addMember($('#mb-email').value); $('#mb-email').value = ''; toast('已添加'); draw(); } catch (e) { showErr(e); }
    }));
    draw();
  }

  // ======================================================================
  // 案例
  // ======================================================================
  function caseCard(c) {
    return `<a class="case-card ${c.result}" href="#/case/${c.id}">
      <div class="case-meta">${resultBadge(c.result)}<b>${esc(c.seller)}</b>${chTag(c.channel)}${c.cohort ? `<span>${esc(c.cohort)}</span>` : ''}<span>${esc(L.fmtDateCN(c.case_date))}</span>
        ${(c.scenes || []).map(s => `<span class="tag gray">${esc(s)}</span>`).join('')}<span class="spacer"></span>${statusBadge(c.review_status)}</div>
      ${c.stuck_point ? `<div class="case-point">${esc(c.stuck_point)}</div>` : ''}
    </a>`;
  }

  function viewCases() {
    const f = S.caseFilter || (S.caseFilter = { week: '', result: '', ch: '', seller: '', status: '' });
    const weeks = uniq(S.cases.map(c => c.review_week)).sort().reverse();
    const sellers = uniq(S.cases.map(c => c.seller)).sort((a, b) => a.localeCompare(b, 'zh'));
    const opt = (v, t, cur) => `<option value="${esc(v)}" ${cur === v ? 'selected' : ''}>${esc(t)}</option>`;
    app().innerHTML = `
      <div class="page-head"><div><h1>案例</h1><p class="sub">销售每周提交成交和流失的案例，复盘会上一起过</p></div><a class="btn" href="#/case/new">+ 提交案例</a></div>
      <div class="toolbar">
        <select id="cf-week" style="width:auto">${opt('', '全部周', f.week)}${weeks.map(w => opt(w, L.weekLabel(w), f.week)).join('')}</select>
        <div class="chips" id="cf-result">${[['', '全部'], ['won', '成交'], ['lost', '流失']].map(([k, t]) => `<button class="chip ${f.result === k ? 'on' : ''}" data-r="${k}">${t}</button>`).join('')}</div>
        <select id="cf-ch" style="width:auto">${opt('', '全部渠道', f.ch)}${L.CASE_CHANNELS.map(k => opt(k, L.CHANNELS[k].label, f.ch)).join('')}</select>
        <select id="cf-seller" style="width:auto">${opt('', '全部销售', f.seller)}${sellers.map(s => opt(s, s, f.seller)).join('')}</select>
        <select id="cf-status" style="width:auto">${opt('', '全部状态', f.status)}${opt('pending', '待复盘', f.status)}${opt('discussed', '已讨论', f.status)}</select>
      </div>
      <div id="cf-body"></div>`;
    const draw = () => {
      const list = S.cases.filter(c => (!f.week || c.review_week === f.week) && (!f.result || c.result === f.result)
        && (!f.ch || c.channel === f.ch) && (!f.seller || c.seller === f.seller) && (!f.status || c.review_status === f.status));
      const body = $('#cf-body');
      if (!S.cases.length) {
        body.innerHTML = '<div class="card empty">还没有案例<br>成交的、流失的都可以交，复盘会上对比着看最有收获<br><a class="btn" href="#/case/new" style="margin-top:8px">提交第一个案例</a></div>';
        return;
      }
      if (!list.length) { body.innerHTML = '<div class="empty">没有符合条件的案例</div>'; return; }
      const byWeek = new Map();
      for (const c of list) { if (!byWeek.has(c.review_week)) byWeek.set(c.review_week, []); byWeek.get(c.review_week).push(c); }
      body.innerHTML = [...byWeek.keys()].sort().reverse().map(w => {
        const cs = byWeek.get(w);
        return `<h2>${esc(L.weekLabel(w))} <small>成交 ${cs.filter(c => c.result === 'won').length} · 流失 ${cs.filter(c => c.result === 'lost').length}</small>
          <a href="#/review/${w}" style="font-size:13px;font-weight:400;margin-left:auto">去这周的复盘会 →</a></h2>
          <div class="case-list">${cs.map(caseCard).join('')}</div>`;
      }).join('');
    };
    draw();
    const bind = (id, key) => $(id).addEventListener('change', e => { f[key] = e.target.value; draw(); });
    bind('#cf-week', 'week'); bind('#cf-ch', 'ch'); bind('#cf-seller', 'seller'); bind('#cf-status', 'status');
    $$('#cf-result [data-r]').forEach(b => b.addEventListener('click', () => {
      f.result = b.dataset.r;
      $$('#cf-result .chip').forEach(x => x.classList.toggle('on', x === b));
      draw();
    }));
  }

  function chatHtml(dialogue) {
    const list = L.normDialogue(dialogue);
    if (!list.length) return '<div class="faint">没有聊天记录</div>';
    return `<div class="chat">${list.map(x => `<div class="bubble-row ${x.role}"><div>${x.role !== 'note' ? `<div class="who">${L.ROLES[x.role]}</div>` : ''}<div class="bubble">${esc(x.text)}</div></div></div>`).join('')}</div>`;
  }

  // AI 分析结果；withActions 时每条建议带「据此改话术」按钮
  function analysisHtml(a, withActions, withConfirmation) {
    if (!a) return '';
    const list = arr => (arr && arr.length ? `<ul>${arr.map(x => `<li>${esc(x)}</li>`).join('')}</ul>` : '<span class="faint">无</span>');
    return `<div class="analysis">
      ${contextUsedHtml(a)}
      ${knowledgeCitationsHtml(a, a.context_used)}
      ${a.summary ? `<div>${esc(a.summary)}</div>` : ''}
      ${a.key_moment ? `<h4>关键转折</h4><div>${esc(a.key_moment)}</div>` : ''}
      <h4>做得好的</h4>${list(a.highlights)}
      <h4>可以改进的</h4>${list(a.problems)}
      ${(a.used_script_candidates || []).length ? `<h4>可能用到的话术（AI 待确认）</h4><div class="faint">核对聊天和当时版本，手动确认后才计入实际使用。</div><ul>${a.used_script_candidates.map(x => `<li>${scriptRef(x.script_id)}${x.evidence ? ` · 对话依据：${esc(x.evidence)}` : ''}${withConfirmation && scriptById(x.script_id) ? ` <button type="button" class="linkbtn" data-use-candidate="${esc(x.script_id)}">核对使用版本</button>` : ''}</li>`).join('')}</ul>` : ''}
      ${(a.recommended_script_ids || []).length ? `<h4>下次建议使用</h4><div>${a.recommended_script_ids.map(id => scriptRef(id)).join('、')}</div><div class="faint">建议不计入本次实际使用。</div>` : ''}
      ${(a.script_ids || []).length ? `<h4>旧版 AI 关联（待核对）</h4><div>${a.script_ids.map(id => scriptRef(id)).join('、')}</div><div class="faint">旧版分析混合了实际使用与建议，请在编辑案例时核对已有的话术记录。</div>` : ''}
      ${a.suggestions && a.suggestions.length ? `<h4>话术建议</h4><ul>${a.suggestions.map((s, i) => `<li>
          ${s.target_note ? '待核对修改目标 · ' : s.script_id && scriptById(s.script_id) ? `改 ${scriptRef(s.script_id)}：` : `新增「${esc(s.scene)}」：`}${esc(s.change)}
          ${s.target_note ? `<div class="note warn">${esc(s.target_note)}</div>` : ''}
          ${knowledgeCitationsHtml(s, a.context_used)}
          ${s.reason ? `<span class="faint">（${esc(s.reason)}）</span>` : ''}
          ${withActions ? ` <button class="linkbtn" data-sug="${i}">据此改话术</button>` : ''}</li>`).join('')}</ul>` : ''}
    </div>`;
  }

  // 根据一个案例去改话术：先选改哪条，再打开编辑器
  async function fixScriptFromCase(c, extra) {
    extra = extra || {};
    const cands = uniq((c.script_refs || []).map(r => Number(r.id))
      .concat(S.scripts.filter(s => s.status !== 'archived' && (c.scenes || []).includes(s.scene)).map(s => s.id)))
      .map(scriptById).filter(Boolean);
    let target = extra.script_id ? scriptById(extra.script_id) : undefined;
    if (target === undefined) {
      const pick = await choiceBox('改哪条话术？', cands.length ? '这个案例相关的话术：' : '这个案例还没关联话术，可以新增一条。',
        cands.map(s => ({ label: `「${s.scene}」v${s.version}`, value: s.id, cls: 'plain' }))
          .concat([{ label: '新增一条话术', value: 'new', cls: 'ghost' }, { label: '取消', value: null, cls: 'plain' }]));
      if (pick == null) return null;
      target = pick === 'new' ? null : scriptById(pick);
    }
    return openScriptEditor(target || null, {
      prefill: target ? {} : { stage: extra.stage || '', scene: extra.scene || (c.scenes || [])[0] || '' },
      change_note: extra.change_note || '',
      hint: extra.hint || '',
      review_id: extra.review_id || (reviewByWeek(c.review_week) || {}).id || null,
      case_ids: [c.id],
    });
  }

  // 版本冲突时问一下要不要覆盖（写文字的地方用，避免悄悄盖掉别人刚写的内容）
  function askOverwrite(what) {
    return choiceBox(`${what}刚被别人改过`, '可以先看看别人写了什么（你写的内容会显示在旁边），也可以直接用你的覆盖。',
      [{ label: '先看看别人写的', value: false, cls: 'plain' }, { label: '用我的覆盖', value: true, cls: 'danger' }]);
  }
  function unsavedNote(text) {
    return text ? `<div class="note warn pre" style="margin-top:8px">你刚才写的（还没保存）：\n${esc(text)}</div>` : '';
  }

  // 只改案例的一两个字段（复盘会上用）。版本冲突时：勾选这类小改动自动拿最新的再改一次；
  // 传了 opts.ask 的先问用户，用户选择不覆盖时返回 null
  async function updateCase(id, mutate, opts) {
    for (let attempt = 0; attempt < 2; attempt++) {
      const cur = caseById(id);
      if (!cur) throw DB.makeError('not_found', '案例不存在');
      const d = clone(cur);
      mutate(d);
      try { const saved = await DB.saveCase(d); upsert(S.cases, saved); return saved; } catch (e) {
        if (e.code !== 'conflict' || attempt) throw e;
        S.cases = await DB.listCases();
        if (opts && opts.ask && !(await askOverwrite(opts.ask))) return null;
      }
    }
  }

  function viewCase(id) {
    const c = caseById(id);
    if (!c) return notFound('这个案例不存在，可能已经被删除');
    const refs = (c.script_refs || []).map(r => scriptRef(r.id, r.version)).join('、');
    app().innerHTML = `
      <a class="back" href="#/cases">← 案例</a>
      <div class="page-head">
        <div class="case-meta" style="font-size:14px">${resultBadge(c.result)}<span>案例 #${c.id}</span><b>${esc(c.seller)}</b>${chTag(c.channel)}${c.cohort ? `<span>${esc(c.cohort)}</span>` : ''}
          <span>${esc(L.fmtDateCN(c.case_date))}</span><a href="#/review/${c.review_week}">归入 ${esc(L.weekLabel(c.review_week))} 的复盘会</a>${statusBadge(c.review_status)}</div>
        <div class="row"><button class="btn danger sm" id="cs-del">删除</button><a class="btn sm" href="#/case/${c.id}/edit">编辑</a></div>
      </div>
      <div class="grid-side">
        <div class="card"><h3>聊天记录 <small>${L.normDialogue(c.dialogue).length} 条</small></h3><div class="chat-scroll">${chatHtml(c.dialogue)}</div></div>
        <div class="stack">
          <div class="card">
            ${c.profile ? `<div class="detail-block"><div class="label">客户画像</div><div class="pre">${esc(c.profile)}</div></div>` : ''}
            <div class="detail-block"><div class="label">${c.result === 'won' ? '成交的关键' : '在哪一步流失的'}</div><div class="pre">${esc(c.stuck_point) || '<span class="faint">没填</span>'}</div></div>
            ${c.self_review ? `<div class="detail-block"><div class="label">销售自己的复盘</div><div class="pre">${esc(c.self_review)}</div></div>` : ''}
            <div class="detail-block"><div class="label">涉及场景</div>${(c.scenes || []).length ? c.scenes.map(s => `<span class="tag gray">${esc(s)}</span>`).join(' ') : '<span class="faint">没标</span>'}</div>
            <div class="detail-block"><div class="label">用到的话术</div>${refs || '<span class="faint">没标</span>'}</div>
            <div class="faint" style="margin-top:12px">${esc(c.created_by_name)} 提交于 ${esc(L.fmtDateTime(c.created_at))}${c.version > 1 ? ` · ${esc(c.updated_by_name)} 最后修改于 ${esc(L.fmtDateTime(c.updated_at))}` : ''}</div>
          </div>
          ${c.ai_analysis ? `<div class="card"><h3>AI 分析</h3>${analysisHtml(c.ai_analysis, true)}</div>` : ''}
          <div class="card"><h3>复盘会结论</h3>
            <textarea id="cs-disc" rows="4" placeholder="会上讨论出来的结论、下次怎么做">${esc(c.discussion)}</textarea>
            <div class="row" style="margin-top:8px"><button class="btn plain sm" id="cs-fix">据此改话术</button><span class="spacer"></span><button class="btn sm" id="cs-disc-save">保存，标记已讨论</button></div>
          </div>
        </div>
      </div>`;
    $('#cs-disc').addEventListener('input', () => markDirty('disc'));
    $('#cs-disc-save').addEventListener('click', () => withBusy($('#cs-disc-save'), '保存中…', async () => {
      try {
        const v = $('#cs-disc').value.trim();
        const saved = await updateCase(id, d => { d.discussion = v; d.review_status = 'discussed'; }, { ask: '这个案例的结论' });
        clearDirty('disc');
        viewCase(id);
        if (saved) toast('已保存'); else $('#cs-disc').insertAdjacentHTML('afterend', unsavedNote(v));
      } catch (e) { showErr(e); }
    }));
    $('#cs-fix').addEventListener('click', () => fixScriptFromCase(c, { change_note: c.discussion ? `案例#${c.id} 复盘结论：${c.discussion}` : '' }));
    $$('[data-sug]').forEach(b => b.addEventListener('click', () => {
      const s = c.ai_analysis.suggestions[+b.dataset.sug];
      fixScriptFromCase(c, {
        script_id: s.script_id && scriptById(s.script_id) ? s.script_id : null, scene: s.scene,
        hint: `AI 建议：${s.change}${s.reason ? `\n依据：${s.reason}` : ''}`,
        change_note: `${s.change}${s.reason ? `（${s.reason}）` : ''}`,
      });
    }));
    $('#cs-del').addEventListener('click', async () => {
      if (!(await confirmBox('删除这个案例？', '删除后找不回来。', '删除', true))) return;
      try { await DB.deleteCase(id); S.cases = S.cases.filter(x => x.id !== id); toast('已删除'); location.hash = '#/cases'; } catch (e) { showErr(e); }
    });
  }

  function viewCaseEdit(id) {
    const orig = id ? caseById(id) : null;
    if (id && !orig) return notFound('这个案例不存在');
    const d = orig ? clone(orig) : L.emptyCase(S.me);
    d.dialogue = L.normDialogue(d.dialogue);
    let images = [];
    const backHref = orig ? '#/case/' + id : '#/cases';
    const cohorts = uniq(S.cases.map(c => c.cohort));
    const sellers = uniq(S.cases.map(c => c.seller).concat(S.me));

    app().innerHTML = `
      <a class="back" href="${backHref}">← ${orig ? '返回案例' : '案例'}</a>
      <div class="page-head"><h1>${orig ? '编辑案例 #' + id : '提交案例'}</h1></div>
      <form id="ce" class="stack" autocomplete="off">
        <div class="card">
          <h3>基本信息</h3>
          <div class="field"><label>结果 <b>*</b></label><div class="radio-group" id="ce-result">
            <label class="won"><input type="radio" name="result" value="won">成交</label>
            <label class="lost"><input type="radio" name="result" value="lost">流失</label></div></div>
          <div class="inline-fields">
            <div class="field"><label>销售 <b>*</b></label><input type="text" name="seller" list="ce-sellers" value="${esc(d.seller)}"></div>
            <div class="field"><label>渠道</label><select name="channel">${channelOptions(L.CASE_CHANNELS, d.channel)}</select></div>
            <div class="field"><label>期次</label><input type="text" name="cohort" list="ce-cohorts" value="${esc(d.cohort)}" placeholder="如：10月第2期"></div>
            <div class="field"><label>日期 <b>*</b></label><input type="date" name="case_date" value="${esc(d.case_date)}"></div>
            <div class="field"><label>归到哪周的复盘会</label><select name="review_week"></select></div>
          </div>
          <div class="field" style="margin-bottom:0"><label>客户画像</label><input type="text" name="profile" value="${esc(d.profile)}" placeholder="如：30 岁上班族，想转行，之前买过别家的课"></div>
        </div>

        <div class="card">
          <h3>聊天记录 <small>手机号、微信号会自动打码；截图只用来识别，不会保存</small></h3>
          <div class="tabs"><button type="button" data-tab="shot" class="on">上传截图</button><button type="button" data-tab="text">粘贴文字</button></div>
          <div data-pane="shot">
            <div class="dropzone" id="ce-drop" tabindex="0">点击选择、拖进来，或者直接 ⌘V / Ctrl+V 粘贴聊天截图<br><span class="faint">可以一次传多张，按聊天顺序排好</span></div>
            <input type="file" id="ce-file" accept="image/*" multiple hidden>
            <div class="thumbs" id="ce-thumbs"></div>
            <div class="row" style="margin-top:10px"><button type="button" class="btn" id="ce-ocr" disabled>识别成对话</button><span class="hint" id="ce-ocr-st"></span></div>
          </div>
          <div data-pane="text" hidden>
            <textarea id="ce-raw" rows="8" placeholder="把聊天文字贴进来。每行开头写了「客户：」「销售：」的会直接识别，其他格式交给 AI 整理"></textarea>
            <div class="row" style="margin-top:10px"><button type="button" class="btn" id="ce-parse">整理成对话</button><span class="hint" id="ce-parse-st"></span></div>
          </div>
          <div style="margin-top:18px">
            <div class="row between"><div class="label" style="margin:0">对话 <span class="faint">（可以直接改；角色识别错了就在左边改过来）</span></div><button type="button" class="linkbtn danger" id="ce-clear">清空</button></div>
            <div id="ce-dlg" style="margin-top:8px"></div>
            <button type="button" class="btn plain sm" id="ce-add-line">+ 加一条</button>
          </div>
        </div>

        <div class="card">
          <h3>复盘</h3>
          <div class="field"><label id="ce-sp-label"></label><textarea name="stuck_point" rows="3">${esc(d.stuck_point)}</textarea></div>
          <div class="field"><label>我自己的复盘（做得好的 / 下次想改的）</label><textarea name="self_review" rows="3">${esc(d.self_review)}</textarea></div>
          ${knowledgeScopeHtml('ce-knowledge')}
          <div class="row"><button type="button" class="btn ghost" id="ce-ai">AI 分析这个案例</button><span class="hint">自动标出场景，列出待确认的话术和改进建议</span></div>
          <div id="ce-analysis" style="margin-top:12px"></div>
          <div class="field" style="margin-top:14px"><label>涉及场景</label><div class="chips" id="ce-scenes"></div>
            <input type="text" id="ce-scene-in" list="ce-scene-list" placeholder="输入场景后按回车，如：嫌贵" style="margin-top:6px;max-width:320px"></div>
          <div class="field" style="margin-bottom:0"><label>实际用到的话术</label><div class="hint">只记录聊天中实际使用的话术；请选当时版本，不确定可选「版本未知」。</div><div class="chips" id="ce-refs"></div>
            <div id="ce-refs-note"></div>
            <select id="ce-ref-add" style="margin-top:6px;max-width:420px"><option value="">+ 从话术库里选，或重新核对版本</option>
              ${S.scripts.map(s => `<option value="${s.id}">${esc((s.stage || '其他') + ' · ' + s.scene)}${s.status === 'archived' ? '（已停用）' : ''}</option>`).join('')}</select>
            <div id="ce-ref-picker" hidden style="margin-top:8px">
              <div class="row"><select id="ce-ref-version" style="width:auto"></select><button type="button" class="btn plain sm" id="ce-ref-confirm" disabled>确认实际使用</button></div>
              <div class="hint" id="ce-ref-status"></div><div class="faint" id="ce-ref-preview" style="margin-top:8px"></div>
            </div></div>
        </div>

        <div id="ce-err"></div>
        <div class="row"><span class="spacer"></span><a class="btn plain" href="${backHref}">取消</a><button type="submit" class="btn" id="ce-save">保存案例</button></div>
      </form>
      <datalist id="ce-sellers">${sellers.map(s => `<option value="${esc(s)}">`).join('')}</datalist>
      <datalist id="ce-cohorts">${cohorts.map(s => `<option value="${esc(s)}">`).join('')}</datalist>
      <datalist id="ce-scene-list">${allScenes().map(s => `<option value="${esc(s)}">`).join('')}</datalist>`;

    const form = $('#ce');
    bindKnowledgeScope('ce-knowledge');
    const dirty = () => markDirty('case');
    form.addEventListener('input', dirty);

    // 结果
    const drawResult = () => {
      $$('#ce-result label').forEach(l => l.classList.toggle('on', l.querySelector('input').value === d.result));
      $$('#ce-result input').forEach(i => { i.checked = i.value === d.result; });
      $('#ce-sp-label').textContent = d.result === 'won' ? '成交的关键是什么' : '在哪一步流失的、为什么';
    };
    $$('#ce-result input').forEach(i => i.addEventListener('change', () => { d.result = i.value; drawResult(); }));
    drawResult();

    // 归到哪周：默认日期所在周，也可以放到下一周
    const drawWeeks = () => {
      const w = L.weekStart(form.case_date.value) || d.review_week;
      const opts = uniq([w, L.addDays(w, 7), d.review_week]);
      if (!opts.includes(d.review_week)) d.review_week = w;
      form.review_week.innerHTML = opts.map(x => `<option value="${x}" ${x === d.review_week ? 'selected' : ''}>${esc(L.weekLabel(x))}${x === w ? '（日期所在周）' : x === L.addDays(w, 7) ? '（下一周）' : ''}</option>`).join('');
    };
    form.case_date.addEventListener('change', () => { d.review_week = L.weekStart(form.case_date.value) || d.review_week; drawWeeks(); });
    form.review_week.addEventListener('change', () => { d.review_week = form.review_week.value; });
    drawWeeks();

    // 标签页
    $$('.tabs button', form).forEach(b => b.addEventListener('click', () => {
      $$('.tabs button', form).forEach(x => x.classList.toggle('on', x === b));
      $$('[data-pane]', form).forEach(p => { p.hidden = p.dataset.pane !== b.dataset.tab; });
    }));

    // 对话编辑
    const drawDlg = () => {
      const box = $('#ce-dlg');
      box.innerHTML = d.dialogue.length ? d.dialogue.map((x, i) => `<div class="dlg-row ${x.role}">
          <select data-i="${i}">${Object.entries(L.ROLES).map(([k, t]) => `<option value="${k}" ${x.role === k ? 'selected' : ''}>${t}</option>`).join('')}</select>
          <textarea data-i="${i}" rows="1">${esc(x.text)}</textarea>
          <button type="button" class="icon-btn" data-del="${i}" title="删除这条">×</button></div>`).join('')
        : '<div class="faint" style="padding:6px 0 10px">还没有对话。上传截图或粘贴文字后点识别，也可以手动加</div>';
      $$('select', box).forEach(s => s.addEventListener('change', () => { d.dialogue[+s.dataset.i].role = s.value; s.parentNode.className = 'dlg-row ' + s.value; }));
      $$('textarea', box).forEach(t => { autosize(t); t.addEventListener('input', () => { d.dialogue[+t.dataset.i].text = t.value; autosize(t); }); });
      $$('[data-del]', box).forEach(b => b.addEventListener('click', () => { d.dialogue.splice(+b.dataset.del, 1); dirty(); drawDlg(); }));
    };
    drawDlg();
    $('#ce-add-line').addEventListener('click', () => {
      const last = d.dialogue[d.dialogue.length - 1];
      d.dialogue.push({ role: last && last.role === 'customer' ? 'sales' : 'customer', text: '' });
      drawDlg();
      const tas = $$('#ce-dlg textarea');
      tas[tas.length - 1].focus();
    });
    $('#ce-clear').addEventListener('click', async () => {
      if (d.dialogue.length && !(await confirmBox('清空对话？', '', '清空', true))) return;
      d.dialogue = []; dirty(); drawDlg();
    });
    const addDialogue = async (list) => {
      if (!list.length) return toast('没识别出对话内容', true);
      if (d.dialogue.length) {
        const how = await choiceBox('已经有对话了', `新识别出 ${list.length} 条，怎么处理？`,
          [{ label: '取消', value: null, cls: 'plain' }, { label: '替换掉', value: 'replace', cls: 'plain' }, { label: '接在后面', value: 'append' }]);
        if (!how) return;
        d.dialogue = how === 'replace' ? list : d.dialogue.concat(list);
      } else d.dialogue = list;
      dirty(); drawDlg();
      toast(`识别出 ${list.length} 条，检查一下角色对不对`);
    };

    // 截图
    const drawThumbs = () => {
      $('#ce-thumbs').innerHTML = images.map((u, i) => `<figure><img src="${u}" alt="截图 ${i + 1}"><button type="button" data-rm="${i}">×</button></figure>`).join('');
      $$('#ce-thumbs [data-rm]').forEach(b => b.addEventListener('click', () => { images.splice(+b.dataset.rm, 1); drawThumbs(); }));
      $('#ce-ocr').disabled = !images.length;
      $('#ce-ocr-st').textContent = images.length ? `${images.length} 张（长截图会自动切开）` : '';
    };
    const addFiles = async files => {
      if (!files.length) return;
      try { images = images.concat(await prepareImages(files)); drawThumbs(); } catch (e) { showErr(e); }
    };
    const drop = $('#ce-drop');
    drop.addEventListener('click', () => $('#ce-file').click());
    $('#ce-file').addEventListener('change', e => { addFiles([...e.target.files]); e.target.value = ''; });
    drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('drag'); });
    drop.addEventListener('dragleave', () => drop.classList.remove('drag'));
    drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('drag'); addFiles([...e.dataTransfer.files]); });
    const onPaste = e => {
      if (!document.body.contains(form)) return document.removeEventListener('paste', onPaste);
      const imgs = imagesFromClipboard(e);
      if (!imgs.length) return;
      e.preventDefault();
      $$('.tabs button', form)[0].click();
      addFiles(imgs);
    };
    document.addEventListener('paste', onPaste);
    $('#ce-ocr').addEventListener('click', () => withBusy($('#ce-ocr'), '识别中…', async () => {
      try {
        await addDialogue(await AI.transcribeImages(images, form.seller.value.trim()));
        images = []; drawThumbs();
      } catch (e) { showErr(e); }
    }));

    // 粘贴文字
    $('#ce-parse').addEventListener('click', () => withBusy($('#ce-parse'), '整理中…', async () => {
      const raw = $('#ce-raw').value.trim();
      if (!raw) return toast('先粘贴聊天文字', true);
      try {
        const local = L.parseDialogue(L.maskSensitive(raw));
        await addDialogue(local || await AI.transcribeText(raw, form.seller.value.trim()));
        $('#ce-raw').value = '';
      } catch (e) { showErr(e); }
    }));

    // 场景
    const drawScenes = () => {
      $('#ce-scenes').innerHTML = d.scenes.length ? d.scenes.map((s, i) => `<span class="chip-x">${esc(s)}<button type="button" data-i="${i}">×</button></span>`).join('') : '<span class="faint">还没标</span>';
      $$('#ce-scenes [data-i]').forEach(b => b.addEventListener('click', () => { d.scenes.splice(+b.dataset.i, 1); dirty(); drawScenes(); }));
    };
    drawScenes();
    $('#ce-scene-in').addEventListener('keydown', e => {
      if (e.key !== 'Enter' || e.isComposing) return;
      e.preventDefault();
      const v = e.target.value.trim();
      if (v && !d.scenes.includes(v)) { d.scenes.push(v); dirty(); drawScenes(); }
      e.target.value = '';
    });

    // AI 候选不写入 script_refs，只有销售确认使用及版本才记录；旧关联保留并提示核对。
    const drawRefs = () => {
      $('#ce-refs').innerHTML = d.script_refs.length ? d.script_refs.map((r, i) => `<span class="chip-x">${scriptRef(r.id, r.version)}<button type="button" data-i="${i}">×</button></span>`).join('') : '<span class="faint">还没标</span>';
      $$('#ce-refs [data-i]').forEach(b => b.addEventListener('click', () => { d.script_refs.splice(+b.dataset.i, 1); dirty(); drawRefs(); }));
      $('#ce-refs-note').innerHTML = d.script_refs.some(r => !r.confirmed)
        ? '<div class="hint">有旧版关联尚未核对：已保留在原有统计中，可能包含当时 AI 的建议。请重新选择并确认真实使用版本，没用过的请移除。</div>' : '';
    };
    drawRefs();
    let refRequest = 0, refVersions = [];
    const drawRefPreview = () => {
      const value = $('#ce-ref-version').value;
      $('#ce-ref-confirm').disabled = !value;
      const version = refVersions.find(v => Number(v.version) === Number(value));
      $('#ce-ref-preview').innerHTML = version ? replyMessagesHtml((version.snapshot || {}).reply || '') : value === 'unknown' ? '本案例会计入该话术总次数，版本单列为未知。' : '';
    };
    const selectRef = async id => {
      const request = ++refRequest;
      const s = scriptById(id);
      $('#ce-ref-add').value = s ? String(s.id) : '';
      $('#ce-ref-picker').hidden = !s;
      refVersions = [];
      if (!s) return;
      $('#ce-ref-version').innerHTML = '<option value="">正在读取版本…</option>';
      $('#ce-ref-status').textContent = '';
      drawRefPreview();
      let history = [], error;
      try { history = await DB.listScriptVersions(s.id); } catch (e) { error = e; }
      if (!document.body.contains(form) || request !== refRequest) return;
      refVersions = [{ version: s.version, snapshot: s }].concat(history.filter(v => Number(v.version) !== Number(s.version)));
      $('#ce-ref-version').innerHTML = '<option value="">请选择当时使用的版本</option><option value="unknown">版本未知</option>'
        + refVersions.map(v => `<option value="${Number(v.version)}">v${Number(v.version)}${Number(v.version) === Number(s.version) ? '（当前版本）' : '（历史版本）'}${v.changed_at ? ' · ' + esc(L.fmtDateCN(v.changed_at)) : ''}</option>`).join('');
      $('#ce-ref-status').textContent = error ? '历史版本读取失败，可重新选择话术重试，或确认版本未知。' : '确认已实际使用后再添加；再次选择已有话术可以修正版本。';
      drawRefPreview();
    };
    $('#ce-ref-add').addEventListener('change', e => selectRef(Number(e.target.value)));
    $('#ce-ref-version').addEventListener('change', drawRefPreview);
    $('#ce-ref-confirm').addEventListener('click', () => {
      const s = scriptById($('#ce-ref-add').value), value = $('#ce-ref-version').value;
      if (!s || !value || (value !== 'unknown' && !refVersions.some(v => Number(v.version) === Number(value)))) return;
      const ref = { id: s.id, version: value === 'unknown' ? null : Number(value), confirmed: true };
      const index = d.script_refs.findIndex(r => Number(r.id) === s.id);
      if (index >= 0) d.script_refs[index] = ref; else d.script_refs.push(ref);
      selectRef(null); dirty(); drawRefs();
    });

    const collect = () => {
      for (const k of ['seller', 'channel', 'cohort', 'case_date', 'profile', 'stuck_point', 'self_review']) d[k] = form[k].value.trim();
      d.review_week = form.review_week.value || L.weekStart(d.case_date);
      d.dialogue = L.normDialogue(d.dialogue).map(x => ({ role: x.role, text: L.maskSensitive(x.text) }));
      d.profile = L.maskSensitive(d.profile);
    };

    // AI 分析
    const drawAnalysis = () => {
      $('#ce-analysis').innerHTML = d.ai_analysis ? `<div class="card" style="background:var(--bg);border:none">${analysisHtml(d.ai_analysis, false, true)}</div>` : '';
      $$('#ce-analysis [data-use-candidate]').forEach(b => b.addEventListener('click', () => selectRef(Number(b.dataset.useCandidate))));
    };
    drawAnalysis();
    $('#ce-ai').addEventListener('click', () => withBusy($('#ce-ai'), '分析中…', async () => {
      collect();
      if (!d.dialogue.length && !d.stuck_point) return toast('先填聊天记录或者关键节点', true);
      try {
        const a = await AI.analyzeCase(d, S.scripts, S.bg, $('#ce-knowledge').value);
        if (!document.body.contains(form)) return;
        d.ai_analysis = a;
        d.scenes = uniq(d.scenes.concat(a.scenes));
        dirty();
        drawScenes(); drawAnalysis();
        toast('分析完成，场景已标注；话术候选需核对版本后手动确认');
      } catch (e) { showErr(e); }
    }));

    // 保存
    form.addEventListener('submit', e => {
      e.preventDefault();
      withBusy($('#ce-save'), '保存中…', async () => {
        collect();
        const errs = L.validateCase(d);
        if (errs.length) { $('#ce-err').innerHTML = `<div class="note err">${errs.map(esc).join('<br>')}</div>`; return; }
        $('#ce-err').innerHTML = '';
        try {
          let saved;
          try { saved = await DB.saveCase(d); } catch (err) {
            if (err.code !== 'conflict') throw err;
            const pick = await choiceBox('这个案例刚被别人改过', '可以放弃你的修改，看看别人改了什么；也可以用你的版本覆盖。',
              [{ label: '放弃我的修改', value: 'reload', cls: 'plain' }, { label: '用我的版本覆盖', value: 'force', cls: 'danger' }]);
            S.cases = await DB.listCases();
            if (pick !== 'force') { clearDirty('case'); location.hash = '#/case/' + d.id; return; }
            d.version = (caseById(d.id) || d).version;
            saved = await DB.saveCase(d);
          }
          upsert(S.cases, saved);
          clearDirty('case');
          toast('已保存');
          location.hash = '#/case/' + saved.id;
        } catch (err) { showErr(err); }
      });
    });
  }

  // ======================================================================
  // 复盘会
  // ======================================================================
  async function ensureReview(week) {
    const r = reviewByWeek(week);
    if (r) return r;
    try {
      const saved = await DB.saveReview({ id: null, week, summary: '', todos: [], ai_brief: null, version: 1 });
      upsert(S.reviews, saved);
      return saved;
    } catch (e) {
      if (e.code !== 'conflict') throw e;
      S.reviews = await DB.listReviews();
      return reviewByWeek(week);
    }
  }
  // 只改复盘会的某个部分，冲突规则同 updateCase
  async function updateReview(week, mutate, opts) {
    for (let attempt = 0; attempt < 2; attempt++) {
      const r = await ensureReview(week);
      const d = clone(r);
      mutate(d);
      try { const saved = await DB.saveReview(d); upsert(S.reviews, saved); return saved; } catch (e) {
        if (e.code !== 'conflict' || attempt) throw e;
        S.reviews = await DB.listReviews();
        if (opts && opts.ask && !(await askOverwrite(opts.ask))) return null;
      }
    }
  }

  function enterPresent() {
    document.body.classList.add('present');
    if (document.documentElement.requestFullscreen) document.documentElement.requestFullscreen().catch(() => {});
  }
  function exitPresent() {
    document.body.classList.remove('present');
    if (document.fullscreenElement) document.exitFullscreen().catch(() => {});
  }

  async function viewReview(weekParam) {
    const alive = viewToken();
    const thisWeek = L.weekStart(L.todayISO());
    let week = weekParam ? L.weekStart(weekParam) : thisWeek;
    if (!weekParam) {
      // 本周还没有案例、上周有，就默认看上周（复盘会常在周初开）
      const prev = L.addDays(thisWeek, -7);
      if (!S.cases.some(c => c.review_week === thisWeek) && S.cases.some(c => c.review_week === prev)) week = prev;
    }
    const weeks = uniq(S.cases.map(c => c.review_week).concat(S.reviews.map(r => r.week), thisWeek, week)).sort().reverse();
    const cases = () => S.cases.filter(c => c.review_week === week);
    const won = cases().filter(c => c.result === 'won').length;
    const prevReview = S.reviews.filter(r => r.week < week).sort((a, b) => (a.week < b.week ? 1 : -1))[0];

    app().innerHTML = `
      <div class="page-head">
        <div><h1>话术复盘会</h1>
          <div class="week-switch" style="margin-top:8px">
            <button class="btn plain sm" id="rv-prev">‹ 上一周</button>
            <select id="rv-week">${weeks.map(w => `<option value="${w}" ${w === week ? 'selected' : ''}>${esc(L.weekLabel(w))}${w === thisWeek ? '（本周）' : ''}</option>`).join('')}</select>
            <button class="btn plain sm" id="rv-next">下一周 ›</button>
          </div></div>
        <div class="row"><a class="btn plain" href="#/case/new">+ 提交案例</a><button class="btn ghost" id="rv-present">投屏模式</button></div>
      </div>
      <div class="stats">
        <div class="stat"><div class="v">${cases().length}</div><div class="k">本周案例 · 成交 ${won} / 流失 ${cases().length - won}</div></div>
        <div class="stat"><div class="v" id="rv-stat-disc"></div><div class="k">已讨论</div></div>
        <div class="stat"><div class="v" id="rv-stat-changes">–</div><div class="k">这次复盘改了几条话术</div></div>
        <div class="stat"><div class="v" id="rv-stat-bad">–</div><div class="k">话术助手被标「不好用」</div></div>
      </div>
      ${prevReview && (prevReview.todos || []).length ? `<div class="card" id="rv-prev-todos"></div>` : ''}
      <section class="section"><h2>① 会前 AI 分析 <small>先看整体，再逐个过案例</small></h2>${knowledgeScopeHtml('rv-knowledge')}<p class="hint" style="margin-bottom:12px">本周有多个训练营时，选择通用资料，或确认所选训练营资料适用于本次讨论。</p><div id="rv-brief"></div></section>
      <section class="section"><h2>② 逐个过案例 <small>同一场景下，成交和流失放在一起对比</small></h2><div id="rv-cases"></div></section>
      <section class="section"><h2>③ 话术助手里「不好用」的提问 <small>通常说明话术库缺了什么</small></h2><div id="rv-bad"><span class="muted">加载中…</span></div></section>
      <section class="section"><h2>④ 这次复盘改了哪些话术</h2><div id="rv-changes"><span class="muted">加载中…</span></div></section>
      <section class="section"><h2>⑤ 复盘总结和待办 <small>待办会在下次复盘会开头检查</small></h2><div id="rv-summary"></div></section>`;

    bindKnowledgeScope('rv-knowledge');
    $('#rv-week').addEventListener('change', e => { location.hash = '#/review/' + e.target.value; });
    $('#rv-prev').addEventListener('click', () => { location.hash = '#/review/' + L.addDays(week, -7); });
    $('#rv-next').addEventListener('click', () => { location.hash = '#/review/' + L.addDays(week, 7); });
    $('#rv-present').addEventListener('click', () => (document.body.classList.contains('present') ? exitPresent() : enterPresent()));
    const review = () => reviewByWeek(week);
    const stats = () => {
      const cs = cases();
      $('#rv-stat-disc').textContent = `${cs.filter(c => c.review_status !== 'pending').length}/${cs.length}`;
    };
    stats();

    // 上次的待办
    const drawPrevTodos = () => {
      if (!alive()) return;
      const box = $('#rv-prev-todos');
      if (!box) return;
      const r = reviewByWeek(prevReview.week);
      const todos = r.todos || [];
      box.innerHTML = `<h3>上次复盘（${esc(L.weekLabel(r.week))}）的待办 <small>完成 ${todos.filter(t => t.done).length}/${todos.length}</small></h3>
        ${todos.map((t, i) => `<label class="row" style="padding:3px 0;cursor:pointer"><input type="checkbox" data-i="${i}" ${t.done ? 'checked' : ''}>
          <span style="${t.done ? 'text-decoration:line-through;color:var(--muted)' : ''}">${esc(t.text)}</span>${t.owner ? `<span class="tag gray">${esc(t.owner)}</span>` : ''}</label>`).join('')}`;
      $$('[data-i]', box).forEach(cb => cb.addEventListener('change', async () => {
        const target = L.todoTarget(r, +cb.dataset.i), done = cb.checked;
        $$('input', box).forEach(input => { input.disabled = true; });
        try { await updateReview(r.week, x => { x.todos = L.setTodoDone(x, target, done); }); }
        catch (e) { showErr(e); }
        finally { drawPrevTodos(); }
      }));
    };
    if (prevReview) drawPrevTodos();

    // ① 会前 AI 分析
    let badLogs = [];
    let badLogsReady = false, badLogsError = null;
    const drawBrief = () => {
      if (!alive()) return;
      const box = $('#rv-brief');
      const b = review() && review().ai_brief;
      if (!b) {
        box.innerHTML = `<div class="card"><p class="muted" style="margin-bottom:10px">${cases().length
          ? `让 AI 先把这周 ${cases().length} 个案例过一遍：按场景对比成交和流失的做法，并提出话术修改建议，会上逐条讨论决定改不改。`
          : '这周还没有案例。销售提交案例后，就可以生成会前分析。'}</p>
          <button class="btn" id="rv-gen" ${cases().length ? '' : 'disabled'}>生成会前分析</button></div>`;
      } else {
        const open = b.proposals.filter(p => p.state === 'open').length;
        box.innerHTML = `<div class="card">
          ${SalesDB.mode === 'demo' ? '<div class="note warn" style="margin-bottom:10px">演示模式：下面是本地模拟的分析，接入 AI 后才是真实结果</div>' : ''}
          ${contextUsedHtml(b)}
          ${knowledgeCitationsHtml(b, b.context_used)}
          ${salesInternalHtml(b.internal_notes)}
          <div class="pre">${esc(b.overview)}</div>
          ${b.scenes.length ? `<h3 style="margin-top:14px">分场景对比</h3>${b.scenes.map(s => `<div style="margin-bottom:8px"><b>${esc(s.scene)}</b>：${esc(s.contrast)}${s.takeaway ? `<div class="muted">→ ${esc(s.takeaway)}</div>` : ''}</div>`).join('')}` : ''}
          ${b.questions.length ? `<h3 style="margin-top:14px">值得会上讨论的问题</h3><ul style="padding-left:18px">${b.questions.map(q => `<li>${esc(q)}</li>`).join('')}</ul>` : ''}
          <div class="faint" style="margin-top:12px">生成于 ${esc(L.fmtDateTime(b.generated_at))} · <button class="linkbtn" id="rv-gen">重新生成</button></div>
        </div>
        ${b.proposals.length ? `<h3 style="margin-top:16px">话术修改建议 <small>${open ? `还有 ${open} 条没定` : '都处理完了'}</small></h3>${b.proposals.map(proposalHtml).join('')}` : ''}`;
        $$('[data-apply]', box).forEach(btn => btn.addEventListener('click', () => applyProposal(btn.dataset.apply)));
        $$('[data-reject]', box).forEach(btn => btn.addEventListener('click', () => setProposal(btn.dataset.reject, { state: 'rejected' })));
        $$('[data-reopen]', box).forEach(btn => btn.addEventListener('click', () => setProposal(btn.dataset.reopen, { state: 'open' })));
      }
      const gen = $('#rv-gen', box);
      if (gen) gen.disabled = !cases().length || !badLogsReady;
      if (gen) gen.addEventListener('click', async () => {
        if (!badLogsReady) return toast('请先成功加载本周的反馈，再生成分析', true);
        if (b && !(await confirmBox('重新生成会前分析？', '现有的分析和建议的处理状态会被替换（已经改过的话术不受影响）。', '重新生成'))) return;
        await withBusy($('#rv-gen'), '分析中，大约要半分钟…', async () => {
          try {
            const brief = await AI.reviewBrief({ cases: cases(), scripts: S.scripts, badLogs, bg: S.bg, knowledgeProject: $('#rv-knowledge').value });
            await updateReview(week, r => { r.ai_brief = brief; });
            drawBrief();
          } catch (e) { showErr(e); }
        });
      });
    };
    const proposalHtml = p => {
      const s = scriptById(p.applied_script_id || p.script_id);
      const prepared = prepareCustomerReply(p.new_reply, p.internal_notes);
      const blocked = prepared.withheld_messages.length > 0;
      const stateTag = p.state === 'applied' ? `<span class="badge won">已采纳${p.applied_version ? ' · v' + p.applied_version : ''}</span>` : p.state === 'rejected' ? '<span class="badge discussed">不采纳</span>' : '';
      return `<div class="proposal ${p.state}">
        <div class="row between"><div class="row"><b>${esc(p.scene || (s && s.scene) || '')}</b>${s ? `<span class="muted">改 ${scriptRef(s.id)}</span>` : p.target_note ? '<span class="tag gray">需核对目标</span>' : '<span class="tag">新增话术</span>'}${stateTag}</div>
          <div class="row">${p.state === 'open'
            ? `<button class="btn sm" data-apply="${esc(p.key)}" ${blocked ? 'disabled' : ''}>${s ? '采纳，去改话术' : p.target_note ? '先确认修改目标' : '采纳，加进话术库'}</button><button class="btn plain sm" data-reject="${esc(p.key)}">不采纳</button>`
            : `<button class="linkbtn" data-reopen="${esc(p.key)}">重新讨论</button>`}</div></div>
        ${p.applied_script_id ? '<div class="hint">此前采纳的话术仍保留在话术库；重新讨论后再次采纳会修改同一条话术。需要恢复内容时，可在话术详情中恢复历史版本。</div>' : ''}
        ${p.problem ? `<div class="muted" style="margin:6px 0">${esc(p.problem)}</div>` : ''}
        ${p.target_note ? `<div class="note warn">${esc(p.target_note)}</div>` : ''}
        ${blocked ? '<div class="note warn">建议正文含有内部说明或待核实内容，请重新生成会前分析后再采纳。</div>' : `<div class="message-preview"><div class="label">建议的发送消息</div>${replyMessagesHtml(prepared.text)}</div>`}
        ${salesInternalHtml(prepared.internal_notes, blocked ? `<details class="original-reply"><summary>查看需要整理的原稿</summary><div class="pre">${esc(p.new_reply)}</div></details>` : '')}
        ${knowledgeCitationsHtml(p, review() && review().ai_brief && review().ai_brief.context_used)}
        ${s && !blocked ? replyChangesHtml(s.reply, p.new_reply) : ''}
        <div class="why">依据：${esc(p.reason)}${p.case_ids.length ? ' · ' + p.case_ids.map(id => `<a href="#/case/${id}">案例#${id}</a>`).join(' ') : ''}</div>
      </div>`;
    };
    const setProposal = async (key, patch) => {
      try {
        await updateReview(week, r => {
          const p = r.ai_brief && r.ai_brief.proposals.find(x => x.key === key);
          if (!p) throw DB.makeError('conflict', '这条建议已被替换，请核对最新的会前分析');
          Object.assign(p, patch);
        });
        drawBrief();
        return true;
      } catch (e) { showErr(e); return false; }
    };
    const applyingProposals = new Set();
    const chooseProposalTarget = p => new Promise(resolve => {
      const m = openModal(`<h2>先确认修改目标</h2><p class="note warn">${esc(p.target_note)}</p>
        <div class="field" style="margin-top:12px"><label for="proposal-target">要修改的话术</label><select id="proposal-target"><option value="">请选择</option><option value="new">明确新增一条话术</option>${S.scripts.filter(s => s.status !== 'archived').map(s => `<option value="${s.id}">${esc(s.scene)} · ${esc((L.CHANNELS[s.channel] || L.CHANNELS.all).label)} · v${s.version}</option>`).join('')}</select></div>
        <div class="modal-actions"><button type="button" class="btn plain" data-close>取消</button><button type="button" class="btn" id="proposal-confirm" disabled>确认目标，继续编辑</button></div>`, {
        small: true, onCancel: () => { m.close(); resolve(null); },
      });
      const select = $('#proposal-target', m.el), confirm = $('#proposal-confirm', m.el);
      select.addEventListener('change', () => { confirm.disabled = !select.value; });
      confirm.addEventListener('click', () => { const value = select.value; if (!value) return; m.close(); resolve(value); });
    });
    const applyProposal = async key => {
      if (applyingProposals.has(key)) return;
      const b = review() && review().ai_brief;
      const p = b && b.proposals.find(x => x.key === key);
      if (!p) { toast('这条建议已经不在了，可能有人重新生成了分析', true); drawBrief(); return; }
      if (p.state !== 'open') return;
      if (prepareCustomerReply(p.new_reply, p.internal_notes).withheld_messages.length) return toast('建议正文含有内部说明，请重新生成会前分析后再采纳', true);
      applyingProposals.add(key);
      try {
        let s = W.proposalTarget(p, S.scripts);
        if (p.target_note && !p.applied_script_id) {
          const target = await chooseProposalTarget(p);
          if (!target) return;
          s = target === 'new' ? null : scriptById(Number(target));
          if (target !== 'new' && !s) throw DB.makeError('not_found', '所选话术已不存在，请重新选择');
        }
        const r = await ensureReview(week);
        const saved = await openScriptEditor(s, {
          prefill: prepareAiScriptPrefill({
            stage: p.stage || (s ? s.stage : ''), scene: s ? s.scene : p.scene,
            customer_says: p.new_customer_says || (s ? s.customer_says : ''), reply: p.new_reply,
            tips: p.new_tips || (s ? s.tips : ''), internal_notes: p.internal_notes,
          }),
          change_note: [p.problem, p.reason].filter(Boolean).join('\n'),
          hint: p.target_note ? (s ? '已选择修改「' + s.scene + '」，请核对内容后保存。' : '已选择新增话术，请核对内容后保存。') : 'AI 建议的版本已经填好，可以按会上讨论的结果再改。',
          allowUnchanged: !!p.applied_script_id,
          review_id: r.id, case_ids: p.case_ids,
        });
        if (saved) {
          const application = { applied_script_id: saved.id, applied_version: saved.version };
          if (p.target_note) Object.assign(application, { script_id: s ? s.id : null, target_note: '' });
          // 先保留已保存的话术编号，状态写入失败后本页重试也不会重复新增。
          Object.assign(p, application);
          const synced = await setProposal(key, Object.assign({ state: 'applied' }, application));
          if (!synced) {
            const latest = review() && review().ai_brief;
            const current = latest && latest.proposals.find(x => x.key === key);
            if (current) Object.assign(current, application);
            drawBrief();
            toast('话术已保存，采纳状态未保存成功；可再次采纳以重试。', true);
          }
          drawChanges();
        }
      } catch (e) { showErr(e); } finally { applyingProposals.delete(key); }
    };
    drawBrief();

    // ② 逐个过案例
    const drawCases = () => {
      if (!alive()) return;
      const box = $('#rv-cases');
      // 保留没保存的结论草稿和展开状态
      const drafts = {}, opened = new Set();
      $$('[data-case]', box).forEach(el => {
        if (!el.dataset.fresh) drafts[el.dataset.case] = $('[data-disc]', el).value;
        if (el.open) opened.add(el.dataset.case);
      });
      const cs = cases();
      const earlier = S.cases.filter(c => c.review_week < week && c.review_status === 'pending');
      const groups = L.groupCasesByScene(cs);
      const card = c => `<details class="rc ${c.review_status !== 'pending' ? 'done' : ''}" data-case="${c.id}" ${opened.has(String(c.id)) ? 'open' : ''}>
          <summary><div class="case-meta"><b>${esc(c.seller)}</b>${chTag(c.channel)}${c.cohort ? `<span>${esc(c.cohort)}</span>` : ''}<span class="faint">#${c.id}</span><span class="spacer"></span>${statusBadge(c.review_status)}</div>
            ${c.stuck_point ? `<div class="case-point">${esc(c.stuck_point)}</div>` : ''}</summary>
          <div class="rc-body">
            ${L.normDialogue(c.dialogue).length ? `<div class="chat-scroll" style="max-height:340px">${chatHtml(c.dialogue)}</div>` : ''}
            ${c.self_review ? `<div class="detail-block" style="margin-top:10px"><div class="label">销售自己的复盘</div><div class="pre">${esc(c.self_review)}</div></div>` : ''}
            ${(c.script_refs || []).length ? `<div class="detail-block" style="margin-top:10px"><div class="label">用到的话术</div>${c.script_refs.map(r => scriptRef(r.id, r.version)).join('、')}</div>` : ''}
            ${c.ai_analysis ? `<details style="margin-top:10px"><summary class="muted" style="cursor:pointer">AI 分析</summary>${analysisHtml(c.ai_analysis)}</details>` : ''}
            <div class="field" style="margin:12px 0 8px"><label>会上的结论</label><textarea rows="3" data-disc placeholder="这个案例给我们的启发、下次怎么做">${esc(drafts[c.id] != null ? drafts[c.id] : c.discussion)}</textarea></div>
            <div class="row"><button class="btn sm" data-act="save">保存结论</button><button class="btn plain sm" data-act="fix">据此改话术</button>
              <button class="linkbtn" data-act="skip">${c.review_status === 'skipped' ? '恢复待复盘' : '这个不讨论了'}</button><span class="spacer"></span><a href="#/case/${c.id}" style="font-size:13px">打开案例</a></div>
          </div>
        </details>`;
      box.innerHTML = (cs.length ? groups.map(g => {
        const libs = S.scripts.filter(s => s.status !== 'archived' && s.scene === g.scene);
        return `<div class="scene-group">
          <div class="row between"><h3 style="margin:0">${esc(g.scene)} <small>成交 ${g.won.length} · 流失 ${g.lost.length}</small></h3>
            <div class="row" style="font-size:13px">${libs.map(s => `话术库：${scriptRef(s.id)}`).join(' · ')}</div></div>
          <div class="vs"><div class="vs-col won"><h4>成交</h4>${g.won.map(card).join('') || '<div class="faint">没有</div>'}</div>
            <div class="vs-col lost"><h4>流失</h4>${g.lost.map(card).join('') || '<div class="faint">没有</div>'}</div></div>
        </div>`;
      }).join('') : '<div class="card empty" style="padding:30px">这周还没有案例<br><a href="#/case/new">去提交一个</a></div>')
        + (earlier.length ? `<details class="card" style="margin-top:14px"><summary style="cursor:pointer"><b>之前几周还没讨论的案例（${earlier.length} 个）</b></summary>
            <div class="case-list" style="margin-top:12px">${earlier.map(c => `<div class="row" style="align-items:stretch"><div style="flex:1">${caseCard(c)}</div><button class="btn plain sm" data-move="${c.id}" style="align-self:center">移到这周</button></div>`).join('')}</div></details>` : '');

      $$('[data-case]', box).forEach(el => {
        const id = Number(el.dataset.case);
        $('[data-disc]', el).addEventListener('input', () => markDirty('disc' + id));
        $$('[data-act]', el).forEach(btn => btn.addEventListener('click', async () => {
          const act = btn.dataset.act;
          const text = $('[data-disc]', el).value.trim();
          if (act === 'fix') {
            const r = await ensureReview(week).catch(showErr);
            const saved = await fixScriptFromCase(caseById(id), { review_id: r && r.id, change_note: text ? `案例#${id} 复盘结论：${text}` : '' });
            if (saved) drawChanges();
            return;
          }
          let lost = '';
          await withBusy(btn, '保存中…', async () => {
            try {
              const saved = await updateCase(id, d => {
                if (act === 'save') { d.discussion = text; d.review_status = 'discussed'; }
                else d.review_status = d.review_status === 'skipped' ? 'pending' : 'skipped';
              }, act === 'save' ? { ask: '这个案例的结论' } : null);
              if (act === 'save') {
                clearDirty('disc' + id);
                el.dataset.fresh = '1'; // 重画时用数据库里的结论，不用草稿
                if (saved) toast('已保存'); else lost = text;
              }
            } catch (e) { showErr(e); }
          });
          drawCases(); stats();
          if (lost) { const card = $(`[data-case="${id}"] [data-disc]`); if (card) card.insertAdjacentHTML('afterend', unsavedNote(lost)); }
        }));
      });
      $$('[data-move]', box).forEach(btn => btn.addEventListener('click', async () => {
        try { await updateCase(Number(btn.dataset.move), d => { d.review_week = week; }); drawCases(); stats(); } catch (e) { showErr(e); }
      }));
    };
    drawCases();

    // ③ 话术助手的「不好用」
    const drawBad = () => {
      if (!alive()) return;
      const box = $('#rv-bad');
      if (badLogsError) {
        $('#rv-stat-bad').textContent = '—';
        box.innerHTML = `<div class="note err">本周反馈加载失败：${esc(badLogsError.message)} <button class="btn plain sm" id="rv-bad-retry">重试</button></div>`;
        $('#rv-bad-retry').addEventListener('click', loadBadLogs);
        return;
      }
      if (!badLogsReady) {
        $('#rv-stat-bad').textContent = '—';
        box.innerHTML = '<span class="muted">加载中…</span>';
        return;
      }
      $('#rv-stat-bad').textContent = badLogs.length;
      box.innerHTML = badLogs.length ? `<div class="card">${badLogs.map((l, i) => `<div class="log-item" style="cursor:default">
          <div class="row between"><div><b>客户说：</b>${esc(l.message)}</div><button class="btn plain sm" data-add="${i}">补一条话术</button></div>
          <div class="faint">${l.feedback_note ? '反馈：' + esc(l.feedback_note) + ' · ' : ''}${esc(l.created_by_name)} · ${esc(L.fmtDateTime(l.created_at))}${l.result && l.result.scene ? ' · AI 判断的场景：' + esc(l.result.scene) : ''}</div>
        </div>`).join('')}</div>` : '<div class="faint">这周没有</div>';
      $$('[data-add]', box).forEach(btn => btn.addEventListener('click', async () => {
        const l = badLogs[+btn.dataset.add];
        try {
          const existing = l.result && l.result.scene ? S.scripts.find(s => s.status !== 'archived' && s.scene === l.result.scene) : null;
          const prefill = existing ? {} : prepareAiScriptPrefill(l.result && l.result.new_script, { customer_says: l.message });
          if (!prefill) return toast('这份草稿含有内部说明，请在话术助手重新生成后再补进话术库', true);
          const r = await ensureReview(week);
          const saved = await openScriptEditor(existing || null, {
            prefill,
            change_note: `话术助手里被标记不好用：「${l.message}」${l.feedback_note ? '，反馈：' + l.feedback_note : ''}`,
            hint: existing ? `客户原话：${l.message}` : '',
            review_id: r.id,
          });
          if (saved) drawChanges();
        } catch (e) { showErr(e); }
      }));
    };

    // ④ 这次复盘改了哪些话术
    async function drawChanges() {
      const box = $('#rv-changes');
      if (!box || !alive()) return;
      const r = review();
      let list = [];
      if (r) { try { list = await DB.listVersionsByReview(r.id); } catch (e) { if (alive()) box.innerHTML = `<div class="note err">${esc(e.message)}</div>`; return; } }
      if (!alive()) return;
      $('#rv-stat-changes').textContent = uniq(list.map(v => v.script_id)).length;
      box.innerHTML = list.length ? `<div class="card"><table class="mini" style="width:100%"><tr><th>话术</th><th>版本</th><th>为什么改</th><th>谁改的</th></tr>
          ${list.map(v => `<tr><td>${scriptById(v.script_id) ? `<a href="#/script/${v.script_id}">${esc(v.snapshot.scene)}</a>` : esc(v.snapshot.scene)}</td><td>v${v.version}${v.action === 'insert' ? ' 新增' : ''}</td>
            <td class="pre">${esc(v.change_note)}</td><td class="faint">${esc(v.changed_by_name)}<br>${esc(L.fmtDateTime(v.changed_at))}</td></tr>`).join('')}</table></div>`
        : '<div class="faint">还没有。会上讨论出结论后，点案例下面的「据此改话术」或者 AI 建议的「采纳」，改动都会记在这里。</div>';
    }
    drawChanges();

    // ⑤ 总结和待办
    const r0 = review();
    const createTodoId = () => window.crypto && window.crypto.randomUUID ? window.crypto.randomUUID() : `todo_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    let todos = clone((r0 && r0.todos) || []);
    const drawSummary = () => {
      if (!alive()) return;
      const r = review();
      $('#rv-summary').innerHTML = `<div class="card">
        <div class="field"><label>本次复盘总结</label><textarea id="rv-sum" rows="4" placeholder="这周最重要的发现、大家达成的共识">${esc(r ? r.summary : '')}</textarea></div>
        <div class="label">待办</div>
        <div id="rv-todos"></div>
        <div class="row" style="margin-top:6px"><button class="btn plain sm" id="rv-todo-add">+ 加一条待办</button><span class="spacer"></span>
          ${r && r.version > 1 ? `<span class="faint">${esc(r.updated_by_name)} 保存于 ${esc(L.fmtDateTime(r.updated_at))}</span>` : ''}<button class="btn" id="rv-sum-save">保存总结和待办</button></div>
      </div>`;
      const drawTodos = () => {
        $('#rv-todos').innerHTML = todos.length ? todos.map((t, i) => `<div class="todo ${t.done ? 'done' : ''}">
            <input type="checkbox" data-i="${i}" data-k="done" ${t.done ? 'checked' : ''}>
            <input type="text" data-i="${i}" data-k="text" value="${esc(t.text)}" placeholder="要做的事，如：下周试用新版「嫌贵」话术">
            <input type="text" class="owner" data-i="${i}" data-k="owner" value="${esc(t.owner)}" placeholder="负责人">
            <button class="icon-btn" data-rm="${i}">×</button></div>`).join('') : '<div class="faint" style="margin-bottom:6px">还没有待办</div>';
        $$('#rv-todos [data-k]').forEach(el => el.addEventListener(el.type === 'checkbox' ? 'change' : 'input', () => {
          todos[+el.dataset.i][el.dataset.k] = el.type === 'checkbox' ? el.checked : el.value;
          if (el.type === 'checkbox') el.parentNode.classList.toggle('done', el.checked);
          markDirty('summary');
        }));
        $$('#rv-todos [data-rm]').forEach(b => b.addEventListener('click', () => { todos.splice(+b.dataset.rm, 1); markDirty('summary'); drawTodos(); }));
      };
      drawTodos();
      $('#rv-sum').addEventListener('input', () => markDirty('summary'));
      $('#rv-todo-add').addEventListener('click', () => {
        todos.push({ id: createTodoId(), text: '', owner: '', done: false });
        markDirty('summary');
        drawTodos();
        const ins = $$('#rv-todos input[data-k=text]');
        ins[ins.length - 1].focus();
      });
      $('#rv-sum-save').addEventListener('click', () => withBusy($('#rv-sum-save'), '保存中…', async () => {
        const summary = $('#rv-sum').value.trim();
        const clean = L.normTodos(todos, createTodoId);
        try {
          const saved = await updateReview(week, r => { r.summary = summary; r.todos = clean; }, { ask: '这周的总结和待办' });
          todos = clone((review() && review().todos) || []);
          clearDirty('summary');
          drawSummary();
          if (saved) toast('已保存');
          else {
            const mine = [summary, ...clean.map(t => `☐ ${t.text}${t.owner ? '（' + t.owner + '）' : ''}`)].filter(Boolean).join('\n');
            $('#rv-sum').insertAdjacentHTML('afterend', unsavedNote(mine));
          }
        } catch (e) { showErr(e); }
      }));
    };
    drawSummary();

    // 完整读取本周反馈；读取失败不伪装成“没有反馈”。
    async function loadBadLogs() {
      badLogsReady = false; badLogsError = null;
      drawBad();
      const gen = $('#rv-gen');
      if (gen) gen.disabled = true;
      try {
        badLogs = await DB.listAiLogs({
          since: new Date(week + 'T00:00:00+08:00').toISOString(),
          until: new Date(L.addDays(week, 7) + 'T00:00:00+08:00').toISOString(),
          feedback: 'bad', all: true,
        });
        badLogsReady = true;
      } catch (e) { badLogsError = e; }
      if (alive()) {
        drawBad();
        const currentGen = $('#rv-gen');
        if (currentGen) currentGen.disabled = !cases().length || !badLogsReady;
      }
    }
    await loadBadLogs();
  }

  // ======================================================================
  // 演示模式：示例数据（都是编的通用例子）
  // ======================================================================
  async function fillDemo() {
    if (S.scripts.length || S.cases.length) {
      if (!(await confirmBox('填入示例数据', '会在现有演示数据后面追加几条示例话术和案例。', '继续'))) return;
    }
    const SCRIPTS = [
      { stage: '邀约入群', scene: '问是不是要收费', customer_says: '进群要钱吗？\n是不是要卖课', reply: '进群是免费的哈，3 天体验营全程不收费，群里每天有学习任务和直播答疑。你先跟着学，觉得有用再说～', tips: '先打消防备，别一上来就讲正价课' },
      { stage: '社群预热', scene: '没时间学', customer_says: '最近太忙了，怕跟不上', reply: '理解，大部分同学都是上班族。体验营每天的任务大概 20 分钟，直播有回放，下班后看也来得及。\n你一般几点比较有空？我把每天的任务提前发你', tips: '给出具体时长，再问对方的时间，推动对方做出承诺' },
      { stage: '直播中', scene: '担心零基础学不会', customer_says: '我完全零基础能学会吗', reply: '可以的，营里一半以上同学都是零基础。课程从装软件开始一步步带，每节课都有助教答疑，作业卡住了随时在群里问', tips: '用比例和服务细节回答，别只说「肯定能」' },
      { stage: '直播后追单', scene: '嫌贵', customer_says: '太贵了\n能不能便宜点', reply: '能理解，这个价格确实要好好考虑。方便说说你主要是担心学完用不上，还是这段时间预算比较紧？\n如果是担心效果，我发你几位和你情况差不多的学员现在的成果，你看看', tips: '先问清楚是哪种「贵」，别直接降价；不承诺额外优惠' },
      { stage: '直播后追单', scene: '再考虑一下', customer_says: '我再考虑考虑\n回头再说', reply: '好的～方便说下主要在犹豫哪一点吗？课程内容、时间安排还是价格？我把对应的信息整理给你，你考虑起来也方便', tips: '把模糊的「考虑」变成具体的顾虑' },
      { stage: '定金转尾款', scene: '交了定金不补尾款', customer_says: '定金交了，尾款我再想想', reply: '定金已经帮你锁定名额啦。尾款截止是【尾款截止时间】，开营前会先发预习资料。\n你现在主要卡在哪一步？我帮你一起看看', tips: '提醒截止时间，但别施压' },
    ];
    try {
      const ids = {};
      for (const s of SCRIPTS) {
        const saved = await DB.saveScript(Object.assign(L.emptyScript(), s, { change_note: '示例数据' }));
        ids[s.scene] = saved;
      }
      const thisWeek = L.weekStart(L.todayISO());
      const day = n => L.addDays(thisWeek, n);
      const ref = scene => [{ id: ids[scene].id, version: ids[scene].version }];
      const CASES = [
        { result: 'won', seller: '销售A', channel: 'xiaoe', cohort: '示例第1期', case_date: day(1), scenes: ['嫌贵'], script_refs: ref('嫌贵'),
          stuck_point: '客户说贵之后，先问清楚是担心用不上，再发了两个同行业学员的成果，当天补了尾款',
          dialogue: [{ role: 'customer', text: '课是挺好的，就是有点贵' }, { role: 'sales', text: '能理解，这个价格确实要好好考虑。方便说说你主要是担心学完用不上，还是这段时间预算比较紧？' },
            { role: 'customer', text: '主要怕学完用不上，我是做行政的' }, { role: 'sales', text: '明白～我发你两位也是做行政的学员现在做的东西，你看看' },
            { role: 'note', text: '[销售发了两张学员作品截图]' }, { role: 'customer', text: '哦这个挺实用的，那我补尾款吧' }] },
        { result: 'lost', seller: '销售B', channel: 'xiaoe', cohort: '示例第1期', case_date: day(1), scenes: ['嫌贵'], script_refs: [],
          stuck_point: '客户说贵，直接给了限时优惠，客户觉得被催单，之后就不回了',
          self_review: '当时太急了，应该先问问客户为什么觉得贵',
          dialogue: [{ role: 'customer', text: '太贵了吧' }, { role: 'sales', text: '今天下单可以再减 200，名额只剩最后几个了！' },
            { role: 'customer', text: '我再看看吧' }, { role: 'sales', text: '优惠今晚 12 点就结束了哦' }, { role: 'note', text: '[客户之后没有再回复]' }] },
        { result: 'lost', seller: '销售A', channel: 'bilibili', cohort: '示例B站专场', case_date: day(2), scenes: ['再考虑一下'], script_refs: [],
          stuck_point: '客户说再考虑，只回了「好的」，没有追问具体在犹豫什么',
          dialogue: [{ role: 'customer', text: '我再考虑考虑' }, { role: 'sales', text: '好的，有问题随时找我' }] },
        { result: 'won', seller: '销售B', channel: 'bilibili', cohort: '示例B站专场', case_date: day(2), scenes: ['担心零基础学不会'], script_refs: ref('担心零基础学不会'),
          stuck_point: '用「一半以上是零基础」+ 助教答疑打消了顾虑',
          dialogue: [{ role: 'customer', text: '我一点基础都没有，能学会吗' }, { role: 'sales', text: '可以的，营里一半以上同学都是零基础，每节课都有助教答疑' }, { role: 'customer', text: '那我报一个' }] },
        { result: 'lost', seller: '销售B', channel: 'xiaoe', cohort: '示例第0期', case_date: L.addDays(thisWeek, -5), scenes: ['没时间学'], script_refs: [],
          stuck_point: '客户说忙，销售没有给出每天具体要花多少时间', dialogue: [{ role: 'customer', text: '最近太忙了' }, { role: 'sales', text: '那等你有空再说～' }] },
      ];
      for (const c of CASES) {
        await DB.saveCase(Object.assign(L.emptyCase(c.seller), c, { review_week: L.weekStart(c.case_date) }));
      }
      const log = await DB.createAiLog({ message: '你们和网上的免费教程有什么区别', context: '', channel: 'all', result: { scene: '未匹配', intent: '', replies: [], gap: true } });
      await DB.setAiFeedback(log.id, 'bad', '话术库里没有这个问题');
      await loadAll();
      toast('已填入示例数据');
      route();
    } catch (e) { showErr(e); }
  }

  boot().catch(renderError);
})();
