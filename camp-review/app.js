/* 训练营周复盘 · 界面层：路由、页面渲染、事件
   计算全部交给 logic.js（window.CampLogic），读写交给 db.js（window.CampDB） */
(function () {
  'use strict';

  const L = window.CampLogic;
  const R = window.CampReview;
  const ReviewUI = window.CampReviewUI;
  const DB = window.CampDB;
  const $app = document.getElementById('app');

  const DEPARTMENTS = ['内容', '运营', '交付', '产品', '其他'];   // 与 skill-hub 一致
  const CH_ORDER = ['xiaoe', 'bilibili'];
  const KINDS = ['invites', 'community', 'conversions'];
  const TEXT_KEYS = new Set(['date', 'slot', 'channel', 'group']);
  const NULL_TOKEN_RE = /^(?:[\/\-—–_~]+|空|无|暂无|null|n\/a|na)$/i;   // 与 logic.js 的占位符一致
  const WRITABLE = ['id', 'version', 'channel', 'name', 'live_date', 'closed', 'summary', 'data'];
  const HIST_ACTION = { insert: '新建', update: '修改', delete: '删除' };

  const state = {
    me: null,              // { user:{id,email}, profile:{display_name,department} }
    access: null,          // checkAccess() 的结果
    cohorts: [],
    loaded: false,
    dash: { week: '', channel: 'all' },
    listChannel: 'all',
    lessons: { channel: 'all', query: '' },
    openDetails: new Set(),
    dirty: false,          // 有未保存的修改
    ed: null,              // 录入 / 编辑页状态
    present: false,
    reviewSaving: false,
  };
  let charts = [];
  let currentHash = location.hash || '#/';
  let routeSeq = 0;
  const mCache = new Map();

  // ---------- 小工具 ----------
  const esc = s => String(s ?? '').replace(/[&<>"']/g,
    c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const clone = o => JSON.parse(JSON.stringify(o));
  const isYMD = s => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ''));
  const pad = n => String(n).padStart(2, '0');
  const sameId = (a, b) => a != null && b != null && String(a) === String(b);

  function safe(fn, fallback) {
    try { return fn(); } catch (e) { console.error('[camp-review]', e); return fallback; }
  }
  const pct = x => safe(() => L.fmtPct(x), '—');
  const int = n => safe(() => L.fmtInt(n), n == null ? '—' : String(n));

  function flash(msg, ms = 2600) {
    const el = document.getElementById('flash');
    el.textContent = msg; el.style.display = 'block';
    clearTimeout(el._t); el._t = setTimeout(() => { el.style.display = 'none'; }, ms);
  }

  function dateCN(d) { return isYMD(d) ? safe(() => L.fmtDateCN(d), d) : (d || '—'); }
  function dateWeek(d) {
    if (!isYMD(d)) return d || '未填日期';
    return `${dateCN(d)} ${safe(() => L.weekdayCN(d), '')}`.trim();
  }
  function shortMD(d) { return isYMD(d) ? `${+d.slice(5, 7)}/${+d.slice(8, 10)}` : String(d || ''); }
  function fmtTime(iso) {
    if (!iso) return '';
    const d = new Date(iso);
    if (isNaN(d)) return '';
    return `${d.getMonth() + 1}月${d.getDate()}日 ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  }
  function addDays(ymd, n) {
    const [y, m, d] = ymd.split('-').map(Number);
    const t = new Date(Date.UTC(y, m - 1, d + n));
    return `${t.getUTCFullYear()}-${pad(t.getUTCMonth() + 1)}-${pad(t.getUTCDate())}`;
  }
  function yearOf(ymd) { return isYMD(ymd) ? Number(ymd.slice(0, 4)) : new Date().getFullYear(); }
  function weekOf(ymd) { return isYMD(ymd) ? safe(() => L.weekStart(ymd), '') : ''; }

  // ---------- 配置适配（logic.js 的列配置可能是 [key,label] 或 {key,label,type}） ----------
  function normCols(list) {
    return (list || []).map(c => Array.isArray(c)
      ? { key: c[0], label: c[1], type: TEXT_KEYS.has(c[0]) ? 'text' : 'int' }
      : { key: c.key, label: c.label, type: c.type || (TEXT_KEYS.has(c.key) ? 'text' : 'int') });
  }
  function chCfg(ch) {
    return (L.CHANNELS && L.CHANNELS[ch]) ||
      { key: ch, label: ch || '未知渠道', color: '#9ca3af', conversionColumns: [], attendanceLabel: '直播出勤人数' };
  }
  function chLabel(ch) { return chCfg(ch).label; }
  function attLabel(ch) { return chCfg(ch).attendanceLabel || '直播出勤人数'; }
  function colsFor(kind, ch) {
    if (kind === 'invites') return normCols(L.INVITE_COLUMNS);
    if (kind === 'community') {
      const raw = typeof L.COMMUNITY_COLUMNS === 'function' ? L.COMMUNITY_COLUMNS(ch) : L.COMMUNITY_COLUMNS;
      return normCols(raw).map(c => c.key === 'attendance' ? { ...c, label: attLabel(ch) } : c);
    }
    const cols = normCols(chCfg(ch).conversionColumns);
    return cols.some(c => c.key === 'date') ? cols : [{ key: 'date', label: '日期', type: 'text' }].concat(cols);
  }
  function convCols(ch) { return colsFor('conversions', ch).filter(c => c.key !== 'date'); }

  const CORE_FALLBACK = [
    { key: 'inviteRate', label: '邀约入群率' }, { key: 'attendRate', label: '直播到课率' },
    { key: 'liveConvRate', label: '直播转化率' }, { key: 'communityConvRate', label: '社群转化率' },
  ];
  function coreMetrics() {
    const list = Array.isArray(L.CORE_METRICS) && L.CORE_METRICS.length ? L.CORE_METRICS : CORE_FALLBACK;
    return list.map(m => {
      const key = typeof m === 'string' ? m : m.key;
      const fb = CORE_FALLBACK.find(f => f.key === key) || {};
      const label = typeof m === 'string' ? fb.label : (m.label || m.name || m.title || fb.label);
      return { key, label: label || key, formula: typeof m === 'object' ? m.formula : null };
    });
  }
  // 每个核心指标的分子 / 分母
  const PARTS = {
    inviteRate: m => [m.joinedKnown, m.reachKnown],
    attendRate: m => [m.attendance, m.members],
    liveConvRate: m => [m.liveDeals, m.attendance],
    communityConvRate: m => [m.totalDeals, m.members],
  };
  function formulaText(key, ch) {
    // 优先用 logic.js 里按渠道写好的说明
    const def = coreMetrics().find(d => d.key === key);
    const f = def && def.formula;
    if (typeof f === 'string' && f) return f;
    if (f && typeof f === 'object' && f[ch]) return f[ch];
    const x = ch !== 'bilibili';
    return ({
      inviteRate: '进群人数 ÷ 触达人数。只算填了触达人数的邀约渠道，朋友圈这类没有触达数的不计入。',
      attendRate: `${attLabel(ch)} ÷ 直播当天社群人数。`,
      liveConvRate: x ? '直播中付了定金、之后补齐尾款的人数 ÷ 直播出勤人数。'
        : '直播中全款成交人数 ÷ 直播观看人数。',
      communityConvRate: x ? '总成交 ÷ 直播当天社群人数。总成交 = 补齐尾款 + 1v1 追单（直接付款）+ 老学员。'
        : '总成交 ÷ 直播当天社群人数。总成交 = 直播全款成交 + 1v1 追单（直接付款）+ 老学员 − 退款。',
    })[key] || '';
  }

  // ---------- 数据整理 ----------
  function emptyRow(kind) {
    const r = safe(() => L.emptyRow(kind), null);
    if (r) return r;
    const keys = colsFor(kind, 'xiaoe').map(c => c.key);
    return Object.fromEntries(keys.map(k => [k, TEXT_KEYS.has(k) ? '' : null]));
  }
  function normCohort(c) {
    const src = c || {};
    const d = src.data || {};
    const out = { ...src, data: { ...d } };
    KINDS.forEach(k => {
      out.data[k] = (Array.isArray(d[k]) ? d[k] : []).map(r => Object.assign(emptyRow(k), r));
    });
    out.data.unparsed = Array.isArray(d.unparsed) ? d.unparsed.map(String) : [];
    out.name = out.name || '';
    out.summary = out.summary || '';
    out.closed = !!out.closed;
    out.live_date = out.live_date ? String(out.live_date).slice(0, 10) : '';
    return out;
  }
  const isEmptyRow = r => Object.values(r).every(v => v === null || v === undefined || String(v).trim() === '');
  function pickWritable(c) {
    const out = {};
    WRITABLE.forEach(k => { if (c[k] !== undefined) out[k] = c[k]; });
    return out;
  }
  function forSave(c) {
    const out = pickWritable(clone(c));
    out.name = String(out.name || '').trim();
    out.data = out.data || {};
    KINDS.forEach(k => { out.data[k] = (out.data[k] || []).filter(r => !isEmptyRow(r)); });
    return out;
  }
  function metricsOf(c) {
    const k = c.id != null ? `${c.id}|${c.version}|${c.updated_at}` : null;
    if (k && mCache.has(k)) return mCache.get(k);
    const m = safe(() => L.computeMetrics(c), null);
    if (k) mCache.set(k, m);
    return m;
  }
  function filterCh(list, ch) { return ch === 'all' ? list : list.filter(c => c.channel === ch); }
  function byChannelThenDate(a, b) {
    return (CH_ORDER.indexOf(a.channel) - CH_ORDER.indexOf(b.channel)) || String(a.live_date).localeCompare(String(b.live_date));
  }
  function replaceInCache(c) {
    const i = state.cohorts.findIndex(x => sameId(x.id, c.id));
    if (i >= 0) state.cohorts[i] = c; else state.cohorts.unshift(c);
  }
  async function loadCohorts() {
    const list = await DB.listCohorts();
    state.cohorts = (list || []).map(normCohort);
    state.loaded = true;
  }

  // ---------- 通用界面 ----------
  function chTag(ch) { return `<span class="ch-tag ch-${esc(ch)}"><i></i>${esc(chLabel(ch))}</span>`; }
  function statusBadge(m) { return `<span class="badge ${esc(m.status)}">${esc(m.statusText)}</span>`; }
  function loadingHTML(t) { return `<div class="loading">${esc(t || '加载中…')}</div>`; }
  function listHTML(items) { return `<ul>${items.map(t => `<li>${esc(t)}</li>`).join('')}</ul>`; }

  function openModal({ title, html, actions, cancelText }) {
    const root = document.getElementById('modal-root');
    root.innerHTML = `<div class="modal-mask"><div class="modal" role="dialog" aria-modal="true" aria-labelledby="modal-title">
      <h3 id="modal-title">${esc(title)}</h3><div class="modal-body">${html}</div>
      <div class="modal-actions">
        ${cancelText ? `<button class="btn ghost" data-m="cancel">${esc(cancelText)}</button>` : ''}
        ${actions.map((a, i) => `<button class="btn ${a.cls || ''}" data-m="${i}">${esc(a.label)}</button>`).join('')}
      </div></div></div>`;
    const onKey = e => { if (e.key === 'Escape') close(); };
    function close() { root.innerHTML = ''; document.removeEventListener('keydown', onKey); }
    document.addEventListener('keydown', onKey);
    root.querySelectorAll('[data-m]').forEach(b => {
      b.onclick = async () => {
        if (b.dataset.m === 'cancel') return close();
        const a = actions[+b.dataset.m];
        root.querySelectorAll('button').forEach(x => { x.disabled = true; });
        try { await a.run(); close(); }
        catch (err) {
          root.querySelectorAll('button').forEach(x => { x.disabled = false; });
          handleErr(err, '操作失败');
        }
      };
    });
    return close;
  }
  function showConflict({ onReload, onOverwrite }) {
    openModal({
      title: '这期数据刚被别人改过',
      html: `<p>你打开之后，有同事保存过这期数据。可以这样处理：</p>
        <ul><li><b>重新加载</b>：换成同事保存的最新版本，你这次的修改不保留。</li>
        <li><b>用我的版本覆盖</b>：以你现在填的内容为准，同事那次改动会被替换（修改记录里还能找回）。</li></ul>`,
      cancelText: '先不处理',
      actions: [
        { label: '重新加载（放弃我的修改）', cls: 'ghost', run: onReload },
        { label: '用我的版本覆盖', cls: 'danger', run: onOverwrite },
      ],
    });
  }

  function handleErr(err, prefix) {
    const code = err && err.code;
    const msg = (err && err.message) || String(err);
    if (code === 'not_member' || code === 'tables_missing') {
      state.access = { ok: false, reason: code, message: msg };
      state.dirty = false;
      route();
      return;
    }
    if (code === 'auth') {
      flash('登录状态失效了，请重新登录', 4000);
      state.me = null; state.access = null; state.dirty = false;
      route();
      return;
    }
    flash((prefix ? prefix + '：' : '') + msg, 5000);
  }

  // 公式说明浮层（ⓘ），挂在 body 上，避免被卡片裁掉
  const $tip = document.getElementById('tip');
  let tipOwner = null;
  let tipShownAt = 0;
  function showTip(el) {
    const text = el.getAttribute('data-tip');
    if (!text) return;
    if (tipOwner !== el || $tip.style.display !== 'block') tipShownAt = Date.now();
    tipOwner = el;
    $tip.textContent = text;
    $tip.style.display = 'block';
    const r = el.getBoundingClientRect();
    const w = $tip.offsetWidth, h = $tip.offsetHeight;
    let left = Math.min(Math.max(8, r.left - 12), window.innerWidth - w - 8);
    let top = r.bottom + 8;
    if (top + h > window.innerHeight - 8) top = Math.max(8, r.top - h - 8);
    $tip.style.left = left + 'px';
    $tip.style.top = top + 'px';
  }
  function hideTip() { $tip.style.display = 'none'; tipOwner = null; }
  document.addEventListener('mouseover', e => { const t = e.target.closest && e.target.closest('[data-tip]'); if (t) showTip(t); });
  document.addEventListener('mouseout', e => { const t = e.target.closest && e.target.closest('[data-tip]'); if (t && t === tipOwner) hideTip(); });
  document.addEventListener('focusin', e => { if (e.target.matches && e.target.matches('[data-tip]')) showTip(e.target); });
  document.addEventListener('focusout', e => { if (e.target === tipOwner) hideTip(); });
  document.addEventListener('click', e => {
    const t = e.target.closest && e.target.closest('[data-tip]');
    // 触屏点一下会先触发 mouseover / focus 再触发 click，刚弹出的不要马上收起
    if (t) { if (tipOwner === t && $tip.style.display === 'block' && Date.now() - tipShownAt > 400) hideTip(); else showTip(t); }
    else if (tipOwner) hideTip();
  });
  window.addEventListener('scroll', () => { if (tipOwner) hideTip(); }, { passive: true });

  // ---------- 认证与顶栏 ----------
  async function loadSession() {
    try { state.me = await DB.getSession(); } catch (e) { state.me = null; }
    renderNav();
  }
  let authTimer = null;
  function onAuthChange() {
    clearTimeout(authTimer);
    authTimer = setTimeout(async () => {
      const prevId = state.me && state.me.user && state.me.user.id;
      let s = null;
      try { s = await DB.getSession(); } catch (e) { s = null; }
      const id = s && s.user && s.user.id;
      if (id === prevId) { if (s) { state.me = s; renderNav(); } return; }
      state.me = s; state.access = null; state.loaded = false; state.cohorts = []; state.dirty = false;
      renderNav(); route();
    }, 60);
  }
  function renderNav() {
    const ok = !!(state.me && state.access && state.access.ok);
    document.getElementById('nav').hidden = !ok;
    document.getElementById('user-area').hidden = !state.me;
    document.getElementById('nav-members').hidden = !ok;
    const p = (state.me && state.me.profile) || {};
    const email = (state.me && state.me.user && state.me.user.email) || '';
    document.getElementById('nav-user').textContent =
      state.me ? `${p.display_name || email}${p.department ? ' · ' + p.department : ''}` : '';
  }
  function setActiveNav(path) {
    document.querySelectorAll('[data-nav]').forEach(a => {
      const n = a.dataset.nav;
      a.classList.toggle('active', n === path || (n === '/new' && path.startsWith('/edit/')));
    });
  }
  async function doLogout() {
    if (state.dirty && !confirm('有未保存的修改，确定退出吗？')) return;
    try { await DB.signOut(); } catch (e) { /* 忽略 */ }
    state.me = null; state.access = null; state.cohorts = []; state.loaded = false; state.dirty = false;
    renderNav(); route();
  }
  document.getElementById('nav-logout').onclick = doLogout;

  // ---------- 路由 ----------
  function parseHash() {
    const h = (location.hash || '#/').slice(1);
    const i = h.indexOf('?');
    const path = (i >= 0 ? h.slice(0, i) : h) || '/';
    return { path, params: new URLSearchParams(i >= 0 ? h.slice(i + 1) : '') };
  }
  window.addEventListener('hashchange', () => {
    if ((location.hash || '#/') === currentHash) return;
    if (state.reviewSaving) { history.replaceState(null, '', currentHash); flash('复盘正在保存，请稍候'); return; }
    if (state.dirty && !confirm('有未保存的修改，确定离开这一页吗？')) {
      history.replaceState(null, '', currentHash);
      return;
    }
    state.dirty = false;
    route();
  });
  window.addEventListener('beforeunload', e => {
    if (state.dirty) { e.preventDefault(); e.returnValue = ''; }
  });

  function destroyCharts() {
    charts.forEach(c => { try { c.destroy(); } catch (e) { /* 忽略 */ } });
    charts = [];
  }

  async function route() {
    const seq = ++routeSeq;
    currentHash = location.hash || '#/';
    destroyCharts(); hideTip();
    $app.classList.remove('refreshing');
    const { path, params } = parseHash();
    if (state.present && path !== '/') exitPresent(false);

    if (!state.me) { renderNav(); return viewAuth('login'); }
    if (!state.access) {
      $app.innerHTML = loadingHTML('正在检查权限…');
      try { state.access = await DB.checkAccess(); }
      catch (err) { state.access = { ok: false, reason: (err && err.code) || 'error', code: err && err.code, message: err && err.message }; }
      if (seq !== routeSeq) return;
      // 登录状态已失效（令牌过期等）：回登录页
      if (!state.access.ok && (state.access.code === 'auth' || state.access.reason === 'auth')) {
        state.me = null; state.access = null;
        renderNav();
        return viewAuth('login');
      }
    }
    renderNav();
    if (!state.access.ok) return viewNoAccess();

    setActiveNav(path);
    if (path === '/list') return viewList(seq);
    if (path === '/review') return viewLessons(seq, params);
    if (path === '/new') return viewNew(seq);
    if (path === '/members') return viewMembers(seq);
    const m = path.match(/^\/edit\/([^/]+)$/);
    if (m) return viewEdit(decodeURIComponent(m[1]), seq);
    if (params.get('week')) {
      state.dash.week = params.get('week');
      state.dash.channel = ['xiaoe', 'bilibili'].includes(params.get('ch')) ? params.get('ch') : 'all';
    }
    return viewDashboard(seq);
  }
  function go(hash) {
    if ((location.hash || '#/') === hash) route();
    else location.hash = hash;
  }

  // ---------- 视图：登录 / 注册 ----------
  function viewAuth(mode) {
    const isLogin = mode === 'login';
    $app.innerHTML = `
      <div class="auth-wrap"><div class="auth-card">
        <div class="auth-brand">训练营周复盘</div>
        <div class="auth-sub">每期训练营的邀约、到课、转化数据，复盘会上一起看</div>
        <div class="seg" role="tablist">
          <button type="button" class="${isLogin ? 'active' : ''}" data-auth="login">登录</button>
          <button type="button" class="${isLogin ? '' : 'active'}" data-auth="register">注册</button>
        </div>
        <form id="auth-form" autocomplete="on">
          ${isLogin ? '' : `
          <label class="field"><span>姓名（同事们看到的名字）</span>
            <input name="display_name" required maxlength="20" placeholder="你的名字"></label>
          <label class="field"><span>部门</span>
            <select name="department">${DEPARTMENTS.map(d => `<option>${esc(d)}</option>`).join('')}</select></label>`}
          <label class="field"><span>邮箱</span>
            <input name="email" type="email" required placeholder="you@example.com" autocomplete="email"></label>
          <label class="field"><span>密码${isLogin ? '' : '（至少 6 位）'}</span>
            <input name="password" type="password" required minlength="6" autocomplete="${isLogin ? 'current-password' : 'new-password'}"></label>
          <button class="btn" id="auth-btn">${isLogin ? '登录' : '注册并进入'}</button>
        </form>
        <div class="auth-hint">和 Skill 市集共用账号，已经注册过的直接登录</div>
      </div></div>`;
    $app.querySelectorAll('[data-auth]').forEach(b => { b.onclick = () => viewAuth(b.dataset.auth); });
    document.getElementById('auth-form').onsubmit = async e => {
      e.preventDefault();
      const fd = new FormData(e.target);
      const email = String(fd.get('email') || '').trim();
      const password = String(fd.get('password') || '');
      const btn = document.getElementById('auth-btn');
      btn.disabled = true;
      try {
        if (isLogin) await DB.signIn(email, password);
        else await DB.signUp(email, password, String(fd.get('display_name') || '').trim(), String(fd.get('department') || ''));
      } catch (err) {
        btn.disabled = false;
        const msg = String((err && err.message) || err);
        if (/Invalid login/i.test(msg)) flash('邮箱或密码不对');
        else if (/already registered/i.test(msg)) flash('这个邮箱已经注册过了，直接登录吧');
        else flash(msg, 5000);
        return;
      }
      await loadSession();
      if (!state.me) {
        btn.disabled = false;
        if (!isLogin) { flash('注册成功，请用刚才的邮箱和密码登录', 4000); viewAuth('login'); }
        else flash('登录没有成功，请重试');
        return;
      }
      state.access = null; state.loaded = false;
      route();
    };
  }

  // ---------- 视图：没有权限 / 未初始化 ----------
  function viewNoAccess() {
    const a = state.access || {};
    const email = (state.me && state.me.user && state.me.user.email) || '';
    let title, body;
    if (a.reason === 'not_member') {
      title = '账号还没开通';
      body = `你的账号（${esc(email)}）还没开通复盘系统，请让已开通的同事在『成员』页添加你的邮箱`;
    } else if (a.reason === 'tables_missing') {
      title = '数据库还没初始化';
      body = '数据库还没初始化，请按 <code>camp-review/SETUP.md</code> 执行 <code>setup.sql</code>';
    } else {
      title = '暂时连不上数据库';
      body = esc(a.message || '请检查网络后重试');
    }
    $app.innerHTML = `<div class="note-card"><h2>${esc(title)}</h2><p>${body}</p>
      <div class="actions"><button class="btn" id="btn-recheck">重新检查</button>
      <button class="btn ghost" id="btn-switch">换个账号登录</button></div></div>`;
    document.getElementById('btn-recheck').onclick = () => { state.access = null; route(); };
    document.getElementById('btn-switch').onclick = doLogout;
  }

  // ---------- 视图：复盘看板 ----------
  async function viewDashboard(seq) {
    if (state.loaded) $app.classList.add('refreshing');
    else $app.innerHTML = loadingHTML();
    try { await loadCohorts(); }
    catch (err) {
      if (seq !== routeSeq) return;
      $app.classList.remove('refreshing');
      if (!state.loaded) $app.innerHTML = `<div class="note-card"><h2>加载失败</h2><p>${esc(err.message || err)}</p></div>`;
      handleErr(err, '加载失败');
      return;
    }
    if (seq !== routeSeq) return;
    $app.classList.remove('refreshing');
    renderDashboard();
    const focusId = parseHash().params.get('cohort');
    if (focusId) {
      const panel = $app.querySelector(`[data-cohort="${CSS.escape(focusId)}"]`);
      if (panel) (panel.querySelector('.review-slot') || panel).scrollIntoView({ block: 'start' });
    }
  }

  // ---------- 视图：全部经验沉淀 ----------
  async function viewLessons(seq, params) {
    $app.innerHTML = loadingHTML('正在加载经验沉淀…');
    try { await loadCohorts(); }
    catch (err) {
      if (seq !== routeSeq) return;
      $app.innerHTML = `<div class="note-card"><h2>经验暂时加载失败</h2><p>${esc(err.message || err)}</p><button class="btn" id="retry-lessons">重新加载</button></div>`;
      document.getElementById('retry-lessons').onclick = () => route();
      handleErr(err, '加载失败'); return;
    }
    if (seq !== routeSeq) return;
    state.lessons = { channel: CH_ORDER.includes(params.get('ch')) ? params.get('ch') : 'all', query: params.get('q') || '' };
    $app.innerHTML = `<div class="page-head"><div><h1>经验沉淀</h1><p class="review-hint">汇总所有期次已保存的经验，按直播日期从新到旧排列。</p></div></div>
      <div class="lesson-toolbar">
        <label class="review-field">渠道<select id="lesson-channel"><option value="all">全部渠道</option>${CH_ORDER.map(ch => `<option value="${ch}" ${state.lessons.channel === ch ? 'selected' : ''}>${esc(chLabel(ch))}</option>`).join('')}</select></label>
        <label class="review-field lesson-search">查找经验<input id="lesson-search" type="search" value="${esc(state.lessons.query)}" placeholder="搜索经验内容、期次名称或日期"></label>
      </div><div id="lesson-results"></div>`;
    const update = () => {
      const p = new URLSearchParams();
      if (state.lessons.channel !== 'all') p.set('ch', state.lessons.channel);
      if (state.lessons.query) p.set('q', state.lessons.query);
      const h = '#/review' + (p.size ? '?' + p.toString() : '');
      history.replaceState(null, '', h); currentHash = h;
      renderLessonsResults();
    };
    document.getElementById('lesson-channel').onchange = e => { state.lessons.channel = e.target.value; update(); };
    document.getElementById('lesson-search').oninput = e => { state.lessons.query = e.target.value; update(); };
    renderLessonsResults();
  }
  function renderLessonsResults() {
    const all = state.cohorts.map(c => ({ c, lessons: R.normalizeReview(c.data.review).lessons })).filter(x => x.lessons)
      .sort((a, b) => String(b.c.live_date).localeCompare(String(a.c.live_date)) || String(b.c.id).localeCompare(String(a.c.id), undefined, { numeric: true }));
    const q = state.lessons.query.trim().toLocaleLowerCase();
    const rows = all.filter(({ c, lessons }) => (state.lessons.channel === 'all' || c.channel === state.lessons.channel) &&
      (!q || [lessons, c.name, c.live_date, chLabel(c.channel)].join('\n').toLocaleLowerCase().includes(q)));
    document.getElementById('lesson-results').innerHTML = `<p class="lesson-count" role="status">${rows.length} 期经验${rows.length !== all.length ? ` · 共 ${all.length} 期已沉淀` : ''}</p>
      ${rows.length ? `<div class="lesson-entries">${rows.map(({ c, lessons }) => `<article class="section lesson-entry" data-lesson-id="${esc(c.id)}">
        <div class="lesson-entry-head"><div><div class="ph-line1">${chTag(c.channel)}<h2>${esc(c.name || '未命名')}</h2></div><p class="review-hint">直播 ${esc(dateWeek(c.live_date))}</p></div>
          <a class="btn small ghost" data-open-cohort href="#/?week=${esc(weekOf(c.live_date))}&ch=${esc(c.channel)}&cohort=${encodeURIComponent(c.id)}">查看本期复盘</a></div>
        <div class="summary-text">${esc(lessons)}</div>
      </article>`).join('')}</div>` : `<div class="empty-state"><h2>${all.length ? '没有找到匹配的经验' : '还没有记录经验沉淀'}</h2><p>${all.length ? '换个关键词，或选择其他渠道再看看。' : '在看板点击「编辑本期复盘」，填写「本期经验沉淀」并保存，就会出现在这里。'}</p>${all.length ? '' : '<a class="btn" href="#/">去复盘看板</a>'}</div>`}`;
  }

  function syncDashHash() {
    const h = `#/?week=${state.dash.week}${state.dash.channel !== 'all' ? '&ch=' + state.dash.channel : ''}`;
    history.replaceState(null, '', h);
    currentHash = h;
  }

  function renderDashboard() {
    destroyCharts(); hideTip();
    if (!state.cohorts.length) {
      $app.innerHTML = `${reviewGuideHTML(false)}<div class="empty-state">
        <h2>还没有任何一期数据</h2>
        <p>把运营统计表复制粘贴进来，系统会自动算出邀约入群率、直播到课率和两个转化率。</p>
        <a class="btn" href="#/new">去录入第一期</a></div>`;
      return;
    }
    const filtered = filterCh(state.cohorts, state.dash.channel);
    const weeks = safe(() => L.groupByWeek(filtered), []) || [];
    if (weeks.length && !weeks.some(w => w.week === state.dash.week)) state.dash.week = weeks[0].week;
    const idx = weeks.findIndex(w => w.week === state.dash.week);
    const wk = weeks[idx];
    const older = weeks[idx + 1], newer = weeks[idx - 1];

    const chOpts = [['all', '全部'], ['xiaoe', chLabel('xiaoe')], ['bilibili', chLabel('bilibili')]];
    const bar = `
      <div class="dash-bar">
        <div class="week-nav">
          <button class="btn-icon" data-week="${older ? esc(older.week) : ''}" ${older ? '' : 'disabled'} title="${older ? esc(older.label) : '没有更早的数据'}" aria-label="上一周">‹<span class="wk-txt"> 上一周</span></button>
          <label class="week-select" title="选择周">
            <select id="week-select" aria-label="选择周" ${weeks.length ? '' : 'disabled'}>
              ${weeks.length ? weeks.map(w => `<option value="${esc(w.week)}" ${w.week === state.dash.week ? 'selected' : ''}>${esc(w.label)}</option>`).join('')
                : '<option>暂无数据</option>'}
            </select><span class="caret">▾</span>
          </label>
          <button class="btn-icon" data-week="${newer ? esc(newer.week) : ''}" ${newer ? '' : 'disabled'} title="${newer ? esc(newer.label) : '已经是最近一周'}" aria-label="下一周"><span class="wk-txt">下一周 </span>›</button>
        </div>
        <div class="seg" role="group" aria-label="渠道筛选">
          ${chOpts.map(([k, t]) => `<button type="button" data-ch="${k}" class="${state.dash.channel === k ? 'active' : ''}" aria-pressed="${state.dash.channel === k}">${esc(t)}</button>`).join('')}
        </div>
        <div class="spacer"></div>
        <button class="btn ghost no-present" id="btn-present" type="button">投屏模式</button>
      </div>`;

    if (!wk) {
      $app.innerHTML = `${bar}${reviewGuideHTML(false)}<div class="empty-state"><h2>${esc(chLabel(state.dash.channel))}还没有数据</h2>
        <p>切换到「全部」看看其他渠道，或者录入这个渠道的第一期。</p><a class="btn" href="#/new">录入数据</a></div>`;
      bindDashboard(null);
      return;
    }

    const list = wk.cohorts.slice().sort(byChannelThenDate);
    const counts = CH_ORDER.map(ch => [ch, list.filter(c => c.channel === ch).length]).filter(x => x[1]);
    const meta = `本周 ${list.length} 期${counts.length > 1 ? '：' + counts.map(([ch, n]) => `${chLabel(ch)} ${n} 期`).join(' · ') : ''}`;
    const channelsInWeek = new Set(list.map(c => c.channel));
    const td = trendData(wk.week);

    $app.innerHTML = `
      ${bar}
      ${reviewGuideHTML(true)}
      <div class="dash-meta">${esc(meta)}</div>
      ${list.map(panelHTML).join('')}
      ${channelsInWeek.size > 1 ? compareHTML(list) : ''}
      ${trendHTML(td)}`;
    bindDashboard(list);
    drawTrends(td);
  }

  function reviewGuideHTML(hasData) {
    return `<section class="review-guide" aria-labelledby="review-guide-title">
      <div class="review-guide-head"><h2 id="review-guide-title">每次复盘，按这个流程走</h2>
        <p>会前把数据和判断写好；会上先看数据、再看趋势，最后讨论复盘与行动。</p></div>
      <div class="review-guide-steps">
        <div class="review-guide-step"><h3><span>01</span>会前准备 <small>运营负责人</small></h3>
          <ol>
            <li><b>粘贴并核对数据</b>：确认本期日期、渠道、人数及是否收口。</li>
            <li><b>填写本期复盘</b>：写清异常、原因假设和待讨论的问题。</li>
            <li><b>填写经验沉淀</b>：记录可复用的做法及适用条件。</li>
            <li><b>评估上期行动</b>：补齐执行情况、数据效果及继续或调整的依据。</li>
            <li><b>拟定下期行动</b>：写具体动作、目标指标和负责人，并保存复盘。</li>
          </ol>
          <p class="guide-note">首期没有上期行动时，可跳过该项。</p>
          <a class="btn small ghost no-present" href="#/new">录入 / 更新数据</a>
        </div>
        <div class="review-guide-step"><h3><span>02</span>会上复盘 <small>一起核对与讨论</small></h3>
          <ol>
            <li><b>先过本期数据</b>：核对直播、一对一、总转化人数及各环节转化率。</li>
            <li><b>再看趋势变化</b>：对照上期和最近八期，找出变化最大的环节。</li>
            <li><b>然后沟通复盘</b>：结合数据讨论原因、上期行动效果和本期经验。</li>
            <li><b>确认下一步</b>：决定哪些做法继续、调整或停止，确认下期行动与目标。</li>
          </ol>
          <p class="guide-note">指标上涨不等于行动有效，结合执行证据判断。</p>
          <button class="btn small ghost no-present" type="button" data-guide-trends ${hasData ? '' : 'disabled'}>查看八期趋势</button>
        </div>
        <div class="review-guide-step"><h3><span>03</span>会后跟进 <small>行动负责人</small></h3>
          <ol>
            <li><b>保存会议共识</b>：把讨论结果补回本期复盘、经验和行动记录。</li>
            <li><b>按约定执行</b>：明确完成时间，保留执行证据并持续更新数据。</li>
            <li><b>下期回来验证</b>：对照目标与实际效果，决定继续采用还是调整。</li>
          </ol>
          <p class="guide-note">下期会自动带出同渠道本期制定的行动。</p>
          <a class="btn small ghost no-present" href="#/review">回看历史经验</a>
        </div>
      </div>
    </section>`;
  }

  // 看板上正在写结论时，切换周 / 渠道 / 投屏前先确认
  function okToDiscard() {
    if (state.reviewSaving) { flash('复盘正在保存，请稍候'); return false; }
    if (!state.dirty) return true;
    if (!confirm('本期复盘还有未保存的修改，确定放弃吗？')) return false;
    state.dirty = false;
    return true;
  }

  function bindDashboard(list) {
    $app.querySelectorAll('[data-week]').forEach(b => {
      b.onclick = () => {
        if (!b.dataset.week || !okToDiscard()) return;
        state.dash.week = b.dataset.week; syncDashHash(); renderDashboard();
      };
    });
    const sel = document.getElementById('week-select');
    if (sel) sel.onchange = () => {
      if (!okToDiscard()) { sel.value = state.dash.week; return; }
      state.dash.week = sel.value; syncDashHash(); renderDashboard();
    };
    $app.querySelectorAll('[data-ch]').forEach(b => {
      b.onclick = () => {
        if (!okToDiscard()) return;
        state.dash.channel = b.dataset.ch;
        const weeks = safe(() => L.groupByWeek(filterCh(state.cohorts, state.dash.channel)), []) || [];
        if (weeks.length && !weeks.some(w => w.week === state.dash.week)) state.dash.week = weeks[0].week;
        if (state.dash.week) syncDashHash();
        renderDashboard();
      };
    });
    const pb = document.getElementById('btn-present');
    if (pb) pb.onclick = () => { if (okToDiscard()) enterPresent(); };
    const trendJump = $app.querySelector('[data-guide-trends]');
    if (trendJump) trendJump.onclick = () => $app.querySelector('.trends')?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    if (!list) return;

    // 同时只打开一份复盘草稿；切换前确认，避免保存一份时丢掉另一份。
    $app.querySelectorAll('[data-edit-review], [data-edit-summary]').forEach(b => {
      b.onclick = () => openReview(b.dataset.editReview || b.dataset.editSummary);
    });
  }

  function openReview(id) {
    if (!okToDiscard()) return;
    renderDashboard();
    const c = state.cohorts.find(x => sameId(x.id, id));
    const slot = $app.querySelector(`[data-review-slot="${CSS.escape(String(id))}"]`);
    if (!c || !slot) return;
    ReviewUI.edit(slot, c, state.cohorts, {
      onDirty: () => { state.dirty = true; },
      onCancel: () => { if (okToDiscard()) renderDashboard(); },
      onSave: (review, summary) => saveReview(c, review, summary),
    });
    slot.scrollIntoView({ block: 'start' });
  }

  async function saveReview(c, review, summary) {
    const seq = routeSeq;
    const patch = base => ({ ...base, summary: summary === c.summary ? base.summary : summary,
      data: { ...base.data, review } });
    state.reviewSaving = true;
    const done = saved => {
      replaceInCache(normCohort(saved));
      if (seq !== routeSeq) return;
      state.dirty = false;
      flash('本期复盘已保存，下期会自动带出行动');
      renderDashboard();
    };
    try {
      done(await DB.updateCohort(forSave(patch(c))));
    } catch (err) {
      if (err && err.code === 'conflict') {
        showConflict({
          onReload: async () => {
            state.reviewSaving = true;
            try {
              replaceInCache(normCohort(await DB.getCohort(c.id)));
              if (seq === routeSeq) { state.dirty = false; renderDashboard(); flash('已换成最新版本'); }
            } finally { state.reviewSaving = false; }
          },
          // 只覆盖本次复盘；保留同事最新的表格和未被本次修改的总评。
          onOverwrite: async () => {
            state.reviewSaving = true;
            try {
              const latest = normCohort(await DB.getCohort(c.id));
              done(await DB.updateCohort(forSave(patch(latest))));
            } finally { state.reviewSaving = false; }
          },
        });
      } else if (err && ['auth', 'not_member', 'tables_missing'].includes(err.code)) {
        // 保存被拒绝时保留编辑器中的文字，避免权限处理的路由切换丢掉复盘草稿。
        flash('复盘未保存：' + (err.message || '请核对登录状态和访问权限'), 6000);
      } else handleErr(err, '保存失败');
    } finally { state.reviewSaving = false; }
  }

  // 一期的面板
  function panelHTML(c) {
    const m = metricsOf(c);
    const head = `
      <div class="panel-head">
        <div class="ph-main">
          <div class="ph-line1">${chTag(c.channel)}<h2 class="ph-title">${esc(c.name || '未命名')}</h2></div>
          <div class="ph-line2"><span>直播 ${esc(dateWeek(c.live_date))}</span>${m ? statusBadge(m) : ''}</div>
          ${c.closed ? '' : '<div class="ph-volatile">这期数据还可能变化</div>'}
        </div>
        <div class="ph-side">
          <a class="btn small ghost no-present" href="#/edit/${encodeURIComponent(c.id)}">编辑数据</a>
          <div class="ph-meta">${c.updated_by_name ? esc(c.updated_by_name) + ' ' : ''}最后更新于 ${esc(fmtTime(c.updated_at))}</div>
        </div>
      </div>`;
    if (!m) {
      return `<section class="panel ch-${esc(c.channel)}">${head}<div class="muted-note">这期数据算不出指标，请到编辑页检查表格</div></section>`;
    }
    const cmp = safe(() => L.compareCohort(c, state.cohorts), null);
    const ins = safe(() => window.CampInsights.buildInsights(c, state.cohorts), []) || [];
    return `
      <section class="panel ch-${esc(c.channel)}" data-cohort="${esc(c.id)}">
        ${head}
        ${dealKpisHTML(c)}
        ${kpiStripHTML(c, m, cmp)}
        <div class="panel-row funnel-row">
          <div class="block"><div class="block-title">转化漏斗</div>${funnelHTML(m)}</div>
          <div class="block">${dealsHTML(c, m)}</div>
        </div>
        <div class="panel-row two">
          <div class="block"><div class="block-title">数据洞察 · 本期优先讨论</div>${insightsHTML(ins)}</div>
          <div class="block"><div class="block-title">本期复盘结论</div>
            <div class="summary-slot" data-summary-slot="${esc(c.id)}">${summaryViewHTML(c)}</div></div>
        </div>
        <div class="review-slot" data-review-slot="${esc(c.id)}">${ReviewUI.render(c, state.cohorts)}</div>
        ${detailsHTML(c, m)}
      </section>`;
  }

  function dealKpisHTML(c) {
    const prev = R.previousCohort(c, state.cohorts);
    const m = metricsOf(c);
    const defs = [
      ['liveDeals', '直播转化人数', c.channel === 'bilibili' ? '直播中全款成交' : '直播定金后补齐尾款'],
      ['directDeals', '一对一转化人数', '直接付款 · 1v1 追单成交'],
      ['totalDeals', '总转化人数', c.channel === 'bilibili' ? '直播 + 1v1 + 老学员 − 退款' : '直播 + 1v1 + 老学员'],
    ];
    return `<div class="deal-kpis">${defs.map(([key, label, note]) => {
      const complete = R.metricValue(c, key), before = prev ? R.metricValue(prev, key) : null;
      const liveKey = c.channel === 'bilibili' ? 'live_full' : 'balance';
      const cols = key === 'liveDeals' ? [liveKey] : key === 'directDeals' ? ['direct'] : [liveKey, 'direct', 'alumni'];
      const hasSome = c.data.conversions.some(row => cols.some(col => L.parseNum(row[col]) != null));
      const v = complete ?? (hasSome ? key === 'directDeals' ? m.sums.direct : m[key] : null);
      const change = complete != null && before != null ? `较上期 ${complete - before > 0 ? '+' : ''}${int(complete - before)} 人`
        : complete == null ? (hasSome ? '已录入合计 · 缺项待补齐' : '尚未录入') : '暂无完整的上期数据';
      return `<div class="deal-kpi ${key === 'totalDeals' ? 'total' : ''}" data-deal-metric="${key}"><h3>${label}</h3>
        <div class="deal-number">${esc(int(v))}<small>人</small></div><p>${esc(change)}${c.closed ? '' : ' · 尚未收口'}</p><p>${esc(note)}</p></div>`;
    }).join('')}</div>`;
  }

  function kpiStripHTML(c, m, cmp) {
    const defs = coreMetrics();
    const hasPrev = !!(cmp && cmp.prev);
    let worst = null, best = null;
    if (hasPrev) {
      defs.forEach(d => {
        const x = cmp.metrics && cmp.metrics[d.key];
        if (!x || x.delta == null) return;
        if (x.delta <= -0.005 && (!worst || x.delta < worst.delta)) worst = { key: d.key, delta: x.delta };
        if (x.delta >= 0.005 && (!best || x.delta > best.delta)) best = { key: d.key, delta: x.delta };
      });
    }
    return `<div class="kpis">${defs.map(d => kpiHTML(d, c, m, cmp,
      worst && worst.key === d.key ? 'worst' : best && best.key === d.key ? 'best' : '')).join('')}</div>`;
  }

  function bigPct(v) {
    const s = pct(v);
    if (v == null || s === '—') return '<span class="na">—</span>';
    return s.endsWith('%') ? `${esc(s.slice(0, -1))}<span class="pct">%</span>` : esc(s);
  }
  function fracText(parts) {
    const [a, b] = parts;
    if (a == null && b == null) return '—';
    return `${int(a)} / ${int(b)} 人`;
  }

  function kpiHTML(def, c, m, cmp, flag) {
    const v = m[def.key];
    const parts = PARTS[def.key] ? PARTS[def.key](m) : [null, null];
    const x = (cmp && cmp.metrics && cmp.metrics[def.key]) || {};
    const hasPrev = !!(cmp && cmp.prev);
    let delta;
    if (!hasPrev) delta = '<div class="kpi-delta flat">首期数据，暂无对比</div>';
    else if (x.delta == null) {
      delta = `<div class="kpi-delta flat">较上期：${v == null ? '本期数据不全' : '上期没有这项数据'}</div>`;
    } else {
      const f = safe(() => L.fmtDelta(x.delta), { text: '', dir: 'flat' });
      const prevName = x.prevName || (cmp.prev && cmp.prev.name) || '';
      delta = `<div class="kpi-delta ${esc(f.dir)}" title="上期：${esc(prevName)}">较上期 <b>${esc(f.text)}</b></div>`;
    }
    const n = x.avgCount != null ? x.avgCount : (cmp ? cmp.avgCount : 0);
    const sub = [];
    if (hasPrev && x.prev != null) sub.push(`上期 ${pct(x.prev)}`);
    if (n > 0 && x.avg != null) sub.push(`近 ${n} 期均值 ${pct(x.avg)}`);
    const flagHTML = flag === 'worst' ? '<span class="kpi-flag down">↓ 掉得最多</span>'
      : flag === 'best' ? '<span class="kpi-flag up">↑ 涨得最多</span>' : '';
    return `
      <div class="kpi ${flag}">
        <div class="kpi-top">
          <span class="kpi-label">${esc(def.label)}</span>
          <span class="tip" tabindex="0" role="button" aria-label="${esc(def.label)}怎么算" data-tip="${esc(formulaText(def.key, c.channel))}">i</span>
          ${flagHTML}
        </div>
        <div class="kpi-val">${bigPct(v)}</div>
        <div class="kpi-frac">${esc(fracText(parts))}</div>
        ${delta}
        ${sub.length ? `<div class="kpi-sub">${esc(sub.join(' · '))}</div>` : ''}
      </div>`;
  }

  function funnelHTML(m) {
    const steps = (m.funnel || []).filter(Boolean);
    if (!steps.length || steps.every(s => !s.value)) return '<div class="muted-note">数据还不够，画不出漏斗</div>';
    const first = steps[0].value || 0;
    const restMax = Math.max(0, ...steps.slice(1).map(s => s.value || 0));
    // 触达量远大于后续人数时单独缩放，连续条形配合文字说明缩放口径。
    const separateScale = restMax > 0 && first > restMax * 3;
    const max = separateScale ? restMax : Math.max(first, restMax);
    const rows = steps.map((s, i) => {
      let w = 0;
      if (s.value > 0) w = separateScale ? (i === 0 ? 100 : s.value / max * 80) : s.value / max * 100;
      const style = w > 0 ? `width:max(4px, ${w.toFixed(2)}%)` : 'width:0';
      const rate = i === 0 ? '' : `<span class="f-arrow">→</span>${s.rateFromPrev == null ? '—' : esc(pct(s.rateFromPrev))}`;
      return `<div class="f-row">
        <div class="f-label" title="${esc(s.label)}">${esc(s.label)}</div>
        <div class="f-track"><div class="f-bar" style="${style}"></div></div>
        <div class="f-val">${esc(int(s.value))}<span class="u">人</span></div>
        <div class="f-rate">${rate}</div>
      </div>`;
    }).join('');
    return `<div class="funnel">
        <div class="f-row f-head"><div>步骤</div><div></div><div class="f-val">人数</div><div class="f-rate">比上一步</div></div>
        ${rows}
      </div>
      ${separateScale ? '<div class="f-note">触达人数单独缩放显示，后续步骤按同一比例绘制；实际人数以右侧数字为准。</div>' : ''}`;
  }

  function dealsHTML(c, m) {
    const isX = c.channel !== 'bilibili';
    const s = m.sums || {};
    const ex = m.extra || {};
    const direct = ex.direct != null ? ex.direct : s.direct;
    const alumni = ex.alumni != null ? ex.alumni : s.alumni;
    const total = ex.totalDeals != null ? ex.totalDeals : m.totalDeals;
    const rows = [
      [isX ? '直播成交（补齐尾款）' : '直播成交（全款）', m.liveDeals, ''],
      ['1v1 追单', direct, m.directShare != null ? `占 ${pct(m.directShare)}` : ''],
      ['老学员', alumni, ''],
    ];
    if (!isX && s.refund) rows.push(['退款', -s.refund, '']);
    const fmt = v => v == null ? '—' : (v < 0 ? '−' + int(-v) : int(v)) + ' 人';
    return `<div class="block-title">成交构成</div>
      <div class="deals-total"><span class="dt-num">${esc(int(total))}</span><span class="dt-unit">人 · 总成交</span></div>
      <ul class="deals-list">${rows.map(r => `<li><span>${esc(r[0])}</span>${r[2] ? `<em>${esc(r[2])}</em>` : ''}<b>${esc(fmt(r[1]))}</b></li>`).join('')}</ul>
      ${isX && m.pendingBalance > 0 ? `<div class="deals-pending">还有 ${esc(int(m.pendingBalance))} 人付了定金、没补尾款</div>` : ''}`;
  }

  function insightsHTML(ins) {
    if (!ins.length) return '<div class="muted-note">暂时没有要特别提醒的</div>';
    const icon = { good: '↑', bad: '↓', warn: '!', info: 'i' };
    return `<ul class="insights">${ins.map(i => `
      <li class="ins ${esc(i.level)}"><span class="ins-ic" aria-hidden="true">${icon[i.level] || '·'}</span><span>${i.title ? `<strong class="ins-title">${esc(i.title)}</strong>` : ''}${esc(i.text)}</span></li>`).join('')}</ul>`;
  }

  function summaryViewHTML(c) {
    return `<div class="summary-text ${c.summary ? '' : 'empty'}">${c.summary ? esc(c.summary) : '还没写结论'}</div>
      <button class="btn small ghost no-present" type="button" data-edit-summary="${esc(c.id)}">编辑结论</button>`;
  }

  function detailsHTML(c, m) {
    const open = state.openDetails.has(String(c.id));
    return `<details class="details" data-details="${esc(c.id)}" ${open ? 'open' : ''}>
      <summary>明细数据</summary>
      <div class="details-body">
        <h4>邀约渠道</h4>${inviteTableHTML(c, m)}
        <h4>社群每日</h4>${communityTableHTML(c, m)}
        <h4>每日成交</h4>${conversionTableHTML(c, m)}
        ${c.channel !== 'bilibili' ? `<h4>定金 → 尾款</h4>${depositHTML(m)}` : ''}
      </div>
    </details>`;
  }
  function inviteTableHTML(c, m) {
    const rows = m.inviteByChannel || [];
    if (!rows.length) return '<div class="muted-note">没有邀约数据</div>';
    return `<div class="table-wrap"><table class="tbl">
      <thead><tr><th>邀约时间</th><th>渠道</th><th class="num">触达</th><th class="num">进群</th><th class="num">入群率</th></tr></thead>
      <tbody>${rows.map(r => `<tr><td>${esc(safe(() => L.fmtInviteDate(r.date, yearOf(c.live_date), c.live_date), r.date) || '—')}</td><td class="wrap">${esc(r.channel)}</td>
        <td class="num">${esc(int(r.reach))}</td><td class="num">${esc(int(r.joined))}</td><td class="num">${esc(pct(r.rate))}</td></tr>`).join('')}</tbody>
      <tfoot><tr><td colspan="2">合计</td><td class="num">${esc(int(m.reachKnown))}</td><td class="num">${esc(int(m.joinedTotal))}</td><td class="num">${esc(pct(m.inviteRate))}</td></tr></tfoot>
    </table></div>
    <div class="tbl-note">合计入群率只算有触达数的渠道：${esc(int(m.joinedKnown))} / ${esc(int(m.reachKnown))}</div>`;
  }
  function communityTableHTML(c, m) {
    const rows = m.communityDaily || [];
    if (!rows.length) return '<div class="muted-note">没有社群数据</div>';
    return `<div class="table-wrap"><table class="tbl">
      <thead><tr><th>日期</th><th>群</th><th class="num">社群人数</th><th class="num">领取</th><th class="num">领取率</th>
        <th class="num">打卡</th><th class="num">打卡率</th><th class="num">${esc(attLabel(c.channel).replace(/人数$/, ''))}</th><th class="num">到课率</th></tr></thead>
      <tbody>${rows.map(r => `<tr><td>${esc(dateCN(r.date))}</td><td>${esc(r.group)}</td>
        <td class="num">${esc(int(r.members))}</td><td class="num">${esc(int(r.claimed))}</td><td class="num">${esc(pct(r.claimRate))}</td>
        <td class="num">${esc(int(r.checkins))}</td><td class="num">${esc(pct(r.checkinRate))}</td>
        <td class="num">${esc(int(r.attendance))}</td><td class="num">${esc(pct(r.attendRate))}</td></tr>`).join('')}</tbody>
    </table></div>`;
  }
  function conversionTableHTML(c, m) {
    const rows = c.data.conversions || [];
    if (!rows.length) return '<div class="muted-note">没有成交数据</div>';
    const cols = convCols(c.channel);
    const sums = m.sums || {};
    return `<div class="table-wrap"><table class="tbl">
      <thead><tr><th>日期</th>${cols.map(col => `<th class="num">${esc(col.label)}</th>`).join('')}</tr></thead>
      <tbody>${rows.map(r => `<tr><td>${esc(dateCN(r.date))}</td>${cols.map(col => `<td class="num">${esc(int(r[col.key]))}</td>`).join('')}</tr>`).join('')}</tbody>
      <tfoot><tr><td>合计</td>${cols.map(col => `<td class="num">${esc(int(sums[col.key] != null ? sums[col.key] : sumKey(rows, col.key)))}</td>`).join('')}</tr></tfoot>
    </table></div>`;
  }
  function depositHTML(m) {
    const s = m.sums || {};
    const rate = m.balanceRate;
    const w = rate == null ? 0 : Math.max(0, Math.min(100, rate * 100));
    return `<div class="progress-line">定金 ${esc(int(s.deposit))} · 退定金 ${esc(int(s.deposit_refund))} · 已补尾款 ${esc(int(s.balance))} · 待补 ${esc(int(m.pendingBalance))}</div>
      <div class="progress" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${w.toFixed(0)}"><div class="progress-bar" style="width:${w.toFixed(1)}%"></div></div>
      <div class="progress-cap">尾款补齐率 ${esc(pct(rate))}（已补尾款 ÷ 扣掉退定金后的定金人数 ${esc(int(m.netDeposit))}）</div>`;
  }
  function sumKey(rows, key) {
    return rows.reduce((t, r) => t + (typeof r[key] === 'number' && isFinite(r[key]) ? r[key] : 0), 0);
  }

  // 同一周两个渠道并排
  function compareHTML(list) {
    const items = list.map(c => ({ c, m: metricsOf(c) })).filter(x => x.m);
    const rows = coreMetrics().map(d => ({ label: d.label, vals: items.map(x => x.m[d.key]), fmt: pct }))
      .concat([{ label: '总成交人数', vals: items.map(x => x.m.totalDeals), fmt: int }]);
    return `<section class="section">
      <div class="section-head"><h3>${esc(chLabel('xiaoe'))} vs ${esc(chLabel('bilibili'))}</h3><span class="section-sub">本周两个渠道放在一起看，较高的一项加粗</span></div>
      <div class="table-wrap"><table class="tbl cmp">
        <thead><tr><th>指标</th>${items.map(x => `<th class="num">${chTag(x.c.channel)}<span class="cmp-name">${esc(x.c.name)}</span></th>`).join('')}</tr></thead>
        <tbody>${rows.map(r => {
          const nums = r.vals.filter(v => v != null);
          const max = nums.length > 1 ? Math.max(...nums) : null;
          return `<tr><td>${esc(r.label)}</td>${r.vals.map(v => `<td class="num ${max != null && v === max ? 'best' : ''}">${esc(r.fmt(v))}</td>`).join('')}</tr>`;
        }).join('')}</tbody>
      </table></div></section>`;
  }

  const DEAL_TRENDS = [
    { key: 'liveDeals', label: '直播转化人数', unit: '人' },
    { key: 'directDeals', label: '一对一转化人数', unit: '人' },
    { key: 'totalDeals', label: '总转化人数', unit: '人' },
  ];
  const trendMetrics = () => DEAL_TRENDS.concat(coreMetrics().map(d => ({ ...d, unit: '%' })));
  const trendValue = (p, def) => R.metricValue(p.c, def.key);
  const trendFormat = (v, def) => def.unit === '人' ? (v == null ? '—' : int(v) + ' 人') : pct(v);

  // 每渠道截至所选周最近 8 条记录。同日多期按 ID 排序，各占一个位置，不覆盖或汇总。
  function trendData(weekStartStr) {
    const end = addDays(weekStartStr, 6);
    const pool = filterCh(state.cohorts, state.dash.channel).filter(c => isYMD(c.live_date) && c.live_date <= end);
    const points = {}, channels = [], slots = new Map();
    CH_ORDER.forEach(ch => {
      const arr = pool.filter(c => c.channel === ch).sort((a, b) => a.live_date.localeCompare(b.live_date) ||
        String(a.id).localeCompare(String(b.id), undefined, { numeric: true })).slice(-8);
      if (!arr.length) return;
      channels.push(ch);
      const ordinal = {};
      points[ch] = arr.map(c => {
        const order = ordinal[c.live_date] || 0;
        ordinal[c.live_date] = order + 1;
        const key = c.live_date + '|' + order;
        slots.set(key, { key, date: c.live_date, order });
        return { c, m: metricsOf(c), key };
      }).filter(p => p.m);
    });
    const axis = Array.from(slots.values()).sort((a, b) => a.date.localeCompare(b.date) || a.order - b.order);
    axis.forEach(s => { s.label = shortMD(s.date) + (axis.some(a => a.date === s.date && a.order > 0) ? ` (${s.order + 1})` : ''); });
    return { channels, points, axis, weekStart: weekStartStr, weekEnd: end };
  }
  function trendTableHTML(td) {
    const rows = td.channels.flatMap(ch => td.points[ch]).sort((a, b) =>
      b.c.live_date.localeCompare(a.c.live_date) || CH_ORDER.indexOf(a.c.channel) - CH_ORDER.indexOf(b.c.channel) ||
      String(b.c.id).localeCompare(String(a.c.id), undefined, { numeric: true }));
    return `<div class="table-wrap"><table class="tbl">
      <thead><tr><th>直播日期</th><th>渠道</th><th>期次</th>${trendMetrics().map(m => `<th class="num">${esc(m.label)}</th>`).join('')}</tr></thead>
      <tbody>${rows.map(p => `<tr data-trend-cohort="${esc(p.c.id)}"><td>${esc(dateCN(p.c.live_date))}</td><td>${chTag(p.c.channel)}</td><td class="wrap">${esc(p.c.name)}${p.c.closed ? '' : '<span class="trend-open">未收口</span>'}</td>
        ${trendMetrics().map(d => `<td class="num" data-trend-value="${d.key}">${esc(trendFormat(trendValue(p, d), d))}</td>`).join('')}</tr>`).join('')}</tbody>
    </table></div>`;
  }
  function trendHTML(td) {
    if (!td.axis.length) return '';
    const hasChart = typeof window.Chart !== 'undefined';
    const legend = `<div class="legend">
      ${td.channels.length > 1 ? td.channels.map(ch => `<span class="lg"><i class="line-key" style="background:${esc(chCfg(ch).color)}"></i>${esc(chLabel(ch))}</span>`).join('') : ''}
      ${hasChart ? '<span class="lg lg-note"><i class="dot-key"></i>大圆点 = 当前这周</span>' : ''}</div>`;
    const cards = defs => defs.map(d => `<div class="trend-card"><div class="trend-title">${esc(d.label)}</div>
      <div class="trend-canvas"><canvas id="trend-${esc(d.key)}" role="img" aria-label="${esc(d.label)}最近八期走势"></canvas></div></div>`).join('');
    const body = hasChart
      ? `<h4 class="trend-group-title">转化人数</h4><div class="trend-grid counts">${cards(DEAL_TRENDS)}</div>
        <h4 class="trend-group-title">各环节转化率</h4><div class="trend-grid">${cards(coreMetrics())}</div>
        <details class="details trend-table"><summary>看数据表</summary><div class="details-body">${trendTableHTML(td)}</div></details>`
      : `<div class="trend-table">${trendTableHTML(td)}</div>`;
    return `<section class="section trends">
      <div class="section-head"><h3>最近 8 期走势</h3>${legend}</div>
      <p class="review-hint">每个渠道截至 ${esc(dateCN(td.weekEnd))} 最近 8 期。人数缺项显示为空，真实 0 会保留；未收口数据仍可能变化。${td.axis.some(s => s.order > 0) ? '同日多期分别显示，日期后的括号表示当天第几条记录。' : ''}</p>
      ${body}</section>`;
  }
  function drawTrends(td) {
    if (!td.axis.length || typeof window.Chart === 'undefined') return;
    const big = state.present;
    const Chart = window.Chart;
    Chart.defaults.font.family = '-apple-system, "PingFang SC", "Microsoft YaHei", sans-serif';
    Chart.defaults.color = '#6b7280';
    trendMetrics().forEach(def => {
      const canvas = document.getElementById('trend-' + def.key);
      if (!canvas) return;
      const datasets = td.channels.map(ch => {
        const color = chCfg(ch).color;
        return {
          label: chLabel(ch),
          // 只放该渠道实际存在的期次，让缺失指标产生断点，其他渠道的空位不截断曲线。
          data: td.points[ch].map(p => ({ x: p.key, y: trendValue(p, def), cohort: p.c })),
          borderColor: color, backgroundColor: color,
          pointBackgroundColor: color, pointBorderColor: '#fff', pointBorderWidth: 2,
          pointRadius: td.points[ch].map(p => p.c.live_date >= td.weekStart && p.c.live_date <= td.weekEnd ? (big ? 8 : 6.5) : (big ? 4.5 : 3.5)),
          pointHoverRadius: big ? 9 : 7, pointHitRadius: 14,
          borderWidth: 2, tension: 0.25, spanGaps: false,
        };
      });
      try {
        charts.push(new Chart(canvas, {
          type: 'line',
          data: { labels: td.axis.map(s => s.key), datasets },
          options: {
            responsive: true, maintainAspectRatio: false, animation: false,
            interaction: { mode: 'nearest', intersect: false },
            layout: { padding: { top: 6, right: 8 } },
            plugins: {
              legend: { display: false },
              tooltip: {
                filter: item => item.raw && item.raw.y != null,
                backgroundColor: '#111827', padding: 10, boxPadding: 4,
                titleFont: { size: big ? 15 : 12 }, bodyFont: { size: big ? 15 : 12 },
                callbacks: {
                  title: items => items.length ? dateCN(items[0].raw.cohort.live_date) : '',
                  label: item => {
                    const c = item.raw.cohort;
                    return `${chLabel(c.channel)} · ${c.name}：${trendFormat(item.raw.y, def)}${c.closed ? '' : '（未收口）'}`;
                  },
                },
              },
            },
            scales: {
              x: { grid: { display: false }, border: { color: '#e5e7eb' }, ticks: { font: { size: big ? 14 : 11 }, callback: value => td.axis[value]?.label || '' } },
              y: {
                beginAtZero: true, grace: '12%',
                grid: { color: '#f0f1f3' }, border: { display: false },
                ticks: { maxTicksLimit: 5, font: { size: big ? 14 : 11 }, ...(def.unit === '人' ? { precision: 0 } : {}),
                  callback: v => def.unit === '人' ? `${int(v)} 人` : `${Math.round(v * 1000) / 10}%` },
              },
            },
          },
        }));
      } catch (e) { console.error('[camp-review] chart', e); }
    });
  }

  // 明细折叠状态（toggle 不冒泡，用捕获）
  $app.addEventListener('toggle', e => {
    const d = e.target;
    if (!d.dataset || d.dataset.details == null) return;
    if (d.open) state.openDetails.add(d.dataset.details); else state.openDetails.delete(d.dataset.details);
  }, true);

  // ---------- 投屏模式 ----------
  async function enterPresent() {
    state.present = true;
    document.body.classList.add('present');
    document.documentElement.classList.add('present');
    try {
      if (document.documentElement.requestFullscreen && !document.fullscreenElement) {
        await document.documentElement.requestFullscreen();
      }
    } catch (e) { /* 浏览器不允许全屏时，只放大字号 */ }
    renderDashboard();
  }
  function exitPresent(rerender = true) {
    if (!state.present) return;
    state.present = false;
    document.body.classList.remove('present');
    document.documentElement.classList.remove('present');
    if (document.fullscreenElement && document.exitFullscreen) document.exitFullscreen().catch(() => {});
    if (rerender && parseHash().path === '/' && state.me && state.access && state.access.ok) renderDashboard();
  }
  document.addEventListener('fullscreenchange', () => { if (!document.fullscreenElement && state.present) exitPresent(); });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && state.present && !document.querySelector('.modal-mask')) exitPresent();
  });
  document.getElementById('present-exit').onclick = () => exitPresent();

  // ---------- 视图：全部期次 ----------
  async function viewList(seq) {
    if (state.loaded) $app.classList.add('refreshing');
    else $app.innerHTML = loadingHTML();
    try { await loadCohorts(); }
    catch (err) { if (seq === routeSeq) { $app.classList.remove('refreshing'); handleErr(err, '加载失败'); } return; }
    if (seq !== routeSeq) return;
    $app.classList.remove('refreshing');
    renderList();
  }
  function renderList() {
    const ch = state.listChannel;
    const list = filterCh(state.cohorts, ch).slice().sort((a, b) => String(b.live_date).localeCompare(String(a.live_date)));
    const defs = coreMetrics();
    const chOpts = [['all', '全部'], ['xiaoe', chLabel('xiaoe')], ['bilibili', chLabel('bilibili')]];
    $app.innerHTML = `
      <div class="page-head"><h1>全部期次</h1><div class="spacer"></div><a class="btn" href="#/new">录入数据</a></div>
      <div class="list-bar">
        <div class="seg" role="group" aria-label="渠道筛选">
          ${chOpts.map(([k, t]) => `<button type="button" data-lch="${k}" class="${ch === k ? 'active' : ''}" aria-pressed="${ch === k}">${esc(t)}</button>`).join('')}
        </div>
        <span class="hint">共 ${list.length} 期，按直播日期从新到旧</span>
      </div>
      ${list.length ? `<div class="card-block" style="padding:0"><div class="table-wrap"><table class="tbl list">
        <thead><tr><th>渠道</th><th>期次名称</th><th>直播日期</th><th>状态</th>
          ${defs.map(d => `<th class="num">${esc(d.label)}</th>`).join('')}<th class="num">总成交</th><th>最后更新</th><th>操作</th></tr></thead>
        <tbody>${list.map(c => {
          const m = metricsOf(c);
          return `<tr>
            <td>${chTag(c.channel)}</td>
            <td class="name">${esc(c.name || '未命名')}</td>
            <td>${esc(dateWeek(c.live_date))}</td>
            <td>${m ? statusBadge(m) : '—'}</td>
            ${defs.map(d => `<td class="num">${esc(m ? pct(m[d.key]) : '—')}</td>`).join('')}
            <td class="num">${esc(m ? int(m.totalDeals) : '—')}</td>
            <td><div>${esc(fmtTime(c.updated_at))}</div><div class="who">${esc(c.updated_by_name || '')}</div></td>
            <td><div class="ops">
              <a href="#/?week=${esc(weekOf(c.live_date))}">看板查看</a>
              <a href="#/edit/${encodeURIComponent(c.id)}">编辑</a>
              <button type="button" class="btn link del" data-del="${esc(c.id)}">删除</button>
            </div></td></tr>`;
        }).join('')}</tbody></table></div></div>`
        : `<div class="empty-state"><h2>${ch === 'all' ? '还没有任何一期数据' : esc(chLabel(ch)) + '还没有数据'}</h2>
            <p>录入后会按直播日期列在这里。</p><a class="btn" href="#/new">去录入第一期</a></div>`}`;
    $app.querySelectorAll('[data-lch]').forEach(b => { b.onclick = () => { state.listChannel = b.dataset.lch; renderList(); }; });
    $app.querySelectorAll('[data-del]').forEach(b => {
      b.onclick = async () => {
        const c = state.cohorts.find(x => sameId(x.id, b.dataset.del));
        if (!c) return;
        if (!confirm(`确定删除「${c.name || '未命名'}」吗？\n删除后可以在数据库历史表里找回。`)) return;
        b.disabled = true;
        try {
          await DB.deleteCohort(c.id);
          state.cohorts = state.cohorts.filter(x => !sameId(x.id, c.id));
          flash('已删除');
          renderList();
        } catch (err) { b.disabled = false; handleErr(err, '删除失败'); }
      };
    });
  }

  // ---------- 视图：录入 / 编辑 ----------
  let draftSeq = 0;
  function makeDraft(cohort, warnings) {
    const c = normCohort(cohort);
    delete c.warnings;
    return {
      key: 'd' + (++draftSeq),
      cohort: c,
      warnings: (warnings || []).slice(),
      parsed: { name: c.name, summary: c.summary, closed: c.closed, review: c.data.review ? clone(c.data.review) : null },
      target: 'update',     // 撞上已有的期时默认更新已有记录
      matchId: null,
      saved: false,
    };
  }
  const curDraft = () => state.ed && state.ed.drafts[state.ed.active];
  function findExisting(c) {
    if (!c.channel || !isYMD(c.live_date)) return null;
    return state.cohorts.find(x => x.channel === c.channel && x.live_date === c.live_date && !sameId(x.id, c.id)) || null;
  }
  function applyExisting(d, ex) {
    d.cohort.name = ex.name || d.cohort.name;
    d.cohort.summary = ex.summary || d.cohort.summary;
    d.cohort.closed = !!ex.closed;
    if (ex.data.review) d.cohort.data.review = clone(ex.data.review);
    else delete d.cohort.data.review;
  }
  function restoreParsedReview(d) {
    if (d.parsed.review) d.cohort.data.review = clone(d.parsed.review);
    else delete d.cohort.data.review;
  }
  // 新录入的期：检查是否撞上已有的同渠道同日期的期
  function syncExisting(d) {
    if (!state.ed || state.ed.mode !== 'new' || d.saved || d.cohort.id != null) return;
    const ex = findExisting(d.cohort);
    const id = ex ? ex.id : null;
    if (!sameId(id, d.matchId)) {
      d.matchId = id;
      if (ex && d.target === 'update') applyExisting(d, ex);
      else restoreParsedReview(d);
    }
  }
  function markDirty() {
    if (!state.dirty) { state.dirty = true; schedulePreview(); }
  }

  async function viewNew(seq) {
    $app.innerHTML = loadingHTML();
    try { await loadCohorts(); }
    catch (err) { if (seq === routeSeq) handleErr(err, '加载已有期次失败'); }
    if (seq !== routeSeq) return;
    state.ed = { mode: 'new', step: 'paste', channel: 'xiaoe', liveDate: '', pasteText: '', globalWarnings: [], drafts: [], active: 0, invalid: {}, saving: false };
    renderEditor();
  }

  async function viewEdit(id, seq) {
    $app.innerHTML = loadingHTML();
    let c;
    try {
      const realId = /^\d+$/.test(id) ? Number(id) : id;
      const res = await Promise.all([DB.getCohort(realId), loadCohorts().catch(() => null)]);
      c = res[0];
    } catch (err) {
      if (seq !== routeSeq) return;
      if (err && ['not_member', 'tables_missing', 'auth'].includes(err.code)) return handleErr(err);
      $app.innerHTML = `<div class="note-card"><h2>没找到这期数据</h2><p>${esc((err && err.message) || '可能已经被删除了')}</p>
        <div class="actions"><a class="btn" href="#/list">回到全部期次</a></div></div>`;
      return;
    }
    if (seq !== routeSeq) return;
    if (!c) {
      $app.innerHTML = '<div class="note-card"><h2>没找到这期数据</h2><p>可能已经被删除了。</p><div class="actions"><a class="btn" href="#/list">回到全部期次</a></div></div>';
      return;
    }
    const d = makeDraft(c, []);
    state.ed = { mode: 'edit', step: 'edit', drafts: [d], active: 0, invalid: {}, globalWarnings: [], saving: false, history: null };
    renderEditor();
  }

  function renderEditor() {
    const ed = state.ed;
    if (!ed) return;
    hideTip();
    if (ed.mode === 'new' && ed.step === 'paste') return renderPasteStep();
    const d = curDraft();
    $app.innerHTML = `
      <div class="page-head">
        <h1>${ed.mode === 'new' ? '录入数据' : '编辑数据'}</h1>
        ${ed.mode === 'new' ? '<button class="btn ghost small" type="button" id="back-paste">← 重新粘贴</button>'
          : `<a class="btn ghost small" href="#/?week=${esc(weekOf(d.cohort.live_date))}">在看板里查看</a>`}
      </div>
      ${ed.globalWarnings.length ? `<div class="warn-box"><b>识别提示</b>${listHTML(ed.globalWarnings)}</div>` : ''}
      ${ed.drafts.length > 1 ? `<div class="tabs" id="ed-tabs">${tabsHTML()}</div>` : ''}
      <div class="ed-layout">
        <div class="ed-main" id="ed-main">${draftHTML(d)}</div>
        <aside class="ed-side"><div class="preview" id="preview"></div></aside>
      </div>`;
    bindEditor();
    updatePreview();
  }

  function renderPasteStep() {
    const ed = state.ed;
    const radio = ch => `
      <label class="radio-card ch-${ch}"><input type="radio" name="p-ch" value="${ch}" ${ed.channel === ch ? 'checked' : ''}>
        <span><b><i></i>${esc(chLabel(ch))}</b><small>${ch === 'xiaoe' ? '直播中收定金，之后补尾款' : '直播中直接付全款'}</small></span></label>`;
    $app.innerHTML = `
      <div class="page-head"><h1>录入数据</h1></div>
      <section class="ed-card">
        <span class="field-label">渠道</span>
        <div class="radio-cards">${radio('xiaoe')}${radio('bilibili')}</div>
        <label class="field paste-date"><span>本期直播日期</span>
          <input id="p-date" type="date" value="${esc(ed.liveDate || '')}" aria-describedby="p-date-hint"></label>
        <p class="hint" id="p-date-hint">录入一期时，先选直播日期，表格里不用写期次标题。批量录入多期时留空，系统按表格内容识别日期。</p>
        <label class="field-label" for="p-text">粘贴这期的运营统计表（包含各张表的表头）</label>
        <textarea id="p-text" class="paste-area" spellcheck="false"
          placeholder="选好本期直播日期后，直接复制邀约、社群、成交表到这里。&#10;请保留列名（如渠道、触达人数、社群总人数），不需要添加「小鹅通0301」这样的期次标题。">${esc(ed.pasteText)}</textarea>
        ${ed.globalWarnings.length ? `<div class="warn-box" style="margin-top:12px"><b>没识别出数据</b>${listHTML(ed.globalWarnings)}</div>` : ''}
        <div class="actions">
          <button class="btn" type="button" id="btn-parse">识别数据</button>
          <button class="btn link" type="button" id="btn-manual">不粘贴，手动填写</button>
        </div>
        <p class="hint" style="margin-top:12px">没有期次标题时，系统按所选渠道和日期生成名称；已有同渠道、同日期的期次会提示更新。识别后仍可修改名称和日期。</p>
      </section>`;
    $app.querySelectorAll('input[name="p-ch"]').forEach(r => { r.onchange = () => { ed.channel = r.value; }; });
    const date = document.getElementById('p-date');
    date.oninput = date.onchange = () => { ed.liveDate = date.value; state.dirty = !!ed.liveDate || ed.pasteText.trim() !== ''; };
    const ta = document.getElementById('p-text');
    ta.oninput = () => { ed.pasteText = ta.value; state.dirty = !!ed.liveDate || ta.value.trim() !== ''; };
    document.getElementById('btn-parse').onclick = doParse;
    document.getElementById('btn-manual').onclick = () => {
      const c = safe(() => L.emptyCohort(ed.channel), null) || { channel: ed.channel, name: '', live_date: '', closed: false, summary: '', data: {} };
      if (ed.liveDate) { c.live_date = ed.liveDate; c.name = `${chLabel(ed.channel)} ${dateCN(ed.liveDate)} 期`; }
      const d = makeDraft({ ...c, id: null, version: 1 }, []);
      KINDS.forEach(k => { if (!d.cohort.data[k].length) d.cohort.data[k].push(emptyRow(k)); });
      ed.drafts = [d]; ed.active = 0; ed.globalWarnings = []; ed.step = 'edit';
      syncExisting(d);
      state.dirty = true;
      renderEditor();
    };
  }

  function doParse() {
    const ed = state.ed;
    const text = ed.pasteText || '';
    if (!text.trim()) { flash('先把表格内容粘贴进来'); return; }
    let res;
    try { res = L.parsePaste(text, { channel: ed.channel, liveDate: ed.liveDate, defaultYear: yearOf(ed.liveDate) }); }
    catch (err) { console.error(err); flash('识别出错了：' + (err.message || err), 5000); return; }
    const cohorts = (res && res.cohorts) || [];
    if (!cohorts.length) {
      ed.globalWarnings = (res && res.warnings && res.warnings.length) ? res.warnings : ['没找到表头（比如「渠道」「社群总人数」「已付尾款」这样的列名），请连同表头一起复制'];
      renderEditor();
      return;
    }
    ed.globalWarnings = (res.warnings || []).slice();
    ed.drafts = cohorts.map(p => makeDraft({ ...p, id: null, version: 1 }, p.warnings));
    ed.drafts.forEach(syncExisting);
    // 同一次粘贴里有重复的期
    const seen = {};
    ed.drafts.forEach(d => {
      const k = d.cohort.channel + '|' + d.cohort.live_date;
      if (seen[k]) d.warnings.push(`这次粘贴里还有一期「${seen[k]}」也是同渠道、同直播日期，请确认是不是重复了`);
      else seen[k] = d.cohort.name || '未命名';
    });
    ed.active = 0; ed.step = 'edit'; ed.invalid = {};
    state.dirty = true;
    renderEditor();
    flash(`识别出 ${cohorts.length} 期，请逐期核对后保存`);
  }

  function tabsHTML() {
    const ed = state.ed;
    return ed.drafts.map((d, i) => `
      <button type="button" class="tab ch-${esc(d.cohort.channel)} ${i === ed.active ? 'active' : ''}" data-tab="${i}">
        <i></i><span class="t-name">${esc(d.cohort.name || '未命名')}</span>
        <span class="t-sub">${esc(isYMD(d.cohort.live_date) ? shortMD(d.cohort.live_date) : '缺日期')}</span>
        ${d.saved ? '<span class="pill ok">已保存</span>' : (d.warnings.length ? `<span class="pill warn">${d.warnings.length}</span>` : '')}
      </button>`).join('');
  }
  function refreshTabs() {
    const el = document.getElementById('ed-tabs');
    if (el) el.innerHTML = tabsHTML();
  }

  function existBoxHTML(d) {
    const ed = state.ed;
    if (ed.mode === 'new' && !d.saved && d.cohort.id == null) {
      const ex = findExisting(d.cohort);
      if (!ex) return '';
      return `<div class="exist-box">
        <b>已经有一期同渠道、同直播日期的数据：「${esc(ex.name || '未命名')}」${ex.updated_at ? `（${esc(ex.updated_by_name || '')} ${esc(fmtTime(ex.updated_at))} 更新）` : ''}</b>
        <label><input type="radio" name="target-${esc(d.key)}" data-target="update" ${d.target === 'update' ? 'checked' : ''}> 更新已有的这期（保留复盘结论）</label>
        <label><input type="radio" name="target-${esc(d.key)}" data-target="new" ${d.target === 'new' ? 'checked' : ''}> 另存为新的一期</label>
      </div>`;
    }
    const ex = findExisting(d.cohort);
    return ex ? `<div class="warn-box"><b>注意</b>已经有一期同渠道、同直播日期的数据：「${esc(ex.name || '未命名')}」，确认不是重复录入。</div>` : '';
  }
  function refreshExistBox() {
    const el = document.getElementById('exist-box');
    if (el) el.innerHTML = existBoxHTML(curDraft());
  }

  function draftHTML(d) {
    const ed = state.ed;
    const c = d.cohort;
    const isEdit = ed.mode === 'edit';
    return `
      ${d.warnings.length ? `<div class="warn-box"><b>识别时发现 ${d.warnings.length} 个地方需要核对</b>${listHTML(d.warnings)}</div>` : ''}
      <div id="exist-box">${existBoxHTML(d)}</div>
      <section class="ed-card">
        <h3>基本信息</h3>
        <div class="form-grid">
          <label class="field"><span>期次名称</span><input id="f-name" value="${esc(c.name)}" placeholder="例如：小鹅通0315" maxlength="60"></label>
          <label class="field"><span>渠道</span><select id="f-channel">
            ${CH_ORDER.map(ch => `<option value="${ch}" ${c.channel === ch ? 'selected' : ''}>${esc(chLabel(ch))}</option>`).join('')}
          </select></label>
          <label class="field"><span>直播日期</span><input id="f-date" type="date" value="${esc(isYMD(c.live_date) ? c.live_date : '')}"></label>
        </div>
        <label class="check"><input type="checkbox" id="f-closed" ${c.closed ? 'checked' : ''}> 这期已经收完了（尾款和追单都不会再变）</label>
      </section>
      <section class="ed-card">
        <h3>邀约入群 <span class="count" data-count="invites"></span></h3>
        <p class="hint">邀约时间可填单日（如 8.19）或日期范围（如 8.19-23、8.28-9.3）；触达、进群人数填这段时间的合计。</p>
        <div id="tbl-invites">${tableHTML('invites')}</div>
      </section>
      <section class="ed-card">
        <h3>社群数据 <span class="count" data-count="community"></span></h3>
        <p class="hint">直播当天那一行填上${esc(attLabel(c.channel))}，系统用它算直播到课率</p>
        <div id="tbl-community">${tableHTML('community')}</div>
      </section>
      <section class="ed-card">
        <h3>每日成交 <span class="count" data-count="conversions"></span></h3>
        <p class="hint">${c.channel === 'bilibili' ? '直播成交填直播中付全款的人数；直接付款是之后 1v1 追单成交的人数。'
          : '尾款记在付定金的这一期里，之后几天补进来的尾款继续往下加行；直接付款是 1v1 追单成交的人数。'}</p>
        <div id="tbl-conversions">${tableHTML('conversions')}</div>
      </section>
      <section class="ed-card" id="unparsed-sec" ${c.data.unparsed.length ? '' : 'hidden'}>${unparsedHTML(c)}</section>
      <section class="ed-card">
        <h3>复盘结论</h3>
        <textarea id="f-summary" class="input" placeholder="这期做得好的、要改进的、下期准备试的…">${esc(c.summary)}</textarea>
      </section>
      ${isEdit ? `
      <section class="ed-card">
        <details id="repaste-box"><summary>重新粘贴覆盖表格</summary>
          <p class="hint">把最新的统计表整块粘贴进来，会替换上面三张表和「未识别的行」；期次名称、复盘结论、「已经收完了」会保留。点保存之前都不会写进数据库。</p>
          <textarea id="repaste-text" class="paste-area" style="min-height:160px;margin-top:10px" spellcheck="false" placeholder="粘贴这一期的统计表"></textarea>
          <div class="actions" style="margin-top:10px"><button class="btn small" type="button" id="btn-repaste">识别并替换表格</button></div>
        </details>
      </section>
      <section class="ed-card">
        <details id="history-box"><summary>修改记录</summary><div id="history-body" class="hint">展开后加载…</div></details>
      </section>
      <section class="ed-card danger-zone">
        <button class="btn danger small" type="button" id="btn-delete">删除这期</button>
        <span class="hint">删除后可以在数据库历史表里找回</span>
      </section>` : ''}
      <div class="bottom-save"><button class="btn" type="button" id="btn-save-bottom">${esc(saveLabel())}</button></div>`;
  }

  function unparsedHTML(c) {
    return `<h3>未识别的行 <span class="count">${c.data.unparsed.length} 行</span></h3>
      <p class="hint">如果里面的数字有用，请手动填进上面的表格；没用的点 × 删掉。</p>
      <ul class="unparsed-list">${c.data.unparsed.map((t, i) => `<li><span>${esc(t)}</span>
        <button type="button" class="row-del" data-del-unparsed="${i}" title="删除这一行" aria-label="删除这一行">×</button></li>`).join('')}</ul>`;
  }
  function refreshUnparsed() {
    const sec = document.getElementById('unparsed-sec');
    const c = curDraft().cohort;
    if (!sec) return;
    sec.hidden = !c.data.unparsed.length;
    sec.innerHTML = unparsedHTML(c);
  }

  const invKey = (d, kind, i, key) => `${d.key}|${kind}|${i}|${key}`;
  function tableHTML(kind) {
    const d = curDraft();
    const c = d.cohort;
    const cols = colsFor(kind, c.channel);
    const rows = c.data[kind];
    const head = cols.map(col => `<th class="${col.type}">${esc(col.label)}</th>`).join('') + '<th class="act"></th>';
    const body = rows.length ? rows.map((r, i) => `<tr>${cols.map(col => cellHTML(d, kind, i, col, r[col.key])).join('')}
        <td class="act"><button type="button" class="row-del" data-del-row="${i}" data-kind="${kind}" title="删除这一行" aria-label="删除这一行">×</button></td></tr>`).join('')
      : `<tr class="empty-row"><td colspan="${cols.length + 1}">还没有数据，点下面「新增一行」</td></tr>`;
    const foot = kind === 'conversions' && rows.length ? `<tfoot id="conv-foot">${convFootHTML(cols, rows)}</tfoot>` : '';
    return `<div class="table-wrap"><table class="grid-edit">
        <thead><tr>${head}</tr></thead><tbody>${body}</tbody>${foot}</table></div>
      <button type="button" class="btn small ghost add-row" data-add-row="${kind}">+ 新增一行</button>`;
  }
  function convFootHTML(cols, rows) {
    return `<tr>${cols.map((col, j) => j === 0 ? '<td>合计</td>'
      : `<td>${col.type === 'int' ? esc(int(sumKey(rows, col.key))) : ''}</td>`).join('')}<td></td></tr>`;
  }
  function cellHTML(d, kind, i, col, v) {
    const raw = state.ed.invalid[invKey(d, kind, i, col.key)];
    const val = raw != null ? raw : (v == null ? '' : String(v));
    const base = `data-kind="${kind}" data-row="${i}" data-key="${esc(col.key)}" aria-label="${esc(col.label)}" value="${esc(val)}"`;
    if (col.type === 'int') {
      return `<td><input class="cell int ${raw != null ? 'invalid' : ''}" inputmode="numeric" autocomplete="off" ${base}${raw != null ? ' title="这里需要填数字"' : ''}></td>`;
    }
    const cls = col.key === 'date' ? 'date' : (col.key === 'channel' ? 'wide' : '');
    const ph = col.key === 'date' ? (kind === 'invites' ? ' placeholder="如 8.19-23"' : ' placeholder="如 3.15"') : '';
    return `<td><input class="cell ${cls}" autocomplete="off" ${base}${ph}></td>`;
  }
  function rerenderTable(kind) {
    const el = document.getElementById('tbl-' + kind);
    if (el) el.innerHTML = tableHTML(kind);
  }
  function refreshConvFoot() {
    const el = document.getElementById('conv-foot');
    const c = curDraft().cohort;
    if (el) el.innerHTML = convFootHTML(colsFor('conversions', c.channel), c.data.conversions);
  }
  function toNum(raw) {
    let v = safe(() => L.parseNum(raw), null);
    if (v && typeof v === 'object' && 'value' in v) v = v.value;
    return typeof v === 'number' && isFinite(v) ? v : null;
  }

  function bindEditor() {
    const ed = state.ed;
    const main = document.getElementById('ed-main');
    const bp = document.getElementById('back-paste');
    if (bp) bp.onclick = () => {
      if (ed.drafts.some(d => d.saved)) { flash('已经有期次保存过了，再识别一次会生成新的待保存数据'); }
      ed.step = 'paste'; ed.globalWarnings = []; renderEditor();
    };
    const tabs = document.getElementById('ed-tabs');
    if (tabs) tabs.onclick = e => {
      const b = e.target.closest('[data-tab]');
      if (!b) return;
      ed.active = +b.dataset.tab; renderEditor();
    };
    document.querySelector('.ed-side').onclick = e => { if (e.target.closest('#btn-save')) doSave(false); };
    document.getElementById('btn-save-bottom').onclick = () => doSave(false);

    main.addEventListener('input', e => {
      const t = e.target;
      const d = curDraft();
      const c = d.cohort;
      if (t.classList.contains('cell')) return onCellInput(t);
      if (t.id === 'f-name') { c.name = t.value; markDirty(); refreshTabs(); return; }
      if (t.id === 'f-summary') { c.summary = t.value; markDirty(); return; }
    });
    main.addEventListener('change', e => {
      const t = e.target;
      const d = curDraft();
      const c = d.cohort;
      if (t.classList.contains('cell') && t.dataset.key === 'date') return normalizeCellDate(t);
      if (t.id === 'f-channel') {
        c.channel = t.value;
        // 换渠道后成交表的列变了：这一期里已经不显示的列，标红记录一并清掉（别的期不动）
        const keep = new Set(colsFor('conversions', c.channel).map(x => x.key));
        Object.keys(ed.invalid).forEach(k => {
          const p = k.split('|');
          if (p[0] === d.key && p[1] === 'conversions' && !keep.has(p[3])) delete ed.invalid[k];
        });
        markDirty(); syncExisting(d); renderEditor(); return;
      }
      if (t.id === 'f-date') {
        c.live_date = t.value; markDirty(); syncExisting(d);
        refreshExistBox(); refreshTabs(); syncBasicInputs(); schedulePreview(); return;
      }
      if (t.id === 'f-closed') { c.closed = t.checked; markDirty(); schedulePreview(); return; }
      if (t.dataset.target) {
        d.target = t.dataset.target;
        const ex = findExisting(c);
        if (d.target === 'update' && ex) applyExisting(d, ex);
        else { c.name = d.parsed.name; c.summary = d.parsed.summary; c.closed = d.parsed.closed; restoreParsedReview(d); }
        markDirty(); syncBasicInputs(); refreshTabs(); schedulePreview();
      }
    });
    main.addEventListener('click', e => {
      const t = e.target.closest('button');
      if (!t) return;
      const d = curDraft();
      const c = d.cohort;
      if (t.dataset.addRow) {
        const kind = t.dataset.addRow;
        c.data[kind].push(emptyRow(kind));
        markDirty(); rerenderTable(kind); schedulePreview();
        const inputs = document.querySelectorAll(`#tbl-${kind} tbody tr:last-child input`);
        if (inputs[0]) inputs[0].focus();
        return;
      }
      if (t.dataset.delRow != null && t.dataset.kind) {
        const kind = t.dataset.kind;
        const at = +t.dataset.delRow;
        c.data[kind].splice(at, 1);
        // 标红格子的原文跟着行号前移，删掉的那行丢弃
        const prefix = `${d.key}|${kind}|`;
        const moved = {};
        Object.keys(ed.invalid).forEach(k => {
          if (!k.startsWith(prefix)) return;
          const [row, ...rest] = k.slice(prefix.length).split('|');
          const n = +row;
          if (n > at) moved[`${prefix}${n - 1}|${rest.join('|')}`] = ed.invalid[k];
          if (n >= at) delete ed.invalid[k];
        });
        Object.assign(ed.invalid, moved);
        markDirty(); rerenderTable(kind); schedulePreview();
        return;
      }
      if (t.dataset.delUnparsed != null) {
        c.data.unparsed.splice(+t.dataset.delUnparsed, 1);
        markDirty(); refreshUnparsed(); schedulePreview();
        return;
      }
      if (t.id === 'btn-repaste') return doRepaste();
      if (t.id === 'btn-delete') return doDelete();
      if (t.dataset.loadHist != null) return loadHistoryVersion(+t.dataset.loadHist);
    });
    const hb = document.getElementById('history-box');
    if (hb) hb.addEventListener('toggle', () => { if (hb.open) loadHistory(); });
  }

  // 改了基本信息后同步输入框（切换「更新 / 另存」时会改名称、结论）
  function syncBasicInputs() {
    const c = curDraft().cohort;
    const n = document.getElementById('f-name'); if (n && n.value !== c.name) n.value = c.name;
    const s = document.getElementById('f-summary'); if (s && s.value !== c.summary) s.value = c.summary;
    const k = document.getElementById('f-closed'); if (k) k.checked = !!c.closed;
  }

  function onCellInput(t) {
    const ed = state.ed;
    const d = curDraft();
    const kind = t.dataset.kind, i = +t.dataset.row, key = t.dataset.key;
    const row = d.cohort.data[kind][i];
    if (!row) return;
    const col = colsFor(kind, d.cohort.channel).find(x => x.key === key);
    const ik = invKey(d, kind, i, key);
    if (col && col.type === 'int') {
      const raw = t.value.trim();
      const v = raw === '' ? null : toNum(raw);
      row[key] = v;
      const bad = raw !== '' && v == null && !NULL_TOKEN_RE.test(raw.replace(/\s/g, '')) && !/[%％]$/.test(raw);
      t.classList.toggle('invalid', bad);
      t.title = bad ? '这里需要填数字' : '';
      if (bad) ed.invalid[ik] = raw; else delete ed.invalid[ik];
    } else {
      row[key] = t.value;
    }
    markDirty();
    if (kind === 'conversions') refreshConvFoot();
    schedulePreview();
  }
  function normalizeCellDate(t) {
    const d = curDraft();
    const row = d.cohort.data[t.dataset.kind][+t.dataset.row];
    if (!row) return;
    const raw = t.value.trim();
    if (!raw) { row.date = ''; t.classList.remove('invalid'); schedulePreview(); return; }
    const isInvite = t.dataset.kind === 'invites';
    const year = yearOf(d.cohort.live_date);
    const v = String(safe(() => isInvite ? L.normalizeInviteDate(raw, year, d.cohort.live_date) : L.normalizeDate(raw, year), raw) || raw);
    row.date = v;
    t.value = v;
    const bad = isInvite ? !safe(() => L.isInviteDate(v, year, d.cohort.live_date), false) : !isYMD(v);
    t.classList.toggle('invalid', bad);
    t.title = bad ? (isInvite ? '没认出这个邀约时间，可以写成 8.19、8.19-23 或 8.28-9.3' : '没认出这个日期，可以写成 3.15 或 2026-03-15') : '';
    markDirty(); schedulePreview();
  }

  // 实时预览
  let pvTimer = null;
  function schedulePreview() { clearTimeout(pvTimer); pvTimer = setTimeout(updatePreview, 120); }
  function saveLabel() {
    const ed = state.ed;
    if (!ed) return '保存';
    if (ed.mode === 'edit') return '保存修改';
    const left = ed.drafts.filter(d => !d.saved).length;
    return ed.drafts.length > 1 ? `全部保存（${left} 期）` : '保存';
  }
  function invalidCount(d) {
    return Object.keys(state.ed.invalid).filter(k => k.startsWith(d.key + '|')).length;
  }
  // 保存前校验：hard 必须改完才能保存，soft 提醒后可以照样保存
  function checkDraft(d) {
    const c = d.cohort;
    const hard = [], soft = [];
    const detailed = typeof L.validateCohortDetailed === 'function' ? safe(() => L.validateCohortDetailed(c), null) : null;
    if (Array.isArray(detailed)) {
      detailed.forEach(x => (x.level === 'error' ? hard : soft).push(x.text));
    } else {
      (safe(() => L.validateCohort(c), []) || []).forEach(t => soft.push(t));
    }
    // 必填项兜底
    const req = [];
    if (!String(c.name || '').trim()) req.push('请填写期次名称');
    if (!CH_ORDER.includes(c.channel)) req.push('请选择渠道');
    if (!isYMD(c.live_date)) req.push('请填写直播日期');
    if (req.length && !hard.length) hard.push(...req);
    const n = invalidCount(d);
    if (n) hard.push(`有 ${n} 个格子填的不是数字（标红的格子），请改正`);
    return { hard, soft: soft.filter(t => !hard.includes(t)) };
  }
  function updatePreview() {
    const ed = state.ed;
    const el = document.getElementById('preview');
    const d = curDraft();
    if (!el || !d) return;
    const c = d.cohort;
    const m = safe(() => L.computeMetrics(c), null);
    const chk = checkDraft(d);
    const exId = ed.mode === 'new' && d.target === 'update' ? (findExisting(c) || {}).id : null;
    const others = state.cohorts.filter(x => !sameId(x.id, c.id) && !sameId(x.id, exId));
    const cmp = isYMD(c.live_date) ? safe(() => L.compareCohort(c, others), null) : null;
    const defs = coreMetrics();
    const kpis = defs.map(def => {
      const val = m ? m[def.key] : null;
      const x = cmp && cmp.prev && cmp.metrics ? cmp.metrics[def.key] : null;
      const f = x && x.delta != null ? safe(() => L.fmtDelta(x.delta), null) : null;
      return `<div class="pv-kpi">
        <div class="pv-label">${esc(def.label)}</div>
        <div class="pv-val">${esc(pct(val))}</div>
        <div class="pv-frac">${esc(m && PARTS[def.key] ? fracText(PARTS[def.key](m)) : '—')}</div>
        ${f ? `<div class="pv-delta ${esc(f.dir)}">较上期 ${esc(f.text)}</div>` : ''}
      </div>`;
    }).join('');
    el.innerHTML = `
      <div class="pv-head"><b>实时预览</b>${m ? statusBadge(m) : ''}</div>
      <div class="pv-grid">${kpis}</div>
      ${m ? `<div class="pv-deals">总成交 ${esc(int(m.totalDeals))} 人 · 1v1 追单 ${esc(int(m.sums ? m.sums.direct : null))} 人${c.channel !== 'bilibili' && m.pendingBalance > 0 ? ` · 待补尾款 ${esc(int(m.pendingBalance))} 人` : ''}</div>` : ''}
      ${chk.hard.length ? `<div class="pv-list bad"><b>改完这些才能保存</b>${listHTML(chk.hard)}</div>` : ''}
      ${chk.soft.length ? `<div class="pv-list warn"><b>保存前请核对</b>${listHTML(chk.soft)}</div>` : ''}
      ${!chk.hard.length && !chk.soft.length ? '<div class="pv-ok">✓ 校验通过</div>' : ''}
      ${m && m.warnings && m.warnings.length ? `<div class="pv-list warn"><b>数据提醒</b>${listHTML(m.warnings)}</div>` : ''}
      <button class="btn pv-save" type="button" id="btn-save" ${ed.saving ? 'disabled' : ''}>${ed.saving ? '保存中…' : esc(saveLabel())}</button>
      ${state.dirty ? '<div class="pv-dirty">有未保存的修改</div>' : '<div class="pv-clean">没有未保存的修改</div>'}`;
    const bb = document.getElementById('btn-save-bottom');
    if (bb) { bb.textContent = ed.saving ? '保存中…' : saveLabel(); bb.disabled = !!ed.saving; }
  }

  // 保存
  async function saveDraft(d) {
    const payload = forSave(d.cohort);
    let saved;
    if (payload.id != null) {
      saved = await DB.updateCohort(payload);
    } else {
      const ex = state.ed.mode === 'new' && d.target === 'update' ? findExisting(d.cohort) : null;
      if (ex) {
        // 表格导入仅更新表格，保留已有复盘；使用缓存版本号让并发修改触发冲突。
        if (ex.data.review) payload.data.review = clone(ex.data.review);
        else delete payload.data.review;
        saved = await DB.updateCohort({ ...payload, id: ex.id, version: ex.version });
      }
      else { delete payload.id; delete payload.version; saved = await DB.createCohort(payload); }
    }
    d.cohort = normCohort(saved);
    d.saved = true;
    replaceInCache(d.cohort);
  }
  async function doSave(skipConfirm) {
    const ed = state.ed;
    if (!ed || ed.saving) return;
    const targets = ed.drafts.filter(d => !d.saved);
    if (!targets.length) return finishSave();
    if (!skipConfirm) {
      for (const d of targets) {
        const hard = checkDraft(d).hard;
        if (hard.length) {
          if (ed.drafts.indexOf(d) !== ed.active) { ed.active = ed.drafts.indexOf(d); renderEditor(); }
          flash((targets.length > 1 ? `「${d.cohort.name || '未命名'}」：` : '') + hard.join('；'), 5000);
          return;
        }
      }
      const soft = [];
      targets.forEach(d => checkDraft(d).soft.forEach(msg =>
        soft.push((targets.length > 1 ? `「${d.cohort.name || '未命名'}」` : '') + msg)));
      if (soft.length && !confirm(`还有 ${soft.length} 条提醒：\n\n· ${soft.slice(0, 8).join('\n· ')}${soft.length > 8 ? '\n……' : ''}\n\n确定照这样保存吗？`)) return;
    }
    ed.saving = true; updatePreview();
    for (const d of targets) {
      try { await saveDraft(d); }
      catch (err) {
        ed.saving = false;
        if (ed.drafts.indexOf(d) !== ed.active) ed.active = ed.drafts.indexOf(d);
        renderEditor();
        if (err && err.code === 'conflict') conflictFor(d);
        else handleErr(err, (targets.length > 1 ? `「${d.cohort.name || '未命名'}」` : '') + '保存失败');
        return;
      }
    }
    ed.saving = false;
    finishSave();
  }
  function conflictFor(d) {
    const ex = d.cohort.id != null ? null : findExisting(d.cohort);
    const id = d.cohort.id != null ? d.cohort.id : (ex && ex.id);
    if (id == null) { flash('保存冲突，请刷新页面后重试', 5000); return; }
    showConflict({
      onReload: async () => {
        const latest = normCohort(await DB.getCohort(id));
        replaceInCache(latest);
        d.cohort = latest; d.warnings = []; d.target = 'update'; d.matchId = latest.id;
        d.reviewRestoreIntent = false;
        d.parsed = { name: latest.name, summary: latest.summary, closed: latest.closed, review: latest.data.review ? clone(latest.data.review) : null };
        Object.keys(state.ed.invalid).forEach(k => { if (k.startsWith(d.key + '|')) delete state.ed.invalid[k]; });
        state.ed.history = null;
        if (state.ed.drafts.every(x => x === d || x.saved)) state.dirty = false;
        renderEditor();
        flash('已换成最新版本');
      },
      onOverwrite: async () => {
        const latest = await DB.getCohort(id);
        const payload = forSave(d.cohort);
        // 表格编辑不改结构化复盘；只有主动载入历史版本时才恢复旧的复盘。
        if (!d.reviewRestoreIntent) {
          if (latest.data.review) payload.data.review = clone(latest.data.review);
          else delete payload.data.review;
        }
        const saved = await DB.updateCohort({ ...payload, id, version: latest.version });
        d.cohort = normCohort(saved); d.saved = true;
        replaceInCache(d.cohort);
        setTimeout(() => doSave(true), 0);   // 继续保存剩下的期
      },
    });
  }
  function finishSave() {
    const ed = state.ed;
    state.dirty = false;
    state.loaded = false;
    const first = ed.drafts[0].cohort;
    flash(ed.drafts.length > 1 ? `${ed.drafts.length} 期都已保存` : '已保存');
    state.dash.week = weekOf(first.live_date);
    state.dash.channel = 'all';
    go(`#/?week=${state.dash.week}`);
  }

  // 已有期次：重新粘贴、修改记录、删除
  function doRepaste() {
    const d = curDraft();
    const c = d.cohort;
    const text = (document.getElementById('repaste-text') || {}).value || '';
    if (!text.trim()) { flash('先把表格内容粘贴进来'); return; }
    let res;
    try { res = L.parsePaste(text, { channel: c.channel, defaultYear: yearOf(c.live_date) }); }
    catch (err) { flash('识别出错了：' + (err.message || err), 5000); return; }
    const p = res && res.cohorts && res.cohorts[0];
    if (!p) { flash('没识别出数据，请连同表头一起复制', 4000); return; }
    const np = normCohort(p);
    c.data = { ...c.data, invites: np.data.invites, community: np.data.community, conversions: np.data.conversions, unparsed: np.data.unparsed };
    const warns = (p.warnings || []).concat(res.warnings || []);
    if (res.cohorts.length > 1) warns.unshift(`粘贴内容里识别出 ${res.cohorts.length} 期，只用了第一期「${p.name || '未命名'}」`);
    if (isYMD(p.live_date) && p.live_date !== c.live_date) warns.unshift(`粘贴的数据推断出的直播日期是 ${p.live_date}，和这期填的 ${c.live_date || '（空）'} 不一样，请确认`);
    if (p.channel && p.channel !== c.channel) warns.unshift(`粘贴的数据看起来是${chLabel(p.channel)}的，这期渠道是${chLabel(c.channel)}，请确认`);
    d.warnings = warns;
    Object.keys(state.ed.invalid).forEach(k => { if (k.startsWith(d.key + '|')) delete state.ed.invalid[k]; });
    state.dirty = true;
    renderEditor();
    flash('表格已替换，核对无误后点「保存修改」');
  }
  async function loadHistory() {
    const ed = state.ed;
    const body = document.getElementById('history-body');
    const c = curDraft().cohort;
    if (!body || c.id == null) return;
    body.textContent = '加载中…';
    try { ed.history = (await DB.listHistory(c.id)) || []; }
    catch (err) { body.textContent = '加载失败：' + ((err && err.message) || err); return; }
    if (!ed.history.length) { body.textContent = '还没有修改记录'; return; }
    body.classList.remove('hint');
    body.innerHTML = `<p class="hint">点「载入这个版本」会把那时的数据填进上面的编辑区，确认后再点保存才会生效。</p>
      <ul class="history-list">${ed.history.map((h, i) => `<li>
        <span class="h-time">${esc(fmtTime(h.changed_at))}</span>
        <span class="h-act">${esc(HIST_ACTION[h.action] || h.action)}</span>
        <span class="h-who">${esc(h.changed_by_name || '未知')}${h.snapshot && h.snapshot.version != null ? ` · 第 ${esc(h.snapshot.version)} 版` : ''}</span>
        ${h.snapshot ? `<button type="button" class="btn small ghost" data-load-hist="${i}">载入这个版本</button>` : ''}
      </li>`).join('')}</ul>`;
  }
  function loadHistoryVersion(i) {
    const ed = state.ed;
    const h = ed.history && ed.history[i];
    if (!h || !h.snapshot) return;
    if (state.dirty && !confirm('编辑区里有还没保存的修改，载入旧版本会替换掉它们，继续吗？')) return;
    const d = curDraft();
    const s = h.snapshot;
    d.reviewRestoreIntent = true;
    const keep = { id: d.cohort.id, version: d.cohort.version, created_at: d.cohort.created_at, updated_at: d.cohort.updated_at, updated_by_name: d.cohort.updated_by_name };
    d.cohort = normCohort({
      ...d.cohort, ...keep,
      name: s.name != null ? s.name : d.cohort.name,
      channel: s.channel || d.cohort.channel,
      live_date: s.live_date || d.cohort.live_date,
      closed: !!s.closed,
      summary: s.summary || '',
      data: s.data || {},
    });
    d.warnings = [];
    Object.keys(ed.invalid).forEach(k => { if (k.startsWith(d.key + '|')) delete ed.invalid[k]; });
    state.dirty = true;
    renderEditor();
    window.scrollTo({ top: 0, behavior: 'smooth' });
    flash(`已载入 ${fmtTime(h.changed_at)} 的版本，确认无误后点「保存修改」`, 4000);
  }
  async function doDelete() {
    const c = curDraft().cohort;
    if (c.id == null) return;
    if (!confirm(`确定删除「${c.name || '未命名'}」吗？\n删除后可以在数据库历史表里找回。`)) return;
    try {
      await DB.deleteCohort(c.id);
      state.cohorts = state.cohorts.filter(x => !sameId(x.id, c.id));
      state.dirty = false;
      flash('已删除');
      go('#/list');
    } catch (err) { handleErr(err, '删除失败'); }
  }

  // ---------- 视图：成员 ----------
  async function viewMembers(seq) {
    $app.innerHTML = loadingHTML();
    let list;
    try { list = (await DB.listMembers()) || []; }
    catch (err) { if (seq === routeSeq) { $app.innerHTML = ''; handleErr(err, '加载成员失败'); } return; }
    if (seq !== routeSeq) return;
    const myEmail = String((state.me && state.me.user && state.me.user.email) || '').toLowerCase();
    $app.innerHTML = `
      <div class="page-head"><h1>成员</h1></div>
      <section class="card-block">
        <p class="lead">对方先用这个邮箱在 Skill 市集或本页注册账号，你把邮箱加进来后他就能看到复盘数据。</p>
        <form class="inline-form" id="m-form">
          <input class="input" type="email" id="m-email" required placeholder="同事的邮箱，例如 name@example.com" autocomplete="off">
          <button class="btn" id="m-add">添加成员</button>
        </form>
        <div class="table-wrap"><table class="tbl">
          <thead><tr><th>邮箱</th><th>添加时间</th><th>添加人</th><th></th></tr></thead>
          <tbody>${list.length ? list.map(mb => {
            const self = String(mb.email || '').toLowerCase() === myEmail;
            return `<tr><td>${esc(mb.email)}${self ? ' <span class="muted">（你）</span>' : ''}</td>
              <td>${esc(fmtTime(mb.added_at))}</td><td>${esc(mb.added_by_name || '—')}</td>
              <td>${self ? '' : `<button type="button" class="btn small danger" data-rm="${esc(mb.email)}">移除</button>`}</td></tr>`;
          }).join('') : '<tr><td colspan="4" class="muted">还没有成员</td></tr>'}</tbody>
        </table></div>
      </section>`;
    document.getElementById('m-form').onsubmit = async e => {
      e.preventDefault();
      const input = document.getElementById('m-email');
      const email = input.value.trim().toLowerCase();
      if (!email) return;
      const btn = document.getElementById('m-add');
      btn.disabled = true;
      try { await DB.addMember(email); flash(`已添加 ${email}`); viewMembers(routeSeq); }
      catch (err) { btn.disabled = false; handleErr(err, '添加失败'); }
    };
    $app.querySelectorAll('[data-rm]').forEach(b => {
      b.onclick = async () => {
        if (!confirm(`确定移除 ${b.dataset.rm} 吗？移除后他就看不到复盘数据了。`)) return;
        b.disabled = true;
        try { await DB.removeMember(b.dataset.rm); flash('已移除'); viewMembers(routeSeq); }
        catch (err) { b.disabled = false; handleErr(err, '移除失败'); }
      };
    });
  }

  // ---------- 启动 ----------
  async function boot() {
    if (!L || !DB) {
      $app.innerHTML = '<div class="note-card"><h2>页面没加载完整</h2><p>有脚本文件没加载成功，请刷新重试。</p></div>';
      return;
    }
    try { await DB.init(window.CAMP_REVIEW_CONFIG || {}); }
    catch (err) {
      $app.innerHTML = `<div class="note-card"><h2>连接数据库失败</h2><p>${esc((err && err.message) || err)}</p>
        <div class="actions"><button class="btn" onclick="location.reload()">刷新重试</button></div></div>`;
      return;
    }
    document.getElementById('demo-bar').hidden = DB.mode !== 'demo';
    try { if (typeof DB.onAuthChange === 'function') DB.onAuthChange(onAuthChange); } catch (e) { /* 忽略 */ }
    await loadSession();
    route();
  }
  boot();
})();
