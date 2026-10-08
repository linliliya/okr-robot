// 数据层：Supabase 适配器 + 演示模式（localStorage）适配器，两者对外行为一致
// 所有错误都是 Error，带 err.code：
//   conflict | tables_missing | not_member | auth | network | unknown | not_found | invalid
(function (root) {
  'use strict';

  const DEMO_KEY = 'camp-review-demo-v1';
  const CHANNEL_KEYS = ['xiaoe', 'bilibili'];
  const HISTORY_LIMIT = 50;
  const EMAIL_RE = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

  const MSG = {
    tables_missing: '数据库还没初始化，请按 camp-review/SETUP.md 执行 setup.sql',
    not_member: '你的账号还没开通复盘系统，请让已开通的同事在『成员』页添加你的邮箱',
    network: '网络连接失败，请检查网络后重试',
    server_down: '数据库暂时连不上，可能在休眠，请稍后重试（管理员可参考 camp-review/SETUP.md）',
    need_login: '请先登录',
    expired: '登录已过期，请重新登录',
    conflict: '这期数据刚被别人改过',
  };

  // ---------- 小工具 ----------
  const clone = x => (x === undefined ? undefined : JSON.parse(JSON.stringify(x)));
  const normEmail = s => String(s == null ? '' : s).replace(/[\s　]+/g, '').toLowerCase();
  const nowISO = () => new Date().toISOString();

  function makeError(code, message, cause) {
    const err = new Error(message);
    err.code = code;
    if (cause !== undefined) err.cause = cause;
    Object.defineProperty(err, '__campdb', { value: true });
    return err;
  }
  const isCampError = e => !!(e && e.__campdb);

  function notMemberMsg(email) {
    return email
      ? `你的账号（${email}）还没开通复盘系统，请让已开通的同事在『成员』页添加你的邮箱`
      : MSG.not_member;
  }

  function validDate(s) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    if (!m) return false;
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    return d.getUTCFullYear() === +m[1] && d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[3];
  }

  function toId(id) {
    const n = Number(id);
    if (!Number.isInteger(n) || n <= 0) throw makeError('invalid', '期次编号不对');
    return n;
  }

  // data 统一成四个数组都在的对象（其余字段原样保留）
  function normData(d) {
    const src = d && typeof d === 'object' && !Array.isArray(d) ? d : {};
    const out = Object.assign({}, src);
    for (const k of ['invites', 'community', 'conversions', 'unparsed']) {
      out[k] = Array.isArray(src[k]) ? src[k] : [];
    }
    return out;
  }

  // 保存前校验并只挑出可写字段（两个适配器共用）
  function cohortPayload(c) {
    if (!c || typeof c !== 'object') throw makeError('invalid', '没有要保存的数据');
    if (!CHANNEL_KEYS.includes(c.channel)) throw makeError('invalid', '请选择渠道（小鹅通 / B站）');
    const name = String(c.name == null ? '' : c.name).trim();
    if (!name) throw makeError('invalid', '请填写期次名称');
    const live_date = String(c.live_date == null ? '' : c.live_date).trim();
    if (!validDate(live_date)) throw makeError('invalid', '请填写正确的直播日期');
    let data;
    try {
      data = normData(clone(c.data || {}));
    } catch (e) {
      throw makeError('invalid', '表格数据格式不对，没法保存', e);
    }
    return {
      channel: c.channel,
      name,
      live_date,
      closed: !!c.closed,
      summary: String(c.summary == null ? '' : c.summary),
      data,
    };
  }

  // 数据库行（或历史快照）→ 前端 Cohort 对象
  function toCohort(row, names) {
    const r = row || {};
    const n = names || {};
    return {
      id: r.id == null ? null : Number(r.id),
      channel: r.channel,
      name: r.name || '',
      live_date: String(r.live_date || '').slice(0, 10),
      closed: !!r.closed,
      summary: r.summary || '',
      version: Number(r.version) || 1,
      data: normData(r.data),
      created_at: r.created_at || null,
      updated_at: r.updated_at || null,
      created_by_name: n.created_by_name || '',
      updated_by_name: n.updated_by_name || '',
    };
  }

  function validateSignUp(email, password, display_name) {
    if (!String(display_name || '').trim()) throw makeError('auth', '请填写姓名');
    if (!EMAIL_RE.test(email)) throw makeError('auth', '邮箱格式不对');
    if (String(password || '').length < 6) throw makeError('auth', '密码至少 6 位');
  }

  // ---------- 错误归类 ----------
  function isTablesMissing(e) {
    const code = String((e && e.code) || '');
    const msg = String((e && e.message) || '');
    return ['42P01', '42883', 'PGRST200', 'PGRST202', 'PGRST205'].includes(code)
      || /relation .* does not exist/i.test(msg)
      || /Could not find (the table|the function|a relationship)/i.test(msg);
  }

  function isNetworkError(e) {
    if (!e) return false;
    const msg = String(e.message || e);
    return e.name === 'AuthRetryableFetchError'
      || /Failed to fetch|NetworkError|Load failed|fetch failed|Network request failed|ERR_NETWORK|ECONNREFUSED|ENOTFOUND/i.test(msg);
  }

  // 数据库请求的错误 → 统一错误；ctx.duplicate 是唯一键冲突时的提示
  function toCampError(e, ctx, status) {
    if (isCampError(e)) return e;
    const c = ctx || {};
    const code = String((e && e.code) || '');
    const msg = String((e && (e.message || e.error_description)) || e || '');

    if (isNetworkError(e) || status === 0) return makeError('network', MSG.network, e);
    if (/^PGRST00[0-3]$/.test(code) || (status >= 500 && !code)) return makeError('network', MSG.server_down, e);
    if (isTablesMissing(e)) return makeError('tables_missing', MSG.tables_missing, e);
    if (/^PGRST3\d\d$/.test(code) || /JWT/i.test(msg)) return makeError('auth', MSG.expired, e);
    if (code === '42501' || /row-level security|permission denied/i.test(msg)) {
      if (status === 401) return makeError('auth', MSG.need_login, e);
      return makeError('not_member', MSG.not_member, e);
    }
    if (code === '23505') return makeError('invalid', c.duplicate || '数据重复了，请刷新后再试', e);
    if (/^22/.test(code) || code === '23502' || code === '23514') {
      return makeError('invalid', '数据格式不对，没保存成功：' + msg, e);
    }
    if (code === 'P0001') return makeError('invalid', msg, e);
    return makeError('unknown', '出错了：' + msg, e);
  }

  // 登录 / 注册的错误 → 中文提示（与 Skill 市集一致）
  function toAuthError(e) {
    if (isCampError(e)) return e;
    if (isNetworkError(e)) return makeError('network', MSG.network, e);
    const msg = String((e && e.message) || e || '');
    const status = e && e.status;
    let text;
    if (/Invalid login/i.test(msg)) text = '邮箱或密码不对';
    else if (/already (been )?registered|already exists/i.test(msg)) text = '这个邮箱已经注册过了，直接登录吧';
    else if (/not confirmed|confirm/i.test(msg)) text = '账号需要邮箱验证才能登录：请管理员在 Supabase 后台关闭邮箱验证（见 skill-hub/SETUP.md 第 4 步）';
    else if (/Password should|weak.?password/i.test(msg)) text = '密码太简单了，至少 6 位';
    else if (/rate limit|too many/i.test(msg) || status === 429) text = '操作太频繁了，请过几分钟再试';
    else if (/invalid format|validate email/i.test(msg)) text = '邮箱格式不对';
    else if (/Signups not allowed/i.test(msg)) text = '注册功能被关闭了，请联系管理员';
    else if (status >= 500) return makeError('network', MSG.server_down, e);
    else text = '登录失败：' + msg;
    return makeError('auth', text, e);
  }

  // ============================================================
  // Supabase 适配器
  // ============================================================
  const COHORT_COLS = 'id, channel, name, live_date, closed, summary, data, version, created_at, updated_at, '
    + 'creator:profiles!camp_cohorts_created_by_fkey(display_name), '
    + 'editor:profiles!camp_cohorts_updated_by_fkey(display_name)';
  const HISTORY_COLS = 'id, cohort_id, action, changed_at, snapshot, '
    + 'changer:profiles!camp_cohort_history_changed_by_fkey(display_name)';
  const MEMBER_COLS = 'email, added_at, adder:profiles!camp_review_members_added_by_fkey(display_name)';

  function createSupabaseAdapter(cfg) {
    let sb = null;
    let loadError;
    const lib = root.supabase;
    if (lib && typeof lib.createClient === 'function') {
      try { sb = lib.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY); } catch (e) { loadError = e; }
    }
    function client() {
      if (!sb) throw makeError('network', 'Supabase 组件没加载成功，请检查网络后刷新页面', loadError);
      return sb;
    }

    // 执行一次请求：出错抛统一错误，成功返回整个响应
    async function exec(fn, ctx) {
      let res;
      try { res = await fn(client()); } catch (e) { throw toCampError(e, ctx); }
      if (res && res.error) throw toCampError(res.error, ctx, res.status);
      return res || {};
    }
    const run = async (fn, ctx) => (await exec(fn, ctx)).data;

    const rowToCohort = row => toCohort(row, {
      created_by_name: row && row.creator && row.creator.display_name,
      updated_by_name: row && row.editor && row.editor.display_name,
    });

    async function rawSession() {
      let res;
      try { res = await client().auth.getSession(); } catch (e) { throw toAuthError(e); }
      if (res.error) {
        if (isNetworkError(res.error)) throw makeError('network', MSG.network, res.error);
        return null;   // 刷新令牌失效等：当作未登录
      }
      return (res.data && res.data.session) || null;
    }
    async function myEmail() {
      const s = await rawSession();
      return s && s.user ? normEmail(s.user.email) : '';
    }

    const api = {
      async getSession() {
        const session = await rawSession();
        if (!session || !session.user) return null;
        const user = { id: session.user.id, email: session.user.email || '' };
        let profile = null;
        try {
          const { data, error } = await client().from('profiles')
            .select('display_name, department').eq('id', user.id).maybeSingle();
          if (!error) profile = data;
        } catch (_) { /* 资料拿不到不影响登录 */ }
        return {
          user,
          profile: {
            display_name: (profile && profile.display_name) || user.email,
            department: (profile && profile.department) || '',
          },
        };
      },

      // 只在「登录的人变了」时回调（登录、退出、切换账号），令牌刷新不打扰页面
      onAuthChange(cb) {
        if (!sb || typeof cb !== 'function') return () => {};
        let lastId;
        const { data } = sb.auth.onAuthStateChange((event, session) => {
          const id = (session && session.user && session.user.id) || null;
          if (event === 'INITIAL_SESSION') { lastId = id; return; }
          if (id === lastId) return;
          lastId = id;
          // 放到下一轮再调：在 supabase 回调里直接发请求会卡住
          setTimeout(async () => {
            let s = null;
            if (id) {
              try { s = await api.getSession(); } catch (_) {
                const email = session.user.email || '';
                s = { user: { id, email }, profile: { display_name: email, department: '' } };
              }
            }
            try { cb(s, event); } catch (e) { console.error(e); }
          }, 0);
        });
        return () => data && data.subscription && data.subscription.unsubscribe();
      },

      async signIn(email, password) {
        const e = normEmail(email);
        if (!e || !password) throw makeError('auth', '请填写邮箱和密码');
        try {
          const { error } = await client().auth.signInWithPassword({ email: e, password });
          if (error) throw error;
        } catch (err) { throw toAuthError(err); }
        return api.getSession();
      },

      async signUp(email, password, display_name, department) {
        const e = normEmail(email);
        validateSignUp(e, password, display_name);
        let data;
        try {
          const res = await client().auth.signUp({
            email: e,
            password,
            options: { data: {
              display_name: String(display_name).trim(),
              department: String(department || '').trim() || '其他',
            } },
          });
          if (res.error) throw res.error;
          data = res.data || {};
        } catch (err) { throw toAuthError(err); }
        // 开了邮箱验证时，重复注册不会报错，只会返回空 identities
        if (data.user && Array.isArray(data.user.identities) && data.user.identities.length === 0) {
          throw makeError('auth', '这个邮箱已经注册过了，直接登录吧');
        }
        if (!data.session) {
          throw makeError('auth', '注册成功，但需要先点邮件里的验证链接才能登录（管理员可在 Supabase 后台关闭邮箱验证）');
        }
        return api.getSession();
      },

      async signOut() {
        try {
          const { error } = await client().auth.signOut();
          if (error) throw error;
        } catch (_) {
          // 网络不通时至少清掉本机登录状态
          try { await client().auth.signOut({ scope: 'local' }); } catch (__) { /* 忽略 */ }
        }
      },

      async checkAccess() {
        try {
          const session = await rawSession();
          if (!session) return { ok: false, reason: 'error', code: 'auth', message: MSG.need_login };
          const c = client();
          const [rpc, probe] = await Promise.all([
            c.rpc('is_camp_review_member'),
            c.from('camp_cohorts').select('id').limit(1),
          ]);
          for (const r of [rpc, probe]) {
            if (r.error && isTablesMissing(r.error)) {
              return { ok: false, reason: 'tables_missing', code: 'tables_missing', message: MSG.tables_missing };
            }
          }
          if (rpc.error) {
            const err = toCampError(rpc.error, null, rpc.status);
            return { ok: false, reason: err.code === 'not_member' ? 'not_member' : 'error', code: err.code, message: err.message };
          }
          if (rpc.data !== true) {
            return { ok: false, reason: 'not_member', code: 'not_member', message: notMemberMsg(session.user.email) };
          }
          if (probe.error) {
            const err = toCampError(probe.error, null, probe.status);
            return { ok: false, reason: err.code === 'not_member' ? 'not_member' : 'error', code: err.code, message: err.message };
          }
          return { ok: true };
        } catch (e) {
          const err = toCampError(e);
          return { ok: false, reason: err.code === 'tables_missing' ? 'tables_missing' : 'error', code: err.code, message: err.message };
        }
      },

      async listCohorts() {
        const PAGE = 1000;
        const all = [];
        for (let from = 0; ; from += PAGE) {
          const res = await exec(c => c.from('camp_cohorts')
            .select(COHORT_COLS, { count: 'exact' })
            .order('live_date', { ascending: false })
            .order('id', { ascending: false })
            .range(from, from + PAGE - 1));
          const rows = res.data || [];
          all.push(...rows);
          const total = typeof res.count === 'number' ? res.count : all.length;
          if (!rows.length || all.length >= total) break;
        }
        return all.map(rowToCohort);
      },

      async getCohort(id) {
        const n = toId(id);
        const row = await run(c => c.from('camp_cohorts').select(COHORT_COLS).eq('id', n).maybeSingle());
        if (!row) throw makeError('not_found', '没找到这期数据（可能已被删除）');
        return rowToCohort(row);
      },

      async createCohort(cohort) {
        const payload = cohortPayload(cohort);
        const row = await run(c => c.from('camp_cohorts').insert(payload).select(COHORT_COLS).single());
        return rowToCohort(row);
      },

      // 乐观锁：只有数据库里的 version 还等于手上的 version 才写入
      async updateCohort(cohort) {
        if (!cohort || cohort.id == null) throw makeError('invalid', '这期还没保存过，请用新建');
        const id = toId(cohort.id);
        const version = Number(cohort.version);
        if (!Number.isInteger(version)) throw makeError('invalid', '缺少版本号，请重新加载这期再保存');
        const payload = cohortPayload(cohort);
        const rows = await run(c => c.from('camp_cohorts').update(payload)
          .eq('id', id).eq('version', version).select(COHORT_COLS));
        if (rows && rows.length) return rowToCohort(rows[0]);
        const latest = await run(c => c.from('camp_cohorts').select(COHORT_COLS).eq('id', id).maybeSingle());
        if (!latest) throw makeError('not_found', '这期已经被删除了');
        const err = makeError('conflict', MSG.conflict);
        err.latest = rowToCohort(latest);
        throw err;
      },

      async deleteCohort(id) {
        const n = toId(id);
        const rows = await run(c => c.from('camp_cohorts').delete().eq('id', n).select('id'));
        if (!rows || !rows.length) throw makeError('not_found', '这期已经不存在了（可能已被别人删除）');
        return { id: n };
      },

      async listHistory(cohortId) {
        const n = toId(cohortId);
        const rows = await run(c => c.from('camp_cohort_history').select(HISTORY_COLS)
          .eq('cohort_id', n).order('id', { ascending: false }).limit(HISTORY_LIMIT));
        return (rows || []).map(h => ({
          id: h.id,
          action: h.action,
          changed_at: h.changed_at,
          changed_by_name: (h.changer && h.changer.display_name) || '',
          snapshot: toCohort(h.snapshot),
        }));
      },

      async listMembers() {
        const [rows, me] = await Promise.all([
          run(c => c.from('camp_review_members').select(MEMBER_COLS).order('added_at', { ascending: true })),
          myEmail(),
        ]);
        return (rows || []).map(m => ({
          email: m.email,
          added_at: m.added_at,
          added_by_name: (m.adder && m.adder.display_name) || '',
          is_self: m.email === me,
        }));
      },

      async addMember(email) {
        const e = normEmail(email);
        if (!EMAIL_RE.test(e)) throw makeError('invalid', '邮箱格式不对');
        const [row, me] = await Promise.all([
          run(c => c.from('camp_review_members').insert({ email: e }).select(MEMBER_COLS).single(),
            { duplicate: '这个邮箱已经是成员了' }),
          myEmail(),
        ]);
        return {
          email: row.email,
          added_at: row.added_at,
          added_by_name: (row.adder && row.adder.display_name) || '',
          is_self: row.email === me,
        };
      },

      async removeMember(email) {
        const e = normEmail(email);
        if (!e) throw makeError('invalid', '邮箱不能为空');
        if (e === await myEmail()) throw makeError('invalid', '不能把自己移出成员，避免把自己锁在外面');
        const rows = await run(c => c.from('camp_review_members').delete().eq('email', e).select('email'));
        if (!rows || !rows.length) throw makeError('not_found', '这个邮箱已经不在成员名单里了');
        return { email: e };
      },
    };
    return api;
  }

  // ============================================================
  // 演示模式适配器：数据存在当前浏览器 localStorage，不内置任何业务数据
  // ============================================================
  function getLocalStorage() {
    try {
      const s = root.localStorage;
      if (!s) return null;
      const k = '__camp_review_probe__';
      s.setItem(k, '1');
      s.removeItem(k);
      return s;
    } catch (_) {
      return null;
    }
  }

  function createDemoAdapter() {
    const USER = { id: 'demo-user', email: 'demo@example.com' };
    const PROFILE = { display_name: '演示用户', department: '运营' };
    const NAMES = { 'demo-user': '演示用户' };
    const ls = getLocalStorage();
    let memoryStore = null;   // localStorage 不可用时退回内存（刷新即丢）
    let signedIn = true;
    const listeners = new Set();

    function emptyStore() {
      const t = nowISO();
      return {
        seq: { cohort: 0, history: 0 },
        cohorts: [],
        history: [],
        members: [
          { email: USER.email, added_by: USER.id, added_at: t },
          { email: 'teammate@example.com', added_by: USER.id, added_at: t },
        ],
      };
    }

    function load() {
      let s = null;
      if (ls) {
        try {
          const raw = ls.getItem(DEMO_KEY);
          if (raw) s = JSON.parse(raw);
        } catch (_) { s = null; }
      } else if (memoryStore) {
        s = clone(memoryStore);
      }
      if (!s || typeof s !== 'object') return emptyStore();
      const base = emptyStore();
      return {
        seq: Object.assign(base.seq, s.seq),
        cohorts: Array.isArray(s.cohorts) ? s.cohorts : [],
        history: Array.isArray(s.history) ? s.history : [],
        members: Array.isArray(s.members) ? s.members : base.members,
      };
    }

    function save(s) {
      if (!ls) { memoryStore = clone(s); return; }
      try {
        ls.setItem(DEMO_KEY, JSON.stringify(s));
      } catch (e) {
        throw makeError('unknown', '浏览器存储空间不够或被禁用了，演示数据没保存下来', e);
      }
    }

    // 与真库触发器一致：每次写入记一条整行快照；每期只留最近 50 条（真库也只读 50 条）
    function logHistory(s, action, row) {
      s.seq.history += 1;
      s.history.push({
        id: s.seq.history,
        cohort_id: row.id,
        action,
        snapshot: clone(row),
        changed_by: signedIn ? USER.id : null,
        changed_at: nowISO(),
      });
      const mine = s.history.filter(h => h.cohort_id === row.id);
      if (mine.length > HISTORY_LIMIT) {
        const drop = new Set(mine.slice(0, mine.length - HISTORY_LIMIT).map(h => h.id));
        s.history = s.history.filter(h => !drop.has(h.id));
      }
    }

    const rowToCohort = row => toCohort(row, {
      created_by_name: NAMES[row.created_by] || '',
      updated_by_name: NAMES[row.updated_by] || '',
    });

    function requireLogin() {
      if (!signedIn) throw makeError('auth', MSG.need_login);
    }
    function session() {
      return signedIn ? { user: clone(USER), profile: clone(PROFILE) } : null;
    }
    function notify(event) {
      const s = session();
      for (const cb of listeners) {
        setTimeout(() => { try { cb(clone(s), event); } catch (e) { console.error(e); } }, 0);
      }
    }
    const tick = () => Promise.resolve();

    const api = {
      async getSession() { await tick(); return session(); },

      onAuthChange(cb) {
        if (typeof cb !== 'function') return () => {};
        listeners.add(cb);
        return () => listeners.delete(cb);
      },

      async signIn(email, password) {
        await tick();
        if (!normEmail(email) || !password) throw makeError('auth', '请填写邮箱和密码');
        const changed = !signedIn;
        signedIn = true;
        if (changed) notify('SIGNED_IN');
        return session();
      },

      async signUp(email, password, display_name) {
        await tick();
        validateSignUp(normEmail(email), password, display_name);
        const changed = !signedIn;
        signedIn = true;
        if (changed) notify('SIGNED_IN');
        return session();
      },

      async signOut() {
        await tick();
        if (!signedIn) return;
        signedIn = false;
        notify('SIGNED_OUT');
      },

      async checkAccess() {
        await tick();
        return signedIn ? { ok: true } : { ok: false, reason: 'error', code: 'auth', message: MSG.need_login };
      },

      async listCohorts() {
        await tick(); requireLogin();
        return load().cohorts.slice()
          .sort((a, b) => (a.live_date < b.live_date ? 1 : a.live_date > b.live_date ? -1 : b.id - a.id))
          .map(r => rowToCohort(clone(r)));
      },

      async getCohort(id) {
        await tick(); requireLogin();
        const n = toId(id);
        const row = load().cohorts.find(r => r.id === n);
        if (!row) throw makeError('not_found', '没找到这期数据（可能已被删除）');
        return rowToCohort(clone(row));
      },

      async createCohort(cohort) {
        await tick(); requireLogin();
        const payload = cohortPayload(cohort);
        const s = load();
        const t = nowISO();
        s.seq.cohort += 1;
        const row = Object.assign({ id: s.seq.cohort }, payload, {
          version: 1, created_by: USER.id, updated_by: USER.id, created_at: t, updated_at: t,
        });
        s.cohorts.push(row);
        logHistory(s, 'insert', row);
        save(s);
        return rowToCohort(clone(row));
      },

      async updateCohort(cohort) {
        await tick(); requireLogin();
        if (!cohort || cohort.id == null) throw makeError('invalid', '这期还没保存过，请用新建');
        const id = toId(cohort.id);
        const version = Number(cohort.version);
        if (!Number.isInteger(version)) throw makeError('invalid', '缺少版本号，请重新加载这期再保存');
        const payload = cohortPayload(cohort);
        const s = load();
        const row = s.cohorts.find(r => r.id === id);
        if (!row) throw makeError('not_found', '这期已经被删除了');
        if (row.version !== version) {
          const err = makeError('conflict', MSG.conflict);
          err.latest = rowToCohort(clone(row));
          throw err;
        }
        Object.assign(row, payload, { version: row.version + 1, updated_by: USER.id, updated_at: nowISO() });
        logHistory(s, 'update', row);
        save(s);
        return rowToCohort(clone(row));
      },

      async deleteCohort(id) {
        await tick(); requireLogin();
        const n = toId(id);
        const s = load();
        const idx = s.cohorts.findIndex(r => r.id === n);
        if (idx < 0) throw makeError('not_found', '这期已经不存在了（可能已被别人删除）');
        const [row] = s.cohorts.splice(idx, 1);
        logHistory(s, 'delete', row);
        save(s);
        return { id: n };
      },

      async listHistory(cohortId) {
        await tick(); requireLogin();
        const n = toId(cohortId);
        return load().history.filter(h => h.cohort_id === n)
          .sort((a, b) => b.id - a.id)
          .slice(0, HISTORY_LIMIT)
          .map(h => ({
            id: h.id,
            action: h.action,
            changed_at: h.changed_at,
            changed_by_name: NAMES[h.changed_by] || '',
            snapshot: toCohort(clone(h.snapshot)),
          }));
      },

      async listMembers() {
        await tick(); requireLogin();
        return load().members.slice()
          .sort((a, b) => (a.added_at < b.added_at ? -1 : a.added_at > b.added_at ? 1 : 0))
          .map(m => ({
            email: m.email,
            added_at: m.added_at,
            added_by_name: NAMES[m.added_by] || '',
            is_self: m.email === USER.email,
          }));
      },

      async addMember(email) {
        await tick(); requireLogin();
        const e = normEmail(email);
        if (!EMAIL_RE.test(e)) throw makeError('invalid', '邮箱格式不对');
        const s = load();
        if (s.members.some(m => m.email === e)) throw makeError('invalid', '这个邮箱已经是成员了');
        const m = { email: e, added_by: USER.id, added_at: nowISO() };
        s.members.push(m);
        save(s);
        return { email: e, added_at: m.added_at, added_by_name: NAMES[USER.id], is_self: e === USER.email };
      },

      async removeMember(email) {
        await tick(); requireLogin();
        const e = normEmail(email);
        if (!e) throw makeError('invalid', '邮箱不能为空');
        if (e === USER.email) throw makeError('invalid', '不能把自己移出成员，避免把自己锁在外面');
        const s = load();
        const before = s.members.length;
        s.members = s.members.filter(m => m.email !== e);
        if (s.members.length === before) throw makeError('not_found', '这个邮箱已经不在成员名单里了');
        save(s);
        return { email: e };
      },
    };
    return api;
  }

  // ============================================================
  // 对外入口
  // ============================================================
  let adapter = null;

  function wantDemo(cfg) {
    let search = '';
    try { search = (root.location && root.location.search) || ''; } catch (_) { search = ''; }
    if (/[?&]demo=1(?:&|$)/.test(search)) return true;
    if (!cfg || cfg.demo === true) return true;
    const url = String(cfg.SUPABASE_URL || '').trim();
    return !url || /YOUR-PROJECT/i.test(url) || !String(cfg.SUPABASE_ANON_KEY || '').trim();
  }

  function ensure() {
    if (!adapter) CampDB.init();
    return adapter;
  }

  const CampDB = {
    mode: null,

    init(config) {
      if (adapter) return { mode: CampDB.mode };
      const cfg = config || root.CAMP_REVIEW_CONFIG || {};
      if (wantDemo(cfg)) {
        adapter = createDemoAdapter();
        CampDB.mode = 'demo';
      } else {
        adapter = createSupabaseAdapter(cfg);
        CampDB.mode = 'supabase';
      }
      return { mode: CampDB.mode };
    },

    getSession: () => ensure().getSession(),
    onAuthChange: cb => ensure().onAuthChange(cb),
    signIn: (email, password) => ensure().signIn(email, password),
    signUp: (email, password, display_name, department) => ensure().signUp(email, password, display_name, department),
    signOut: () => ensure().signOut(),
    checkAccess: () => ensure().checkAccess(),
    listCohorts: () => ensure().listCohorts(),
    getCohort: id => ensure().getCohort(id),
    createCohort: cohort => ensure().createCohort(cohort),
    updateCohort: cohort => ensure().updateCohort(cohort),
    deleteCohort: id => ensure().deleteCohort(id),
    listHistory: cohortId => ensure().listHistory(cohortId),
    listMembers: () => ensure().listMembers(),
    addMember: email => ensure().addMember(email),
    removeMember: email => ensure().removeMember(email),
  };

  root.CampDB = CampDB;
  if (typeof module === 'object' && module && module.exports) module.exports = CampDB;
})(typeof window !== 'undefined' ? window : globalThis);
