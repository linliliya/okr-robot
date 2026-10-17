// 数据层：Supabase 适配器 + 演示模式（localStorage）适配器，对外行为一致
// 错误统一是 Error，带 err.code：conflict | tables_missing | knowledge_missing | ai_log_delete_missing | not_member | forbidden | auth | network | config | invalid | not_found | demo | unknown
(function (root) {
  'use strict';

  const DEMO_KEY = 'sales-script-demo-v1';
  const EMAIL_RE = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

  const SCRIPT_FIELDS = ['stage', 'scene', 'customer_says', 'reply', 'tips', 'channel', 'status', 'sort_order',
    'change_note', 'change_review_id', 'change_case_ids'];
  const CASE_FIELDS = ['case_date', 'review_week', 'seller', 'channel', 'cohort', 'result', 'profile', 'dialogue',
    'stuck_point', 'self_review', 'scenes', 'script_refs', 'ai_analysis', 'discussion', 'review_status'];
  const REVIEW_FIELDS = ['week', 'summary', 'todos', 'ai_brief'];
  const KNOWLEDGE_META = 'id,title,project,kind,channel,status,source_note,version,updated_at,char_count,chunk_count';

  const clone = x => (x === undefined ? undefined : JSON.parse(JSON.stringify(x)));
  const nowISO = () => new Date().toISOString();
  const normEmail = s => String(s == null ? '' : s).replace(/[\s　]+/g, '').toLowerCase();
  const pick = (obj, fields) => { const o = {}; for (const f of fields) if (obj[f] !== undefined) o[f] = clone(obj[f]); return o; };
  function aiLogId(value) {
    if ((typeof value !== 'number' && typeof value !== 'string') || !/^\d+$/.test(String(value))
      || !Number.isSafeInteger(Number(value)) || Number(value) <= 0) throw makeError('invalid', '提问记录编号无效，请刷新列表后重试');
    return Number(value);
  }

  function makeError(code, message, cause) {
    const err = new Error(message);
    err.code = code;
    if (cause !== undefined) err.cause = cause;
    Object.defineProperty(err, '__sales', { value: true });
    return err;
  }

  // 保存前统一整理
  function scriptPayload(s) {
    const p = pick(s, SCRIPT_FIELDS);
    p.scene = String(p.scene || '').trim();
    if (!p.scene) throw makeError('invalid', '请填写场景');
    p.stage = String(p.stage || '').trim();
    p.change_case_ids = (p.change_case_ids || []).map(Number).filter(Boolean);
    p.change_review_id = p.change_review_id ? Number(p.change_review_id) : null;
    p.sort_order = Number(p.sort_order) || 0;
    return p;
  }
  function casePayload(c) {
    const p = pick(c, CASE_FIELDS);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(p.case_date || '')) throw makeError('invalid', '请填写日期');
    p.scenes = (p.scenes || []).map(s => String(s).trim()).filter(Boolean);
    p.script_refs = (p.script_refs || []).filter(r => r && r.id);
    return p;
  }
  function knowledgeTools() {
    if (!root.SalesKnowledge) throw makeError('config', '知识库组件加载失败，请刷新页面后重试');
    return root.SalesKnowledge;
  }
  function knowledgePayload(doc) {
    const d = doc || {};
    const p = { title: String(d.title || '').trim(), project: String(d.project || '').trim(),
      kind: d.kind || 'other', channel: d.channel || 'all', content: String(d.content || ''),
      status: d.status || 'active', source_note: String(d.source_note || '').trim() };
    if (!p.title || !p.content.trim()) throw makeError('invalid', '请填写资料标题和正文');
    if (Array.from(p.content).length > 200000) throw makeError('invalid', '每份资料最多 200000 字，请拆成多份保存');
    if (Array.from(p.title).length > 240 || Array.from(p.project).length > 120 || Array.from(p.source_note).length > 1000) {
      throw makeError('invalid', '标题最多 240 字、项目名最多 120 字、来源说明最多 1000 字');
    }
    if (!['detail', 'live', 'faq', 'other'].includes(p.kind) || !['all', 'xiaoe', 'bilibili'].includes(p.channel)
      || !['active', 'archived'].includes(p.status)) throw makeError('invalid', '资料类型、渠道或状态不正确');
    if (d.id != null && d.id !== '') {
      if (!Number.isSafeInteger(Number(d.id)) || Number(d.id) <= 0 || !Number.isSafeInteger(Number(d.version)) || Number(d.version) <= 0) {
        throw makeError('invalid', '资料版本无效，请重新打开后保存');
      }
      p.id = Number(d.id); p.version = Number(d.version);
    }
    return p;
  }
  function knowledgeOptions(opts) {
    const o = opts || {};
    return { project: String(o.project || '').trim(), channel: ['all', 'xiaoe', 'bilibili'].includes(o.channel) ? o.channel : 'all',
      limit: Math.max(1, Math.min(6, Math.floor(Number(o.limit) || 6))) };
  }

  // ---------- Supabase ----------
  function supabaseAdapter(config) {
    const sb = root.InternalAccess?.client || root.supabase.createClient(config.SUPABASE_URL, config.SUPABASE_ANON_KEY);
    let profileNames = null;
    let currentEmail = '';

    function classify(e, fallback) {
      if (e && e.__sales) return e;
      const code = String((e && e.code) || '');
      const msg = String((e && e.message) || e || '');
      if (['42P01', '42883', 'PGRST202', 'PGRST205'].includes(code) || /does not exist|Could not find the (table|function)/i.test(msg)) {
        return makeError('tables_missing', '数据库还没初始化，请按 sales-script/SETUP.md 执行 setup.sql', e);
      }
      if (code === '42501' || /row-level security/i.test(msg)) {
        return makeError('not_member', `你的账号${currentEmail ? '（' + currentEmail + '）' : ''}还没开通，请联系管理员调整账号权限`, e);
      }
      if (/JWT|token|PGRST301|PGRST303/i.test(code + msg)) return makeError('auth', '登录已过期，请重新登录', e);
      if (/Failed to fetch|NetworkError|Load failed|fetch failed/i.test(msg)) return makeError('network', '网络连接失败，请检查网络后重试', e);
      return makeError('unknown', (fallback || '操作失败') + '：' + msg, e);
    }
    async function run(promise, fallback) {
      let res;
      try { res = await promise; } catch (e) { throw classify(e, fallback); }
      if (res.error) throw classify(res.error, fallback);
      return res.data;
    }
    async function knowledgeRun(promise, fallback) {
      try { return await run(promise, fallback); }
      catch (e) {
        if (e.code === 'tables_missing') throw makeError('knowledge_missing', '知识库还没初始化，请在 Supabase SQL Editor 执行 sales-script/knowledge-setup.sql', e);
        const cause = e.cause || {};
        if (cause.message === 'knowledge_conflict') throw makeError('conflict', '这份资料刚被别人改过，请重新打开后再保存', e);
        if (cause.message === 'knowledge_not_found') throw makeError('not_found', '找不到这份资料', e);
        if (['22023', '22P02', '23514'].includes(cause.code)) throw makeError('invalid', '资料内容或版本不正确，请检查后重新保存', e);
        throw e;
      }
    }

    async function names() {
      if (profileNames) return profileNames;
      const rows = await run(sb.from('profiles').select('id, display_name'), '读取成员姓名失败');
      profileNames = {};
      for (const r of rows || []) profileNames[r.id] = r.display_name || '';
      return profileNames;
    }
    function withNames(row, n) {
      if (!row) return row;
      return Object.assign({}, row, {
        created_by_name: n[row.created_by] || '',
        updated_by_name: n[row.updated_by] || '',
        changed_by_name: n[row.changed_by] || '',
      });
    }

    // Supabase 一次最多返回 1000 行，分页拉全
    async function fetchAll(build, fallback) {
      const PAGE = 1000;
      const out = [];
      for (let from = 0; ; from += PAGE) {
        const rows = await run(build().range(from, from + PAGE - 1), fallback);
        out.push(...(rows || []));
        if (!rows || rows.length < PAGE) return out;
      }
    }

    async function save(table, fields, obj, payload) {
      const n = await names();
      if (obj.id) {
        const rows = await run(sb.from(table).update(payload).eq('id', obj.id).eq('version', obj.version).select(), '保存失败');
        if (!rows || !rows.length) throw makeError('conflict', '刚被别人改过');
        return withNames(rows[0], n);
      }
      const rows = await run(sb.from(table).insert(payload).select(), '保存失败');
      return withNames(rows[0], n);
    }

    return {
      mode: 'supabase',
      async getSession() {
        const { data } = await sb.auth.getSession();
        const user = data && data.session && data.session.user;
        if (!user) return null;
        currentEmail = user.email || '';
        const { data: prof } = await sb.from('profiles').select('display_name, department').eq('id', user.id).maybeSingle();
        return { user: { id: user.id, email: user.email }, profile: prof || { display_name: user.email, department: '' } };
      },
      onAuthChange(cb) { sb.auth.onAuthStateChange((event) => cb(event)); },
      async signIn(email, password) {
        const { error } = await sb.auth.signInWithPassword({ email: normEmail(email), password });
        if (error) throw makeError('auth', /Invalid login/i.test(error.message) ? '邮箱或密码不对' : error.message);
      },
      async signUp(email, password, display_name, department) {
        if (!String(display_name || '').trim()) throw makeError('auth', '请填写姓名');
        if (!EMAIL_RE.test(normEmail(email))) throw makeError('auth', '邮箱格式不对');
        if (String(password || '').length < 6) throw makeError('auth', '密码至少 6 位');
        const { data, error } = await sb.auth.signUp({
          email: normEmail(email), password,
          options: { data: { display_name: String(display_name).trim(), department } },
        });
        if (error) throw makeError('auth', /registered/i.test(error.message) ? '这个邮箱已经注册过了，直接登录就行' : error.message);
        if (!data.session) throw makeError('auth', '注册成功，请去邮箱点验证链接后再登录');
      },
      async signOut() { await sb.auth.signOut(); profileNames = null; },
      async checkAccess() {
        try {
          const ok = await run(sb.rpc('is_sales_member'), '检查权限失败');
          if (ok) return { ok: true };
          const e = classify({ code: '42501' });
          return { ok: false, reason: 'not_member', message: e.message };
        } catch (e) {
          return { ok: false, reason: e.code === 'tables_missing' ? 'tables_missing' : 'error', message: e.message };
        }
      },

      async listScripts() {
        const n = await names();
        const rows = await fetchAll(() => sb.from('sales_scripts').select('*').order('stage').order('sort_order').order('id'), '读取话术库失败');
        return rows.map(r => withNames(r, n));
      },
      async saveScript(s) { return save('sales_scripts', SCRIPT_FIELDS, s, scriptPayload(s)); },
      async listKnowledgeDocuments() {
        const out = [], page = 1000;
        for (let from = 0; ; from += page) {
          const rows = await knowledgeRun(sb.from('sales_knowledge_documents').select(KNOWLEDGE_META)
            .order('updated_at', { ascending: false }).order('id', { ascending: false }).range(from, from + page - 1), '读取知识库失败');
          out.push(...(rows || []));
          if (!rows || rows.length < page) return out;
        }
      },
      async getKnowledgeDocument(id) {
        const row = await knowledgeRun(sb.from('sales_knowledge_documents').select('*').eq('id', id).maybeSingle(), '读取资料失败');
        if (!row) throw makeError('not_found', '找不到这份资料');
        return row;
      },
      async saveKnowledgeDocument(doc) {
        const data = await knowledgeRun(sb.rpc('save_sales_knowledge', { p_document: knowledgePayload(doc) }), '保存资料失败');
        return Array.isArray(data) ? data[0] : data;
      },
      async searchKnowledge(query, opts) {
        const o = knowledgeOptions(opts);
        return (await knowledgeRun(sb.rpc('search_sales_knowledge', { p_terms: knowledgeTools().queryTerms(query),
          p_project: o.project, p_channel: o.channel, p_limit: o.limit }), '检索知识库失败')) || [];
      },
      async listScriptVersions(scriptId) {
        const n = await names();
        const rows = await run(sb.from('sales_script_versions').select('*').eq('script_id', scriptId).order('version', { ascending: false }).limit(100), '读取历史版本失败');
        return rows.map(r => withNames(r, n));
      },
      async listVersionsByReview(reviewId) {
        const n = await names();
        const rows = await run(sb.from('sales_script_versions').select('*').eq('review_id', reviewId).order('changed_at', { ascending: false }), '读取改动记录失败');
        return rows.map(r => withNames(r, n));
      },

      async listCases() {
        const n = await names();
        const rows = await fetchAll(() => sb.from('sales_cases').select('*').order('case_date', { ascending: false }).order('id', { ascending: false }), '读取案例失败');
        return rows.map(r => withNames(r, n));
      },
      async saveCase(c) { return save('sales_cases', CASE_FIELDS, c, casePayload(c)); },
      async deleteCase(id) { await run(sb.from('sales_cases').delete().eq('id', id), '删除失败'); },

      async listReviews() {
        const n = await names();
        const rows = await fetchAll(() => sb.from('sales_reviews').select('*').order('week', { ascending: false }), '读取复盘会失败');
        return rows.map(r => withNames(r, n));
      },
      async saveReview(r) {
        try {
          return await save('sales_reviews', REVIEW_FIELDS, r, pick(r, REVIEW_FIELDS));
        } catch (e) {
          // 两个人同时第一次保存同一周：唯一约束冲突，当作版本冲突处理
          if (e.cause && e.cause.code === '23505') throw makeError('conflict', '这周的复盘刚被别人创建了');
          throw e;
        }
      },

      // opts: { since, until, feedback, createdBy, limit, all }
      // all: true 用于复盘会，分页读取全部匹配记录；普通列表仍默认最多 200 条。
      async listAiLogs(opts) {
        const o = opts || {};
        const n = await names();
        const build = () => {
          let q = sb.from('sales_ai_logs').select('*').order('created_at', { ascending: false }).order('id', { ascending: false });
          if (o.since) q = q.gte('created_at', o.since);
          if (o.until) q = q.lt('created_at', o.until);
          if (o.feedback) q = q.eq('feedback', o.feedback);
          if (o.createdBy) q = q.eq('created_by', o.createdBy);
          return q;
        };
        const rows = o.all === true
          ? await fetchAll(build, '读取提问记录失败')
          : await run(build().limit(o.limit || 200), '读取提问记录失败');
        return rows.map(r => withNames(r, n));
      },
      async createAiLog(log) {
        const rows = await run(sb.from('sales_ai_logs').insert(pick(log, ['message', 'context', 'channel', 'result'])).select(), '记录失败');
        return rows[0];
      },
      async setAiFeedback(id, feedback, note) {
        await run(sb.from('sales_ai_logs').update({ feedback, feedback_note: note || '' }).eq('id', id), '保存反馈失败');
      },
      async deleteAiLog(id) {
        const target = aiLogId(id);
        let deleted;
        try { deleted = await run(sb.rpc('delete_sales_ai_log', { p_id: target }), '删除提问记录失败'); }
        catch (e) {
          const cause = e.cause || {};
          if (e.code === 'tables_missing') throw makeError('ai_log_delete_missing', '提问记录删除功能尚未配置，请在 Supabase SQL Editor 执行 sales-script/ai-log-delete-setup.sql', e);
          if (cause.message === 'ai_log_auth_required') throw makeError('auth', '登录已过期，请重新登录', e);
          if (cause.message === 'ai_log_not_found') throw makeError('not_found', '这条提问记录已不存在，请刷新列表', e);
          if (cause.message === 'ai_log_not_owner') throw makeError('forbidden', '只能删除自己创建的提问记录', e);
          if (cause.message === 'ai_log_delete_denied') throw makeError('ai_log_delete_missing', '记录未删除，删除权限尚未就绪，请执行 sales-script/ai-log-delete-setup.sql 后重试', e);
          if (cause.code === '42501' && cause.message !== 'sales_members_only') throw makeError('forbidden', '没有权限删除这条提问记录，请确认账号权限和删除功能配置', e);
          if (['22023', '22P02'].includes(cause.code)) throw makeError('invalid', '提问记录编号无效，请刷新列表后重试', e);
          throw e;
        }
        if ((typeof deleted !== 'number' && typeof deleted !== 'string') || Number(deleted) !== target) {
          throw makeError('unknown', '服务器未确认删除成功，请刷新列表核实后重试');
        }
        return { id: target };
      },

      async getSetting(key) {
        const rows = await run(sb.from('sales_settings').select('value').eq('key', key), '读取设置失败');
        return rows && rows[0] ? rows[0].value : '';
      },
      async setSetting(key, value) {
        await run(sb.from('sales_settings').upsert({ key, value: String(value || ''), updated_at: nowISO() }), '保存设置失败');
      },

      async listMembers() {
        const n = await names();
        const rows = await run(sb.from('sales_members').select('*').order('added_at'), '读取成员失败');
        return rows.map(r => ({ email: r.email, added_at: r.added_at, added_by_name: n[r.added_by] || '' }));
      },
      async addMember(email) {
        const e = normEmail(email);
        if (!EMAIL_RE.test(e)) throw makeError('invalid', '邮箱格式不对');
        try {
          await run(sb.rpc('add_sales_member', { p_email: e }), '添加失败');
        } catch (err) {
          // 数据库函数里 raise 的中文提示原样给用户看
          if (err.cause && err.cause.code === 'P0001') throw makeError('invalid', err.cause.message);
          throw err;
        }
      },
      async removeMember(email) { await run(sb.from('sales_members').delete().eq('email', normEmail(email)), '移除失败'); },

      // 调 AI 中转函数
      async invokeAIStream(body, opts) {
        const o = opts || {};
        if (!root.SalesStream) throw makeError('config', '流式组件未加载，请刷新页面');
        if (o.signal && o.signal.aborted) throw root.SalesStream.abortError();
        const session = sb.auth.getSession();
        const { data, error: sessionError } = await (o.signal ? new Promise((resolve, reject) => {
          const cleanup = () => o.signal.removeEventListener('abort', stop);
          const stop = () => { cleanup(); reject(root.SalesStream.abortError()); };
          o.signal.addEventListener('abort', stop, { once: true });
          Promise.resolve(session).then(value => { cleanup(); resolve(value); }, error => { cleanup(); reject(error); });
          if (o.signal.aborted) stop();
        }) : session);
        if (sessionError || !data || !data.session || !data.session.access_token) throw makeError('auth', '登录已过期，请重新登录');
        if (o.signal && o.signal.aborted) throw root.SalesStream.abortError();
        const headers = {
          'Content-Type': 'application/json', 'Accept': 'text/event-stream',
          Authorization: 'Bearer ' + data.session.access_token, apikey: config.SUPABASE_ANON_KEY,
        };
        if (config.AI_REGION) headers['x-region'] = config.AI_REGION;
        let response;
        try {
          response = await fetch(config.SUPABASE_URL.replace(/\/$/, '') + '/functions/v1/' + encodeURIComponent(config.AI_FUNCTION || 'sales-ai'), {
            method: 'POST', headers, body: JSON.stringify({ ...body, stream: true }), signal: o.signal,
          });
        } catch (e) {
          if (o.signal && o.signal.aborted) throw root.SalesStream.abortError();
          throw makeError('network', '无法连接 AI 中转服务，请检查网络后重试。暂时无法判断 ChatGPT 或 DeepSeek 的状态。', e);
        }
        return root.SalesStream.readResponse(response, o);
      },
      async invokeAI(body) {
        const headers = config.AI_REGION ? { 'x-region': config.AI_REGION } : undefined;
        let res;
        try {
          res = await sb.functions.invoke(config.AI_FUNCTION || 'sales-ai', { body, headers });
        } catch (e) {
          throw makeError('network', '无法连接 AI 中转服务，请检查网络后重试。暂时无法判断 ChatGPT 或 DeepSeek 的状态。', e);
        }
        if (res.error) {
          let msg = res.error.message || '';
          const ctx = res.error.context;
          let payload = null;
          if (ctx && typeof ctx.json === 'function') {
            try { payload = await ctx.json(); msg = (payload && (payload.error || payload.msg || payload.message)) || msg; } catch (_) { /* 用默认文案 */ }
            if (ctx.status === 404) msg = 'AI 中转函数还没部署，请按 sales-script/SETUP.md 第 4 步操作';
          }
          const network = /Failed to send a request|Failed to fetch|NetworkError|Load failed/i.test(msg);
          if (network) msg = '无法连接 AI 中转服务，请检查网络或函数部署。暂时无法判断 ChatGPT 或 DeepSeek 的状态。';
          const err = makeError(network ? 'network' : ctx && ctx.status === 401 ? 'auth' : 'ai', msg || 'AI 调用失败');
          if (payload && payload.ai_status) err.ai_status = payload.ai_status;
          throw err;
        }
        return res.data;
      },
    };
  }

  // ---------- 演示模式（数据只在本机浏览器） ----------
  function demoAdapter() {
    const USER = { id: 'demo-user', email: 'demo@example.com' };
    const NAME = '演示用户';
    const empty = () => ({ scripts: [], versions: [], cases: [], reviews: [], logs: [], knowledgeDocuments: [], knowledgeChunks: [], settings: {}, seq: 1 });
    let mem = null;

    function load() {
      if (mem) return mem;
      try { mem = JSON.parse(localStorage.getItem(DEMO_KEY) || 'null'); } catch (e) { mem = null; }
      if (!mem || typeof mem !== 'object') mem = empty();
      for (const k of ['scripts', 'versions', 'cases', 'reviews', 'logs', 'knowledgeDocuments', 'knowledgeChunks']) if (!Array.isArray(mem[k])) mem[k] = [];
      if (!mem.settings || typeof mem.settings !== 'object') mem.settings = {};
      if (!mem.seq) mem.seq = 1;
      return mem;
    }
    function persist() {
      try { localStorage.setItem(DEMO_KEY, JSON.stringify(mem)); } catch (e) { /* 存不了就只留在内存里 */ }
    }
    const named = r => Object.assign({}, clone(r), { created_by_name: NAME, updated_by_name: NAME, changed_by_name: NAME });

    function save(listKey, obj, payload, onSaved) {
      const db = load();
      const list = db[listKey];
      let row;
      if (obj.id) {
        const i = list.findIndex(x => x.id === obj.id);
        if (i < 0) throw makeError('not_found', '找不到这条记录');
        if (list[i].version !== obj.version) throw makeError('conflict', '刚被别人改过');
        row = Object.assign({}, list[i], payload, { version: list[i].version + 1, updated_at: nowISO(), updated_by: USER.id });
        list[i] = row;
      } else {
        row = Object.assign({ id: db.seq++, created_at: nowISO(), created_by: USER.id }, payload, { version: 1, updated_at: nowISO(), updated_by: USER.id });
        list.push(row);
      }
      if (onSaved) onSaved(row, !obj.id);
      persist();
      return named(row);
    }

    return {
      mode: 'demo',
      async getSession() { return { user: USER, profile: { display_name: NAME, department: '' } }; },
      onAuthChange() {},
      async signIn() {}, async signUp() {}, async signOut() {},
      async checkAccess() { return { ok: true }; },

      async listScripts() { return load().scripts.map(named); },
      async listKnowledgeDocuments() {
        return load().knowledgeDocuments.slice().sort((a, b) => b.updated_at.localeCompare(a.updated_at) || b.id - a.id)
          .map(d => pick(d, KNOWLEDGE_META.split(',')));
      },
      async getKnowledgeDocument(id) {
        const row = load().knowledgeDocuments.find(d => Number(d.id) === Number(id));
        if (!row) throw makeError('not_found', '找不到这份资料');
        return clone(row);
      },
      async saveKnowledgeDocument(doc) {
        const payload = knowledgePayload(doc), parts = knowledgeTools().chunkText(payload.content);
        delete payload.id; delete payload.version;
        payload.char_count = Array.from(payload.content).length; payload.chunk_count = parts.length;
        // 在副本上组装正文与分块，浏览器存储成功后一起替换，避免容量不足时假装保存成功。
        const next = clone(load()), id = doc && doc.id ? Number(doc.id) : null;
        const index = id ? next.knowledgeDocuments.findIndex(d => d.id === id) : -1;
        if (id && index < 0) throw makeError('not_found', '找不到这份资料');
        const previous = index >= 0 ? next.knowledgeDocuments[index] : null;
        if (previous && previous.version !== Number(doc.version)) throw makeError('conflict', '这份资料刚被别人改过，请重新打开后再保存');
        const row = Object.assign({}, previous || { id: next.seq++, created_at: nowISO(), created_by: USER.id }, payload,
          { version: previous ? previous.version + 1 : 1, updated_at: nowISO(), updated_by: USER.id });
        if (index >= 0) next.knowledgeDocuments[index] = row; else next.knowledgeDocuments.push(row);
        next.knowledgeChunks = next.knowledgeChunks.filter(c => c.document_id !== row.id).concat(parts.map((content, i) => ({
          chunk_id: next.seq++, document_id: row.id, content, ordinal: i + 1, version: row.version,
        })));
        try { localStorage.setItem(DEMO_KEY, JSON.stringify(next)); }
        catch (e) { throw makeError('unknown', '浏览器没有保存成功，可能本地存储空间不足。请缩短资料或切换正式模式', e); }
        mem = next;
        return named(row);
      },
      async searchKnowledge(query, opts) {
        const db = load(), documents = new Map(db.knowledgeDocuments.map(d => [d.id, d]));
        const chunks = db.knowledgeChunks.map(c => {
          const d = documents.get(c.document_id);
          return d && Object.assign({}, c, pick(d, ['title', 'project', 'kind', 'channel', 'status', 'source_note']));
        }).filter(Boolean);
        return knowledgeTools().rankChunks(chunks, query, knowledgeOptions(opts))
          .map(c => pick(c, ['chunk_id', 'document_id', 'title', 'project', 'kind', 'content', 'ordinal', 'version', 'source_note']));
      },
      async saveScript(s) {
        return save('scripts', s, scriptPayload(s), (row, isNew) => {
          load().versions.push({
            id: load().seq++, script_id: row.id, version: row.version, action: isNew ? 'insert' : 'update',
            snapshot: clone(row), change_note: row.change_note || '', review_id: row.change_review_id || null,
            case_ids: row.change_case_ids || [], changed_by: USER.id, changed_at: nowISO(),
          });
        });
      },
      async listScriptVersions(id) {
        return load().versions.filter(v => v.script_id === id).sort((a, b) => b.version - a.version).map(named);
      },
      async listVersionsByReview(reviewId) {
        return load().versions.filter(v => v.review_id === reviewId).sort((a, b) => (a.changed_at < b.changed_at ? 1 : -1)).map(named);
      },

      async listCases() {
        return load().cases.slice().sort((a, b) => (a.case_date < b.case_date ? 1 : a.case_date > b.case_date ? -1 : b.id - a.id)).map(named);
      },
      async saveCase(c) { return save('cases', c, casePayload(c)); },
      async deleteCase(id) { const db = load(); db.cases = db.cases.filter(c => c.id !== id); persist(); },

      async listReviews() { return load().reviews.slice().sort((a, b) => (a.week < b.week ? 1 : -1)).map(named); },
      async saveReview(r) {
        if (!r.id && load().reviews.some(x => x.week === r.week)) throw makeError('conflict', '这周的复盘刚被别人创建了');
        return save('reviews', r, pick(r, REVIEW_FIELDS));
      },

      async listAiLogs(opts) {
        const o = opts || {};
        const t = x => new Date(x).getTime();
        const rows = load().logs
          .filter(l => (!o.since || t(l.created_at) >= t(o.since)) && (!o.until || t(l.created_at) < t(o.until))
            && (!o.feedback || l.feedback === o.feedback) && (!o.createdBy || l.created_by === o.createdBy))
          .sort((a, b) => (a.created_at < b.created_at ? 1 : a.created_at > b.created_at ? -1 : b.id - a.id));
        return (o.all === true ? rows : rows.slice(0, o.limit || 200)).map(named);
      },
      async createAiLog(log) {
        const db = load();
        const row = Object.assign({ id: db.seq++, feedback: '', feedback_note: '', created_by: USER.id, created_at: nowISO() },
          pick(log, ['message', 'context', 'channel', 'result']));
        db.logs.push(row);
        persist();
        return named(row);
      },
      async setAiFeedback(id, feedback, note) {
        const l = load().logs.find(x => x.id === id);
        if (l) { l.feedback = feedback; l.feedback_note = note || ''; persist(); }
      },
      async deleteAiLog(id) {
        const target = aiLogId(id), db = load();
        const row = db.logs.find(l => Number(l.id) === target);
        if (!row) throw makeError('not_found', '这条提问记录已不存在，请刷新列表');
        if (row.created_by !== USER.id) throw makeError('forbidden', '只能删除自己创建的提问记录');
        const next = clone(db);
        next.logs = next.logs.filter(l => Number(l.id) !== target);
        try { localStorage.setItem(DEMO_KEY, JSON.stringify(next)); }
        catch (e) { throw makeError('unknown', '浏览器未能保存删除操作，记录仍然保留，请检查本地存储后重试', e); }
        mem = next;
        return { id: target };
      },

      async getSetting(key) { return load().settings[key] || ''; },
      async setSetting(key, value) { load().settings[key] = String(value || ''); persist(); },

      async listMembers() { return [{ email: USER.email, added_at: nowISO(), added_by_name: '' }]; },
      async addMember() { throw makeError('invalid', '演示模式下不能添加成员'); },
      async removeMember() {},

      async invokeAI() { throw makeError('demo', '演示模式'); },

      // 演示模式专用
      resetDemo() { mem = empty(); persist(); },
    };
  }

  const SalesDB = {
    mode: null,
    init(config) {
      const demo = /[?&]demo=1(?:&|$)/.test(root.location ? root.location.search : '');
      if (!demo) {
        if (!config || !String(config.SUPABASE_URL || '').trim() || !String(config.SUPABASE_ANON_KEY || '').trim()) {
          throw makeError('config', '系统配置不完整，请检查 config.js 中的 Supabase 地址和公开密钥');
        }
        if (!root.supabase || typeof root.supabase.createClient !== 'function') {
          throw makeError('network', '数据库组件加载失败，请检查网络后刷新页面');
        }
      }
      const impl = demo ? demoAdapter() : supabaseAdapter(config);
      Object.assign(SalesDB, impl);
      SalesDB.mode = impl.mode;
      return { mode: impl.mode };
    },
    makeError,
  };
  root.SalesDB = SalesDB;
})(window);
