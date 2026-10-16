// 纯函数：日期、脱敏、聊天解析、分组、统计、文本对比。不碰 DOM、不碰网络。
(function (root) {
  'use strict';

  const CHANNELS = {
    all: { key: 'all', label: '通用', color: '#6B7280' },
    xiaoe: { key: 'xiaoe', label: '小鹅通', color: '#2F6BFF' },
    bilibili: { key: 'bilibili', label: 'B站', color: '#FB7299' },
    other: { key: 'other', label: '其他', color: '#6B7280' },
  };
  const SCRIPT_CHANNELS = ['all', 'xiaoe', 'bilibili'];
  const CASE_CHANNELS = ['xiaoe', 'bilibili', 'other'];
  const DEFAULT_STAGES = ['邀约入群', '社群预热', '直播中', '直播后追单', '定金转尾款', '售后与退款', '其他'];
  const RESULTS = { won: '成交', lost: '流失' };
  const ROLES = { customer: '客户', sales: '销售', note: '备注' };
  const UNSORTED_SCENE = '未分类';

  // ---------- 日期 ----------
  function pad(n) { return String(n).padStart(2, '0'); }
  function parseISO(s) {
    const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(s || ''));
    return m ? new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])) : null;
  }
  function toISO(d) { return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; }
  function todayISO() {
    const d = new Date();
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  }
  function addDays(s, n) {
    const d = parseISO(s);
    if (!d) return '';
    d.setUTCDate(d.getUTCDate() + n);
    return toISO(d);
  }
  // 所在周的周一
  function weekStart(s) {
    const d = parseISO(s);
    if (!d) return '';
    const back = (d.getUTCDay() + 6) % 7;
    d.setUTCDate(d.getUTCDate() - back);
    return toISO(d);
  }
  function fmtDateCN(s) {
    const d = parseISO(s);
    return d ? `${d.getUTCMonth() + 1}月${d.getUTCDate()}日` : '';
  }
  function weekLabel(week) {
    return week ? `${fmtDateCN(week)}–${fmtDateCN(addDays(week, 6))}` : '';
  }
  function fmtDateTime(iso) {
    if (!iso) return '';
    const d = new Date(iso);
    if (isNaN(d)) return '';
    return `${d.getMonth() + 1}月${d.getDate()}日 ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  }

  // ---------- 脱敏：手机号、邮箱、身份证、微信号 ----------
  function maskSensitive(text) {
    return String(text == null ? '' : text)
      .replace(/(?<![\dA-Za-z])\d{17}[\dXx](?![\dA-Za-z])/g, '[身份证号]')
      .replace(/(?<!\d)(?:\+?86[-\s]?)?1[3-9]\d[-\s]?\d{4}[-\s]?\d{4}(?!\d)/g, '[手机号]')
      .replace(/[\w.+-]+@[\w-]+(?:\.[\w-]+)+/g, '[邮箱]')
      .replace(/wxid_[A-Za-z0-9_-]+/g, '[微信号]')
      .replace(/((?:微信号?|vx|VX|Vx|wx|WX|v信|V信|威信)\s*[:：]?\s*)[A-Za-z][-_A-Za-z0-9]{5,19}/g, '$1[微信号]');
  }

  // ---------- 聊天记录 ----------
  const CUSTOMER_PREFIX = '客户|用户|学员|顾客|对方|家长';
  const SALES_PREFIX = '销售|我|顾问|老师|助教|班主任|课程顾问';
  const LINE_RE = new RegExp(`^\\s*(${CUSTOMER_PREFIX}|${SALES_PREFIX}|备注)\\s*[：:]\\s*(.*)$`);
  const CUSTOMER_RE = new RegExp(`^(${CUSTOMER_PREFIX})$`);

  // 按「客户：/销售：」前缀识别；一行都认不出来时返回 null（交给 AI 整理）
  function parseDialogue(text) {
    const out = [];
    let hit = 0;
    for (const raw of String(text || '').replace(/\r\n?/g, '\n').split('\n')) {
      const line = raw.trim();
      if (!line) continue;
      const m = LINE_RE.exec(line);
      if (m) {
        hit++;
        const role = m[1] === '备注' ? 'note' : CUSTOMER_RE.test(m[1]) ? 'customer' : 'sales';
        out.push({ role, text: m[2].trim() });
      } else if (out.length) {
        out[out.length - 1].text += (out[out.length - 1].text ? '\n' : '') + line;
      } else {
        out.push({ role: 'note', text: line });
      }
    }
    return hit ? out.filter(x => x.text) : null;
  }

  function normDialogue(list) {
    if (!Array.isArray(list)) return [];
    return list
      .map(x => ({
        role: ROLES[x && x.role] ? x.role : 'note',
        text: String((x && x.text) || '').trim(),
      }))
      .filter(x => x.text);
  }

  function dialogueToText(list) {
    return normDialogue(list).map(x => `${ROLES[x.role]}：${x.text}`).join('\n');
  }

  // ---------- 话术库 ----------
  // 显式换行表示按顺序发送的独立消息；只忽略空白行，保留其余原文。
  function splitReplyMessages(text) {
    return String(text == null ? '' : text).split(/\r\n|\r|\n/).filter(line => line.trim());
  }

  function emptyScript() {
    return {
      id: null, stage: '', scene: '', customer_says: '', reply: '', tips: '',
      channel: 'all', status: 'active', sort_order: 0,
      change_note: '', change_review_id: null, change_case_ids: [], version: 1,
    };
  }

  function validateScript(s) {
    const errs = [];
    if (!String(s.scene || '').trim()) errs.push('请填写场景（客户的问题或异议，如「嫌贵」）');
    if (!String(s.reply || '').trim()) errs.push('请填写推荐回复');
    if (!SCRIPT_CHANNELS.includes(s.channel)) errs.push('适用渠道不对');
    return errs;
  }

  // 阶段排序：常用阶段按固定顺序，其余按名字
  function sortStages(stages) {
    const uniq = [...new Set(stages.map(s => s || '其他'))];
    return uniq.sort((a, b) => {
      const ia = DEFAULT_STAGES.indexOf(a), ib = DEFAULT_STAGES.indexOf(b);
      if (ia >= 0 && ib >= 0) return ia - ib;
      if (ia >= 0) return a === '其他' ? 1 : -1;
      if (ib >= 0) return b === '其他' ? -1 : 1;
      return a.localeCompare(b, 'zh');
    });
  }

  function groupScriptsByStage(scripts) {
    const map = new Map();
    for (const s of scripts) {
      const k = s.stage || '其他';
      if (!map.has(k)) map.set(k, []);
      map.get(k).push(s);
    }
    return sortStages([...map.keys()]).map(stage => ({
      stage,
      scripts: map.get(stage).sort((a, b) => (a.sort_order - b.sort_order) || String(a.scene).localeCompare(String(b.scene), 'zh') || (a.id - b.id)),
    }));
  }

  // 关键词搜索（空格分隔，全部命中才算）
  function searchScripts(scripts, q) {
    const words = String(q || '').toLowerCase().split(/\s+/).filter(Boolean);
    if (!words.length) return scripts;
    return scripts.filter(s => {
      const hay = [s.stage, s.scene, s.customer_says, s.reply, s.tips].join('\n').toLowerCase();
      return words.every(w => hay.includes(w));
    });
  }

  function scriptChannelMatch(s, channel) {
    return !channel || channel === 'all' || s.channel === 'all' || s.channel === channel;
  }

  // ---------- 案例 ----------
  function emptyCase(seller) {
    const today = todayISO();
    return {
      id: null, case_date: today, review_week: weekStart(today), seller: seller || '',
      channel: 'xiaoe', cohort: '', result: 'lost', profile: '', dialogue: [],
      stuck_point: '', self_review: '', scenes: [], script_refs: [], ai_analysis: null,
      discussion: '', review_status: 'pending', version: 1,
    };
  }

  function validateCase(c) {
    const errs = [];
    if (!RESULTS[c.result]) errs.push('请选择成交还是流失');
    if (!parseISO(c.case_date)) errs.push('请填写日期');
    if (!String(c.seller || '').trim()) errs.push('请填写销售');
    if (!normDialogue(c.dialogue).length && !String(c.stuck_point || '').trim()) {
      errs.push('至少填一下聊天记录或者「关键节点」');
    }
    return errs;
  }

  // 复盘会按场景分组：每个案例只放在它的第一个场景下，成交和流失并排
  function groupCasesByScene(cases) {
    const map = new Map();
    for (const c of cases) {
      const k = (c.scenes && c.scenes[0]) || UNSORTED_SCENE;
      if (!map.has(k)) map.set(k, { scene: k, won: [], lost: [] });
      map.get(k)[c.result === 'won' ? 'won' : 'lost'].push(c);
    }
    return [...map.values()].sort((a, b) => {
      if (a.scene === UNSORTED_SCENE) return 1;
      if (b.scene === UNSORTED_SCENE) return -1;
      // 成交和流失都有的场景最值得对比，排前面
      const both = x => (x.won.length && x.lost.length ? 1 : 0);
      return both(b) - both(a) || (b.won.length + b.lost.length) - (a.won.length + a.lost.length);
    });
  }

  // 每条话术（及每个版本）在案例里被用到的次数和结果
  function scriptUsage(cases) {
    const map = new Map();
    for (const c of cases) {
      const seen = new Set();
      for (const r of Array.isArray(c.script_refs) ? c.script_refs : []) {
        const id = Number(r && r.id);
        if (!id || seen.has(id)) continue;
        seen.add(id);
        if (!map.has(id)) map.set(id, { won: 0, lost: 0, byVersion: {} });
        const u = map.get(id);
        const key = c.result === 'won' ? 'won' : 'lost';
        u[key]++;
        const v = Number(r.version) || 0;
        if (!u.byVersion[v]) u.byVersion[v] = { won: 0, lost: 0 };
        u.byVersion[v][key]++;
      }
    }
    return map;
  }

  function usageText(u) {
    if (!u) return '';
    const n = u.won + u.lost;
    return n ? `案例里用到 ${n} 次 · 成交 ${u.won}` : '';
  }

  // ---------- 复盘待办 ----------
  function normTodos(list, createId) {
    return (Array.isArray(list) ? list : [])
      .filter(t => t && String(t.text || '').trim())
      .map(t => ({
        id: typeof t.id === 'string' && t.id ? t.id : createId(),
        text: String(t.text).trim(), owner: String(t.owner || '').trim(), done: !!t.done,
      }));
  }

  // 记录用户看到的任务身份；旧数据没有 ID 时，必须同时锁定当时的复盘版本。
  function todoTarget(review, index) {
    const t = (review.todos || [])[index];
    if (!t) return null;
    return { id: t.id || null, index, version: review.version, text: t.text, owner: t.owner, done: t.done };
  }

  function setTodoDone(review, target, done) {
    const todos = review.todos || [];
    let index = -1;
    if (target && target.id) {
      const matches = todos.map((t, i) => t.id === target.id ? i : -1).filter(i => i >= 0);
      if (matches.length === 1) index = matches[0];
    } else if (target && review.version === target.version) {
      const t = todos[target.index];
      if (t && !t.id && t.text === target.text && t.owner === target.owner && t.done === target.done) index = target.index;
    }
    if (index < 0) {
      const err = new Error('待办已被其他人更新，请确认最新列表后再操作');
      err.code = 'stale_todo';
      throw err;
    }
    return todos.map((t, i) => i === index ? Object.assign({}, t, { done: !!done }) : t);
  }

  // ---------- 文本对比（按句子） ----------
  function splitSentences(text) {
    return String(text || '').match(/[^。！？!?；;\n]+[。！？!?；;]*|\n/g) || [];
  }

  function diffText(a, b) {
    const x = splitSentences(a), y = splitSentences(b);
    const n = x.length, m = y.length;
    const dp = Array.from({ length: n + 1 }, () => new Array(m + 1).fill(0));
    for (let i = n - 1; i >= 0; i--) {
      for (let j = m - 1; j >= 0; j--) {
        dp[i][j] = x[i] === y[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
      }
    }
    const out = [];
    let i = 0, j = 0;
    const push = (type, text) => {
      const last = out[out.length - 1];
      if (last && last.type === type) last.text += text; else out.push({ type, text });
    };
    while (i < n && j < m) {
      if (x[i] === y[j]) { push('same', x[i]); i++; j++; }
      else if (dp[i + 1][j] >= dp[i][j + 1]) { push('del', x[i]); i++; }
      else { push('add', y[j]); j++; }
    }
    while (i < n) push('del', x[i++]);
    while (j < m) push('add', y[j++]);
    return out;
  }

  // ---------- 给 AI 的文本 ----------
  function scriptToPrompt(s) {
    const lines = [`[#${s.id}] 阶段：${s.stage || '其他'}｜场景：${s.scene}｜适用：${CHANNELS[s.channel] ? CHANNELS[s.channel].label : '通用'}`];
    if (s.customer_says) lines.push(`客户常说：${s.customer_says.replace(/\n+/g, ' / ')}`);
    lines.push('推荐回复（以下消息按顺序发送）：');
    splitReplyMessages(s.reply).forEach((message, i) => lines.push(`消息 ${i + 1}：${message}`));
    if (s.tips) lines.push(`要点/禁忌：${s.tips}`);
    return lines.join('\n');
  }

  function scriptsToPrompt(scripts) {
    const active = scripts.filter(s => s.status !== 'archived');
    return active.length ? active.map(scriptToPrompt).join('\n\n') : '（话术库还是空的）';
  }

  function caseToPrompt(c) {
    const head = [
      `[案例#${c.id || '新'}] ${RESULTS[c.result] || ''}｜${CHANNELS[c.channel] ? CHANNELS[c.channel].label : ''}${c.cohort ? '｜' + c.cohort : ''}｜销售：${c.seller || ''}`,
    ];
    if (c.profile) head.push(`客户画像：${c.profile}`);
    if (c.scenes && c.scenes.length) head.push(`涉及场景：${c.scenes.join('、')}`);
    if (c.stuck_point) head.push(`关键节点：${c.stuck_point}`);
    if (c.self_review) head.push(`销售自评：${c.self_review}`);
    const d = dialogueToText(c.dialogue);
    if (d) head.push('聊天记录：\n' + d);
    return head.join('\n');
  }

  // 从模型输出里取 JSON（兼容包在 ``` 里的情况）
  function extractJSON(text) {
    const s = String(text || '').replace(/^\s*```(?:json)?\s*/i, '').replace(/\s*```\s*$/, '');
    try { return JSON.parse(s); } catch (e) { /* 再试一次 */ }
    const a = s.indexOf('{'), b = s.lastIndexOf('}');
    if (a >= 0 && b > a) {
      try { return JSON.parse(s.slice(a, b + 1)); } catch (e) { /* 放弃 */ }
    }
    return null;
  }

  // 长文档切块（按空行分段，尽量不切断一段话）
  function splitDoc(text, maxLen) {
    const limit = maxLen || 3000;
    const paras = String(text || '').replace(/\r\n?/g, '\n').split(/\n\s*\n/).map(p => p.trim()).filter(Boolean);
    const chunks = [];
    let cur = '';
    const flush = () => { if (cur) chunks.push(cur); cur = ''; };
    for (const p of paras) {
      if (p.length > limit) {
        flush();
        let buf = '';
        for (let line of p.split('\n')) {
          if (buf && buf.length + line.length + 1 > limit) { chunks.push(buf); buf = ''; }
          // 单行就超长（没有换行的大段文字），只能硬切
          while (line.length > limit) { chunks.push(line.slice(0, limit)); line = line.slice(limit); }
          buf += (buf ? '\n' : '') + line;
        }
        if (buf) chunks.push(buf);
        continue;
      }
      if (cur && cur.length + p.length + 2 > limit) flush();
      cur += (cur ? '\n\n' : '') + p;
    }
    flush();
    return chunks;
  }

  // 把 AI 拆出来的条目整理成话术对象
  function normImportedScript(x) {
    const s = emptyScript();
    s.stage = String((x && x.stage) || '').trim();
    s.scene = String((x && x.scene) || '').trim();
    s.customer_says = Array.isArray(x && x.customer_says) ? x.customer_says.join('\n') : String((x && x.customer_says) || '').trim();
    s.reply = Array.isArray(x && x.reply) ? x.reply.join('\n') : String((x && x.reply) || '');
    s.tips = Array.isArray(x && x.tips) ? x.tips.join('\n') : String((x && x.tips) || '').trim();
    s.channel = SCRIPT_CHANNELS.includes(x && x.channel) ? x.channel : 'all';
    return s;
  }

  root.SalesLogic = {
    CHANNELS, SCRIPT_CHANNELS, CASE_CHANNELS, DEFAULT_STAGES, RESULTS, ROLES, UNSORTED_SCENE,
    todayISO, addDays, weekStart, weekLabel, fmtDateCN, fmtDateTime,
    maskSensitive, parseDialogue, normDialogue, dialogueToText,
    splitReplyMessages, emptyScript, validateScript, sortStages, groupScriptsByStage, searchScripts, scriptChannelMatch,
    emptyCase, validateCase, groupCasesByScene, scriptUsage, usageText,
    normTodos, todoTarget, setTodoDone,
    diffText, scriptToPrompt, scriptsToPrompt, caseToPrompt, extractJSON, splitDoc, normImportedScript,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = root.SalesLogic;
})(typeof window !== 'undefined' ? window : globalThis);
