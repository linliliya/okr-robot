// 关键词检索工具：不使用 embedding，也不改写资料或话术原文。
(function (root) {
  'use strict';
  const MAX_DOCUMENT_CHARS = 200000;
  const CHUNK_SIZE = 900, CHUNK_OVERLAP = 100, MAX_HITS = 6, MAX_CONTEXT_CHARS = 6000;
  const chars = value => Array.from(String(value == null ? '' : value));
  const length = value => chars(value).length;

  function chunkText(text) {
    const input = chars(text);
    if (input.length > MAX_DOCUMENT_CHARS) throw new RangeError('每份资料最多 200000 字，请拆成多份保存');
    const chunks = [];
    for (let start = 0; start < input.length;) {
      let end = Math.min(start + CHUNK_SIZE, input.length);
      if (end < input.length) {
        const newline = input.slice(start + 600, end).lastIndexOf('\n');
        if (newline >= 0) end = start + 600 + newline + 1;
      }
      const content = input.slice(start, end).join('');
      if (content.trim()) chunks.push(content);
      if (end === input.length) break;
      start = end - CHUNK_OVERLAP;
    }
    return chunks;
  }

  function queryTerms(text) {
    const tokens = String(text || '').toLowerCase().slice(0, 4000).match(/[\u3400-\u9fff]+|[a-z0-9_]+/g) || [];
    const terms = new Set();
    const add = term => { if (term && terms.size < 64) terms.add(term); };
    for (const token of tokens) {
      if (/^[a-z0-9_]+$/.test(token)) { if (token.length >= 2) add(token.slice(0, 64)); continue; }
      if (token.length <= 16) add(token);
      for (let i = 0; i + 1 < token.length; i++) add(token.slice(i, i + 2));
    }
    return [...terms];
  }

  function matchScore(title, content, terms) {
    const heading = String(title || '').toLowerCase(), body = String(content || '').toLowerCase();
    return terms.reduce((score, term) => {
      const weight = Math.min(length(term), 8);
      return score + (heading.includes(term) ? weight * 2 : 0) + (body.includes(term) ? weight : 0);
    }, 0);
  }
  const boundedLimit = value => Math.max(1, Math.min(MAX_HITS, Math.floor(Number(value) || MAX_HITS)));
  const contextCost = c => length(c.content) + length(c.title) + length(c.project) + length(c.source_note) + 100;

  function rankChunks(chunks, query, opts) {
    const o = opts || {}, project = String(o.project || '').trim(), channel = o.channel || 'all';
    const terms = queryTerms(query);
    if (!terms.length) return [];
    const scored = (chunks || []).filter(c => c && (!c.status || c.status === 'active')
      && (!c.project || c.project === project) && (!c.channel || c.channel === 'all' || c.channel === channel))
      .map((chunk, i) => ({ chunk, i, key: chunk.document_id == null ? 'chunk-' + i : String(chunk.document_id),
        bodyScore: matchScore('', chunk.content, terms), titleScore: matchScore(chunk.title, '', terms) }));
    const bodyDocuments = new Set(scored.filter(x => x.bodyScore > 0).map(x => x.key));
    const firstChunks = new Map();
    for (const item of scored) {
      const first = firstChunks.get(item.key);
      if (!first || Number(item.chunk.ordinal || 0) < Number(first.chunk.ordinal || 0)) firstChunks.set(item.key, item);
    }
    const ranked = scored.filter(x => x.bodyScore > 0 || (x.titleScore > 0 && !bodyDocuments.has(x.key) && firstChunks.get(x.key) === x))
      .sort((a, b) => b.bodyScore - a.bodyScore || b.titleScore - a.titleScore || Number(b.chunk.document_id || 0) - Number(a.chunk.document_id || 0)
        || Number(a.chunk.ordinal || 0) - Number(b.chunk.ordinal || 0) || a.i - b.i);
    const result = [];
    let used = 0;
    for (const { chunk } of ranked) {
      used += contextCost(chunk);
      if (used > MAX_CONTEXT_CHARS) break;
      result.push(chunk);
      if (result.length === boundedLimit(o.limit)) break;
    }
    return result;
  }

  function selectStyleScripts(scripts, query, opts) {
    const o = opts || {}, channel = o.channel || 'all';
    const limit = boundedLimit(o.limit), maxChars = Math.max(0, Number(o.maxChars == null ? 7000 : o.maxChars) || 0);
    const terms = queryTerms(query);
    const ranked = (scripts || []).filter(s => s && s.status !== 'archived' &&
      (channel === 'all' || !s.channel || s.channel === 'all' || s.channel === channel))
      .map((s, i) => ({ s, i, score: matchScore(s.scene, [s.stage, s.customer_says, s.reply, s.tips].join('\n'), terms) }))
      .sort((a, b) => b.score - a.score || a.i - b.i);
    const selected = [];
    let used = 0;
    for (const { s } of ranked) {
      const messageLabels = String(s.reply || '').split(/\r\n|\r|\n/).filter(line => line.trim())
        .reduce((n, _line, i) => n + String(i + 1).length + 5, 0);
      const customerSeparators = (String(s.customer_says || '').match(/\n+/g) || []).length * 3;
      const cost = [s.stage, s.scene, s.customer_says, s.reply, s.tips].reduce((n, value) => n + length(value), 0)
        + messageLabels + customerSeparators + 100;
      if (used + cost > maxChars) continue;
      selected.push(s); used += cost;
      if (selected.length === limit) break;
    }
    return selected;
  }

  // 直接使用原话需要明确且唯一的匹配，不能靠回复正文里的泛词推断客户顾虑。
  function normalizeConcern(text) {
    return String(text || '').normalize('NFKC').toLowerCase()
      .replace(/(?:没有|没|缺乏)(?:任何|一点|什么)?(?:编程|程序设计|计算机|代码|python)?(?:的)?基础|(?:一点|任何)?(?:编程)?基础(?:都|也)?没有|从(?:来)?没(?:有)?(?:学过|接触过)(?:编程|代码|程序设计|python)|(?:编程)?(?:小白|新手)|0基础/g, '零基础')
      .replace(/(?:担心|害怕|怕)(?:自己|我)?(?:会)?(?:学不会|学不懂|跟不上|听不懂)|能不能(?:学会|跟上|听懂)|(?:能|可以)(?:学会|跟上|听懂)(?:吗|么)/g, '怕学不会');
  }
  const concernGroups = [
    ['zero', 'learning', /零基础/g],
    ['learn', 'learning', /怕学不会|学不会|学不懂|跟不上|听不懂|(?:担心|怕)(?:课程)?太难/g],
    ['expensive', 'price', /(?:价格|学费)?(?:太贵|有点贵|比较贵|贵了|嫌贵)|价格(?:太)?高|超(?:出)?预算|预算(?:不够|不足)|买不起/g],
    ['price', 'price', /多少钱|价格多少|怎么收费|如何收费|学费多少/g],
    ['time', 'time', /没有时间(?:学|学习)?|没时间(?:学|学习)?|时间不够|没空(?:学|学习)?|抽不出时间|(?:工作|最近)?太忙|工作忙/g],
    ['refund', 'refund', /退款|退费/g],
    ['trust', 'trust', /(?:怕|担心)被骗|靠谱吗|靠谱不靠谱|信不过/g],
    ['consider', 'consider', /再考虑(?:一下)?|考虑一下|再想想/g],
    ['start_date', 'schedule', /开营日期|开营时间|开课时间|开课日期|什么时候(?:开课|开营)|几号(?:开课|开营)|哪天(?:开课|开营)/g],
    ['class_time', 'schedule', /上课时间|上课时段|什么时候上课|几点(?:上课|开课)/g],
    ['replay_retention', 'duration', /回放(?:保留|能看|可以看)?多久|回放(?:保留期|有效期)/g],
    ['validity', 'duration', /(?:课程|资料)?有效期|(?:课程|资料)(?:有效|能看|可以看)多久/g],
    ['course_duration', 'duration', /(?:课程)?(?:要)?学(?:习)?多久|课程(?:一共)?(?:多长时间|多久)|学习多长时间/g],
    ['content', 'content', /课程内容|学什么|教什么/g],
    ['job', 'outcome', /找(?:到)?工作|就业/g],
    ['freelance', 'outcome', /接单/g],
    ['effect_guarantee', 'outcome', /保证效果|保证结果/g],
    ['learning_guarantee', 'outcome', /保证学会|包学会/g],
  ];
  const trimConcern = text => text.replace(/[\s\p{P}\p{S}]/gu, '')
    .replace(/请问|怎么办|怎么|如何|可以吗|这个|这门|课程|编程|自己|完全|一点|有点|主要|就是|其实|现在|目前|真的|觉得|感觉|可能|非常|比较|学习|担心|害怕|怕|您好|你好|老师|我是|我|也|又|还|很|会|的|了|呢|啊|呀|吗|想|问/g, '');
  function concernInfo(text) {
    const normalized = normalizeConcern(text), features = [], domains = new Set();
    // 否定与反向基础说明宁可交给生成，不能发送相反含义的原话。
    const reversed = /(?:不是|并非|不算)(?:完全)?零基础|(?<!没)有(?:一些|一点|一定|扎实)?(?:编程|代码|计算机|python)?基础|(?:具备|已有)(?:编程)?基础/.test(normalized);
    const negated = /(?:不|没)(?:再|太|怎么|那么)?(?:担心|害怕|怕)|(?:不是|并非|不)(?:因为|说)?(?:零基础|学不会|学不懂|跟不上|听不懂|没时间|太贵|嫌贵)|不(?:太|算|是)?贵/.test(normalized);
    let rest = normalized;
    for (const [feature, domain, pattern] of concernGroups) {
      pattern.lastIndex = 0;
      if (pattern.test(rest)) { features.push(feature); domains.add(domain); }
      pattern.lastIndex = 0;
      rest = rest.replace(pattern, '');
    }
    return { key: features.sort().join('|'), domains, remainder: trimConcern(rest), normalized: trimConcern(normalized), reversed, negated };
  }
  function matchScript(scripts, message, opts) {
    const query = String(message || '').trim(), channel = opts && opts.channel || 'all';
    const none = reason => ({ type: 'none', script: null, candidates: [], reason });
    if (!query || length(query) > 400) return none('客户消息不足或较长，需要结合聊天判断。');
    const q = concernInfo(query);
    if (q.reversed || q.negated) return none('包含否定或反向表达，需要核对实际顾虑。');
    if (q.domains.size > 1) return none('包含多种顾虑，需要结合当前聊天调整。');
    if (/\d|[零一二两三四五六七八九十百千万]+(?:个)?(?:元|块|天|周|月|年|小时|分钟|点|号)|(?:星期|周)[一二三四五六日天]/.test(q.remainder)) {
      return none('包含具体金额或时间条件，需要结合资料核实。');
    }
    const matches = (scripts || []).filter(s => s && s.status !== 'archived' && String(s.reply || '').trim()
      && (channel === 'all' || !s.channel || s.channel === 'all' || s.channel === channel))
      .filter(s => [s.scene].concat(Array.isArray(s.customer_says) ? s.customer_says : String(s.customer_says || '').split(/\r?\n/)).some(text => {
        if (!String(text || '').trim()) return false;
        const candidate = concernInfo(text);
        if (candidate.negated || candidate.reversed || candidate.domains.size > 1) return false;
        const sameWords = q.normalized.length >= 2 && q.normalized === candidate.normalized;
        const sameConcern = q.key && q.key === candidate.key && !q.remainder && !candidate.remainder;
        return sameWords || sameConcern;
      }));
    const unique = matches.filter((s, i) => matches.findIndex(x => Number(x.id) === Number(s.id)) === i);
    const candidates = unique.map(s => ({ id: s.id, scene: s.scene, version: s.version }));
    if (unique.length > 1) return { type: 'ambiguous', script: null, candidates, reason: '多条话术同样符合，需结合当前聊天选择。' };
    if (!unique.length) return none('未找到唯一且明确对应的既定话术。');
    return { type: 'exact', script: unique[0], candidates, reason: '客户说法与这条话术的场景或典型说法明确对应。' };
  }

  const api = { MAX_DOCUMENT_CHARS, CHUNK_SIZE, CHUNK_OVERLAP, MAX_HITS, MAX_CONTEXT_CHARS,
    chunkText, queryTerms, rankChunks, selectStyleScripts, matchScript };
  root.SalesKnowledge = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
