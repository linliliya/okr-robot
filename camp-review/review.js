/* 复盘闭环：行动快照、目标校验与数据对照。纯函数，不读写页面或数据库。 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./logic.js'));
  else root.CampReview = factory(root.CampLogic);
})(typeof window !== 'undefined' ? window : globalThis, function (L) {
  'use strict';

  const METRICS = Object.freeze([
    { key: 'liveDeals', label: '直播成交人数', unit: '人' },
    { key: 'directDeals', label: '1v1 追单成交人数', unit: '人' },
    { key: 'totalDeals', label: '总成交人数', unit: '人' },
    { key: 'attendRate', label: '直播到课率', unit: '%' },
    { key: 'liveConvRate', label: '直播转化率', unit: '%' },
    { key: 'communityConvRate', label: '社群转化率', unit: '%' },
    { key: 'inviteRate', label: '邀约入群率', unit: '%' },
    { key: 'balanceRate', label: '尾款补齐率', unit: '%' },
  ].map(Object.freeze));
  const metric = key => METRICS.find(m => m.key === key);
  const text = value => value == null ? '' : String(value).trim();
  const object = value => value && typeof value === 'object' && !Array.isArray(value) ? value : {};
  const finite = value => typeof value === 'number' && Number.isFinite(value);
  const blank = value => value == null || (typeof value === 'string' && value.trim() === '');
  function number(value) {
    if (blank(value) || (typeof value !== 'number' && typeof value !== 'string')) return null;
    const n = typeof value === 'number' ? value : L.parseNum(value);
    return finite(n) ? n : null;
  }
  function targetValue(value) {
    if (blank(value)) return null;
    if (typeof value === 'number') return value;
    if (typeof value === 'string') {
      const s = value.trim();
      return /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(s) ? Number(s) : s;
    }
    // 保留错误的形状为非数字，交给校验提示，避免 Number(null/false) 变成 0。
    return typeof value === 'boolean' ? value : String(value);
  }
  function ratio(a, b) {
    return finite(a) && finite(b) && b > 0 && finite(a / b) ? a / b : null;
  }
  function completeSum(rows, keys) {
    if (!Array.isArray(rows) || !rows.length) return null;
    const sums = Object.fromEntries(keys.map(key => [key, 0]));
    for (const raw of rows) {
      const row = object(raw);
      for (const key of keys) {
        const n = number(row[key]);
        if (n == null) return null;
        sums[key] += n;
        if (!finite(sums[key])) return null;
      }
    }
    return sums;
  }

  function metricValue(cohort, key) {
    const c = object(cohort);
    if (!metric(key) || !['xiaoe', 'bilibili'].includes(c.channel)) return null;
    const data = object(c.data), rows = data.conversions;
    const liveKey = c.channel === 'bilibili' ? 'live_full' : 'balance';
    if (key === 'liveDeals' || key === 'directDeals') {
      const col = key === 'liveDeals' ? liveKey : 'direct';
      const sums = completeSum(rows, [col]);
      return sums ? sums[col] : null;
    }
    if (key === 'totalDeals') {
      const sums = completeSum(rows, c.channel === 'bilibili'
        ? ['live_full', 'direct', 'alumni', 'refund'] : ['balance', 'direct', 'alumni']);
      if (!sums) return null;
      const total = sums[liveKey] + sums.direct + sums.alumni - (c.channel === 'bilibili' ? sums.refund : 0);
      return finite(total) ? total : null;
    }
    if (key === 'balanceRate') {
      if (c.channel !== 'xiaoe') return null;
      const sums = completeSum(rows, ['balance', 'deposit', 'deposit_refund']);
      return sums ? ratio(sums.balance, sums.deposit - sums.deposit_refund) : null;
    }
    if (key === 'inviteRate') {
      const invites = Array.isArray(data.invites) ? data.invites : [];
      let reach = 0, joined = 0, known = false;
      for (const raw of invites) {
        const row = object(raw);
        if (blank(row.reach) || /^[\/\-—–_~]+$/.test(text(row.reach))) continue;
        const r = number(row.reach), j = number(row.joined);
        if (r == null || j == null) return null;
        known = true; reach += r; joined += j;
      }
      return known ? ratio(joined, reach) : null;
    }
    let m;
    try { m = L.computeMetrics(c); } catch (_) { return null; }
    if (key === 'attendRate') return ratio(m.attendance, m.members);
    if (key === 'liveConvRate') return ratio(metricValue(c, 'liveDeals'), m.attendance);
    if (key === 'communityConvRate') return ratio(metricValue(c, 'totalDeals'), m.members);
    return null;
  }

  function normalizeAction(raw, index) {
    const a = object(raw);
    return {
      id: text(a.id) || 'action-' + (index + 1), title: text(a.title),
      hypothesis: text(a.hypothesis), owner: text(a.owner), metric: text(a.metric),
      target: targetValue(a.target),
    };
  }
  function normalizeReview(raw) {
    const r = object(raw);
    let followup = null;
    if (r.followup && typeof r.followup === 'object' && !Array.isArray(r.followup)) {
      const f = r.followup;
      followup = {
        sourceId: typeof f.sourceId === 'number' && finite(f.sourceId) ? f.sourceId : text(f.sourceId),
        sourceName: text(f.sourceName), sourceDate: text(f.sourceDate), sourceChannel: text(f.sourceChannel),
        items: (Array.isArray(f.items) ? f.items : []).map((rawItem, i) => {
          const item = object(rawItem);
          return {
            action: normalizeAction(item.action, i), baseline: number(item.baseline),
            execution: text(item.execution) || 'pending', decision: text(item.decision) || 'pending', note: text(item.note),
          };
        }),
      };
    }
    return { version: 1, lessons: text(r.lessons), actions: (Array.isArray(r.actions) ? r.actions : []).map(normalizeAction), followup };
  }

  function earlierCohorts(c, all) {
    if (!c || !L.isISODate(c.live_date)) return [];
    return (Array.isArray(all) ? all : []).filter(x => x && x.channel === c.channel &&
      L.isISODate(x.live_date) && x.live_date < c.live_date).slice().sort((a, b) => {
      const dates = b.live_date.localeCompare(a.live_date);
      if (dates) return dates;
      const ai = Number(a.id), bi = Number(b.id);
      if (!blank(a.id) && !blank(b.id) && finite(ai) && finite(bi) && ai !== bi) return bi - ai;
      return text(b.id).localeCompare(text(a.id));
    });
  }
  function previousCohort(c, all) { return earlierCohorts(c, all)[0] || null; }
  function reviewContext(c, all) {
    const review = normalizeReview(object(c && c.data).review);
    if (review.followup) return review;
    const previous = previousCohort(c, all);
    if (!previous) return review;
    const actions = normalizeReview(object(previous.data).review).actions;
    if (!actions.length) return review;
    review.followup = {
      sourceId: previous.id, sourceName: text(previous.name), sourceDate: previous.live_date, sourceChannel: previous.channel,
      items: actions.map(action => ({ action, baseline: metricValue(previous, action.metric), execution: 'pending', decision: 'pending', note: '' })),
    };
    return review;
  }

  function validateReview(raw) {
    const r = normalizeReview(raw), errors = [];
    function length(value, max, label) { if (value.length > max) errors.push(label + '不能超过 ' + max + ' 字'); }
    function action(a, label) {
      if (!a.title) errors.push(label + '请填写行动内容');
      length(a.title, 160, label + '行动内容'); length(a.hypothesis, 1200, label + '预期作用'); length(a.owner, 80, label + '负责人');
      const def = metric(a.metric);
      if (!def) errors.push(label + '请选择有效指标');
      if (!finite(a.target) || a.target < 0) errors.push(label + '请填写非负有限数字目标');
      else if (def && def.unit === '%' && a.target > 1) errors.push(label + '百分比目标应在 0%–100% 之间');
      else if (def && def.unit === '人' && !Number.isInteger(a.target)) errors.push(label + '人数目标必须是整数');
    }
    length(r.lessons, 8000, '本期经验');
    if (r.actions.length > 50) errors.push('每期最多记录 50 项行动');
    const ids = new Set();
    r.actions.forEach((a, i) => {
      action(a, '第 ' + (i + 1) + ' 项：');
      if (ids.has(a.id)) errors.push('第 ' + (i + 1) + ' 项行动编号重复');
      ids.add(a.id);
    });
    if (r.followup) r.followup.items.forEach((item, i) => {
      const label = '上期第 ' + (i + 1) + ' 项：';
      action(item.action, label);
      if (!['pending', 'done', 'partial', 'skipped'].includes(item.execution)) errors.push(label + '执行状态无效');
      if (!['pending', 'keep', 'adjust', 'stop', 'retry'].includes(item.decision)) errors.push(label + '后续决策无效');
      if (item.decision !== 'pending') {
        if (item.execution === 'pending') errors.push(label + '请先填写执行情况再作结论');
        if (!item.note) errors.push(label + '作出结论时请填写观察说明');
      }
      if (item.decision === 'keep' && ['pending', 'skipped'].includes(item.execution)) errors.push(label + '未执行的行动不能决定继续沿用');
      length(item.note, 4000, label + '观察说明');
    });
    return errors;
  }

  function formatMetric(key, value) {
    if (!finite(value) || !metric(key)) return '—';
    return metric(key).unit === '%' ? (value * 100).toFixed(1) + '%' : value.toLocaleString('en-US', { maximumFractionDigits: 3 }) + ' 人';
  }
  function evaluateAction(c, rawItem) {
    const item = object(rawItem), a = normalizeAction(item.action, 0);
    const def = metric(a.metric);
    const targetValid = finite(a.target) && a.target >= 0 && def && (def.unit === '%' ? a.target <= 1 : Number.isInteger(a.target));
    const baseline = number(item.baseline), current = metricValue(c, a.metric), target = targetValid ? a.target : null;
    const difference = baseline != null && current != null ? current - baseline : null;
    const delta = finite(difference) ? difference : null;
    const execution = text(item.execution) || 'pending';
    const executed = execution === 'done' || execution === 'partial';
    let status, met = executed && current != null && target != null ? current >= target : executed ? null : false;
    const source = object(object(c && c.data).review).followup;
    const mismatch = source && source.sourceChannel && source.sourceChannel !== c.channel;
    let detail;
    if (mismatch) { status = 'missing'; met = null; detail = '历史来源渠道与本期不一致，暂不比较'; }
    else if (!executed) { status = 'pending'; detail = execution === 'skipped' ? '本期未执行，不判定达标' : '待填写执行情况，不判定达标'; }
    else if (baseline == null || current == null || target == null) { status = 'missing'; met = null; detail = '基线、本期数据或目标尚不完整，暂不能完成对照'; }
    else if (!c.closed) { status = 'provisional'; detail = '本期尚未收口，当前' + (met ? '达到' : '未达到') + '目标，之后继续核对'; }
    else { status = 'ready'; detail = '本期已收口，' + (met ? '达到' : '未达到') + '目标'; }
    const parts = ['上期 ' + formatMetric(a.metric, baseline) + ' → 本期 ' + formatMetric(a.metric, current), '目标 ≥ ' + formatMetric(a.metric, target)];
    if (delta != null && !mismatch) {
      const value = metric(a.metric) && metric(a.metric).unit === '%' ? (Math.abs(delta) * 100).toFixed(1) + ' 个百分点' : formatMetric(a.metric, Math.abs(delta));
      parts.push('变化 ' + (delta > 0 ? '+' : delta < 0 ? '−' : '') + value);
    }
    return { baseline, current, target, delta: mismatch ? null : delta, met, status, text: parts.join('；') + '。' + detail + '。这些数据用于对照，不能据此判断行动造成了变化。' };
  }
  function lastLessons(c, all, limit = 3) {
    const count = finite(limit) ? Math.max(0, Math.floor(limit)) : 3;
    return earlierCohorts(c, all).map(p => ({ id: p.id, name: text(p.name), live_date: p.live_date, lessons: normalizeReview(object(p.data).review).lessons }))
      .filter(p => p.lessons).slice(0, count);
  }

  return { METRICS, metricValue, normalizeReview, previousCohort, reviewContext, validateReview, formatMetric, evaluateAction, lastLessons };
});
