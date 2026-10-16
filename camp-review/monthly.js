/* 月度复盘按期次直播月份汇总，该期跨月补款仍回到直播月。 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./logic.js'), require('./review.js'));
  else root.CampMonthly = factory(root.CampLogic, root.CampReview);
})(typeof window !== 'undefined' ? window : globalThis, function (L, R) {
  'use strict';

  const CHANNELS = ['xiaoe', 'bilibili'];
  const METRICS = ['totalDeals', 'liveDeals', 'directDeals', 'alumniDeals', 'refunds'];
  const LABELS = { balance: '已付尾款', live_full: '直播全款', direct: '1v1 追单', alumni: '老学员', refund: '退款' };
  const object = value => value && typeof value === 'object' && !Array.isArray(value) ? value : {};
  const pad2 = value => String(value).padStart(2, '0');

  function validMonth(value) {
    return typeof value === 'string' && /^\d{4}-(0[1-9]|1[0-2])$/.test(value) && +value.slice(0, 4) > 0;
  }

  function shiftMonth(value, delta) {
    if (!validMonth(value) || !Number.isInteger(delta)) return '';
    const n = +value.slice(0, 4) * 12 + +value.slice(5, 7) - 1 + delta;
    const year = Math.floor(n / 12), month = n - year * 12 + 1;
    return year >= 1 && year <= 9999 ? String(year).padStart(4, '0') + '-' + pad2(month) : '';
  }

  function monthLabel(value) {
    return validMonth(value) ? +value.slice(0, 4) + '年' + +value.slice(5, 7) + '月' : '';
  }

  function today(at) {
    const date = at == null ? new Date() : new Date(at);
    if (!Number.isFinite(date.getTime())) return '';
    const parts = new Intl.DateTimeFormat('en-CA', {
      timeZone: 'Asia/Shanghai', year: 'numeric', month: '2-digit', day: '2-digit',
    }).formatToParts(date);
    const values = Object.fromEntries(parts.map(part => [part.type, part.value]));
    return values.year + '-' + values.month + '-' + values.day;
  }

  function timestamp(year, month, day) {
    const d = new Date(0);
    d.setUTCFullYear(year, month - 1, day);
    d.setUTCHours(0, 0, 0, 0);
    return d.getTime();
  }

  function validDate(value) {
    if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(value) || !validMonth(value.slice(0, 7))) return false;
    const d = new Date(timestamp(+value.slice(0, 4), +value.slice(5, 7), +value.slice(8, 10)));
    return d.getUTCDate() === +value.slice(8, 10) && d.getUTCMonth() + 1 === +value.slice(5, 7);
  }

  function number(value) {
    if (value == null || (typeof value !== 'number' && typeof value !== 'string') || String(value).trim() === '') return null;
    const n = typeof value === 'number' ? value : L.parseNum(value);
    return typeof n === 'number' && Number.isFinite(n) ? n : null;
  }

  function rowValues(cohort, raw) {
    const row = object(raw);
    const part = { channel: cohort.channel, data: { conversions: [row] } };
    return {
      totalDeals: R.metricValue(part, 'totalDeals'),
      liveDeals: R.metricValue(part, 'liveDeals'),
      directDeals: R.metricValue(part, 'directDeals'),
      alumniDeals: number(row.alumni),
      refunds: cohort.channel === 'bilibili' ? number(row.refund) : 0,
    };
  }

  function emptyMonth(key) {
    return {
      key,
      metrics: Object.fromEntries(METRICS.map(name => [name, { value: null, knownValue: 0, missingRows: 0 }])),
      cohorts: [], rows: [], cohortCount: 0, openCount: 0,
    };
  }

  function buildMonthly(cohorts, options) {
    const opts = object(options);
    const now = validDate(opts.today) ? opts.today : today();
    const selected = validMonth(opts.month) ? opts.month : now.slice(0, 7);
    const channel = CHANNELS.includes(opts.channel) ? opts.channel : 'all';
    const months = Array.from({ length: 12 }, (_, i) => shiftMonth(selected, i - 11)).filter(Boolean).map(emptyMonth);
    const buckets = new Map(months.map(month => [month.key, month]));
    const previousKey = shiftMonth(selected, -1);
    const issues = [];
    const ids = new Map(months.map(month => [month.key, { all: new Set(), open: new Set() }]));
    const seenCohorts = new Set();

    for (const raw of Array.isArray(cohorts) ? cohorts : []) {
      const cohort = object(raw);
      if (!CHANNELS.includes(cohort.channel) || (channel !== 'all' && cohort.channel !== channel)) continue;
      // 已保存期次按 ID 去重，无 ID 的临时记录仅去重同一个对象。
      // 不按直播日期或成交行内容猜测重复，保留同日多场及多行成交。
      const identity = cohort.id == null ? cohort : String(cohort.id);
      if (seenCohorts.has(identity)) continue;
      seenCohorts.add(identity);
      if (!validDate(cohort.live_date)) {
        issues.push({ code: 'invalid_date', cohort, rowIndex: null, date: cohort.live_date == null ? '' : String(cohort.live_date),
          message: '直播日期缺失或无效，请补全年月日后再统计' });
        continue;
      }
      const month = buckets.get(cohort.live_date.slice(0, 7));
      if (month) {
        const key = cohort.id == null ? cohort : String(cohort.id);
        const group = ids.get(month.key);
        if (!group.all.has(key)) month.cohorts.push(cohort);
        group.all.add(key);
        if (!cohort.closed) group.open.add(key);
      }
      const data = object(cohort.data);
      const rows = Array.isArray(data.conversions) ? data.conversions : [];
      if (!rows.length) {
        issues.push({ code: 'empty_conversions', cohort, rowIndex: null, date: '', message: '这期尚未录入每日成交，月合计仍有缺项' });
        // 空表表示尚未录入，不能使同月其他期次的小计变成完整月合计。
        if (month) METRICS.forEach(name => month.metrics[name].missingRows++);
        continue;
      }
      rows.forEach((rawRow, rowIndex) => {
        const row = object(rawRow), values = rowValues(cohort, row);
        const required = cohort.channel === 'bilibili' ? ['live_full', 'direct', 'alumni', 'refund'] : ['balance', 'direct', 'alumni'];
        const fields = required.filter(key => number(row[key]) == null);
        const date = row.date == null ? '' : String(row.date);
        if (fields.length) {
          issues.push({ code: 'missing_values', cohort, rowIndex, date, fields,
            message: '成交人数尚未填完整，请核对' + fields.map(key => LABELS[key]).join('、') + '，没有发生的项目请填 0' });
        }
        if (!month) return;
        month.rows.push({ cohort, rowIndex, date, values });
        METRICS.forEach(name => {
          const metric = month.metrics[name];
          if (values[name] == null) metric.missingRows++;
          else metric.knownValue += values[name];
        });
      });
    }

    months.forEach(month => {
      month.cohortCount = ids.get(month.key).all.size;
      month.openCount = ids.get(month.key).open.size;
      month.cohorts.sort((a, b) => a.live_date.localeCompare(b.live_date));
      month.rows.sort((a, b) => a.cohort.live_date.localeCompare(b.cohort.live_date) || a.date.localeCompare(b.date));
      METRICS.forEach(name => {
        const metric = month.metrics[name];
        metric.value = month.rows.length && !metric.missingRows && Number.isFinite(metric.knownValue) ? metric.knownValue : null;
        if (!Number.isFinite(metric.knownValue)) metric.knownValue = null;
      });
    });
    return {
      month: buckets.get(selected),
      previous: buckets.get(previousKey) || emptyMonth(previousKey),
      months,
      issues,
    };
  }

  return { validMonth, shiftMonth, monthLabel, today, buildMonthly };
});
