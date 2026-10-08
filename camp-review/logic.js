/* 训练营周复盘 · 纯逻辑模块
 * 解析粘贴、计算指标、期次对比、自动分析、校验、格式化。
 * 不碰 DOM、不碰网络。浏览器里挂到 window.CampLogic，Node 里 module.exports。 */
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module && module.exports) module.exports = api;
  if (typeof window !== 'undefined') window.CampLogic = api;
})(this, function () {
  'use strict';

  // ================= 配置 =================

  // 列定义：既能 [key, label] 解构，也能 .key / .label / .type 读取
  function col(key, label, type) {
    var c = [key, label];
    c.key = key;
    c.label = label;
    c.type = type; // 'text' | 'int'（日期列 key 固定为 'date'，type 为 'text'）
    return c;
  }

  var CHANNELS = {
    xiaoe: {
      key: 'xiaoe', label: '小鹅通', color: '#2F6BFF',
      conversionColumns: [
        col('deposit', '已付定金', 'int'), col('low_price', '已付低价档', 'int'),
        col('deposit_refund', '已退定金', 'int'), col('balance', '已付尾款', 'int'),
        col('direct', '直接付款(1v1)', 'int'), col('alumni', '老学员', 'int'),
      ],
      attendanceLabel: '直播出勤人数',
    },
    bilibili: {
      key: 'bilibili', label: 'B站', color: '#FB7299',
      conversionColumns: [
        col('live_full', '直播成交(全款)', 'int'), col('direct', '直接付款(1v1)', 'int'),
        col('refund', '退款', 'int'), col('alumni', '老学员', 'int'),
      ],
      attendanceLabel: '直播观看人数',
    },
  };
  var CHANNEL_KEYS = ['xiaoe', 'bilibili'];

  var INVITE_COLUMNS = [
    col('date', '时间', 'text'), col('slot', '具体时间段', 'text'), col('channel', '渠道', 'text'),
    col('reach', '触达人数', 'int'), col('joined', '进群人数', 'int'),
  ];

  function communityColumns(channel) {
    var ch = CHANNELS[channel] || CHANNELS.xiaoe;
    return [
      col('date', '时间', 'text'), col('group', '群', 'text'), col('members', '社群总人数', 'int'),
      col('claimed', '课程领取/预约人数', 'int'), col('checkins', '作业打卡人数', 'int'),
      col('attendance', ch.attendanceLabel, 'int'), col('add_assistant', '引导加助理号人数', 'int'),
    ];
  }

  // 既能 COMMUNITY_COLUMNS('bilibili') 按渠道取，也能直接当数组遍历（默认小鹅通文案）
  var COMMUNITY_COLUMNS = (function () {
    var base = communityColumns('xiaoe');
    var f = function (channel) { return communityColumns(channel); };
    base.forEach(function (c, i) { f[i] = c; });
    Object.defineProperty(f, 'length', { value: base.length });
    ['map', 'forEach', 'filter', 'find', 'findIndex', 'some', 'every', 'reduce', 'slice', 'indexOf', 'concat', 'join']
      .forEach(function (m) { f[m] = function () { return Array.prototype[m].apply(base, arguments); }; });
    if (typeof Symbol !== 'undefined' && Symbol.iterator) {
      f[Symbol.iterator] = function () { return base[Symbol.iterator](); };
    }
    return f;
  })();

  function conversionColumns(channel) {
    return (CHANNELS[channel] || CHANNELS.xiaoe).conversionColumns.slice();
  }

  // 编辑器用：某张表的完整列（成交表在前面补一列日期）
  function columnsFor(kind, channel) {
    if (kind === 'invites') return INVITE_COLUMNS.slice();
    if (kind === 'community') return communityColumns(channel);
    if (kind === 'conversions') return [col('date', '日期', 'text')].concat(conversionColumns(channel));
    return [];
  }

  var ROW_FIELDS = {
    invites: { text: ['date', 'slot', 'channel'], num: ['reach', 'joined'] },
    community: { text: ['date', 'group'], num: ['members', 'claimed', 'checkins', 'attendance', 'add_assistant'] },
    conversions: {
      text: ['date'],
      num: ['deposit', 'low_price', 'deposit_refund', 'balance', 'live_full', 'refund', 'direct', 'alumni'],
    },
  };
  var SUM_KEYS = ROW_FIELDS.conversions.num;
  var TABLE_NAMES = { invites: '邀约入群', community: '社群数据', conversions: '每日成交' };
  var SEG_NAMES = { invites: '邀约数据', community: '社群数据', conversions: '成交数据' };

  var CORE_METRICS = [
    {
      key: 'inviteRate', label: '邀约入群率', numKey: 'joinedKnown', denKey: 'reachKnown',
      numLabel: { xiaoe: '进群', bilibili: '进群' },
      denLabel: { xiaoe: '触达', bilibili: '触达' },
      formula: {
        xiaoe: '进群人数 ÷ 触达人数。没有触达数的渠道（比如朋友圈）不计入',
        bilibili: '进群人数 ÷ 触达人数。没有触达数的渠道（比如朋友圈）不计入',
      },
    },
    {
      key: 'attendRate', label: '直播到课率', numKey: 'attendance', denKey: 'members',
      numLabel: { xiaoe: '直播出勤', bilibili: '直播观看' },
      denLabel: { xiaoe: '直播当天社群人数', bilibili: '直播当天社群人数' },
      formula: {
        xiaoe: '直播出勤人数 ÷ 直播当天社群人数',
        bilibili: '直播观看人数 ÷ 直播当天社群人数',
      },
    },
    {
      key: 'liveConvRate', label: '直播转化率', numKey: 'liveDeals', denKey: 'attendance',
      numLabel: { xiaoe: '补齐尾款', bilibili: '直播全款成交' },
      denLabel: { xiaoe: '直播出勤', bilibili: '直播观看' },
      formula: {
        xiaoe: '补齐尾款人数 ÷ 直播出勤人数。只算直播期间付了定金、后来补齐尾款的人',
        bilibili: '直播全款成交人数 ÷ 直播观看人数',
      },
    },
    {
      key: 'communityConvRate', label: '社群转化率', numKey: 'totalDeals', denKey: 'members',
      numLabel: { xiaoe: '总成交', bilibili: '总成交' },
      denLabel: { xiaoe: '直播当天社群人数', bilibili: '直播当天社群人数' },
      formula: {
        xiaoe: '总成交人数 ÷ 直播当天社群人数。总成交 = 补齐尾款 + 1v1 追单 + 老学员',
        bilibili: '总成交人数 ÷ 直播当天社群人数。总成交 = 直播全款 + 1v1 追单 + 老学员 − 退款',
      },
    },
  ];

  // ================= 基础工具 =================

  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function isNum(x) { return typeof x === 'number' && isFinite(x); }
  function toInt(x) { var n = parseInt(x, 10); return isFinite(n) ? n : null; }

  // 全角数字/符号转半角，全角空格转空格
  function toHalfWidth(s) {
    return String(s)
      .replace(/[！-～]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
      .replace(/[　 ]/g, ' ');
  }

  // 四舍五入（处理 2.65 这类浮点误差）
  function roundTo(v, digits) {
    var f = Math.pow(10, digits);
    var s = v < 0 ? -1 : 1;
    return s * Math.round(Math.abs(v) * f + 1e-7) / f;
  }

  function ratio(a, b) {
    if (!isNum(a) || !isNum(b) || b <= 0) return null;
    return a / b;
  }

  function sumField(rows, key) {
    var s = 0;
    rows.forEach(function (r) { if (isNum(r[key])) s += r[key]; });
    return s;
  }

  function nonEmpty(cells) { return cells.filter(function (c) { return c !== ''; }); }

  function pushUnique(arr, msg) { if (arr.indexOf(msg) < 0) arr.push(msg); }

  // ================= 数字与日期 =================

  var NULL_TOKEN_RE = /^(?:[\/\-—–_~]+|空|无|暂无|null|n\/a|na)$/i;

  // 返回 { value, bad }：bad=true 表示写了东西但不是数字
  function parseNumInfo(raw) {
    if (raw == null) return { value: null, bad: false };
    if (typeof raw === 'number') return { value: isFinite(raw) ? raw : null, bad: false };
    var s = toHalfWidth(raw).replace(/[\s,]/g, '').replace(/人/g, '');
    if (s === '' || NULL_TOKEN_RE.test(s)) return { value: null, bad: false };
    if (/%$/.test(s)) return { value: null, bad: false }; // 百分比是率，不当人数
    if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(s)) return { value: Number(s), bad: false };
    return { value: null, bad: true };
  }

  function parseNum(raw) { return parseNumInfo(raw).value; }

  // 只包含数字/百分比/占位符的单元格
  function isNumberish(c) {
    var s = toHalfWidth(c).replace(/[\s,]/g, '');
    return /^[+-]?(\d+\.?\d*|\.\d+)%?$/.test(s);
  }

  function isNullToken(c) {
    var s = toHalfWidth(c).replace(/\s/g, '');
    return s === '' || NULL_TOKEN_RE.test(s);
  }

  var MONTH_DAYS = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  function daysInMonth(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }
  function validYMD(y, m, d) {
    if (!(m >= 1 && m <= 12 && d >= 1)) return false;
    return d <= (y ? daysInMonth(y, m) : MONTH_DAYS[m - 1]);
  }

  // 解析日期各部分（按字符串切，绝不 parseFloat：9.10 是 9 月 10 日）
  // 返回 { y|null, m, d } 或 null
  function parseDateParts(raw) {
    if (raw == null) return null;
    var s = toHalfWidth(raw).trim();
    if (!s) return null;
    s = s.replace(/\s+\d{1,2}:\d{2}(:\d{2})?$/, '');                         // 去掉时刻
    s = s.replace(/\s*\(?(周|星期|礼拜)[一二三四五六日天]\)?$/, '').trim();   // 去掉星期
    var m;
    function mk(y, mo, d) { return validYMD(y, mo, d) ? { y: y, m: mo, d: d } : null; }
    if ((m = s.match(/^(\d{4})\s*[-\/.年]\s*(\d{1,2})\s*[-\/.月]\s*(\d{1,2})\s*[日号]?$/))) return mk(+m[1], +m[2], +m[3]);
    if ((m = s.match(/^(\d{4})(\d{2})(\d{2})$/))) return mk(+m[1], +m[2], +m[3]);
    if ((m = s.match(/^(\d{2})(\d{2})(\d{2})$/))) {
      var yy = +m[1];
      return yy >= 20 && yy <= 40 ? mk(2000 + yy, +m[2], +m[3]) : null;
    }
    if ((m = s.match(/^(\d{1,2})\s*[-\/.月]\s*(\d{1,2})\s*[日号]?$/))) return mk(null, +m[1], +m[2]);
    return null;
  }

  function ymdToISO(y, m, d) { return y + '-' + pad2(m) + '-' + pad2(d); }

  function normalizeDate(raw, year) {
    if (raw == null) return '';
    var str = String(raw).trim();
    if (!str) return '';
    var p = parseDateParts(str);
    if (!p) return str;
    var y = p.y || toInt(year) || new Date().getFullYear();
    if (!validYMD(y, p.m, p.d)) return str;
    return ymdToISO(y, p.m, p.d);
  }

  // 邀约可以持续多天；仍存到 date 字符串，兼容已有单日记录和数据库结构。
  // 先识别单日，再尝试拆分范围，避免把 ISO 日期里的连字符当成范围。
  function parseInviteDateParts(raw) {
    var s = toHalfWidth(raw == null ? '' : raw).trim();
    var single = parseDateParts(s);
    if (single) return { start: single, end: null };
    var separators = /[-~–—至到]+/g, match;
    while ((match = separators.exec(s))) {
      var start = parseDateParts(s.slice(0, match.index).trim());
      if (!start) continue;
      var right = s.slice(match.index + match[0].length).trim();
      var end = parseDateParts(right);
      var day = right.match(/^(\d{1,2})\s*[日号]?$/);
      if (!end && day && validYMD(start.y, start.m, +day[1])) {
        end = { y: null, m: start.m, d: +day[1] };
      }
      if (end) return { start: start, end: end };
    }
    return null;
  }

  function inviteDateBounds(raw, year, liveDate) {
    var parts = parseInviteDateParts(raw);
    if (!parts) return null;
    var a = parts.start, b = parts.end;
    var baseYear = toInt(year) || new Date().getFullYear();
    var wraps = b && b.m < a.m;
    // 只允许省略年份的跨月范围跨年；同月日期倒序仍是输入错误。
    function resolve(startYear) {
      var endYear = b ? (b.y || startYear + (wraps ? 1 : 0)) : startYear;
      if (!validYMD(startYear, a.m, a.d) || (b && !validYMD(endYear, b.m, b.d))) return null;
      var start = ymdToISO(startYear, a.m, a.d);
      var end = b ? ymdToISO(endYear, b.m, b.d) : start;
      return start <= end ? { start: start, end: end, range: !!b } : null;
    }
    if (a.y) return resolve(a.y);
    if (b && b.y) return resolve(b.y - (wraps ? 1 : 0));
    if (!isISODate(liveDate)) return resolve(baseYear);
    var anchor = isoToUTC(liveDate), best = null;
    // 先按距离确定年份，再校验日期，避免为迁就非法闰日而偷偷换年。
    [baseYear - 1, baseYear, baseYear + 1].forEach(function (y) {
      var start = Date.UTC(y, a.m - 1, a.d);
      var end = b ? Date.UTC(y + (wraps ? 1 : 0), b.m - 1, b.d) : start;
      var dist = Math.max(start - anchor, anchor - end, 0);
      if (!best || dist < best.dist) best = { year: y, dist: dist };
    });
    return resolve(best ? best.year : baseYear);
  }

  function normalizeInviteDate(raw, year, liveDate) {
    var bounds = inviteDateBounds(raw, year, liveDate);
    if (!bounds) return raw == null ? '' : String(raw).trim();
    return bounds.range ? bounds.start + ' ~ ' + bounds.end : bounds.start;
  }

  function isInviteDate(raw, year, liveDate) {
    return inviteDateBounds(raw, year, liveDate) !== null;
  }

  function fmtInviteDate(raw, year, liveDate) {
    var bounds = inviteDateBounds(raw, year, liveDate);
    if (!bounds) return raw ? String(raw) : '—';
    if (!bounds.range) return fmtDateCN(bounds.start);
    var crossYear = bounds.start.slice(0, 4) !== bounds.end.slice(0, 4);
    function label(s) { return (crossYear ? s.slice(0, 4) + '年' : '') + fmtDateCN(s); }
    return label(bounds.start) + '–' + label(bounds.end);
  }

  function isISODate(s) {
    return typeof s === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(s) &&
      validYMD(+s.slice(0, 4), +s.slice(5, 7), +s.slice(8, 10));
  }
  function isoToUTC(s) { return Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)); }
  function utcToISO(t) {
    var d = new Date(t);
    return ymdToISO(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
  }

  // ================= 格式化 =================

  function fmtPct(x, digits) {
    if (digits == null) digits = 1;
    if (!isNum(x)) return '—';
    return roundTo(x * 100, digits).toFixed(digits) + '%';
  }

  function fmtInt(n) {
    if (!isNum(n)) return '—';
    var neg = n < 0;
    var a = Math.abs(n);
    var s = Number.isInteger(a) ? String(a) : String(roundTo(a, 2));
    var parts = s.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return (neg ? '-' : '') + parts.join('.');
  }

  // 变化不到 0.5 个百分点算持平（留一点浮点余量）
  function isFlat(d) { return Math.abs(d) < 0.005 - 1e-9; }
  function ppText(d) { return roundTo(Math.abs(d) * 100, 1).toFixed(1); }

  function fmtDelta(d) {
    if (!isNum(d)) return { text: '—', dir: 'flat' };
    if (isFlat(d)) return { text: '基本持平', dir: 'flat' };
    return d > 0
      ? { text: '↑ ' + ppText(d) + ' 个百分点', dir: 'up' }
      : { text: '↓ ' + ppText(d) + ' 个百分点', dir: 'down' };
  }

  function fmtDateCN(s) {
    if (!s) return '—';
    if (!isISODate(s)) return String(s);
    return (+s.slice(5, 7)) + '月' + (+s.slice(8, 10)) + '日';
  }

  function weekdayCN(s) {
    if (!isISODate(s)) return '';
    return '周' + '日一二三四五六'.charAt(new Date(isoToUTC(s)).getUTCDay());
  }

  // 'M.D'，用于在同名渠道前面区分日期
  function shortDate(s) {
    if (!isISODate(s)) return s ? String(s) : '';
    return (+s.slice(5, 7)) + '.' + (+s.slice(8, 10));
  }

  function weekStart(s) {
    if (!isISODate(s)) return '';
    var t = isoToUTC(s);
    var dow = new Date(t).getUTCDay();
    return utcToISO(t - ((dow + 6) % 7) * 86400000);
  }

  function weekLabel(ws) {
    if (!isISODate(ws)) return '未填日期';
    return fmtDateCN(ws) + '–' + fmtDateCN(utcToISO(isoToUTC(ws) + 6 * 86400000));
  }

  function channelLabel(ch) { return CHANNELS[ch] ? CHANNELS[ch].label : String(ch || ''); }

  // 周内排序：小鹅通在前，同渠道按直播日期
  function cohortOrder(a, b) {
    var ca = CHANNEL_KEYS.indexOf(a.channel), cb = CHANNEL_KEYS.indexOf(b.channel);
    if (ca !== cb) return (ca < 0 ? 99 : ca) - (cb < 0 ? 99 : cb);
    var da = a.live_date || '', db = b.live_date || '';
    return da < db ? -1 : da > db ? 1 : 0;
  }

  function groupByWeek(cohorts) {
    var map = {};
    (cohorts || []).forEach(function (c) {
      if (!c) return;
      var w = weekStart(c.live_date);
      (map[w] = map[w] || []).push(c);
    });
    return Object.keys(map)
      .sort(function (a, b) {
        if (a === '') return 1;
        if (b === '') return -1;
        return a < b ? 1 : a > b ? -1 : 0;
      })
      .map(function (w) {
        return { week: w, label: weekLabel(w), cohorts: map[w].slice().sort(cohortOrder) };
      });
  }

  // ================= 空对象 / 规范化 =================

  function emptyRow(kind) {
    var f = ROW_FIELDS[kind];
    if (!f) throw new Error('未知的表：' + kind);
    var r = {};
    f.text.forEach(function (k) { r[k] = ''; });
    f.num.forEach(function (k) { r[k] = null; });
    return r;
  }

  function emptyCohort(channel) {
    return {
      id: null,
      channel: CHANNELS[channel] ? channel : 'xiaoe',
      name: '',
      live_date: '',
      closed: false,
      summary: '',
      version: 1,
      data: { invites: [], community: [], conversions: [], unparsed: [] },
      created_at: null,
      updated_at: null,
      updated_by_name: '',
    };
  }

  function normalizeRows(rows, kind) {
    var f = ROW_FIELDS[kind];
    return (Array.isArray(rows) ? rows : []).map(function (r) {
      r = r || {};
      var out = emptyRow(kind);
      f.text.forEach(function (k) { out[k] = r[k] == null ? '' : String(r[k]).trim(); });
      f.num.forEach(function (k) { out[k] = parseNum(r[k]); });
      return out;
    });
  }

  // 补齐缺失的 key、把字符串数字转成数字；返回新对象，不改原对象
  function normalizeCohort(cohort) {
    var c = cohort || {};
    var data = c.data || {};
    var out = {};
    Object.keys(c).forEach(function (k) { out[k] = c[k]; });
    out.id = c.id == null ? null : c.id;
    out.channel = c.channel;
    out.name = c.name == null ? '' : String(c.name);
    out.live_date = c.live_date == null ? '' : String(c.live_date);
    out.closed = !!c.closed;
    out.summary = c.summary == null ? '' : String(c.summary);
    out.version = c.version == null ? 1 : c.version;
    out.data = {
      invites: normalizeRows(data.invites, 'invites'),
      community: normalizeRows(data.community, 'community'),
      conversions: normalizeRows(data.conversions, 'conversions'),
      unparsed: (Array.isArray(data.unparsed) ? data.unparsed : []).map(String),
    };
    // 复盘和行动是持久化业务数据；指标计算归一化不能把它丢掉。
    if (data.review != null) out.data.review = JSON.parse(JSON.stringify(data.review));
    return out;
  }

  // ================= 指标 =================

  function rowName(row, idx) {
    return row && row.date ? fmtDateCN(row.date) : '第 ' + (idx + 1) + ' 行';
  }

  function computeMetrics(cohort) {
    var c = normalizeCohort(cohort);
    var ch = CHANNELS[c.channel] ? c.channel : 'xiaoe';
    var isX = ch === 'xiaoe';
    var attLabel = CHANNELS[ch].attendanceLabel;
    var d = c.data;
    var warnings = [];

    // 邀约
    var reachKnown = null, joinedKnown = null, joinedTotal = null;
    d.invites.forEach(function (r) {
      if (r.reach != null) {
        reachKnown = (reachKnown || 0) + r.reach;
        if (r.joined != null) joinedKnown = (joinedKnown || 0) + r.joined;
      }
      if (r.joined != null) joinedTotal = (joinedTotal || 0) + r.joined;
    });
    var inviteByChannel = d.invites.map(function (r) {
      return { date: r.date, channel: r.channel, reach: r.reach, joined: r.joined, rate: ratio(r.joined, r.reach) };
    });

    // 社群：填了出勤人数的那行就是直播当天
    var liveIdx = [];
    d.community.forEach(function (r, i) { if (r.attendance != null) liveIdx.push(i); });
    var liveRow = liveIdx.length ? d.community[liveIdx[liveIdx.length - 1]] : null;
    if (liveIdx.length > 1) {
      warnings.push('社群数据里有 ' + liveIdx.length + ' 行填了' + attLabel + '，系统按最后一行（' +
        rowName(liveRow, liveIdx[liveIdx.length - 1]) + '）算');
    }
    var members = null;
    if (liveRow && liveRow.members != null) members = liveRow.members;
    else {
      for (var i = d.community.length - 1; i >= 0; i--) {
        if (d.community[i].members != null) { members = d.community[i].members; break; }
      }
    }
    var attendance = liveRow ? liveRow.attendance : null;
    var communityDaily = d.community.map(function (r) {
      return {
        date: r.date, group: r.group, members: r.members,
        claimed: r.claimed, claimRate: ratio(r.claimed, r.members),
        checkins: r.checkins, checkinRate: ratio(r.checkins, r.claimed),
        attendance: r.attendance, attendRate: ratio(r.attendance, r.members),
      };
    });

    // 成交
    var sums = {};
    SUM_KEYS.forEach(function (k) { sums[k] = sumField(d.conversions, k); });

    var netDeposit = null, pendingBalance = null, balanceRate = null, depositRate = null;
    var liveDeals, totalDeals;
    if (isX) {
      netDeposit = sums.deposit - sums.deposit_refund;
      pendingBalance = Math.max(0, netDeposit - sums.balance);
      balanceRate = ratio(sums.balance, netDeposit);
      depositRate = ratio(sums.deposit, attendance);
      liveDeals = sums.balance;
      totalDeals = sums.balance + sums.direct + sums.alumni;
    } else {
      liveDeals = sums.live_full;
      totalDeals = sums.live_full + sums.direct + sums.alumni - sums.refund;
    }
    var directShare = ratio(sums.direct, totalDeals);

    var inviteRate = ratio(joinedKnown, reachKnown);
    var attendRate = ratio(attendance, members);
    var liveConvRate = ratio(liveDeals, attendance);
    var communityConvRate = ratio(totalDeals, members);

    var steps = isX
      ? [['reach', '触达', reachKnown], ['joined', '进群', joinedTotal], ['members', '直播当天在群', members],
        ['attendance', '直播到课', attendance], ['deposit', '付定金', sums.deposit], ['balance', '补尾款', sums.balance]]
      : [['reach', '触达', reachKnown], ['joined', '进群', joinedTotal], ['members', '直播当天在群', members],
        ['attendance', '直播观看', attendance], ['live_full', '直播全款成交', sums.live_full]];
    var funnel = steps.map(function (s, idx) {
      return { key: s[0], label: s[1], value: s[2], rateFromPrev: idx ? ratio(s[2], steps[idx - 1][2]) : null };
    });

    var status, statusText;
    if (c.closed) { status = 'closed'; statusText = '已结束'; }
    else if (isX && pendingBalance > 0) { status = 'collecting'; statusText = '尾款收集中 · 还差 ' + pendingBalance + ' 人'; }
    else { status = 'open'; statusText = '进行中'; }

    if (!d.invites.length) warnings.push('还没有邀约数据，邀约入群率暂时算不出来');
    else if (reachKnown == null) warnings.push('邀约数据里没有触达人数，邀约入群率暂时算不出来');
    if (members == null) warnings.push('没找到直播当天的社群人数，直播到课率和社群转化率暂时算不出来');
    if (attendance == null) warnings.push('社群数据里还没填' + attLabel + '（填在直播当天那一行），直播到课率和直播转化率暂时算不出来');
    if (d.unparsed.length) warnings.push('还有 ' + d.unparsed.length + ' 行未识别的数据，确认没用后可以删掉');

    return {
      channel: ch,
      reachKnown: reachKnown, joinedKnown: joinedKnown, joinedTotal: joinedTotal, inviteRate: inviteRate,
      inviteByChannel: inviteByChannel,
      members: members, attendance: attendance, attendRate: attendRate,
      communityDaily: communityDaily,
      sums: sums,
      netDeposit: netDeposit, pendingBalance: pendingBalance, balanceRate: balanceRate, depositRate: depositRate,
      liveDeals: liveDeals, totalDeals: totalDeals, directShare: directShare,
      liveConvRate: liveConvRate, communityConvRate: communityConvRate,
      funnel: funnel,
      extra: { direct: sums.direct, alumni: sums.alumni, totalDeals: totalDeals },
      status: status, statusText: statusText,
      warnings: warnings,
    };
  }

  // ================= 对比 =================

  function compareCohort(target, allCohorts) {
    var t = target || {};
    var tm = computeMetrics(t);
    var history = (allCohorts || []).filter(function (c) {
      if (!c || c === target) return false;
      if (t.id != null && c.id === t.id) return false;
      return c.channel === t.channel && isISODate(c.live_date) && isISODate(t.live_date) && c.live_date < t.live_date;
    }).sort(function (a, b) { return a.live_date < b.live_date ? 1 : a.live_date > b.live_date ? -1 : 0; });

    var hm = history.map(computeMetrics);
    var prev = history[0] || null;
    var recent = hm.slice(0, 4); // 近 4 期
    var avgCount = 0;
    var metrics = {};
    CORE_METRICS.forEach(function (m) {
      var value = tm[m.key];
      var pv = prev ? hm[0][m.key] : null;
      var vals = recent.map(function (x) { return x[m.key]; }).filter(isNum);
      var avg = vals.length ? vals.reduce(function (a, b) { return a + b; }, 0) / vals.length : null;
      if (vals.length > avgCount) avgCount = vals.length;
      metrics[m.key] = {
        value: value,
        prev: pv,
        delta: isNum(value) && isNum(pv) ? value - pv : null,
        avg: avg,
        deltaAvg: isNum(value) && isNum(avg) ? value - avg : null,
        prevName: prev ? prev.name : null,
        avgCount: vals.length, // 该指标参与均值的期数
      };
    });
    return { prev: prev, avgCount: avgCount, metrics: metrics };
  }

  // ================= 自动分析 =================

  function buildInsights(cohort, metrics, comparison) {
    var c = normalizeCohort(cohort);
    var mt = metrics || computeMetrics(c);
    var cmp = comparison || { prev: null, avgCount: 0, metrics: {} };
    var isX = mt.channel === 'xiaoe';
    var attLabel = (CHANNELS[mt.channel] || CHANNELS.xiaoe).attendanceLabel;
    var out = [];
    function add(level, text) { out.push({ level: level, text: text }); }

    // 1. 和上期比
    if (cmp.prev) {
      var items = CORE_METRICS.map(function (m) { return { m: m, x: cmp.metrics[m.key] }; })
        .filter(function (it) { return it.x && isNum(it.x.delta); });
      var worst = null, best = null;
      items.forEach(function (it) {
        if (isFlat(it.x.delta)) return;
        if (it.x.delta < 0 && (!worst || it.x.delta < worst.x.delta)) worst = it;
        if (it.x.delta > 0 && (!best || it.x.delta > best.x.delta)) best = it;
      });
      if (worst) add('bad', worst.m.label + ' ' + fmtPct(worst.x.value) + '，比上期下降 ' + ppText(worst.x.delta) + ' 个百分点，是本期掉得最多的一环');
      if (best) add('good', best.m.label + ' ' + fmtPct(best.x.value) + '，比上期上升 ' + ppText(best.x.delta) + ' 个百分点，是本期涨得最多的一环');
      items.forEach(function (it) {
        if (it === worst || it === best || isFlat(it.x.delta)) return;
        var up = it.x.delta > 0;
        add(up ? 'good' : 'bad', it.m.label + ' ' + fmtPct(it.x.value) + '，比上期' + (up ? '上升 ' : '下降 ') + ppText(it.x.delta) + ' 个百分点');
      });
      if (!items.length) add('info', '上期（' + (cmp.prev.name || '') + '）的数据不全，暂时没法对比');
      else if (!worst && !best) add('info', '四个核心指标和上期基本持平');
    } else {
      add('info', '这是该渠道第一期数据，下周开始就能看到对比');
    }

    // 2. 和近几期均值比
    CORE_METRICS.forEach(function (m) {
      var x = cmp.metrics && cmp.metrics[m.key];
      if (!x || !(x.avgCount >= 2) || !isNum(x.deltaAvg)) return;
      if (Math.abs(x.deltaAvg) < 0.01 - 1e-9) return;
      var hi = x.deltaAvg > 0;
      add(hi ? 'good' : 'bad', m.label + ' ' + fmtPct(x.value) + '，' + (hi ? '高于' : '低于') + '近 ' + x.avgCount + ' 期均值 ' + fmtPct(x.avg));
    });

    // 3. 待补尾款
    if (isX && mt.pendingBalance > 0) {
      if (c.closed) add('info', '这期已标记结束，有 ' + mt.pendingBalance + ' 位付了定金、最后没补尾款');
      else add('warn', '还有 ' + mt.pendingBalance + ' 位付了定金、没补尾款，记得继续追');
    }

    // 4. 退定金 / 退款
    if (isX && mt.sums.deposit_refund > 0) {
      add('info', '退定金 ' + mt.sums.deposit_refund + ' 人' +
        (mt.sums.deposit > 0 ? '，占定金人数 ' + fmtPct(mt.sums.deposit_refund / mt.sums.deposit) : ''));
    }
    if (!isX && mt.sums.refund > 0) {
      var gross = mt.sums.live_full + mt.sums.direct + mt.sums.alumni;
      add('info', '退款 ' + mt.sums.refund + ' 人' + (gross > 0 ? '，占退款前成交人数 ' + fmtPct(mt.sums.refund / gross) : ''));
    }

    // 5. 1v1 追单
    if (mt.sums.direct > 0) {
      add('info', '1v1 追单成交 ' + mt.sums.direct + ' 人' + (isNum(mt.directShare) ? '，占总成交 ' + fmtPct(mt.directShare) : ''));
    }

    // 6. 邀约渠道对比
    var withRate = mt.inviteByChannel.filter(function (r) { return isNum(r.rate); });
    if (withRate.length >= 2) {
      var nameCount = {};
      withRate.forEach(function (r) { nameCount[r.channel] = (nameCount[r.channel] || 0) + 1; });
      var label = function (r, idx) {
        var n = r.channel || '';
        var parts = parseInviteDateParts(r.date);
        var when = parts && parts.end ? fmtInviteDate(r.date, +c.live_date.slice(0, 4), c.live_date) : shortDate(r.date);
        if (!n) return r.date ? when : '第 ' + (idx + 1) + ' 个渠道';
        return nameCount[n] > 1 && r.date ? when + ' ' + n : n;
      };
      var hiIdx = 0, loIdx = 0;
      withRate.forEach(function (r, idx) {
        if (r.rate > withRate[hiIdx].rate) hiIdx = idx;
        if (r.rate < withRate[loIdx].rate) loIdx = idx;
      });
      if (hiIdx !== loIdx) {
        add('info', '入群率最高的渠道：' + label(withRate[hiIdx], hiIdx) + '（' + fmtPct(withRate[hiIdx].rate) +
          '）；最低：' + label(withRate[loIdx], loIdx) + '（' + fmtPct(withRate[loIdx].rate) + '）');
      }
    }

    // 7. 课程领取率变化
    var claimDays = [];
    mt.communityDaily.forEach(function (r, idx) { if (isNum(r.claimRate)) claimDays.push({ idx: idx, rate: r.claimRate }); });
    if (claimDays.length >= 2) {
      var a = claimDays[0], b = claimDays[claimDays.length - 1];
      var pa = fmtPct(a.rate), pb = fmtPct(b.rate);
      var verb = pa === pb ? '保持在' : (b.rate < a.rate ? '降到' : '升到');
      add('info', '课程领取率从第 ' + (a.idx + 1) + ' 天的 ' + pa + ' ' + verb + '第 ' + (b.idx + 1) + ' 天的 ' + pb);
    }

    // 8. 数据缺失
    if (mt.members == null) add('warn', '没有直播当天的社群人数，直播到课率和社群转化率算不出来');
    if (mt.attendance == null) add('warn', '还没填' + attLabel + '，直播到课率和直播转化率算不出来');
    if (c.data.unparsed.length) add('warn', '有 ' + c.data.unparsed.length + ' 行未识别的数据还没处理，去编辑页确认一下');

    return out;
  }

  // ================= 保存前校验 =================

  // 返回 [{ level: 'error'|'warn', text }]；error 建议拦住保存，warn 只提示
  function validateCohortDetailed(cohort) {
    var out = [];
    function err(t) { out.push({ level: 'error', text: t }); }
    function wrn(t) { out.push({ level: 'warn', text: t }); }
    var raw = cohort || {};
    var c = normalizeCohort(raw);
    var rawData = raw.data || {};

    if (!String(c.name).trim()) err('请填写期次名称');
    if (!CHANNELS[c.channel]) err('请选择渠道（小鹅通 / B站）');
    if (!c.live_date) err('请填写直播日期');
    else if (!isISODate(c.live_date)) err('直播日期格式不对，请重新选一下');

    ['invites', 'community', 'conversions'].forEach(function (kind) {
      var cols = columnsFor(kind, c.channel);
      var labelOf = {};
      cols.forEach(function (cc) { labelOf[cc.key] = cc.label; });
      var rawRows = Array.isArray(rawData[kind]) ? rawData[kind] : [];
      c.data[kind].forEach(function (r, i) {
        var where = TABLE_NAMES[kind] + ' 第 ' + (i + 1) + ' 行';
        if (r.date && (kind === 'invites'
          ? !isInviteDate(r.date, +c.live_date.slice(0, 4), c.live_date) : !isISODate(r.date))) {
          wrn(where + '的日期「' + r.date + '」格式不对，请改成 ' +
            (kind === 'invites' ? '8.19、8.19-23 或 8.28-9.3' : '3.15') + ' 这样的写法');
        }
        ROW_FIELDS[kind].num.forEach(function (k) {
          var label = labelOf[k] || k;
          var rv = rawRows[i] ? rawRows[i][k] : null;
          if (typeof rv === 'string' && parseNumInfo(rv).bad) wrn(where + '的「' + label + '」写的是「' + rv + '」，不是数字');
          var v = r[k];
          if (!isNum(v)) return;
          if (v < 0) err(where + '的「' + label + '」是负数');
          else if (!Number.isInteger(v)) wrn(where + '的「' + label + '」不是整数');
        });
        if (kind === 'invites' && isNum(r.reach) && isNum(r.joined) && r.reach >= 0 && r.joined > r.reach) {
          wrn(where + '进群人数（' + r.joined + '）比触达人数（' + r.reach + '）还多');
        }
        if (kind === 'community' && isNum(r.attendance) && isNum(r.members) && r.attendance > r.members) {
          wrn(where + labelOf.attendance + '（' + r.attendance + '）比社群总人数（' + r.members + '）还多');
        }
      });
    });

    var s = {};
    SUM_KEYS.forEach(function (k) { s[k] = sumField(c.data.conversions, k); });
    if (c.channel === 'xiaoe') {
      var net = s.deposit - s.deposit_refund;
      if (s.deposit_refund > s.deposit) wrn('已退定金合计 ' + s.deposit_refund + ' 人，比已付定金 ' + s.deposit + ' 人还多');
      if (s.balance > net) wrn('已付尾款合计 ' + s.balance + ' 人，比净定金 ' + net + ' 人（定金 − 退定金）还多，请检查');
    } else if (c.channel === 'bilibili') {
      var gross = s.live_full + s.direct + s.alumni;
      if (s.refund > gross) wrn('退款合计 ' + s.refund + ' 人，比成交人数 ' + gross + ' 人还多，请检查');
    }
    return out;
  }

  function validateCohort(cohort) {
    return validateCohortDetailed(cohort).map(function (x) { return x.text; });
  }

  // ================= 粘贴解析 =================

  function cleanCell(v) {
    return String(v == null ? '' : v)
      .replace(/[​-‍﻿]/g, '')
      .replace(/[　 ]/g, ' ')
      .replace(/\s*\n\s*/g, ' ')
      .trim();
  }

  // 按分隔符切分；单元格以引号开头时支持引号包裹（里面可以有分隔符和换行）
  function parseDelimited(text, delim) {
    var rows = [];
    var row = [];
    var n = text.length;
    var i = 0;
    if (!n) return [['']];
    for (;;) {
      var val = null;
      var k = i;
      if (text.charAt(i) === '"') {
        var j = i + 1, buf = '', closed = false;
        while (j < n) {
          var ch = text.charAt(j);
          if (ch === '"') {
            if (text.charAt(j + 1) === '"') { buf += '"'; j += 2; continue; }
            if (j + 1 >= n || text.charAt(j + 1) === delim || text.charAt(j + 1) === '\n') closed = true;
            break;
          }
          buf += ch;
          j++;
        }
        if (closed) { val = buf; k = j + 1; }
      }
      if (val === null) {
        k = i;
        while (k < n && text.charAt(k) !== delim && text.charAt(k) !== '\n') k++;
        val = text.slice(i, k);
      }
      row.push(val);
      if (k >= n) { rows.push(row); break; }
      if (text.charAt(k) === delim) {
        i = k + 1;
        if (i >= n) { row.push(''); rows.push(row); break; }
        continue;
      }
      rows.push(row); // 换行
      row = [];
      i = k + 1;
      if (i >= n) break;
    }
    return rows;
  }

  function splitRows(text) {
    var t = String(text == null ? '' : text).replace(/^﻿/, '').replace(/\r\n?/g, '\n');
    var rows, spaced = false;
    if (t.indexOf('\t') >= 0) {
      rows = parseDelimited(t, '\t');
    } else {
      var csv = parseDelimited(t, ',');
      var isCsv = csv.some(function (r) { return r.filter(function (x) { return cleanCell(x) !== ''; }).length >= 2; });
      spaced = !isCsv;
      rows = isCsv ? csv : t.split('\n').map(function (l) { return l.replace(/[　 ]/g, ' ').split(/ {2,}/); });
    }
    // spaced=true：没有制表符和逗号，按连续空格切的（空单元格会丢，列可能错位）
    return { rows: rows.map(function (r) { return r.map(cleanCell); }), spaced: spaced };
  }

  // 去掉所有空白（含全角空格、单元格内换行），表头和合计行按这个匹配
  function squash(s) { return String(s == null ? '' : s).replace(/\s+/g, ''); }

  var TITLE_RE = /(小鹅通|B站|b站|哔哩|bilibili)/i;
  var TOTAL_RE = /^(总计|合计|小计|共计|汇总|累计|总数|总和|total|sum)(?![a-z])/i;
  function isTotalText(s) { return TOTAL_RE.test(squash(s)); }

  function detectChannel(text) {
    if (/B站|b站|哔哩|bilibili/i.test(text)) return 'bilibili';
    if (/小鹅通/.test(text)) return 'xiaoe';
    return null;
  }

  function isRateHeader(h) {
    var s = h.replace(/[（(][^）)]*[）)]/g, '').trim();
    return /率$/.test(s) || /转化率/.test(s) || /[%％]$/.test(s);
  }

  // 判断表头行属于哪一段：'invites' | 'community' | 'conversions' | null
  function headerKind(cells) {
    var ne = nonEmpty(cells);
    if (ne.length < 2) return null;
    if (ne.some(isNumberish)) return null; // 表头里不会有纯数字
    var sq = ne.map(squash);
    function has(re) { return sq.some(function (c) { return re.test(c); }); }
    if ((has(/渠道/) && has(/触达|进群|入群/)) || (has(/触达/) && has(/进群|入群/))) return 'invites';
    if (has(/社群总人数/) || (has(/群/) && has(/出勤|观看|到课/))) return 'community';
    if (has(/尾款|定金|直播成交|全款|直播付款|直播下单/) || has(/直接付款|1v1|追单|老学员/i)) return 'conversions';
    return null;
  }

  // 认不出是哪张表、但看起来像表头的行：至少 2 格、全是文字，并且有「时间 / 日期 / …人数」这类列名
  function looksLikeHeader(cells) {
    var ne = nonEmpty(cells);
    if (ne.length < 2) return false;
    var texty = ne.every(function (c) {
      return !isNumberish(c) && !isNullToken(c) && !parseInviteDateParts(c) && !/^\d{1,2}[:：]\d{2}/.test(c);
    });
    if (!texty) return false;
    return ne.map(squash).some(function (c) { return /^(时间|日期)$/.test(c) || /(人数|人次|数量)$/.test(c); });
  }

  // 标题行：只有 1 个非空单元格且含渠道词；或者 A 列含渠道词、后面只多了几格文字（不是表头、没有数字）
  function isTitleRow(cells) {
    var ne = nonEmpty(cells);
    if (ne.length === 1) return TITLE_RE.test(ne[0]);
    if (!ne.length || cells[0] === '' || !TITLE_RE.test(cells[0]) || parseDateParts(cells[0])) return false;
    if (headerKind(cells) || looksLikeHeader(cells)) return false;
    return !ne.some(function (c) { return isNumberish(c) || isNullToken(c); });
  }

  // 列名 → 字段。'_rate' 率列（忽略）、'_check' 仅核对、'_sheetTotal' 总转化人数、'_unused' 没用到
  function mapInviteHeader(h) {
    if (isRateHeader(h)) return '_rate';
    if (/时间段/.test(h)) return 'slot';
    if (/时间|日期/.test(h)) return 'date';
    if (/渠道/.test(h)) return 'channel';
    if (/触达/.test(h)) return 'reach';
    if (/进群|入群/.test(h) && !/总计|在群|累计|合计/.test(h)) return 'joined';
    if (/总计|在群|累计|合计/.test(h)) return '_check'; // 累计在群人数，和进群数重复，不报警告
    return '_unused';
  }

  function mapCommunityHeader(h) {
    if (isRateHeader(h)) return '_rate';
    if (/时间|日期/.test(h)) return 'date';
    if (/^(群|群名|群号|群名称|社群|社群名称|群聊|群聊名称)$/.test(h)) return 'group';
    if (/社群总人数|在群人数|社群人数|群人数/.test(h)) return 'members';
    if (/领取|预约/.test(h)) return 'claimed';
    if (/打卡/.test(h)) return 'checkins';
    if (/出勤|观看|到课/.test(h)) return 'attendance';
    if (/助理/.test(h)) return 'add_assistant';
    return '_unused';
  }

  function mapConversionHeader(h) {
    if (isRateHeader(h)) return '_rate';
    if (/时间|日期/.test(h)) return 'date';
    if (/退/.test(h) && /定金/.test(h)) return 'deposit_refund';
    if (/定金/.test(h)) return 'deposit';
    if (/退/.test(h)) return 'refund';                 // 先判断「退」：「全款退款」是退款
    if (/尾款/.test(h)) return /待|未|欠|应/.test(h) ? '_unused' : 'balance'; // 「待补尾款」不是已付尾款
    if (/直播成交|全款|直播付款|直播下单/.test(h)) return 'live_full';
    if (/直接付款|1v1|追单/i.test(h)) return 'direct';
    if (/老学员/.test(h)) return 'alumni';
    if (/^已付/.test(h)) return 'low_price';
    if (/总转化/.test(h)) return '_sheetTotal';
    return '_unused';
  }

  var HEADER_MAPPERS = { invites: mapInviteHeader, community: mapCommunityHeader, conversions: mapConversionHeader };

  function buildSegment(kind, cells) {
    var used = {};
    var fieldIdx = {};
    var map = cells.map(function (h, idx) {
      if (!h) return null;
      var f = HEADER_MAPPERS[kind](squash(h)); // 「社群\n总人数」「直接 付款」也能认出来；labels 保留原文
      if (f !== '_rate' && f !== '_unused') {
        if (used[f]) return '_unused'; // 重复的列只用第一列
        used[f] = true;
        fieldIdx[f] = idx;
      }
      return f;
    });
    return { kind: kind, map: map, labels: cells.slice(), fieldIdx: fieldIdx };
  }

  function cellFor(seg, cells, field) {
    var idx = seg.fieldIdx[field];
    return idx == null ? '' : (cells[idx] || '');
  }

  function colLabel(seg, idx) {
    return (seg && seg.labels[idx]) || '第 ' + (idx + 1) + ' 列';
  }

  function hasDateLike(text) {
    return !!parseTitleDate(text, 2000);
  }

  // 标题里的日期：yyyyMMdd / yyMMdd / yyyy.M.D / M月D日 / M.D / MMdd
  function parseTitleDate(title, defaultYear) {
    var s = toHalfWidth(title);
    var m = s.match(/(^|\D)(\d{4})\s*[.\/\-年]\s*(\d{1,2})\s*[.\/\-月]\s*(\d{1,2})(?!\d)/);
    if (m && validYMD(+m[2], +m[3], +m[4])) {
      var okY = +m[2] >= 2020 && +m[2] <= 2040;
      return { raw: m[0].replace(/^\D/, ''), year: okY ? +m[2] : defaultYear, yearTrusted: okY, explicitYear: okY, m: +m[3], d: +m[4] };
    }
    m = s.match(/(^|\D)(\d{8})(?!\d)/);
    if (m) {
      var y8 = +m[2].slice(0, 4), m8 = +m[2].slice(4, 6), d8 = +m[2].slice(6, 8);
      if (validYMD(y8, m8, d8)) {
        var ok8 = y8 >= 2020 && y8 <= 2040;
        return { raw: m[2], year: ok8 ? y8 : defaultYear, yearTrusted: ok8, explicitYear: ok8, m: m8, d: d8 };
      }
    }
    m = s.match(/(^|\D)(\d{6})(?!\d)/);
    if (m) {
      var yy = +m[2].slice(0, 2), m6 = +m[2].slice(2, 4), d6 = +m[2].slice(4, 6);
      if (validYMD(null, m6, d6)) {
        var ok6 = yy >= 20 && yy <= 40;
        return { raw: m[2], year: ok6 ? 2000 + yy : defaultYear, yearTrusted: ok6, explicitYear: ok6, m: m6, d: d6 };
      }
    }
    m = s.match(/(\d{1,2})\s*月\s*(\d{1,2})/) || s.match(/(?:^|[^\d.])(\d{1,2})[.\/](\d{1,2})(?![\d.])/);
    if (m && validYMD(null, +m[1], +m[2])) {
      return { raw: m[0].replace(/^[^\d]/, ''), year: defaultYear, yearTrusted: true, explicitYear: false, m: +m[1], d: +m[2] };
    }
    m = s.match(/(^|\D)(\d{2})(\d{2})(?!\d)/); // 「小鹅通0315」
    if (m && validYMD(null, +m[2], +m[3])) {
      return { raw: m[2] + m[3], year: defaultYear, yearTrusted: true, explicitYear: false, m: +m[2], d: +m[3] };
    }
    return null;
  }

  // 按标题行切成多期
  function splitBlocks(rows) {
    var items = rows.map(function (cells) {
      return { cells: cells, hk: headerKind(cells), title: isTitleRow(cells) };
    });
    function newBlock(title) { return { title: title || '', lines: [], kinds: {}, lastKind: null, channelHint: null, noTitle: false }; }
    function nextHeader(from) {
      for (var k = from; k < items.length; k++) {
        if (items[k].title) return null;
        if (items[k].hk) return items[k].hk;
      }
      return null;
    }
    // 只有一格文字、带日期（如「0909期」「260916运营」），可以当标题
    function dateTitleText(cells) {
      var ne = nonEmpty(cells);
      if (ne.length !== 1 || isNumberish(ne[0]) || parseInviteDateParts(ne[0])) return null;
      return hasDateLike(ne[0]) ? ne[0] : null;
    }
    // 新开一期前，把上一期末尾的标题样的行拿过来（空行跳过）
    function takeTrailingTitle(block) {
      for (var k = block.lines.length - 1; k >= 0; k--) {
        var ln = block.lines[k];
        if (!nonEmpty(ln.cells).length) continue;
        var t = dateTitleText(ln.cells);
        if (t) { block.lines.splice(k); return t; }
        return null;
      }
      return null;
    }
    var blocks = [];
    var cur = newBlock('');
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      if (it.hk) {
        // 同一种表在后面又出现了（比如成交表之后又来一张邀约表）：前面没复制到标题的新一期
        if (cur.kinds[it.hk] && cur.lastKind !== it.hk) {
          var tt = takeTrailingTitle(cur);
          blocks.push(cur);
          cur = newBlock(tt);
          cur.noTitle = !tt;
          cur.titleNoChannel = !!tt;
        }
        cur.kinds[it.hk] = true;
        cur.lastKind = it.hk;
        cur.lines.push(it);
        continue;
      }
      // 第一期还没开始时，带日期的单格文字当标题（标题里没写渠道名）
      if (!it.title && !blocks.length && !cur.title && !Object.keys(cur.kinds).length && !hasAnyData(cur)) {
        var dt = dateTitleText(it.cells);
        if (dt) { cur.title = dt; cur.titleNoChannel = true; continue; }
      }
      if (it.title) {
        var text = nonEmpty(it.cells)[0];
        var hasKinds = Object.keys(cur.kinds).length > 0;
        var nx = nextHeader(i + 1);
        // 第一个标题：标题前的内容归入这一期；除非标题后面又出现了同样的表（说明前面是完整的另一期，只是没复制到标题）
        if (!blocks.length && !cur.title) {
          if (hasKinds && (nx ? cur.kinds[nx] : hasDateLike(text))) { blocks.push(cur); cur = newBlock(text); }
          else cur.title = text;
          continue;
        }
        // 带日期的标题、或者后面又出现了本期已有的表 → 新的一期；否则当作段说明（如「B站社群数据」）
        if (hasDateLike(text) || (hasKinds && (!nx || cur.kinds[nx]))) {
          blocks.push(cur);
          cur = newBlock(text);
        } else if (!cur.channelHint) {
          cur.channelHint = detectChannel(text);
        }
        continue;
      }
      cur.lines.push(it);
    }
    blocks.push(cur);
    return blocks;
  }

  function hasAnyData(block) {
    return block.lines.some(function (l) { return nonEmpty(l.cells).length > 0; });
  }

  function parseBlock(block, ctx) {
    var headWarnings = [];
    var warnings = [];
    function warn(msg) { pushUnique(warnings, msg); }

    var title = block.title || '';
    var titleChannel = title ? detectChannel(title) : null;
    var titleInfo = title ? parseTitleDate(title, ctx.defaultYear) : null;
    var baseYear = ctx.liveDate ? +ctx.liveDate.slice(0, 4) :
      (titleInfo && titleInfo.yearTrusted ? titleInfo.year : ctx.defaultYear);

    var tables = { invites: [], community: [], conversions: [] };
    var unparsed = [];
    var dateEntries = []; // { row, raw, segName, kind }
    var seg = null;
    var segRows = [];
    var segCount = 0;
    var unusedWarned = {};
    var outside = [];
    var sheetTotals = null;
    var rowTotals = [];
    var inviteTotals = [];
    var columnHint = null;
    var unknownHeader = false; // 上一个表头认不出来：下面的行都进 unparsed
    var convLabels = {};       // 成交表字段 → 原表头文字
    var shortRows = 0, misRows = 0;

    function rowText(cells) {
      var first = -1, last = -1;
      cells.forEach(function (c, k) { if (c !== '') { if (first < 0) first = k; last = k; } });
      return cells.slice(first, last + 1).join(' | ');
    }

    // 按空格切分时，比表头少格子的行：各格类型对得上就照常解析，对不上（错位）就进 unparsed
    function spacedRowState(cells) {
      if (!ctx.spaced || nonEmpty(cells).length >= seg.width) return 'ok';
      var bad = Object.keys(seg.fieldIdx).some(function (f) {
        var v = cells[seg.fieldIdx[f]] || '';
        if (!v) return false;
        if (f === 'date') return !(seg.kind === 'invites' ? parseInviteDateParts(v) : parseDateParts(v)) && !isTotalText(v);
        if (ROW_FIELDS[seg.kind].text.indexOf(f) >= 0) return isNumberish(v) || isNullToken(v);
        if (ROW_FIELDS[seg.kind].num.indexOf(f) >= 0) return parseNumInfo(v).bad;
        return false;
      });
      return bad ? 'bad' : 'short';
    }
    // 返回 true 表示这一行已放进 unparsed
    function dropMisaligned(cells, segName) {
      var st = spacedRowState(cells);
      if (st === 'short') shortRows++;
      if (st !== 'bad') return false;
      misRows++;
      unparsed.push('[' + segName + '] ' + rowText(cells));
      return true;
    }

    // 日期 / 渠道那格写的不是日期，但各数字列正好等于上面各行之和 → 合计行
    function looksLikeTotal(cells) {
      var key = cellFor(seg, cells, 'date') || (seg.kind === 'invites' ? cellFor(seg, cells, 'channel') : '');
      if (!key || (seg.kind === 'invites' ? parseInviteDateParts(key) : parseDateParts(key)) || segRows.length < 2) return false;
      var n = 0, nonzero = false, ok = true;
      ROW_FIELDS[seg.kind].num.forEach(function (f) {
        var idx = seg.fieldIdx[f];
        if (idx == null || !ok) return;
        var v = parseNum(cells[idx]);
        if (v == null) return;
        n++;
        if (v) nonzero = true;
        if (v !== sumField(segRows, f)) ok = false;
      });
      if (!ok || !n || !nonzero) return false;
      warn(SEG_NAMES[seg.kind] + '里「' + key + '」那一行的数字正好是上面各行加起来的，已当作合计行，没算进数据');
      return true;
    }

    function rowDesc(segName, cells) {
      var key = cellFor(seg, cells, 'date') || cellFor(seg, cells, 'channel') || cellFor(seg, cells, 'group');
      return segName + (key ? ' ' + key + ' 那一行' : '里有一行') + '的';
    }

    function readNum(cells, field, segName) {
      var idx = seg.fieldIdx[field];
      if (idx == null) return null;
      var info = parseNumInfo(cells[idx]);
      if (info.bad) warn(rowDesc(segName, cells) + '「' + seg.labels[idx] + '」写的是「' + cells[idx] + '」，不是数字，已当作空值');
      return info.value;
    }

    function checkUnused(cells) {
      cells.forEach(function (v, idx) {
        if (!v || isNullToken(v)) return;
        var f = seg.map[idx];
        if (f === '_unused') {
          if (!unusedWarned[idx]) { unusedWarned[idx] = 1; warn('列「' + seg.labels[idx] + '」有数据，系统暂时没用到'); }
        } else if (f == null) {
          if (!unusedWarned[idx]) { unusedWarned[idx] = 1; warn('第 ' + (idx + 1) + ' 列有数据但没有表头，系统暂时没用到'); }
        }
      });
    }

    // 单个非空单元格是不是「段说明行」
    function isDescription(cells) {
      var idx = -1;
      cells.forEach(function (c, k) { if (c !== '' && idx < 0) idx = k; });
      var v = cells[idx];
      if (isNumberish(v) || (seg && seg.kind === 'invites' ? parseInviteDateParts(v) : parseDateParts(v))) return false;
      if (idx === 0 || !seg) return true;
      // 写在已识别列里的（比如只填了渠道、数字列里写了字）按数据行处理
      var f = seg.map[idx];
      var fields = ROW_FIELDS[seg.kind];
      return fields.num.indexOf(f) < 0 && fields.text.indexOf(f) < 0;
    }

    function stray(cells, needText) {
      var segName = SEG_NAMES[seg.kind];
      var parts = [];
      cells.forEach(function (v, idx) { if (v) parts.push({ label: colLabel(seg, idx), v: v }); });
      unparsed.push('[' + segName + '] ' + parts.map(function (p) { return p.label + '：' + p.v; }).join(' | '));
      if (parts.length === 1) {
        var extra = '';
        if (seg.kind === 'invites' && isNumberish(parts[0].v)) {
          extra = '，和上面加起来的 ' + sumField(segRows, 'joined') + ' 对不上';
        }
        warn(segName + '里有一行只在「' + parts[0].label + '」列写了 ' + parts[0].v + extra + '，系统没法判断它的含义，已放进『未识别的行』');
      } else {
        warn(segName + '里有一行没填' + needText + '（' + parts.map(function (p) { return p.label + ' ' + p.v; }).join('、') +
          '），系统没法判断它的含义，已放进『未识别的行』');
      }
    }

    function isTotalRow(cells) {
      var first = nonEmpty(cells)[0] || '';
      return isTotalText(first) || isTotalText(cellFor(seg, cells, 'date'));
    }

    function addRow(kind, cells, segName) {
      var row = emptyRow(kind);
      ROW_FIELDS[kind].text.forEach(function (k) { row[k] = cellFor(seg, cells, k); });
      ROW_FIELDS[kind].num.forEach(function (k) { row[k] = readNum(cells, k, segName); });
      if (row.date) dateEntries.push({ row: row, raw: row.date, segName: segName, kind: kind });
      tables[kind].push(row);
      segRows.push(row);
      checkUnused(cells);
      return row;
    }

    block.lines.forEach(function (line) {
      var cells = line.cells;
      var ne = nonEmpty(cells);
      if (!ne.length) return;
      if (line.hk) {
        seg = buildSegment(line.hk, cells);
        seg.width = ne.length;
        segRows = [];
        segCount++;
        unknownHeader = false;
        if (seg.fieldIdx.date == null && (line.hk !== 'invites' || seg.fieldIdx.channel == null)) {
          warn(SEG_NAMES[line.hk] + '的表头里没有「时间」或「日期」列，下面的行没法归位，请连这一列一起复制');
        }
        if (line.hk === 'conversions') {
          if (seg.fieldIdx.deposit != null || seg.fieldIdx.balance != null) columnHint = columnHint || 'xiaoe';
          else if (seg.fieldIdx.live_full != null) columnHint = columnHint || 'bilibili';
          Object.keys(seg.fieldIdx).forEach(function (f) { if (!convLabels[f]) convLabels[f] = seg.labels[seg.fieldIdx[f]]; });
        }
        return;
      }
      if (ne.every(isNullToken)) return; // 只有 / - 之类的占位
      if (ne.length === 1 && isDescription(cells)) return;
      // 认不出的表头（比如运营自己加的小表）：结束当前段，下面的行不往上一张表里塞
      if (looksLikeHeader(cells)) {
        seg = null;
        unknownHeader = true;
        warn('这一行看起来是表头，但认不出是哪张表：' + ne.join(' | ') + '。它下面的行已放进『未识别的行』，有用的数字请手动填进表格');
        return;
      }
      if (!seg) {
        var text = rowText(cells);
        unparsed.push(text);
        if (!unknownHeader) outside.push(text);
        return;
      }
      var segName = SEG_NAMES[seg.kind];

      if (seg.kind === 'invites') {
        if (isTotalRow(cells) || looksLikeTotal(cells)) { inviteTotals.push({ cells: cells, seg: seg, rows: segRows.slice() }); return; }
        if (cellFor(seg, cells, 'date') || cellFor(seg, cells, 'channel')) {
          if (!dropMisaligned(cells, segName)) addRow('invites', cells, segName);
          return;
        }
        // 小计行：数字等于上面已解析的进群合计（或触达合计）→ 静默忽略
        var sumJ = sumField(segRows, 'joined'), sumR = sumField(segRows, 'reach');
        var vals = ne.filter(function (c) { return !isNullToken(c); }).map(parseNumInfo);
        if (vals.every(function (v) { return v.value != null && (v.value === sumJ || v.value === sumR); })) return;
        stray(cells, '日期和渠道');
        return;
      }

      if (seg.kind === 'community') {
        if (isTotalRow(cells)) return;
        if (cellFor(seg, cells, 'date')) {
          if (!dropMisaligned(cells, segName)) addRow('community', cells, segName);
          return;
        }
        stray(cells, '日期');
        return;
      }

      // 成交段
      if (isTotalRow(cells) || looksLikeTotal(cells)) {
        sheetTotals = {};
        Object.keys(seg.fieldIdx).forEach(function (f) {
          if (f === 'date' || f === '_check') return;
          var idx = seg.fieldIdx[f];
          sheetTotals[f] = { value: parseNum(cells[idx]), label: seg.labels[idx] };
        });
        return;
      }
      if (cellFor(seg, cells, 'date')) {
        if (dropMisaligned(cells, segName)) return;
        var row = addRow('conversions', cells, segName);
        var tIdx = seg.fieldIdx._sheetTotal;
        if (tIdx != null) rowTotals.push({ row: row, value: parseNum(cells[tIdx]), label: seg.labels[tIdx] });
        return;
      }
      stray(cells, '日期');
    });

    // 渠道：标题 > 段说明里的渠道 > 选择的渠道 > 成交表列名 > 小鹅通
    var channel = titleChannel || block.channelHint || ctx.optChannel || columnHint || 'xiaoe';
    if (titleChannel && ctx.optChannel && titleChannel !== ctx.optChannel) {
      headWarnings.push('标题里写的是' + channelLabel(titleChannel) + '，这期已按' + channelLabel(titleChannel) + '识别（和选的渠道不同）');
    } else if (!titleChannel && !block.channelHint && columnHint && columnHint !== channel) {
      headWarnings.push('成交表的列名看起来是' + channelLabel(columnHint) + '的数据，但现在按' + channelLabel(channel) + '识别，请确认渠道');
    }
    // 这个渠道用不到的成交列（如小鹅通表里的「退款」）：编辑器里看不到，不存进去，提示一下
    // 渠道可能选错时（列名像另一个渠道）先保留，换渠道后还能看到
    if (!columnHint || columnHint === channel) {
      var allowed = conversionColumns(channel).map(function (cc) { return cc.key; });
      SUM_KEYS.forEach(function (f) {
        if (allowed.indexOf(f) >= 0) return;
        var had = false;
        tables.conversions.forEach(function (r) { if (r[f]) had = true; r[f] = null; });
        if (had) warn('列「' + (convLabels[f] || f) + '」有数据，但' + channelLabel(channel) + '的期不用这一列，系统没算进去');
      });
    }

    // 日期：没写年份的，取离锚点最近的年份（处理跨年）
    // 年份没写明时，锚点不能比今天晚一个多月，否则按去年算（比如 1 月补录 12 月那期）
    var yearGuessed = false;
    function implicitYear(m, d) {
      var y = ctx.defaultYear;
      if (ctx.today != null && validYMD(y - 1, m, d) && Date.UTC(y, m - 1, d) > ctx.today + 31 * 86400000) {
        yearGuessed = true;
        return y - 1;
      }
      return y;
    }
    var anchor = ctx.liveDate ? isoToUTC(ctx.liveDate) : null;
    if (titleInfo && !ctx.liveDate) {
      if (!titleInfo.explicitYear) baseYear = titleInfo.year = implicitYear(titleInfo.m, titleInfo.d);
      anchor = Date.UTC(titleInfo.year, titleInfo.m - 1, titleInfo.d);
    }
    if (anchor == null) {
      var liveCandidates = tables.community.filter(function (r) { return r.attendance != null; });
      var probe = liveCandidates.length ? liveCandidates[liveCandidates.length - 1].date : null;
      var probeEntry = null;
      dateEntries.forEach(function (e) { if (!probeEntry && (probe == null || e.raw === probe) && parseDateParts(e.raw)) probeEntry = e; });
      if (!probeEntry) dateEntries.forEach(function (e) { if (!probeEntry && parseDateParts(e.raw)) probeEntry = e; });
      if (probeEntry) {
        var pp = parseDateParts(probeEntry.raw);
        var py = pp.y || (baseYear = implicitYear(pp.m, pp.d));
        if (validYMD(py, pp.m, pp.d)) anchor = Date.UTC(py, pp.m - 1, pp.d);
      }
    }
    dateEntries.forEach(function (e) {
      // 范围必须等直播日期确定后再定年份（1 月直播的邀约可能始于去年）。
      if (e.kind === 'invites' && !parseDateParts(e.raw)) return;
      var p = parseDateParts(e.raw);
      if (!p) {
        warn(e.segName + '里的日期「' + e.raw + '」系统认不出来，请改成 3.15 这样的写法');
        return;
      }
      var y = p.y;
      if (!y) {
        var best = null;
        [baseYear - 1, baseYear, baseYear + 1].forEach(function (cy) {
          if (!validYMD(cy, p.m, p.d)) return;
          var dist = anchor == null ? (cy === baseYear ? 0 : Infinity) : Math.abs(Date.UTC(cy, p.m - 1, p.d) - anchor);
          if (best == null || dist < best.dist) best = { y: cy, dist: dist };
        });
        if (!best) { warn(e.segName + '里的日期「' + e.raw + '」不存在，请检查'); return; }
        y = best.y;
      }
      e.row.date = ymdToISO(y, p.m, p.d);
    });

    // 手选日期优先；未选择时从出勤、首笔直播成交或标题推断。
    var liveDate = '';
    var liveRows = tables.community.filter(function (r) { return r.attendance != null; });
    var lr = liveRows.length ? liveRows[liveRows.length - 1] : null;
    if (lr && isISODate(lr.date)) liveDate = lr.date;
    if (!liveDate) {
      var dr = tables.conversions.filter(function (r) { return (r.deposit || 0) > 0 || (r.live_full || 0) > 0; })[0];
      if (dr && isISODate(dr.date)) liveDate = dr.date;
    }
    var inferredLiveDate = liveDate;
    if (ctx.liveDate) {
      liveDate = ctx.liveDate;
      if (inferredLiveDate && inferredLiveDate !== liveDate) {
        headWarnings.push('表格里推断的直播日期是 ' + inferredLiveDate + '，与选择的 ' + liveDate + ' 不同，已按你选择的日期归期，请核对');
      }
    }
    dateEntries.forEach(function (e) {
      if (e.kind !== 'invites' || parseDateParts(e.raw)) return;
      var inviteAnchor = liveDate || (anchor == null ? '' : utcToISO(anchor));
      var inviteYear = inviteAnchor ? +inviteAnchor.slice(0, 4) : baseYear;
      if (isInviteDate(e.raw, inviteYear, inviteAnchor)) {
        e.row.date = normalizeInviteDate(e.raw, inviteYear, inviteAnchor);
      } else {
        warn(e.segName + '里的日期范围「' + e.raw + '」格式不对或起止顺序有误，请写成 8.19-23 或 8.28-9.3');
      }
    });
    var fromData = !!liveDate;
    if (titleInfo && ctx.liveDate) {
      var selectedMatchesTitle = titleInfo.explicitYear
        ? ymdToISO(titleInfo.year, titleInfo.m, titleInfo.d) === liveDate
        : pad2(titleInfo.m) + '-' + pad2(titleInfo.d) === liveDate.slice(5);
      if (!titleInfo.yearTrusted || !selectedMatchesTitle) {
        headWarnings.push('标题里的日期 ' + titleInfo.raw + ' 与选择的直播日期不一致，已按你选择的 ' + liveDate + ' 归期，请核对');
      }
    } else if (titleInfo) {
      var titleISO = validYMD(titleInfo.year, titleInfo.m, titleInfo.d) ? ymdToISO(titleInfo.year, titleInfo.m, titleInfo.d) : '';
      if (fromData) {
        var mismatch = titleInfo.explicitYear ? titleISO !== liveDate : (pad2(titleInfo.m) + '-' + pad2(titleInfo.d)) !== liveDate.slice(5);
        if (!titleInfo.yearTrusted || mismatch) {
          headWarnings.push('标题里的日期 ' + titleInfo.raw + ' 看起来有误，直播日期已按数据推断为 ' + liveDate + '，请确认');
        }
      } else if (titleISO) {
        liveDate = titleISO;
        if (!titleInfo.yearTrusted) headWarnings.push('标题里的日期 ' + titleInfo.raw + ' 年份看起来有误，直播日期暂按 ' + liveDate + ' 填写，请确认');
      }
    }
    if (!liveDate) headWarnings.push('没找到直播日期，请手动填写');
    else if (yearGuessed && liveDate.slice(0, 4) === String(ctx.defaultYear - 1)) {
      headWarnings.push('表格里没写年份，直播日期按去年推断为 ' + liveDate + '，请确认');
    }
    if (block.noTitle && !ctx.liveDate) headWarnings.push('第 ' + block.index + ' 期是按表格重复出现自动分开的，请核对直播日期');

    // 核对：总计行 / 每行总转化人数
    var isX = channel === 'xiaoe';
    if (sheetTotals) {
      Object.keys(sheetTotals).forEach(function (f) {
        var t = sheetTotals[f];
        if (t.value == null || f === '_sheetTotal') return;
        var calc = sumField(tables.conversions, f);
        if (t.value !== calc) warn('表格里的『总计』行 ' + t.label + ' 写的是 ' + t.value + '，按每天加起来是 ' + calc + '，系统按每天的数字算');
      });
      var st = sheetTotals._sheetTotal;
      if (st && st.value != null) {
        var s = {};
        SUM_KEYS.forEach(function (k) { s[k] = sumField(tables.conversions, k); });
        var deals = isX ? s.balance + s.direct + s.alumni : s.live_full + s.direct + s.alumni;
        var dealsNet = isX ? deals : deals - s.refund;
        if (st.value !== deals && st.value !== dealsNet) {
          warn('表格里的『总计』行 ' + st.label + ' 写的是 ' + st.value + '，按每天的成交加起来是 ' + dealsNet + '，系统按每天的数字算');
        }
      }
    }
    rowTotals.forEach(function (rt) {
      if (rt.value == null) return;
      var r = rt.row;
      var n = function (v) { return v || 0; };
      if (isX) {
        var expect = n(r.balance) + n(r.direct) + n(r.alumni);
        if (rt.value !== expect) {
          warn('成交数据 ' + fmtDateCN(r.date) + ' 那一行「' + rt.label + '」写的是 ' + rt.value + '，按 已付尾款+直接付款+老学员 算是 ' + expect);
        }
      } else {
        var e1 = n(r.live_full) + n(r.direct) + n(r.alumni);
        var e2 = e1 - n(r.refund);
        if (rt.value !== e1 && rt.value !== e2) {
          warn('成交数据 ' + fmtDateCN(r.date) + ' 那一行「' + rt.label + '」写的是 ' + rt.value + '，按 直播成交+直接付款+老学员−退款 算是 ' + e2);
        }
      }
    });
    inviteTotals.forEach(function (t) {
      ['reach', 'joined'].forEach(function (f) {
        var idx = t.seg.fieldIdx[f];
        if (idx == null) return;
        var v = parseNum(t.cells[idx]);
        var calc = sumField(t.rows, f);
        if (v != null && v !== calc) {
          warn('表格里邀约数据的『合计』行 ' + t.seg.labels[idx] + ' 写的是 ' + v + '，按每行加起来是 ' + calc + '，系统按每行的数字算');
        }
      });
    });

    // 段缺失提示
    if (!segCount) {
      if (unparsed.length) headWarnings.push('没找到表头（比如「时间 / 渠道 / 触达人数」这一行），请连表头一起复制；内容都放进了『未识别的行』');
      else headWarnings.push('没识别到表格数据，请连表头一起复制');
    } else {
      outside.forEach(function (t) {
        warn('有一行不在任何表格里（' + (t.length > 30 ? t.slice(0, 30) + '…' : t) + '），已放进『未识别的行』');
      });
      if (!block.kinds.invites) warn('没找到邀约数据表（表头要有「渠道」和「触达人数」或「进群人数」）');
      if (!block.kinds.community) warn('没找到社群数据表（表头要有「社群总人数」）');
      if (!block.kinds.conversions) warn('没找到成交数据表（表头要有「定金 / 尾款」或「直播成交」），还没开始成交可以忽略');
    }

    if (misRows) {
      warn('复制的内容里没有制表符（可能是从聊天软件里复制的），空格子丢了，有 ' + misRows +
        ' 行的列对不上表头，已放进『未识别的行』。请从在线表格里直接复制');
    } else if (shortRows) {
      warn('复制的内容里没有制表符，有 ' + shortRows + ' 行比表头少几格，系统按从左到右对齐，请核对各列的数字');
    }

    var name = title || (channelLabel(channel) + ' ' + (liveDate ? fmtDateCN(liveDate) + ' 期' : '新一期'));
    var cohort = emptyCohort(channel);
    cohort.name = name;
    cohort.live_date = liveDate;
    cohort.data = { invites: tables.invites, community: tables.community, conversions: tables.conversions, unparsed: unparsed };
    // 这次粘贴里实际出现了哪几张表（没出现的表，更新已有期时保留原数据）
    cohort.segments = { invites: !!block.kinds.invites, community: !!block.kinds.community, conversions: !!block.kinds.conversions };
    cohort.warnings = headWarnings.concat(warnings.filter(function (w) { return headWarnings.indexOf(w) < 0; }));
    return cohort;
  }

  // 「今天」的 UTC 零点：opts.today 可以是 Date、'YYYY-MM-DD' 或不传（取当前时间）
  function todayUTC(v) {
    if (typeof v === 'string' && isISODate(v)) return isoToUTC(v);
    var d = v instanceof Date ? v : (typeof v === 'number' ? new Date(v) : new Date());
    if (isNaN(d.getTime())) d = new Date();
    return Date.UTC(d.getFullYear(), d.getMonth(), d.getDate());
  }

  function parsePaste(text, opts) {
    opts = opts || {};
    var selectedDate = opts.liveDate == null ? '' : String(opts.liveDate).trim();
    if (selectedDate && !isISODate(selectedDate)) {
      return { cohorts: [], warnings: ['本期直播日期格式不对，请重新选择'] };
    }
    var ctx = {
      defaultYear: selectedDate ? +selectedDate.slice(0, 4) : (toInt(opts.defaultYear) || new Date().getFullYear()),
      liveDate: selectedDate,
      optChannel: CHANNELS[opts.channel] ? opts.channel : null,
      today: todayUTC(opts.today),
    };
    var split = splitRows(text);
    var rows = split.rows;
    ctx.spaced = split.spaced;
    var hasContent = rows.some(function (r) { return nonEmpty(r).length > 0; });
    if (!hasContent) return { cohorts: [], warnings: ['粘贴的内容是空的'] };

    var warnings = [];
    var blocks = splitBlocks(rows).filter(function (b, idx) { return idx > 0 || b.title || hasAnyData(b); });
    blocks.forEach(function (b, idx) { b.index = idx + 1; });
    var cohorts = blocks.map(function (b) { return parseBlock(b, ctx); });

    // 有多期时，丢掉没有任何数据的期（比如末尾多出来的标题）
    if (cohorts.length > 1) {
      cohorts = cohorts.filter(function (c) {
        var d = c.data;
        var empty = !d.invites.length && !d.community.length && !d.conversions.length && !d.unparsed.length;
        if (empty) warnings.push('「' + c.name + '」下面没有识别到数据，已跳过');
        return !empty;
      });
    }
    if (selectedDate && cohorts.length > 1) {
      return { cohorts: [], warnings: ['你选择了本期直播日期，但粘贴内容识别出 ' + cohorts.length + ' 期。请只粘贴这一期的表格，或清空日期后再批量识别'] };
    }
    return { cohorts: cohorts, warnings: warnings };
  }

  // ================= 导出 =================

  return {
    CHANNELS: CHANNELS,
    CHANNEL_KEYS: CHANNEL_KEYS,
    INVITE_COLUMNS: INVITE_COLUMNS,
    COMMUNITY_COLUMNS: COMMUNITY_COLUMNS,
    CORE_METRICS: CORE_METRICS,
    TABLE_NAMES: TABLE_NAMES,
    columnsFor: columnsFor,
    conversionColumns: conversionColumns,
    parsePaste: parsePaste,
    normalizeDate: normalizeDate,
    normalizeInviteDate: normalizeInviteDate,
    isInviteDate: isInviteDate,
    fmtInviteDate: fmtInviteDate,
    parseNum: parseNum,
    parseNumInfo: parseNumInfo,
    emptyCohort: emptyCohort,
    emptyRow: emptyRow,
    normalizeCohort: normalizeCohort,
    computeMetrics: computeMetrics,
    compareCohort: compareCohort,
    buildInsights: buildInsights,
    validateCohort: validateCohort,
    validateCohortDetailed: validateCohortDetailed,
    weekStart: weekStart,
    weekLabel: weekLabel,
    groupByWeek: groupByWeek,
    fmtPct: fmtPct,
    fmtInt: fmtInt,
    fmtDelta: fmtDelta,
    fmtDateCN: fmtDateCN,
    weekdayCN: weekdayCN,
    channelLabel: channelLabel,
    isISODate: isISODate,
  };
});
