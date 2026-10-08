/* 复盘洞察：只按已填数据作算术拆解和可验证建议，不推测成交原因。 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./logic.js'));
  else root.CampInsights = factory(root.CampLogic);
})(typeof window !== 'undefined' ? window : this, function (L) {
  'use strict';

  var LABELS = {
    deposit: '已付定金', deposit_refund: '已退定金', balance: '已付尾款',
    live_full: '直播全款成交', direct: '1v1 追单', alumni: '老学员', refund: '退款',
  };
  var FIELDS = Object.keys(LABELS);
  function hasValue(value) { return value != null && String(value).trim() !== ''; }
  function count(value) {
    var n = L.parseNum(value);
    return typeof n === 'number' && isFinite(n) && n >= 0 && Number.isInteger(n) ? n : null;
  }
  function rows(data, kind) { return Array.isArray(data[kind]) ? data[kind] : []; }
  function number(value) { return L.fmtInt(Math.round(value * 10) / 10); }
  function signed(value) { return (value > 0 ? '+' : value < 0 ? '−' : '') + number(Math.abs(value)); }
  function pct(value) { return L.fmtPct(value); }
  function parsedScope(current, previous) {
    return current.unparsed || (previous && previous.unparsed)
      ? '这项分析仅基于已识别数据，不代表完整业绩；补齐未识别行后再复核。' : '';
  }
  function today() {
    var d = new Date();
    return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  }

  // 每一条实际成交行都必须填该字段；日期之外整行留空的编辑占位行不计。
  // 不能用 computeMetrics 的 sums 判缺失：它按报表口径把 null 加成 0。
  function inspect(cohort) {
    var c = cohort || {}, data = c.data || {};
    var isX = c.channel === 'xiaoe', liveKey = isX ? 'balance' : 'live_full';
    var conversionRows = rows(data, 'conversions').filter(function (r) {
      return r && typeof r === 'object' && (hasValue(r.date) || FIELDS.some(function (f) { return hasValue(r[f]); }));
    });
    var sums = {}, missing = [];
    FIELDS.forEach(function (f) {
      var total = 0, complete = conversionRows.length > 0;
      conversionRows.forEach(function (r) {
        var value = count(r[f]);
        if (value === null) complete = false;
        else total += value;
      });
      sums[f] = complete ? total : null;
    });
    var required = isX ? ['deposit', 'deposit_refund', 'balance', 'direct', 'alumni'] : ['live_full', 'direct', 'alumni', 'refund'];
    required.forEach(function (f) { if (sums[f] === null) missing.push(LABELS[f]); });

    var community = rows(data, 'community');
    var attendanceRows = community.filter(function (r) { return r && hasValue(r.attendance) && L.parseNum(r.attendance) !== null; });
    var invalidAttendance = community.some(function (r) {
      return r && hasValue(r.attendance) && L.parseNum(r.attendance) !== null && count(r.attendance) === null;
    });
    var badAttendanceText = community.some(function (r) {
      return r && L.parseNumInfo(r.attendance).bad;
    });
    var attendance = null, issues = [];
    if (invalidAttendance || badAttendanceText) issues.push('出勤人数含非整数、负数或无法识别的内容');
    else if (attendanceRows.length > 1) issues.push('有多行出勤数据，尚不能确认用于拆解的同一场直播人数');
    else if (attendanceRows.length === 1) {
      var liveRow = attendanceRows[0];
      attendance = count(liveRow.attendance);
      if (!L.isISODate(liveRow.date) || liveRow.date !== c.live_date) {
        issues.push('出勤记录日期与本期直播日未对齐');
        attendance = null;
      } else if (count(liveRow.members) !== null && attendance > count(liveRow.members)) {
        issues.push('出勤人数大于社群人数，请确认是否包含群外观众');
        attendance = null;
      }
    } else issues.push('缺少可用的直播出勤人数');
    if (attendance === 0) issues.push('直播出勤为 0，不能计算转化效率');

    var unparsed = rows(data, 'unparsed').length;
    if (unparsed) issues.push('仍有 ' + unparsed + ' 行未识别数据');
    var live = sums[liveKey];
    if (attendance !== null && live !== null && live > attendance) {
      issues.push('直播成交超过出勤人数，需核对人数与订单归属');
      attendance = null;
    }
    var grossKnown = [live, sums.direct, sums.alumni].every(function (n) { return n !== null; });
    var gross = grossKnown ? live + sums.direct + sums.alumni : null;
    var net = gross !== null && (isX || sums.refund !== null) ? gross - (isX ? 0 : sums.refund) : null;
    if (net !== null && net < 0) { issues.push('退款人数超过退款前成交人数'); net = null; }
    var pending = null, netDeposit = null, invalidLive = false;
    if (isX && [sums.deposit, sums.deposit_refund, sums.balance].every(function (n) { return n !== null; })) {
      netDeposit = sums.deposit - sums.deposit_refund;
      if (netDeposit < 0 || sums.balance > netDeposit) {
        issues.push('定金、退定金和尾款人数无法对齐');
        invalidLive = true;
        net = null;
      }
      else pending = netDeposit - sums.balance;
    }
    return {
      cohort: c, isX: isX, sums: sums, live: live, attendance: attendance,
      rate: attendance > 0 && live !== null && !invalidLive ? live / attendance : null,
      gross: gross, net: net, pending: pending, netDeposit: netDeposit,
      missing: missing, issues: issues, unparsed: unparsed,
    };
  }

  function buildInsights(cohort, allCohorts) {
    var c = cohort || {}, out = [];
    function add(level, title, text) { out.push({ level: level, title: title, text: text }); }
    if (!L || !L.computeMetrics) {
      add('warn', '分析依赖尚未就绪', '基础指标模块未加载，暂时无法核对数据。请刷新页面后重试。');
      return out;
    }
    if (!L.CHANNELS[c.channel] || !L.isISODate(c.live_date)) {
      add('warn', '先确认渠道和直播日期', '当前缺少有效的渠道或直播日期，无法确定可比期次。请补齐后再看历史拆解。');
      return out;
    }
    var now = today();
    if (c.live_date > now) {
      add('info', '这期直播尚未发生', '直播日期为 ' + c.live_date + '，目前记录不能视为已实现业绩。直播后核对出勤、成交和退款，再复查分析。');
      return out;
    }

    var current = inspect(c);
    var history = (Array.isArray(allCohorts) ? allCohorts : []).filter(function (item) {
      return item && item !== cohort && (c.id == null || item.id !== c.id) && item.channel === c.channel &&
        L.isISODate(item.live_date) && item.live_date < c.live_date && item.live_date <= now;
    }).sort(function (a, b) { return a.live_date < b.live_date ? 1 : a.live_date > b.live_date ? -1 : 0; });
    var previous = history.length ? inspect(history[0]) : null;
    var mature = history.filter(function (item) { return item.closed === true; }).map(inspect).filter(function (item) {
      return item.rate !== null && !item.issues.length && !item.unparsed;
    }).slice(0, 4);

    // 1. 把会限制解读的事实放在前面；30 人仅是展示提醒线，不是显著性标准。
    var caveats = current.issues.slice();
    if (current.missing.length) caveats.push('成交表的「' + current.missing.join('、') + '」有空白或无效人数');
    var unusablePrevious = previous && current.rate !== null && previous.rate === null;
    if (unusablePrevious) caveats.push('上期（' + previous.cohort.live_date + '）的出勤或直播成交尚不可比，暂不能拆解与上期的人数差额');
    if (!c.closed) caveats.push('本期尚未收口，当前成交仍可能因尾款和退款变化');
    if (current.attendance > 0 && current.attendance < 30) {
      caveats.push('出勤只有 ' + number(current.attendance) + ' 人，单个成交就会改变 ' + number(100 / current.attendance) + ' 个百分点');
    }
    if (caveats.length) {
      var hasBlockingGaps = current.issues.length > (current.unparsed ? 1 : 0) || current.missing.length || unusablePrevious;
      add('warn', hasBlockingGaps ? '先核对数据，再判断受影响指标' : current.unparsed ? '仍有未识别行，先看已识别数据' :
        !c.closed ? '本期仍在收口，先看过程' : '小样本下，单个成交就会改变判断',
        caveats.join('；') + '。' + (current.unparsed ?
          '后续分析仅基于已识别数据，不代表完整业绩；请先确认未识别行的含义，补齐后再复核。' : '') +
        (hasBlockingGaps ? '空白不会当作 0，缺失或矛盾字段影响的拆解暂不输出；请核实人数，确认确实没有发生的项目再填 0。' :
          !current.unparsed ? '这些数值尚不足以证明稳定的运营变化；约定下一次复查时间，并对照同样收款进度的期次。' : ''));
    }

    // 2. 已付定金的人群是可核对的跟进名单，不把所有未补者当作必然可追回。
    if (current.isX && current.pending > 0) {
      add(c.closed ? 'info' : 'warn', c.closed ? '已结束期仍有未补尾款，先核对结局' : '先核对 ' + number(current.pending) + ' 位待补尾款',
        '已付定金 ' + number(current.sums.deposit) + ' 人，退定金 ' + number(current.sums.deposit_refund) + ' 人，净定金 ' +
        number(current.netDeposit) + ' 人中已补尾款 ' + number(current.sums.balance) + ' 人，剩余 ' + number(current.pending) +
        ' 人（' + pct(current.pending / current.netDeposit) + '）。未补不等于一定能追回。' + (c.closed ?
          '先把名单区分为确认流失、待退款和仍在跟进，并核对“已结束”标记。' :
          '优先核对这批人的付款与沟通记录，为每人补上下一次跟进时间和结果，再验证回收人数。') + parsedScope(current));
    }

    // 3. 对称分解 Δ(A×r)：规模项 ΔA×平均r，效率项 Δr×平均A；无顺序偏差。
    if (previous && current.rate !== null && previous.rate !== null) {
      var delta = current.live - previous.live;
      var scale = (current.attendance - previous.attendance) * (current.rate + previous.rate) / 2;
      var efficiency = (current.rate - previous.rate) * (current.attendance + previous.attendance) / 2;
      if (Math.abs(scale) > 0.05 || Math.abs(efficiency) > 0.05) {
        var stage = !c.closed || previous.cohort.closed !== true;
        var dominant = Math.abs(scale) >= Math.abs(efficiency) ? '出勤规模项' : '转化效率项';
        add(stage ? 'info' : delta > 0 ? 'good' : delta < 0 ? 'bad' : 'info',
          delta === 0 ? '出勤规模与转化效率的变化互相抵消' : Math.abs(Math.abs(scale) - Math.abs(efficiency)) < 0.05 ?
            '规模与效率对成交差额的贡献相当' : '直播成交差额中，' + dominant + '更大',
          '与上期（' + previous.cohort.live_date + '）相比，' + (current.isX ? '补尾款成交' : '直播全款成交') + '从 ' +
          number(previous.live) + ' 人变为 ' + number(current.live) + ' 人，差额 ' + signed(delta) + ' 人；出勤从 ' +
          number(previous.attendance) + ' 人变为 ' + number(current.attendance) + ' 人，转化率从 ' + pct(previous.rate) + ' 变为 ' + pct(current.rate) +
          '。按两期平均值对称拆解，规模项约 ' + signed(scale) + ' 人，效率项约 ' + signed(efficiency) +
          ' 人。这是算术贡献，不能证明原因；' + (stage ? '至少一期尚未收口，差额只是当前进度。' : '两期均已标记结束，仍需核对人群和记录口径。') +
          (dominant === '出勤规模项' ? '下一步按邀约来源核对实际到课人数，验证规模变化出现在哪个来源。' :
            '下一步对齐订单归属和收款进度，再按来源核对到课者成交，验证效率差异。') + parsedScope(current, previous));
      }
    }

    // 4. 历史基线只用当前期之前的同渠道成熟期，人数加权；机会是情景差额而非预测。
    if (current.rate !== null && mature.length >= 2) {
      var totalAttendance = 0, totalLive = 0;
      mature.forEach(function (item) { totalAttendance += item.attendance; totalLive += item.live; });
      var baseline = totalLive / totalAttendance;
      var scenario = current.attendance * baseline, gap = scenario - current.live;
      if (Math.abs(gap) >= 1) {
        add('info', gap > 0 ? '按成熟期参考测算，仍有情景差额' : '当前直播成交高于成熟期参考',
          '最近 ' + mature.length + ' 个已结束且可用的同渠道历史期，直播成交合计 ' + number(totalLive) + ' 人 / 出勤合计 ' +
          number(totalAttendance) + ' 人，加权转化率 ' + pct(baseline) + '。按本期 ' + number(current.attendance) +
          ' 人出勤测算，对应约 ' + number(scenario) + ' 人成交，较当前 ' + number(current.live) + ' 人' +
          (gap > 0 ? '多约 ' : '少约 ') + number(Math.abs(gap)) + ' 人。样本和人群可能不同，这是情景参考，不是承诺或因果结论；' +
          (!c.closed ? '本期尚未收口，差额还会变化。' : '') + (current.isX ?
            '分别列出下一期目标和本期可跟进尾款，别把历史差额直接当作可追回人数。' :
            '先核对来源与退款归属，再把这个情景作为下一期验证目标。') + parsedScope(current));
      }
    }

    // 5. B站退款无法归因到来源，构成占比必须用退款前分母，不能沿用 directShare。
    if (current.gross > 0 && current.net !== null &&
        (current.sums.direct / current.gross >= 0.3 || (!current.isX && current.sums.refund > 0))) {
      var share = current.sums.direct / current.gross;
      if (share >= 0.3) {
        add('info', '1v1 追单占成交的较大部分',
          '当前 1v1 追单 ' + number(current.sums.direct) + ' 人，占' + (current.isX ? '总成交' : '退款前成交') + ' ' +
          number(current.gross) + ' 人的 ' + pct(share) + '。' + (current.isX ?
            '这说明成交结构包含较多后续追单，不能仅凭占比判断效率好坏。' :
            '退款 ' + number(current.sums.refund) + ' 人后净成交 ' + number(current.net) + ' 人；退款尚未按来源分摊，不能据此计算追单净转化。') +
          (!c.closed ? '本期仍在收口，占比还可能变化。' : '') + '下一期记录追单工时、订单来源和退款归属，验证这部分成交能否稳定复现。' + parsedScope(current));
      } else {
        add(current.net === 0 ? 'warn' : 'info', '退款影响净成交，先核对来源',
          '退款前成交 ' + number(current.gross) + ' 人，退款 ' + number(current.sums.refund) + ' 人（' +
          pct(current.sums.refund / current.gross) + '），净成交剩 ' + number(current.net) +
          ' 人。退款尚未按来源分摊，不能判断来自直播还是后续追单；' + (!c.closed ? '本期仍未收口，净成交还可能变化。' : '') +
          '先逐单核对退款归属、日期和原因分类，再验证净成交变化集中在哪类订单。' + parsedScope(current));
      }
    }

    if (!out.length) {
      add('info', '先积累可比记录，再判断变化',
        '本期已填直播成交 ' + number(current.live) + ' 人，出勤 ' + number(current.attendance) + ' 人；可用的已结束同渠道历史期有 ' + mature.length +
        ' 个。当前没有触发明显的规模、效率或成交结构差额，这不等于已证明稳定。继续保留来源、付款和跟进记录，下一期用同一口径复查。');
    }
    return out.slice(0, 5);
  }

  return { buildInsights: buildInsights };
});
