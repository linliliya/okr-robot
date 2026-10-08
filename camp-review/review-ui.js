/* 看板复盘：展示与编辑。计算和快照归一化由 review.js 负责。 */
(function () {
  'use strict';
  const R = window.CampReview;
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const clone = x => JSON.parse(JSON.stringify(x));
  const executions = { pending: '待确认执行', done: '已执行', partial: '部分执行', skipped: '未执行' };
  const decisions = { pending: '待复盘', keep: '继续采用', adjust: '调整后再试', stop: '停止采用', retry: '继续观察' };
  const metric = key => R.METRICS.find(m => m.key === key);
  const applicable = (c, key) => !!metric(key) && (c.channel !== 'bilibili' || key !== 'balanceRate');
  const metricLabel = key => metric(key)?.label || '未选择指标';
  const actionMeta = a => `${metricLabel(a.metric)}目标 ≥ ${R.formatMetric(a.metric, a.target)}${a.owner ? ' · 负责人：' + a.owner : ''}`;
  const phase = (n, title) => `<h3 class="review-phase"><span>${n}</span>${title}</h3>`;
  const empty = text => `<p class="review-empty">${esc(text)}</p>`;

  function evidence(c, item) {
    const result = R.evaluateAction(c, item);
    const comparable = result.status === 'ready' || result.status === 'provisional';
    const badge = !comparable ? (result.status === 'missing' ? '数据待补齐' : '待确认执行')
      : `${result.status === 'provisional' ? '暂时' : ''}${result.met ? '达到目标' : '未达目标'}`;
    return `<div class="action-evidence">
      <div><span>上期基线</span><b>${esc(R.formatMetric(item.action.metric, result.baseline))}</b></div>
      <div><span>本期数据</span><b>${esc(R.formatMetric(item.action.metric, result.current))}</b></div>
      <div><span>预定目标 ≥</span><b>${esc(R.formatMetric(item.action.metric, result.target))}</b></div>
    </div><p class="review-hint"><b>${esc(badge)}</b> · ${esc(result.text)}</p>`;
  }

  function sourceLabel(review) {
    const f = review.followup;
    return f ? `<p class="review-hint">承接 ${esc(f.sourceName || '上期')} · ${esc(f.sourceDate)} 的行动。基线按本次复盘保存时留存，本期数据随录入更新。</p>` : '';
  }

  function render(c, all) {
    const r = R.reviewContext(c, all);
    const prev = R.previousCohort(c, all);
    const history = R.lastLessons(c, all, 3);
    const items = r.followup?.items || [];
    const evals = items.filter(x => x.decision !== 'pending').length;
    return `<div class="review-head"><div><h3>复盘与行动</h3><p>上期行动 → 本期验证 → 经验沉淀 → 下期行动</p></div>
      <button class="btn small no-present" data-edit-review="${esc(c.id)}">编辑本期复盘</button></div>
      <div class="block review-followup">${phase('01', '上期行动，本期效果')}
        ${sourceLabel(r)}${items.length ? `<p class="review-hint">已复盘 ${evals} / ${items.length} 项。指标达标只说明结果，是否由行动带来，需要结合执行记录判断。</p>
        <div class="review-cards">${items.map(item => `<article class="action-card">
          <h4>${esc(item.action.title)}</h4><p class="review-hint">${esc(actionMeta(item.action))}</p>
          ${item.action.hypothesis ? `<p class="action-copy">原判断：${esc(item.action.hypothesis)}</p>` : ''}
          ${evidence(c, item)}<div class="action-tags"><span>${esc(executions[item.execution])}</span><span>${esc(decisions[item.decision])}</span></div>
          ${item.note ? `<p class="action-copy">${esc(item.note)}</p>` : empty('补充执行证据、数据效果及判断，再决定是否继续。')}
        </article>`).join('')}</div>` : empty(prev ? '上期还没有记录可跟踪的行动。本期先写下行动、目标，下期会自动带出。' : '这是该渠道的首期记录。从本期制定行动，下期开始验证。')}
      </div>
      <div class="panel-row two">
        <div class="block review-lessons">${phase('02', '本期经验沉淀')}
          ${r.lessons ? `<div class="summary-text">${esc(r.lessons)}</div>` : empty('记录什么做法值得保留、适用条件是什么，以及哪类尝试不再重复。')}
          ${history.length ? `<details class="lesson-library"><summary>回看同渠道近 ${history.length} 期经验</summary>${history.map(h => `<article><a href="#/?week=${esc(window.CampLogic.weekStart(h.live_date))}&ch=${esc(c.channel)}">${esc(h.name)} · ${esc(h.live_date)}</a><div class="summary-text">${esc(h.lessons)}</div></article>`).join('')}</details>` : ''}
        </div>
        <div class="block review-plans">${phase('03', '下期准备做什么')}
          ${r.actions.length ? `<div class="review-cards">${r.actions.map(a => `<article class="action-card"><h4>${esc(a.title)}</h4><p class="review-hint">${esc(actionMeta(a))}</p>${a.hypothesis ? `<p class="action-copy">${esc(a.hypothesis)}</p>` : ''}</article>`).join('')}</div>` : empty('把动作写具体，并选一个能验证它的指标和目标值。下一期同渠道复盘时会自动承接。')}
        </div>
      </div>`;
  }

  function edit(slot, c, all, handlers) {
    const draft = clone(R.reviewContext(c, all));
    let summary = c.summary || '';
    let saving = false;
    const options = (map, selected) => Object.entries(map).map(([k, v]) => `<option value="${esc(k)}" ${selected === k ? 'selected' : ''}>${esc(v)}</option>`).join('');
    const newId = () => window.crypto?.randomUUID?.() || `action-${Date.now()}-${Math.random().toString(36).slice(2)}`;
    const paint = () => {
      slot.innerHTML = `<form class="review-editor" novalidate>
        <div class="review-head"><div><h3>编辑本期复盘</h3><p>先核对执行和数据，再沉淀经验、制定下一期行动。</p></div></div>
        <fieldset><label class="review-field">本期复盘结论<textarea class="summary-input" data-review-summary rows="3" placeholder="这期最重要的判断是什么？">${esc(summary)}</textarea></label>
          <div class="block">${phase('01', '上期行动，本期效果')}${sourceLabel(draft)}
          ${draft.followup?.items.length ? draft.followup.items.map((item, i) => `<article class="action-card review-evaluation" data-evaluation="${i}">
            <h4>${esc(item.action.title)}</h4><p class="review-hint">${esc(actionMeta(item.action))}</p><div class="evaluation-evidence">${evidence(c, item)}</div>
            <div class="review-form-grid">
              <label class="review-field">执行情况<select data-eval-field="execution">${options(executions, item.execution)}</select></label>
              <label class="review-field">后续决定<select data-eval-field="decision">${options(decisions, item.decision)}</select></label>
            </div>
            <label class="review-field">执行证据与效果判断<textarea data-eval-field="note" rows="3" placeholder="具体做了什么？数据变化与原假设是否一致？还有哪些因素影响结果？为什么保留或调整？">${esc(item.note)}</textarea></label>
            <button class="btn small ghost" type="button" data-carry-action="${i}">沿用到下期行动</button>
          </article>`).join('') : empty('还没有可承接的上期行动。可以先记录本期经验和下期行动。')}</div>
          <div class="block">${phase('02', '本期经验沉淀')}
            <label class="review-field">可复用的做法与适用条件<textarea data-review-lessons rows="4" placeholder="例如：对已领课但未到场的人分层提醒；先在小鹅通渠道验证，样本还少，下期继续观察。">${esc(draft.lessons)}</textarea></label>
          </div>
          <div class="block">${phase('03', '下期准备做什么')}
            <p class="review-hint">目标按“达到或超过”填写；百分比填 35 表示 35%。负责人和行动假设可选。</p>
            <div class="review-cards">${draft.actions.map((a, i) => `<article class="action-card" data-action="${i}">
              <label class="review-field">行动 ${i + 1}<input data-action-field="title" value="${esc(a.title)}" placeholder="做什么、对谁做、什么时候做" maxlength="160"></label>
              <label class="review-field">为什么做 / 怎样验证<textarea data-action-field="hypothesis" rows="2" placeholder="根据哪项数据提出这个行动？预期改善什么？">${esc(a.hypothesis)}</textarea></label>
              <div class="review-form-grid three">
                <label class="review-field">观察指标<select data-action-field="metric">${!applicable(c, a.metric) ? '<option value="" selected>请选择适用于本渠道的指标</option>' : ''}${R.METRICS.filter(m => applicable(c, m.key)).map(m => `<option value="${m.key}" ${a.metric === m.key ? 'selected' : ''}>${esc(m.label)}</option>`).join('')}</select></label>
                <label class="review-field">目标值（${metric(a.metric)?.unit === '%' ? '%' : '人'}）<input data-action-field="target" type="number" min="0" step="${metric(a.metric)?.unit === '%' ? 'any' : '1'}" ${metric(a.metric)?.unit === '%' ? 'max="100"' : ''} value="${esc(a.target == null ? '' : metric(a.metric)?.unit === '%' ? +(a.target * 100).toFixed(6) : a.target)}"></label>
                <label class="review-field">负责人<input data-action-field="owner" value="${esc(a.owner)}" placeholder="谁来跟进" maxlength="80"></label>
              </div><button type="button" class="btn small ghost" data-remove-action="${i}">移除这条行动</button>
            </article>`).join('')}</div>
            <button class="btn small ghost" type="button" data-add-action>＋ 添加下期行动</button>
          </div>
        </fieldset>
        <div class="review-errors" role="alert" hidden></div>
        <div class="row-actions"><button class="btn" type="submit" data-save-review data-save-summary="${esc(c.id)}">保存本期复盘</button><button class="btn ghost" type="button" data-cancel-review data-cancel-summary="${esc(c.id)}">取消</button></div>
      </form>`;
    };
    const capture = event => {
      const t = event.target;
      if (saving) return;
      if (t.hasAttribute('data-review-summary')) summary = t.value;
      else if (t.hasAttribute('data-review-lessons')) draft.lessons = t.value;
      else if (t.dataset.evalField) {
        const card = t.closest('[data-evaluation]');
        const item = draft.followup.items[+card.dataset.evaluation];
        item[t.dataset.evalField] = t.value;
        if (t.dataset.evalField === 'execution') card.querySelector('.evaluation-evidence').innerHTML = evidence(c, item);
      }
      else if (t.dataset.actionField) {
        const a = draft.actions[+t.closest('[data-action]').dataset.action];
        const key = t.dataset.actionField;
        if (key === 'target') a.target = t.value === '' ? null : Number(t.value) / (metric(a.metric)?.unit === '%' ? 100 : 1);
        else if (key === 'metric') { a.metric = t.value; a.target = null; paint(); }
        else a[key] = t.value;
      } else return;
      handlers.onDirty();
    };
    slot.oninput = capture;
    slot.onchange = capture;
    slot.onclick = event => {
      const b = event.target.closest('button');
      if (!b || saving) return;
      if (b.hasAttribute('data-add-action')) {
        draft.actions.push({ id: newId(), title: '', hypothesis: '', owner: '', metric: 'liveDeals', target: null });
        handlers.onDirty(); paint(); slot.querySelector('[data-action]:last-child input')?.focus();
      } else if (b.hasAttribute('data-remove-action')) {
        draft.actions.splice(+b.dataset.removeAction, 1); handlers.onDirty(); paint();
      } else if (b.hasAttribute('data-carry-action')) {
        const a = draft.followup.items[+b.dataset.carryAction].action;
        draft.actions.push({ ...clone(a), id: newId() }); handlers.onDirty(); paint();
        slot.querySelector('[data-action]:last-child input')?.focus();
      } else if (b.hasAttribute('data-cancel-review')) handlers.onCancel();
    };
    slot.onsubmit = async event => {
      event.preventDefault();
      if (saving) return;
      const errors = R.validateReview(draft);
      draft.actions.forEach((a, i) => { if (!applicable(c, a.metric)) errors.push(`第 ${i + 1} 项行动：请选择适用于本渠道的指标`); });
      const box = slot.querySelector('.review-errors');
      if (errors.length) { box.innerHTML = errors.map(e => `<p>${esc(e)}</p>`).join(''); box.hidden = false; box.scrollIntoView({ block: 'nearest' }); return; }
      box.hidden = true;
      saving = true;
      slot.querySelector('fieldset').disabled = true;
      slot.querySelectorAll('.row-actions button').forEach(b => { b.disabled = true; });
      try { await handlers.onSave(clone(draft), summary); }
      finally {
        saving = false;
        if (slot.isConnected) {
          slot.querySelector('fieldset').disabled = false;
          slot.querySelectorAll('.row-actions button').forEach(b => { b.disabled = false; });
        }
      }
    };
    paint();
    slot.querySelector('textarea')?.focus();
  }
  window.CampReviewUI = { render, edit };
})();
