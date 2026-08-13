/**
 * HealthCheck.js
 * Turns silent wrongness into named, owned, actionable findings.
 *
 * The cockpit's numbers are only as good as the budgets in Drive and the coding in
 * Xero, and both used to fail invisibly: BudgetReader skipped a malformed line
 * behind a Logger.log nobody reads, and an untagged Xero transaction vanished from
 * every total without trace. Every finding here states what is wrong, where, who
 * owns it and what to do.
 *
 * Check ids and severities match dashboard/HEALTH_CHECKS.md.
 *   error   - a number on the dashboard is wrong right now
 *   warning - a number may be wrong, or will be soon
 *   info    - hygiene; nothing is wrong yet
 *
 * Pure over its inputs - no Drive or Xero calls - so it can be exercised offline.
 */

const HEALTH_CATALOGUE = {
  A1: { severity: 'error', category: 'Sheet structure',
    title: 'Budget sheet could not be read',
    action: 'The entire file is invisible on the dashboard until this is fixed.' },
  A2: { severity: 'error', category: 'Sheet structure',
    title: 'Required column missing',
    action: 'Add the column. Without it the whole file is skipped.' },
  A3: { severity: 'warning', category: 'Sheet structure',
    title: 'Unexpected tab',
    action: 'Keep only Budget, Forecast and Submitted_budget. Actuals live in Xero.' },
  A4: { severity: 'warning', category: 'Sheet structure',
    title: 'No Project column',
    action: 'Add it so lines can be attributed across projects.' },
  // A5 (no *Account column) was retired on 2026-08-11. Budgets are set at
  // milestone level, not per account, so its absence is expected rather than a
  // finding - and nothing in the code read a budget line's account anyway.
  A6: { severity: 'info', category: 'Sheet structure',
    title: 'No Submitted_budget tab',
    action: 'No frozen record of what the funder was actually given.' },
  A7: { severity: 'error', category: 'Sheet structure',
    title: 'Forecast column header not recognised',
    action: 'Name it exactly "MMM-MMM YY Forecast", e.g. "Jul-Sep 26 Forecast".' },
  A8: { severity: 'warning', category: 'Sheet structure',
    title: 'Forecast row does not match a budget line',
    action: 'Label it with the milestone, or "Description - Milestone" from the ' +
      'Budget tab. The row\'s forecast is discarded until it resolves.' },
  A9: { severity: 'warning', category: 'Sheet structure',
    title: 'Forecast for a milestone not in the budget',
    action: 'Either the budget line was removed, or the code is a typo.' },
  A10: { severity: 'warning', category: 'Sheet structure',
    title: 'Two forecast rows share one label',
    action: 'Usually a sorted Budget tab: the label formulas now point at the ' +
      'wrong lines. Re-point them and do not sort the Budget tab.' },
  B1: { severity: 'warning', category: 'Data quality',
    title: 'Lines skipped as empty',
    action: 'Set Cost/Income deliberately, or delete the row.' },
  B2: { severity: 'error', category: 'Data quality',
    title: 'Unparseable date',
    action: 'Use DD/MMM/YY. The line is ignored entirely.' },
  B3: { severity: 'error', category: 'Data quality',
    title: 'End date precedes Start date',
    action: 'Check the year: a two-digit year can land in the wrong century. ' +
      'The line is ignored.' },
  B4: { severity: 'error', category: 'Data quality',
    title: 'Lines with no Xero Inventory Item',
    action: 'Add the milestone code, or these fall outside quarterly tracking.' },
  B5: { severity: 'warning', category: 'Data quality',
    title: 'Contribution does not equal Income minus Cost',
    action: 'Recompute the Contribution column, or explain it in Comments.' },
  G1: { severity: 'warning', category: 'Funding',
    title: 'Secured funding exceeds the budgeted cost',
    action: 'Either two applications for the same work both landed, in which case ' +
      'reallocate the surplus, or income is filed against the wrong milestone.' },
  G2: { severity: 'info', category: 'Funding',
    title: 'Proposed source with no Probability',
    action: 'Add Probability to Funding_info (0-100). Without it this ask is left out ' +
      'of expected income entirely, rather than guessed at.' },
  G3: { severity: 'info', category: 'Funding',
    title: 'Cost not counted, a competing application carries it',
    action: 'Expected: one exclusivity group is one piece of work. Remove the ' +
      'Exclusivity group value if these are genuinely separate work.' },
  D1: { severity: 'error', category: 'Xero coding',
    title: 'Actual spend with no Projects tag',
    action: 'Tag these in Xero. Untagged lines are dropped from every total.' },
  D2: { severity: 'warning', category: 'Xero coding',
    title: 'Actual spend with no Funding source tag',
    action: 'Tag these in Xero, or they land in "(unassigned)".' },
  F1: { severity: 'error', category: 'System',
    title: 'Xero is not connected',
    action: 'Run logXeroAuthUrl and authorise. Actuals are stale until you do.' },
  F3: { severity: 'error', category: 'System',
    title: 'Xero credentials not configured',
    action: 'Set them in Project Settings > Script Properties.' },
  F5: { severity: 'info', category: 'System',
    title: 'Refresh summary',
    action: '' }
};

const HEALTH_SEVERITY_RANK = { error: 0, warning: 1, info: 2 };

/**
 * @param {Array} budgets  from readAllBudgets()
 * @param {Array} actualLines  normalised Xero lines
 * @param {Object} ctx  { xeroConnected, exclusion: {count,total}, secretsMissing: [] }
 * @return {Array} findings, most severe first, and by value at risk within severity
 */
function buildHealth(budgets, actualLines, ctx) {
  ctx = ctx || {};
  const out = [];

  function add(id, fields) {
    const spec = HEALTH_CATALOGUE[id];
    if (!spec) return; // never let an unknown id break the refresh
    out.push({
      id: id, severity: spec.severity, category: spec.category, title: spec.title,
      action: spec.action,
      detail: fields.detail || '',
      fundingSource: fields.fundingSource || '',
      project: fields.project || '',
      owner: fields.owner || '',
      row: fields.row || null,
      amount: fields.amount || 0,
      link: fields.link || ''
    });
  }

  (budgets || []).forEach(b => {
    const meta = b.metadata || {};
    const base = { fundingSource: b.name, project: b.projectFolder,
      owner: meta['owner'] || '', link: b.sheetUrl || '' };

    // --- findings the reader already collected ---
    (b.issues || []).forEach(iss => {
      // A missing required column arrives as a parse failure; it is really A2.
      const id = (iss.check === 'A1' && /missing column/i.test(iss.detail || '')) ? 'A2' : iss.check;
      const fields = { detail: iss.detail, row: iss.row };
      for (var k in base) fields[k] = base[k];
      // Value at risk for missing milestone codes is worth more than a count.
      if (id === 'B4') fields.amount = sumCostWhere_(b.lines, l => !l.item);
      add(id, fields);
    });

    // --- A3 / A6: tab hygiene ---
    const allowedTabs = [CONFIG.FUNDING_INFO_TAB, CONFIG.BUDGET_TAB,
      CONFIG.FORECAST_TAB, CONFIG.SUBMITTED_TAB];
    const allowed = {};
    allowedTabs.forEach(t => { allowed[t] = true; });
    (b.tabs || []).forEach(t => {
      if (!allowed[t]) {
        add('A3', withBase_(base, { detail: 'tab "' + t + '" is not one of ' +
          allowedTabs.join(' / ') }));
      }
    });
    if ((b.tabs || []).length && (b.tabs || []).indexOf(CONFIG.SUBMITTED_TAB) === -1) {
      add('A6', withBase_(base, { detail: 'no "' + CONFIG.SUBMITTED_TAB + '" tab' }));
    }

    // --- A9: forecast for a milestone the budget no longer has ---
    // Compare codes with codes. This used to key on the raw Inventory Item cell,
    // which holds "CODE - Name", while forecast keys are the bare code - so it
    // fired on every correctly coded sheet in the org and never on a real fault.
    // Since parseForecastTab_ now resolves labels against the Budget tab before
    // storing them, this should be unreachable: if it ever appears, a forecast key
    // reached the snapshot without going through resolveForecastLabel_.
    const budgetItems = {};
    (b.lines || []).forEach(l => {
      const code = itemCode_(l.item);
      if (code) budgetItems[code] = true;
    });
    const forecastItems = {};
    ['cost', 'income'].forEach(kind => {
      const map = (b.forecast && b.forecast[kind]) || {};
      Object.keys(map).forEach(key => { forecastItems[String(key).split('||')[0]] = true; });
    });
    Object.keys(forecastItems).forEach(code => {
      if (!budgetItems[code]) {
        add('A9', withBase_(base, { detail: 'Forecast references "' + code +
          '", which has no line on the Budget tab' }));
      }
    });

    // --- B5: contribution arithmetic ---
    var b5 = 0;
    (b.lines || []).forEach(l => {
      if (Math.abs((l.income - l.cost) - l.contribution) > 1) b5++;
    });
    if (b5) {
      add('B5', withBase_(base, { detail: b5 +
        ' line(s) where Contribution does not equal Income minus Cost' }));
    }

    // --- E1: secured income beyond the work it pays for ---
    // Two applications for the same thing both landing is a good problem, but it has to
    // surface or the surplus is never reallocated. The same check catches income filed
    // against the wrong milestone, and projected revenue misfiled as secured.
    if (b.status === 'secured') {
      var cost = 0, income = 0;
      (b.lines || []).forEach(l => { cost += l.cost || 0; income += l.income || 0; });
      if (income - cost > 1) {
        add('G1', withBase_(base, { amount: Math.round(income - cost),
          detail: 'secured income ' + Math.round(income) + ' exceeds budgeted cost ' +
            Math.round(cost) + ' by ' + Math.round(income - cost) }));
      }
    }

    // --- E2: a proposed ask with no stated probability ---
    if (b.status === 'proposed' &&
        sourceProbability_(b.status, meta) === null &&
        (b.lines || []).length) {
      add('G2', withBase_(base, { detail: 'no "' + CONFIG.META.probability +
        '" in Funding_info, so this ask is absent from expected income' }));
    }

    // --- E3: cost suppressed because a competing application carries it ---
    var supp = ((ctx || {}).exclusivity || {}).suppressed || {};
    if (supp[b.name]) {
      add('G3', withBase_(base, { detail: 'exclusivity group "' + supp[b.name].group +
        '": the cost is counted on ' + supp[b.name].countedIn + ' instead, so this ' +
        'sheet contributes its ask but not the work behind it' }));
    }
  });

  // --- D1 / D2: Xero coding. An untagged line is silently dropped from totals. ---
  var noProject = 0, noProjectTotal = 0, noSource = 0, noSourceTotal = 0;
  (actualLines || []).forEach(l => {
    if (l.kind !== 'expense') return;
    if (!l.project) { noProject++; noProjectTotal += Number(l.amount) || 0; }
    else if (!l.fundingSource) { noSource++; noSourceTotal += Number(l.amount) || 0; }
  });
  if (noProject) {
    add('D1', { detail: noProject + ' expense line(s) carry no Projects tracking value, ' +
      'so they are excluded from every project and organisation total',
      amount: Math.round(noProjectTotal) });
  }
  if (noSource) {
    add('D2', { detail: noSource + ' expense line(s) carry no Funding source value, ' +
      'so they fall outside the quarterly tracking grid',
      amount: Math.round(noSourceTotal) });
  }

  // --- F: system ---
  if (ctx.xeroConnected === false) {
    add('F1', { detail: 'actuals shown are whatever was last cached' });
  }
  (ctx.secretsMissing || []).forEach(k => {
    add('F3', { detail: 'Script Property "' + k + '" is not set' });
  });
  const excl = ctx.exclusion || { count: 0, total: 0 };
  add('F5', { detail: (budgets || []).length + ' funding source(s) read, ' +
    countLines_(budgets) + ' budget line(s) parsed, ' +
    (actualLines || []).length + ' actual line(s) after excluding ' + excl.count +
    ' on balance-sheet accounts (' + excl.total + ')' });

  out.sort(function (a, b) {
    const s = HEALTH_SEVERITY_RANK[a.severity] - HEALTH_SEVERITY_RANK[b.severity];
    if (s !== 0) return s;
    return (b.amount || 0) - (a.amount || 0); // value at risk first within severity
  });
  return out;
}

/** Short strings for the legacy dataFlags list, errors and warnings only. */
function healthToFlags(health) {
  return (health || [])
    .filter(f => f.severity !== 'info')
    .map(f => (f.severity === 'error' ? '[error] ' : '[warning] ') +
      (f.fundingSource ? f.fundingSource + ': ' : '') + f.title +
      (f.detail ? ' - ' + f.detail : ''));
}

function withBase_(base, fields) {
  const o = {};
  for (var k in base) o[k] = base[k];
  for (var j in fields) o[j] = fields[j];
  return o;
}

function sumCostWhere_(lines, pred) {
  var t = 0;
  (lines || []).forEach(l => { if (pred(l)) t += Number(l.cost) || 0; });
  return Math.round(t);
}

function countLines_(budgets) {
  var n = 0;
  (budgets || []).forEach(b => { n += ((b.lines || []).length); });
  return n;
}
