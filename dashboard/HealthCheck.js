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
  A5: { severity: 'warning', category: 'Sheet structure',
    title: 'No *Account column',
    action: 'Add it, or account-level reporting is impossible for this source.' },
  A6: { severity: 'info', category: 'Sheet structure',
    title: 'No Submitted_budget tab',
    action: 'No frozen record of what the funder was actually given.' },
  A7: { severity: 'error', category: 'Sheet structure',
    title: 'Forecast column header not recognised',
    action: 'Name it exactly "MMM-MMM YY Forecast", e.g. "Jul-Sep 26 Forecast".' },
  A8: { severity: 'warning', category: 'Sheet structure',
    title: 'Forecast row is not a milestone',
    action: 'Use "CODE - Name", matching the Xero Inventory Item on the Budget tab.' },
  A9: { severity: 'info', category: 'Sheet structure',
    title: 'Forecast for a milestone not in the budget',
    action: 'Either the budget line was removed, or the code is a typo.' },
  B1: { severity: 'warning', category: 'Data quality',
    title: 'Lines skipped as empty',
    action: 'Set Cost/Income deliberately, or delete the row.' },
  B2: { severity: 'error', category: 'Data quality',
    title: 'Unparseable date',
    action: 'Use DD/MMM/YY. The line is ignored entirely.' },
  B3: { severity: 'error', category: 'Data quality',
    title: 'End date precedes Start date',
    action: 'Check the year - "30/Jun/01" parses as 2001. The line is ignored.' },
  B4: { severity: 'error', category: 'Data quality',
    title: 'Lines with no Xero Inventory Item',
    action: 'Add the milestone code, or these fall outside quarterly tracking.' },
  B5: { severity: 'warning', category: 'Data quality',
    title: 'Contribution does not equal Income minus Cost',
    action: 'Recompute the Contribution column, or explain it in Comments.' },
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
    const allowed = {};
    allowed[CONFIG.BUDGET_TAB] = true;
    allowed[CONFIG.FORECAST_TAB] = true;
    allowed[CONFIG.SUBMITTED_TAB] = true;
    (b.tabs || []).forEach(t => {
      if (!allowed[t]) {
        add('A3', withBase_(base, { detail: 'tab "' + t + '" is not one of ' +
          CONFIG.BUDGET_TAB + ' / ' + CONFIG.FORECAST_TAB + ' / ' + CONFIG.SUBMITTED_TAB }));
      }
    });
    if ((b.tabs || []).length && (b.tabs || []).indexOf(CONFIG.SUBMITTED_TAB) === -1) {
      add('A6', withBase_(base, { detail: 'no "' + CONFIG.SUBMITTED_TAB + '" tab' }));
    }

    // --- A9: forecast for a milestone the budget no longer has ---
    const budgetItems = {};
    (b.lines || []).forEach(l => { if (l.item) budgetItems[l.item] = true; });
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
