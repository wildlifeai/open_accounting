/**
 * Aggregator.js
 * Joins parsed budgets (from BudgetReader) with Xero actuals (from XeroClient)
 * and produces the snapshot the dashboard renders. This is where the confirmed
 * definitions live:
 *
 *   Secured forecast  = budgets in `secured/` folders only.
 *   Proposed forecast = `secured/` + `proposed/`.
 *   Prioritisation    = shortfall (budget − secured income) AND timing
 *                       (month secured income stops covering forecast spend).
 *
 * Project-level attribution of a funding source's expense uses each project's
 * day-weighted share of that source's operating expense; income and overhead are
 * allocated by the same share. Actuals come pre-split by Xero's Projects tracking
 * category, so actual numbers are exact, not approximated.
 */

function buildSnapshot() {
  const budgets = readAllBudgets();
  const actualSince = earliestBudgetStart_(budgets);
  const actualLines = isXeroConnected() ? fetchXeroActuals(actualSince) : [];

  const projects = {}; // name -> rollup accumulator
  const fundingSources = []; // per-source summary
  const dataFlags = [];
  const budgetByProjSrc = {}; // 'project||source' -> { budget, income, status }
  const actualByProjSrc = {}; // 'project||source' -> actual expense
  const sourceStatus = {};    // source name -> 'secured' | 'proposed'

  function project_(name) {
    if (!projects[name]) {
      projects[name] = { name: name, proposedBudget: 0, securedIncome: 0,
        proposedIncome: 0, actualExpense: 0 };
    }
    return projects[name];
  }

  budgets.forEach(src => {
    // Per-line project attribution is handled in BudgetReader: a line uses its
    // `Project` value when set, otherwise the parent project folder. Falling back
    // to the folder is expected, so it is not flagged here.
    const fc = computeFundingSourceForecast(src.lines);
    const shares = projectExpenseShares_(src.lines); // {project: 0..1}

    Object.keys(shares).forEach(pName => {
      const share = shares[pName];
      const p = project_(pName);
      const expTotal = fc.totalBudgetExpense * share;
      const incTotal = fc.totalBudgetIncome * share;

      // Proposed budget = secured + proposed; secured income only on secured sources.
      p.proposedBudget += expTotal;
      p.proposedIncome += incTotal;
      if (src.status === 'secured') p.securedIncome += incTotal;

      budgetByProjSrc[pName + '||' + src.name] =
        { budget: expTotal, income: incTotal, status: src.status };
    });

    sourceStatus[src.name] = src.status;
    fundingSources.push({ name: src.name, status: src.status,
      project: src.projectFolder, budgetExpense: fc.totalBudgetExpense,
      budgetIncome: fc.totalBudgetIncome });
  });

  // Actuals from Xero, split by project AND funding-source tracking categories.
  actualLines.forEach(l => {
    if (l.kind !== 'expense') return;
    if (!l.project || startsWith_(l.project, CONFIG.ARCHIVE_PREFIX)) return;
    project_(l.project).actualExpense += l.amount;
    const fs = l.fundingSource && !startsWith_(l.fundingSource, CONFIG.ARCHIVE_PREFIX)
      ? l.fundingSource : '(unassigned)';
    const akey = l.project + '||' + fs;
    actualByProjSrc[akey] = (actualByProjSrc[akey] || 0) + l.amount;
  });

  const breakdownRows = buildBreakdownRows_(budgetByProjSrc, actualByProjSrc, sourceStatus);

  // Quarterly tracking grid (baseline + actual per funding source / milestone /
  // quarter). Forecast is layered on at view time from the live Forecast sheet,
  // so GM edits show immediately without a full refresh.
  const tracking = buildTracking_(budgets, actualLines);

  // Per-project rollups (feed the summary cards / org totals).
  const rows = Object.keys(projects).map(name => {
    const p = projects[name];
    return {
      project: name,
      proposedBudget: round_(p.proposedBudget),
      securedIncome: round_(p.securedIncome),
      actualExpense: round_(p.actualExpense),
      unsecuredGap: round_(Math.max(0, p.proposedBudget - p.securedIncome))
    };
  });

  return {
    generatedAt: new Date().toISOString(),
    xeroConnected: isXeroConnected(),
    currentQuarter: currentQuarterLabel(),
    totals: orgTotals_(rows),
    projects: rows,
    fundingSources: fundingSources,
    breakdownRows: breakdownRows,
    tracking: tracking,
    dataFlags: dataFlags
  };
}

/**
 * Finest-grain rows for the Overview breakdown: one per (project, funding source)
 * with budget (forecast spend), secured funding (income on secured sources only),
 * and actual spend. The UI groups these by any combination of project / funding
 * source / status. Built from the union of budget and actual keys so spend that
 * has no matching budget line still shows up.
 */
function buildBreakdownRows_(budgetByProjSrc, actualByProjSrc, sourceStatus) {
  const keys = {};
  Object.keys(budgetByProjSrc).forEach(k => (keys[k] = true));
  Object.keys(actualByProjSrc).forEach(k => (keys[k] = true));

  return Object.keys(keys).map(key => {
    const parts = key.split('||');
    const project = parts[0];
    const source = parts[1];
    const b = budgetByProjSrc[key] || { budget: 0, income: 0 };
    const status = (b.status || sourceStatus[source] || 'unknown');
    return {
      project: project,
      fundingSource: source,
      status: status,
      budget: round_(b.budget || 0),
      secured: round_(status === 'secured' ? (b.income || 0) : 0),
      actual: round_(actualByProjSrc[key] || 0)
    };
  });
}

/**
 * Per funding source: baseline (frozen budget) and actual (from Xero) spend,
 * bucketed by milestone (item code) and quarter. Returns an array ready for the
 * quarterly tracking screen; the live forecast layer is merged in WebApp/UI.
 */
function buildTracking_(budgets, actualLines) {
  // Index actuals: source||itemCode -> { quarter: amount }
  const actualBySourceItem = {};
  actualLines.forEach(l => {
    if (l.kind !== 'expense' || !l.fundingSource) return;
    if (startsWith_(l.fundingSource, CONFIG.ARCHIVE_PREFIX)) return;
    const q = quarterOfMonthKey_(DateMath.monthKey(new Date(l.date)));
    const key = l.fundingSource + '||' + itemCode_(l.item);
    (actualBySourceItem[key] = actualBySourceItem[key] || {});
    actualBySourceItem[key][q] = (actualBySourceItem[key][q] || 0) + l.amount;
  });

  return budgets.map(src => {
    // Group budget lines by milestone (item code).
    const byItem = {};
    src.lines.forEach(l => {
      const code = itemCode_(l.item) || ('(' + (l.milestone || 'unassigned') + ')');
      byItem[code] = byItem[code] || { milestone: l.milestone || code, lines: [] };
      byItem[code].lines.push(l);
    });

    const quarterSet = {};
    const milestones = Object.keys(byItem).map(code => {
      const baseline = roundMap_(bucketToQuarters(distributeByMonth_(byItem[code].lines, 'cost')));
      const actual = roundMap_(actualBySourceItem[src.name + '||' + code] || {});
      Object.keys(baseline).forEach(q => (quarterSet[q] = true));
      Object.keys(actual).forEach(q => (quarterSet[q] = true));
      return { item: code, milestone: byItem[code].milestone, baseline: baseline, actual: actual };
    });

    const quarters = Object.keys(quarterSet).sort((a, b) => quarterSortNum(a) - quarterSortNum(b));
    return { source: src.name, status: src.status, project: src.projectFolder,
      quarters: quarters, milestones: milestones };
  });
}

function roundMap_(map) {
  const out = {};
  Object.keys(map).forEach(k => (out[k] = Math.round(map[k])));
  return out;
}

// ---- helpers ---------------------------------------------------------------

function projectExpenseShares_(lines) {
  const totals = {};
  let grand = 0;
  lines.forEach(l => {
    totals[l.project] = (totals[l.project] || 0) + l.cost;
    grand += l.cost;
  });
  const shares = {};
  if (grand === 0) { // no operating expense — attribute evenly to seen projects
    const names = uniqueProjects_(lines);
    names.forEach(n => (shares[n] = 1 / names.length));
    return shares;
  }
  Object.keys(totals).forEach(p => (shares[p] = totals[p] / grand));
  return shares;
}

function uniqueProjects_(lines) {
  const set = {};
  lines.forEach(l => (set[l.project] = true));
  const names = Object.keys(set);
  return names.length ? names : [CONFIG.DEFAULT_PROJECT];
}

function round_(n) { return Math.round(n); }
function startsWith_(s, p) { return String(s).indexOf(p) === 0; }

function earliestBudgetStart_(budgets) {
  let min = null;
  budgets.forEach(b => b.lines.forEach(l => {
    if (l.start && (!min || l.start < min)) min = l.start;
  }));
  return min || new Date(new Date().getFullYear() - 1, 0, 1);
}

function orgTotals_(rows) {
  const t = { budget: 0, secured: 0, actual: 0, unsecuredGap: 0 };
  rows.forEach(r => {
    t.budget += r.proposedBudget; t.secured += r.securedIncome;
    t.actual += r.actualExpense; t.unsecuredGap += r.unsecuredGap;
  });
  Object.keys(t).forEach(k => (t[k] = round_(t[k])));
  return t;
}
