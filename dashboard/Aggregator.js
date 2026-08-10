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
  const now = new Date();
  const fy = fyBounds_(now);

  const budgets = readAllBudgets();
  const actualSince = earliestBudgetStart_(budgets);
  const actualLines = isXeroConnected() ? fetchXeroActuals(actualSince) : [];

  const projects = {}; // name -> rollup accumulator
  const fundingSources = []; // per-source summary
  const dataFlags = [];
  const budgetByProjSrc = {}; // 'project||source' -> { budget, income, status }
  const actualByProjSrc = {}; // 'project||source' -> actual expense
  const actualByProjSrcFY = {}; // 'project||source' -> FY actual expense
  const sourceStatus = {};    // source name -> 'secured' | 'proposed'

  function project_(name) {
    if (!projects[name]) {
      projects[name] = { name: name, proposedBudget: 0, securedIncome: 0,
        proposedIncome: 0, actualExpense: 0,
        proposedBudgetFY: 0, securedIncomeFY: 0, proposedIncomeFY: 0, actualExpenseFY: 0 };
    }
    return projects[name];
  }

  budgets.forEach(src => {
    // Per-line project attribution is handled in BudgetReader: a line uses its
    // `Project` value when set, otherwise the parent project folder. Falling back
    // to the folder is expected, so it is not flagged here.
    const fc = computeFundingSourceForecast(src.lines);
    const shares = projectExpenseShares_(src.lines); // {project: 0..1}

    const projDates = {};
    src.lines.forEach(l => {
      const p = l.project || CONFIG.DEFAULT_PROJECT;
      if (!projDates[p]) projDates[p] = { start: null, end: null };
      if (l.start && (!projDates[p].start || l.start < projDates[p].start)) projDates[p].start = l.start;
      if (l.end && (!projDates[p].end || l.end > projDates[p].end)) projDates[p].end = l.end;
    });

    const expenseFY = sumMonthsInFY_(fc.expenseByMonth, fy);
    const incomeFY = sumMonthsInFY_(fc.incomeByMonth, fy);

    Object.keys(shares).forEach(pName => {
      const share = shares[pName];
      const p = project_(pName);
      const expTotal = fc.totalBudgetExpense * share;
      const incTotal = fc.totalBudgetIncome * share;
      const expFY = expenseFY * share;
      const incFY = incomeFY * share;

      // Proposed budget = secured + proposed; secured income only on secured sources.
      p.proposedBudget += expTotal;
      p.proposedIncome += incTotal;
      p.proposedBudgetFY += expFY;
      p.proposedIncomeFY += incFY;
      if (src.status === 'secured') {
        p.securedIncome += incTotal;
        p.securedIncomeFY += incFY;
      }

      const dStart = projDates[pName] ? projDates[pName].start : null;
      const dEnd = projDates[pName] ? projDates[pName].end : null;

      budgetByProjSrc[pName + '||' + src.name] =
        { budget: expTotal, income: incTotal, status: src.status,
          budgetFY: expFY, incomeFY: incFY, start: isoOrNull_(dStart), end: isoOrNull_(dEnd) };
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
    const p = project_(l.project);
    p.actualExpense += l.amount;
    
    const isFY = (new Date(l.date) >= fy.start && new Date(l.date) <= fy.end);
    if (isFY) p.actualExpenseFY += l.amount;

    const fs = l.fundingSource && !startsWith_(l.fundingSource, CONFIG.ARCHIVE_PREFIX)
      ? l.fundingSource : '(unassigned)';
    const akey = l.project + '||' + fs;
    actualByProjSrc[akey] = (actualByProjSrc[akey] || 0) + l.amount;
    if (isFY) actualByProjSrcFY[akey] = (actualByProjSrcFY[akey] || 0) + l.amount;
  });

  const breakdownRows = buildBreakdownRows_(budgetByProjSrc, actualByProjSrc, actualByProjSrcFY, sourceStatus);

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
      unsecuredGap: round_(Math.max(0, p.proposedBudget - p.securedIncome)),
      proposedBudgetFY: round_(p.proposedBudgetFY),
      securedIncomeFY: round_(p.securedIncomeFY),
      actualExpenseFY: round_(p.actualExpenseFY),
      unsecuredGapFY: round_(Math.max(0, p.proposedBudgetFY - p.securedIncomeFY))
    };
  });

  // fy and now are already defined at the top
  const coverage = {
    today: now.toISOString(),
    fyLabel: fy.label,
    fyStart: fy.start.toISOString(),
    fyEnd: fy.end.toISOString(),
    forecastStart: isoOrNull_(earliestBudgetStart_(budgets)),
    forecastEnd: isoOrNull_(latestBudgetEnd_(budgets))
  };

  return {
    generatedAt: now.toISOString(),
    xeroConnected: isXeroConnected(),
    currentQuarter: currentQuarterLabel(),
    coverage: coverage,
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
function buildBreakdownRows_(budgetByProjSrc, actualByProjSrc, actualByProjSrcFY, sourceStatus) {
  const keys = {};
  Object.keys(budgetByProjSrc).forEach(k => (keys[k] = true));
  Object.keys(actualByProjSrc).forEach(k => (keys[k] = true));

  return Object.keys(keys).map(key => {
    const parts = key.split('||');
    const project = parts[0];
    const source = parts[1];
    const b = budgetByProjSrc[key] || { budget: 0, income: 0, budgetFY: 0, incomeFY: 0, start: null, end: null };
    const status = (b.status || sourceStatus[source] || 'unknown');
    return {
      project: project,
      fundingSource: source,
      status: status,
      budget: round_(b.budget || 0),
      secured: round_(status === 'secured' ? (b.income || 0) : 0),
      actual: round_(actualByProjSrc[key] || 0),
      budgetFY: round_(b.budgetFY || 0),
      securedFY: round_(status === 'secured' ? (b.incomeFY || 0) : 0),
      actualFY: round_(actualByProjSrcFY[key] || 0),
      start: b.start || null,
      end: b.end || null
    };
  });
}

/**
 * Per funding source: baseline (frozen budget) and actual (from Xero) spend,
 * bucketed by milestone (item code) and quarter. Returns an array ready for the
 * quarterly tracking screen; the live forecast layer is merged in WebApp/UI.
 */
function buildTracking_(budgets, actualLines) {
  // Index Xero actuals by source||itemCode -> { quarter: amount }, split by
  // expense (cost) vs income. Also remember a display name per item code.
  const costActuals = {};
  const incomeActuals = {};
  const actualNames = {}; // source||code -> Xero item name
  actualLines.forEach(l => {
    if (!l.fundingSource || startsWith_(l.fundingSource, CONFIG.ARCHIVE_PREFIX)) return;
    const q = quarterOfMonthKey_(DateMath.monthKey(new Date(l.date)));
    const code = itemCode_(l.item);
    const key = l.fundingSource + '||' + code;
    const bucket = l.kind === 'income' ? incomeActuals : costActuals;
    (bucket[key] = bucket[key] || {});
    bucket[key][q] = (bucket[key][q] || 0) + l.amount;
    if (code && l.itemName && !actualNames[key]) actualNames[key] = l.itemName;
  });

  return budgets.map(src => {
    // Group budget lines by milestone (item code).
    const byItem = {};
    src.lines.forEach(l => {
      const code = itemCode_(l.item) || ('(' + (l.milestone || 'unassigned') + ')');
      byItem[code] = byItem[code] || { milestone: l.milestone || code, lines: [] };
      byItem[code].lines.push(l);
    });

    const milestones = Object.keys(byItem).map(code => {
      const key = src.name + '||' + code;
      return {
        item: code, milestone: byItem[code].milestone, source: src.name,
        project: dominantProject_(byItem[code].lines),
        baseline: roundMap_(bucketToQuarters(distributeByMonth_(byItem[code].lines, 'cost'))),
        actual: roundMap_(costActuals[key] || {}),
        incomeBaseline: roundMap_(bucketToQuarters(distributeByMonth_(byItem[code].lines, 'income'))),
        incomeActual: roundMap_(incomeActuals[key] || {})
      };
    });

    // Add a row for every actual item code NOT in the budget, so unbudgeted
    // income (e.g. cash-received items) and miscoded expenses still show and the
    // totals reconcile to Xero. Codeless actuals go to an "Unassigned" row.
    const prefix = src.name + '||';
    const seenCodes = {};
    Object.keys(costActuals).concat(Object.keys(incomeActuals)).forEach(k => {
      if (k.indexOf(prefix) === 0) seenCodes[k.substring(prefix.length)] = true;
    });
    Object.keys(seenCodes).forEach(code => {
      if (byItem[code]) return; // already a budget milestone
      const key = prefix + code;
      const actual = roundMap_(costActuals[key] || {});
      const incomeActual = roundMap_(incomeActuals[key] || {});
      if (!Object.keys(actual).length && !Object.keys(incomeActual).length) return;
      const blank = (code === '');
      milestones.push({
        item: blank ? '(unassigned)' : code,
        milestone: blank ? 'Unassigned (no product/service)'
          : (actualNames[key] || code) + ' (unbudgeted)',
        source: src.name, project: src.projectFolder, actualOnly: true,
        baseline: {}, actual: actual, incomeBaseline: {}, incomeActual: incomeActual
      });
    });

    const quarterSet = {};
    milestones.forEach(m => [m.baseline, m.actual, m.incomeBaseline, m.incomeActual]
      .forEach(map => Object.keys(map).forEach(q => (quarterSet[q] = true))));
    const quarters = Object.keys(quarterSet).sort((a, b) => quarterSortNum(a) - quarterSortNum(b));
    return { source: src.name, status: src.status, project: src.projectFolder,
      quarters: quarters, milestones: milestones };
  });
}

/** The project carrying the most budgeted cost across a milestone's lines. */
function dominantProject_(lines) {
  const byProject = {};
  lines.forEach(l => (byProject[l.project] = (byProject[l.project] || 0) + (l.cost || 0)));
  let best = null, bestVal = -1;
  Object.keys(byProject).forEach(p => { if (byProject[p] > bestVal) { bestVal = byProject[p]; best = p; } });
  return best || (lines[0] && lines[0].project) || CONFIG.DEFAULT_PROJECT;
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

function latestBudgetEnd_(budgets) {
  let max = null;
  budgets.forEach(b => b.lines.forEach(l => {
    if (l.end && (!max || l.end > max)) max = l.end;
  }));
  return max;
}

function isoOrNull_(d) { return d ? d.toISOString() : null; }

function orgTotals_(rows) {
  const t = { budget: 0, secured: 0, actual: 0, unsecuredGap: 0,
              budgetFY: 0, securedFY: 0, actualFY: 0, unsecuredGapFY: 0 };
  rows.forEach(r => {
    t.budget += r.proposedBudget; t.secured += r.securedIncome;
    t.actual += r.actualExpense; t.unsecuredGap += r.unsecuredGap;
    t.budgetFY += r.proposedBudgetFY; t.securedFY += r.securedIncomeFY;
    t.actualFY += r.actualExpenseFY; t.unsecuredGapFY += r.unsecuredGapFY;
  });
  Object.keys(t).forEach(k => (t[k] = round_(t[k])));
  return t;
}

function sumMonthsInFY_(monthMap, fy) {
  let sum = 0;
  Object.keys(monthMap).forEach(k => {
    const p = k.split('-');
    const d = new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, 1);
    if (d >= fy.start && d <= fy.end) sum += monthMap[k];
  });
  return sum;
}
