/**
 * Tests.js
 * Lightweight checks runnable from the Apps Script editor (Run > runTests).
 * No Drive/Xero needed - they exercise the pure forecast math against synthetic
 * funding-source budget lines (Cost/Income/milestone shape, matching the real
 * `Budget` tab parsed by BudgetReader).
 */

function runTests() {
  var results = [];
  function check(name, cond) { results.push((cond ? 'PASS  ' : 'FAIL  ') + name); }

  var d = function (y, m, day) { return new Date(y, m - 1, day); };

  // Mirrors SPY_26_HAND: milestones with Cost / Income / Contribution.
  var lines = [
    { milestone: 'Data cleaning', start: d(2025, 11, 1), end: d(2026, 1, 1),
      cost: 5600, income: 10000, contribution: 4400, project: 'Spyfish Aotearoa' },
    { milestone: 'Machine learning models', start: d(2026, 1, 1), end: d(2026, 3, 1),
      cost: 14000, income: 15000, contribution: 1000, project: 'Spyfish Aotearoa' },
    { milestone: 'Operational playbook', start: d(2025, 11, 1), end: d(2026, 4, 1),
      cost: 7560, income: 10000, contribution: 2440, project: 'General' }
  ];

  var fc = computeFundingSourceForecast(lines);

  check('expense total = 27,160', Math.abs(fc.totalBudgetExpense - 27160) < 1);
  check('income total = 35,000', Math.abs(fc.totalBudgetIncome - 35000) < 1);
  check('contribution total = 7,840', Math.abs(fc.totalContribution - 7840) < 1);

  // Cost spreads day-weighted across the months each milestone spans.
  check('expense spread over multiple months', Object.keys(fc.expenseByMonth).length >= 5);

  // Project shares by Cost: Spyfish = (5600+14000)/27160, General = 7560/27160.
  var shares = projectExpenseShares_(lines);
  check('Spyfish share ~0.722', Math.abs(shares['Spyfish Aotearoa'] - (19600 / 27160)) < 0.001);
  check('General share ~0.278', Math.abs(shares['General'] - (7560 / 27160)) < 0.001);

  // Financial-year quarter helpers (FY starts April).
  check('May 2026 -> 26/27 Q1', quarterOfMonthKey_('2026-05') === '26/27 Q1');
  check('Jan 2026 -> 25/26 Q4', quarterOfMonthKey_('2026-01') === '25/26 Q4');
  check('itemCode strips name', itemCode_('WW_25_TOI_002 - General management') === 'WW_25_TOI_002');
  var qb = bucketToQuarters({ '2026-01': 100, '2026-02': 50, '2026-04': 30 });
  check('bucketToQuarters sums FY Q4', qb['25/26 Q4'] === 150);
  check('bucketToQuarters next FY Q1', qb['26/27 Q1'] === 30);
  check('quarterSortNum order', quarterSortNum('26/27 Q1') > quarterSortNum('25/26 Q4'));

  // Budget tab: the column header row is located, not assumed, so a key|value
  // metadata block can sit above it (see BUDGET_SHEET_TEMPLATE.md).
  var sheetRows = [
    ['Funding source', 'WW_25_TOI'],
    ['Status', 'secured'],
    ['Contribution policy', 'percent_of_income:40'],
    [],
    ['Description', 'Start', 'End', 'Cost', 'Income', 'Milestone'],
    ['Delivery lead', '03/Nov/25', '02/Aug/26', 20000, 0, 'Delivery']
  ];
  check('header row located below metadata block', findHeaderRow_(sheetRows) === 4);
  check('header row is 0 when there is no metadata block',
    findHeaderRow_([['Description', 'Start', 'End', 'Cost']]) === 0);
  check('no header row returns -1', findHeaderRow_([['Notes', 'x'], ['more', 'y']]) === -1);

  var meta = readMetadataBlock_(sheetRows, 4);
  check('metadata keys are lower-cased', meta['funding source'] === 'WW_25_TOI');
  check('metadata carries contribution policy',
    meta['contribution policy'] === 'percent_of_income:40');
  check('metadata stops at the header row', meta['description'] === undefined);
  check('isoDate_ zero-pads', isoDate_(new Date(2026, 5, 3)) === '2026-06-03');

  // Balance-sheet exclusions match on account code, not label, so renaming an
  // account in Xero cannot silently un-exclude it.
  check('excludes Wages Payable', isExcludedAccount_('Wages Payable - Payroll (814)'));
  check('excludes PAYE Payable', isExcludedAccount_('PAYE Payable (825)'));
  check('excludes after a Xero rename', isExcludedAccount_('Renamed In Xero (814)'));
  check('keeps Salaries', !isExcludedAccount_('Salaries (477)'));
  check('keeps an untagged/blank account', !isExcludedAccount_(''));
  check('keeps an unknown code', !isExcludedAccount_('Something New (999)'));

  // Health findings: severity order, value at risk, and the checks the reader
  // cannot make for itself. See dashboard/HEALTH_CHECKS.md.
  var hBudgets = [{
    name: 'WW_25_TOI', status: 'secured', projectFolder: 'Wildlife Watcher', sheetUrl: '',
    metadata: { owner: 'victor@wildlife.ai' },
    tabs: ['Budget', 'Forecast', 'Xero export'],
    lines: [{ item: 'WW_25_TOI_002', cost: 100, income: 0, contribution: 0 },
            { item: '', cost: 5000, income: 0, contribution: 0 }],
    forecast: { cost: { 'WW_25_TOI_099||26/27 Q1': 50 }, income: {}, comments: {} },
    issues: [{ check: 'B4', detail: '1 line without an item code' },
             { check: 'A1', detail: 'missing column Cost' }]
  }];
  var hActuals = [{ kind: 'expense', project: '', fundingSource: '', amount: 4053 },
                  { kind: 'expense', project: 'General', fundingSource: '', amount: 200 }];
  var health = buildHealth(hBudgets, hActuals,
    { xeroConnected: false, exclusion: { count: 3, total: 3090 }, secretsMissing: [] });
  var hIds = health.map(function (f) { return f.id; });

  check('missing column reported as A2, not A1', hIds.indexOf('A2') !== -1);
  check('A3 flags a tab outside the three allowed', hIds.indexOf('A3') !== -1);
  check('A6 flags the missing Submitted_budget tab', hIds.indexOf('A6') !== -1);
  check('A9 flags a forecast for a milestone not in the budget', hIds.indexOf('A9') !== -1);
  check('B4 carries value at risk, not just a count',
    health.some(function (f) { return f.id === 'B4' && f.amount === 5000; }));
  check('D1 counts only untagged expense',
    health.some(function (f) { return f.id === 'D1' && f.amount === 4053; }));
  check('D2 does not double-count the D1 line',
    health.some(function (f) { return f.id === 'D2' && f.amount === 200; }));
  check('F1 raised when Xero is disconnected', hIds.indexOf('F1') !== -1);
  check('errors sort before warnings before info', (function () {
    var rank = { error: 0, warning: 1, info: 2 };
    for (var i = 1; i < health.length; i++) {
      if (rank[health[i].severity] < rank[health[i - 1].severity]) return false;
    }
    return true;
  })());
  check('owner carried through from sheet metadata',
    health.some(function (f) { return f.owner === 'victor@wildlife.ai'; }));
  check('legacy dataFlags strings exclude info findings',
    healthToFlags(health).length === health.filter(function (f) {
      return f.severity !== 'info'; }).length);

  // Forecast row labels. The Forecast tab may name a budget line by item code,
  // by milestone, by "Description - Milestone", or by a description unique in the
  // file. Mirrors SPY_26_UOA: two milestones, one item code each, with the same
  // two descriptions appearing under both.
  var uoaLines = [
    { description: 'Data Scientist', milestone: 'Baseline model assessment and data ingestion',
      item: 'SPY_26_UOA_001 - Baseline model assessment' },
    { description: 'Project Manager', milestone: 'Baseline model assessment and data ingestion',
      item: 'SPY_26_UOA_001 - Baseline model assessment' },
    { description: 'Data Scientist', milestone: 'Final validation and co-authored manuscript',
      item: 'SPY_26_UOA_002 - Final validation and manuscript' },
    { description: 'Project Manager', milestone: 'Final validation and co-authored manuscript',
      item: 'SPY_26_UOA_002 - Final validation and manuscript' }
  ];
  var uoaMap = buildForecastLabelMap_(uoaLines);
  function resolves(label) { return resolveForecastLabel_(label, uoaMap).code; }

  check('milestone name resolves', resolves('Baseline model assessment and data ingestion')
    === 'SPY_26_UOA_001');
  check('"Description - Milestone" resolves',
    resolves('Data Scientist - Final validation and co-authored manuscript') === 'SPY_26_UOA_002');
  check('bare item code still resolves', resolves('SPY_26_UOA_001') === 'SPY_26_UOA_001');
  check('full "CODE - Name" still resolves',
    resolves('SPY_26_UOA_001 - Baseline model assessment') === 'SPY_26_UOA_001');
  check('label matching ignores case and extra spaces',
    resolves('data scientist  -  final validation and co-authored manuscript')
    === 'SPY_26_UOA_002');
  // A description under two milestones must be refused, not guessed at. These two
  // cost $936 in both milestones, so no amount of cleverness could pick one.
  check('a description spanning two milestones is refused', !resolves('Project Manager'));
  check('and the refusal names both candidates', (function () {
    var e = resolveForecastLabel_('Project Manager', uoaMap).error || '';
    return e.indexOf('SPY_26_UOA_001') !== -1 && e.indexOf('SPY_26_UOA_002') !== -1;
  })());
  check('an unknown label is refused', !resolves('Data Engineer'));
  // One milestone spanning several accounts is not ambiguous: those lines share a
  // code, so the candidate set collapses to one. This is the normal shape.
  check('one milestone over three accounts resolves', resolveForecastLabel_('Data science',
    buildForecastLabelMap_([
      { description: 'Recruitment fees', milestone: 'Data science', item: 'X_001 - Data science' },
      { description: 'Contractor', milestone: 'Data science', item: 'X_001 - Data science' },
      { description: 'Software', milestone: 'Data science', item: 'X_001 - Data science' }
    ])).code === 'X_001');
  // Labels are matched whole, never split, so " - " inside a description is safe.
  check('a description containing " - " is not mis-split', resolveForecastLabel_(
    'Travel - domestic', buildForecastLabelMap_([{ description: 'Travel - domestic',
      milestone: 'Fieldwork', item: 'Y_001 - Fieldwork' }])).code === 'Y_001');
  check('a line with no item code offers no label',
    Object.keys(buildForecastLabelMap_([{ description: 'Thing', milestone: 'M', item: '' }]))
      .length === 0);

  // Forecast merge with FY columns: aggregate 'Up to last FY' + this FY quarters.
  // Forecast overrides live on the milestone itself, read from each funding
  // source's own Forecast tab by BudgetReader.parseForecastTab_.
  var entity = {
    id: 'WW_25_TOI', label: 'WW_25_TOI', type: 'source', source: 'WW_25_TOI',
    status: 'secured', project: 'Wildlife Watcher',
    milestones: [{ item: 'WW_25_TOI_002', milestone: 'General management', source: 'WW_25_TOI',
      baseline: { '25/26 Q3': 4992, '25/26 Q4': 7615,
        '26/27 Q1': 7703, '26/27 Q2': 2792, '26/27 Q3': 1000 },
      actual: { '25/26 Q3': 2000, '25/26 Q4': 2500 },
      costForecast: { '26/27 Q2': 5000 },   // override on one future quarter only
      forecastComment: 'staffing ramp' }]
  };
  var grid = composeTracking(entity, quarterSortNum('26/27 Q1'), 'cost');
  var m = grid.milestones[0];

  // Look columns up by label - index arithmetic is brittle as columns evolve.
  function cellFor(label) {
    for (var i = 0; i < grid.columns.length; i++) {
      if (grid.columns[i].label === label) return m.cells[i];
    }
    return null;
  }

  check('first column is aggregate', grid.columns[0].type === 'aggregate');
  check('aggregate sums prior FY actual', m.cells[0].effective === 4500);
  check('future quarter uses its override', cellFor('26/27 Q2').effective === 5000
    && cellFor('26/27 Q2').hasForecast === true);
  // Regression guard: an unmaintained Forecast tab must fall back to the budget
  // baseline, never to 0. Reading 0 makes a source look certain to underspend.
  check('future quarter with no override falls back to baseline',
    cellFor('26/27 Q3').effective === 1000 && cellFor('26/27 Q3').hasForecast === false);
  check('current quarter uses actual, not baseline', cellFor('26/27 Q1').effective === 0);
  check('expected = 4500 + 0 + 5000 + 1000 + 0', m.expectedTotal === 10500);
  check('baseline total = 24102', m.baselineTotal === 24102);
  check('comment carried through', m.comment === 'staffing ramp');

  // Per-quarter buckets, which let the Overview total any financial year rather than
  // only the current one. A line spanning Jan to Dec 2026 crosses FY25/26 Q4 into
  // FY26/27 Q1-Q3, so the split must be day-weighted and the partition exhaustive.
  var straddler = { milestone: 'M', item: 'X_001 - M', project: 'General',
    start: d(2026, 1, 1), end: d(2026, 12, 31), cost: 12000, income: 12000, contribution: 0 };
  var strMonths = distributeByMonth_([straddler], 'cost');
  var strQ = bucketToQuarters(strMonths);
  function sumMap(m) {
    return Object.keys(m).reduce(function (a, k) { return a + m[k]; }, 0);
  }
  function sumFyQ(m, fy) {
    return Object.keys(m).filter(function (q) { return q.split(' ')[0] === fy; })
      .reduce(function (a, q) { return a + m[q]; }, 0);
  }
  check('quarter buckets sum to the line cost', Math.abs(sumMap(strQ) - 12000) < 0.01);
  check('a straddling line lands in two financial years',
    sumFyQ(strQ, '25/26') > 0 && sumFyQ(strQ, '26/27') > 0);
  check('summing both years reproduces the whole line',
    Math.abs(sumFyQ(strQ, '25/26') + sumFyQ(strQ, '26/27') - 12000) < 0.01);
  // The per-quarter total must agree with the scalar the old FY-only view used, or the
  // FY selector would quietly disagree with every previously reported figure.
  check('per-quarter FY total equals sumMonthsInFY_',
    Math.abs(sumMonthsInFY_(strMonths, fyBounds_(d(2026, 8, 13))) - sumFyQ(strQ, '26/27')) < 0.01);
  var accQ = {};
  addInto_(accQ, { '26/27 Q1': 10, '26/27 Q2': 5 });
  addInto_(accQ, { '26/27 Q1': 7 });
  check('addInto_ accumulates rather than overwrites',
    accQ['26/27 Q1'] === 17 && accQ['26/27 Q2'] === 5);

  // Project-lead scoped access. filterSnapshotForProjects_ is pure precisely so this can
  // run without a second Google account signed in.
  var snap = {
    generatedAt: '2026-08-12T00:00:00Z', xeroConnected: true, currentQuarter: '26/27 Q2',
    coverage: {},
    totals: { budget: 999, secured: 999, actual: 999, unsecuredGap: 999,
              budgetFY: 999, securedFY: 999, actualFY: 999, unsecuredGapFY: 999 },
    projects: [
      { project: 'Spyfish Aotearoa', proposedBudget: 100, securedIncome: 60,
        actualExpense: 40, unsecuredGap: 40, proposedBudgetFY: 10, securedIncomeFY: 6,
        actualExpenseFY: 4, unsecuredGapFY: 4 },
      { project: 'Wildlife Watcher', proposedBudget: 200, securedIncome: 50,
        actualExpense: 70, unsecuredGap: 150, proposedBudgetFY: 20, securedIncomeFY: 5,
        actualExpenseFY: 7, unsecuredGapFY: 15 }
    ],
    fundingSources: [{ project: 'Spyfish Aotearoa', name: 'SPY_26_UOA' },
                     { project: 'Wildlife Watcher', name: 'WW_25_TOI' }],
    breakdownRows: [{ project: 'Spyfish Aotearoa' }, { project: 'Wildlife Watcher' }],
    tracking: [{ id: 'SPY_26_UOA', milestones: [{ project: 'Spyfish Aotearoa' }] },
               { id: 'WW_25_TOI', milestones: [{ project: 'Wildlife Watcher' }] }],
    timeline: [{ project: 'Spyfish Aotearoa' }, { project: 'Wildlife Watcher' }],
    health: [
      { id: 'B4', severity: 'error', project: 'Spyfish Aotearoa',
        fundingSource: 'SPY_26_UOA', title: 'mine', detail: '', amount: 1,
        owner: 'a@wildlife.ai', link: 'https://sheet/spy' },
      { id: 'A1', severity: 'error', project: 'Wildlife Watcher',
        fundingSource: 'WW_25_TOI', title: 'not mine', detail: '', amount: 2,
        owner: 'b@wildlife.ai', link: 'https://sheet/ww' },
      { id: 'D1', severity: 'error', project: '', fundingSource: '',
        title: 'org-wide untagged spend', detail: '', amount: 5000 },
      { id: 'F5', severity: 'info', project: '', fundingSource: '',
        title: 'Refresh summary', detail: '11 funding source(s) read' },
      { id: 'F1', severity: 'error', project: '', fundingSource: '',
        title: 'Xero is not connected', detail: '' }
    ],
    dataFlags: ['stale flag that must be rebuilt from the filtered findings']
  };

  var lead = filterSnapshotForProjects_(JSON.parse(JSON.stringify(snap)),
                                        ['Spyfish Aotearoa']);
  check('scoped: access level is filtered', lead._accessLevel === 'filtered');
  check('scoped: only their project', lead.projects.length === 1 &&
    lead.projects[0].project === 'Spyfish Aotearoa');
  check('scoped: funding sources filtered', lead.fundingSources.length === 1 &&
    lead.fundingSources[0].name === 'SPY_26_UOA');
  check('scoped: breakdown filtered', lead.breakdownRows.length === 1);
  check('scoped: tracking filtered', lead.tracking.length === 1 &&
    lead.tracking[0].id === 'SPY_26_UOA');
  check('scoped: timeline filtered', lead.timeline.length === 1);
  // Totals must be rebuilt, never inherited: 999 would leak the org-wide figure.
  check('scoped: totals recomputed from their project only',
    lead.totals.budget === 100 && lead.totals.secured === 60 &&
    lead.totals.actual === 40 && lead.totals.unsecuredGap === 40);
  check('scoped: FY totals recomputed too',
    lead.totals.budgetFY === 10 && lead.totals.actualFY === 4);

  var hIds2 = lead.health.map(function (f) { return f.id; });
  check('scoped: keeps their own finding', hIds2.indexOf('B4') !== -1);
  check('scoped: hides another project\'s finding', hIds2.indexOf('A1') === -1);
  check('scoped: hides org-wide untagged spend (D1)', hIds2.indexOf('D1') === -1);
  check('scoped: hides the org-wide refresh summary (F5)', hIds2.indexOf('F5') === -1);
  check('scoped: keeps Xero disconnected (F1), their numbers are stale too',
    hIds2.indexOf('F1') !== -1);
  check('scoped: no other sheet link survives', lead.health.every(function (f) {
    return !f.link || f.link.indexOf('/ww') === -1; }));
  check('scoped: no other owner email survives', lead.health.every(function (f) {
    return f.owner !== 'b@wildlife.ai'; }));
  check('scoped: dataFlags rebuilt from the filtered findings, not inherited',
    lead.dataFlags.length === healthToFlags(lead.health).length &&
    lead.dataFlags.join(' ').indexOf('stale flag') === -1);

  var none = filterSnapshotForProjects_(JSON.parse(JSON.stringify(snap)), []);
  check('no access: everything empty', none._accessLevel === 'none' &&
    none.projects.length === 0 && none.health.length === 0 &&
    none.dataFlags.length === 0 && none.totals.budget === 0);

  var admin = filterSnapshotForProjects_(JSON.parse(JSON.stringify(snap)), ['*']);
  check('admin: sees everything', admin._accessLevel === 'admin' &&
    admin.health.length === 5 && admin.projects.length === 2);

  Logger.log(results.join('\n'));
  return results;
}
