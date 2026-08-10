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

  // Forecast merge with FY columns: aggregate 'Up to last FY' + this FY quarters.
  var entity = {
    id: 'WW_25_TOI', label: 'WW_25_TOI', type: 'source', source: 'WW_25_TOI',
    status: 'secured', project: 'Wildlife Watcher',
    milestones: [{ item: 'WW_25_TOI_002', milestone: 'General management', source: 'WW_25_TOI',
      baseline: { '25/26 Q3': 4992, '25/26 Q4': 7615, '26/27 Q1': 7703, '26/27 Q2': 2792 },
      actual: { '25/26 Q3': 2000, '25/26 Q4': 2500 } }]
  };
  var fmap = { amounts: { 'WW_25_TOI||WW_25_TOI_002||26/27 Q1': { cost: 5000 } },
    comments: { 'WW_25_TOI||WW_25_TOI_002': 'staffing ramp' } };
  var grid = composeTracking(entity, fmap, quarterSortNum('26/27 Q1'));
  var m = grid.milestones[0];
  check('first column is aggregate', grid.columns[0].type === 'aggregate');
  check('aggregate sums prior FY actual', m.cells[0].effective === 4500);
  check('current quarter uses override', m.cells[1].effective === 5000 && m.cells[1].hasOverride);
  check('future quarter uses baseline', m.cells[2].effective === 2792);
  check('expected = 4500+5000+2792', m.expectedTotal === 12292);
  check('baseline total = 23102', m.baselineTotal === 23102);
  check('comment carried through', m.comment === 'staffing ramp');

  Logger.log(results.join('\n'));
  return results;
}
