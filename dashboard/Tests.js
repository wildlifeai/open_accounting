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

  // Quarterly helpers.
  check('quarterOfMonthKey 2026-05 -> 2026 Q2', quarterOfMonthKey_('2026-05') === '2026 Q2');
  check('itemCode strips name', itemCode_('WW_25_TOI_002 - General management') === 'WW_25_TOI_002');
  var q = bucketToQuarters({ '2026-01': 100, '2026-02': 50, '2026-04': 30 });
  check('bucketToQuarters sums Q1', q['2026 Q1'] === 150);
  check('bucketToQuarters Q2', q['2026 Q2'] === 30);
  check('quarterSortNum order', quarterSortNum('2026 Q2') > quarterSortNum('2025 Q4'));

  // Forecast merge: past quarter -> actual, current/future -> override or baseline.
  var tsrc = {
    source: 'WW_25_TOI', status: 'secured', project: 'Wildlife Watcher',
    quarters: ['2025 Q4', '2026 Q1', '2026 Q2', '2026 Q3'],
    milestones: [{ item: 'WW_25_TOI_002', milestone: 'General management',
      baseline: { '2025 Q4': 4992, '2026 Q1': 7615, '2026 Q2': 7703, '2026 Q3': 2792 },
      actual: { '2025 Q4': 2000, '2026 Q1': 2500 } }]
  };
  var fmap = { 'WW_25_TOI||WW_25_TOI_002||2026 Q2': { cost: 5000 } };
  var grid = composeTrackingForSource(tsrc, fmap, '2026 Q2');
  var m = grid.milestones[0];
  check('past quarter uses actual', m.cells[0].effective === 2000);
  check('current quarter uses override', m.cells[2].effective === 5000 && m.cells[2].hasOverride);
  check('future quarter uses baseline', m.cells[3].effective === 2792);
  check('expected = 2000+2500+5000+2792', m.expectedTotal === 12292);
  check('actual to date = 4500', m.actualToDate === 4500);

  Logger.log(results.join('\n'));
  return results;
}
