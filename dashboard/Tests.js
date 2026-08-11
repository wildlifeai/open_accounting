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
    ['GM 0.2 FTE', '03/Nov/25', '02/Aug/26', 23101, 0, 'General management']
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

  Logger.log(results.join('\n'));
  return results;
}
