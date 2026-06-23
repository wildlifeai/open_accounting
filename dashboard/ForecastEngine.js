/**
 * ForecastEngine.js
 * Pure forecasting math, ported from create_xero_budget_project.js and reshaped
 * into side-effect-free functions that operate on parsed budget lines rather than
 * live spreadsheets. No SpreadsheetApp / DriveApp here — that makes it unit-testable.
 *
 * A budget line is: { account, start: Date, end: Date, amount: Number,
 *                     milestone, item, project }
 *
 * The three behaviours preserved from the original script:
 *   1. day-weighted distribution of each line across the months it spans
 *   2. overhead (account 500) split across quarters weighted by expense load
 *   3. grant revenue recognition: recognise income as expense is incurred,
 *      defer the remainder (FIFO by income date).
 */

const DateMath = {
  monthKey(d) {
    return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
  },
  firstOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1); },
  lastOfMonth(d) { return new Date(d.getFullYear(), d.getMonth() + 1, 0); },

  monthsBetween(start, end) {
    const out = [];
    let cur = new Date(start.getFullYear(), start.getMonth(), 1);
    const last = new Date(end.getFullYear(), end.getMonth(), 1);
    while (cur <= last) { out.push(new Date(cur)); cur.setMonth(cur.getMonth() + 1); }
    return out;
  },

  /** Inclusive overlap in days between [aStart,aEnd] and [bStart,bEnd]. */
  overlapDays(aStart, aEnd, bStart, bEnd) {
    const start = Math.max(aStart.getTime(), bStart.getTime());
    const end = Math.min(aEnd.getTime(), bEnd.getTime());
    return Math.max(0, (end - start) / 86400000 + 1);
  },

  quartersBetween(start, end) {
    const quarters = [];
    for (let y = start.getFullYear(); y <= end.getFullYear(); y++) {
      for (let q = 1; q <= 4; q++) {
        const qStart = new Date(y, (q - 1) * 3, 1);
        const qEnd = new Date(y, q * 3, 0);
        if (qStart > end) break;
        if (qEnd >= start) quarters.push({ label: y + ' Q' + q, start: qStart, end: qEnd });
      }
    }
    return quarters;
  }
};

// ---- quarterly helpers (for the forecast/tracking layer) -------------------

function quarterLabel_(year, q) { return year + ' Q' + q; }

/** 'YYYY-MM' -> '2026 Q2'. */
function quarterOfMonthKey_(monthKey) {
  const p = monthKey.split('-');
  return quarterLabel_(parseInt(p[0], 10), Math.floor((parseInt(p[1], 10) - 1) / 3) + 1);
}

/** Roll a { 'YYYY-MM': n } month map up into { '2026 Q2': n }. */
function bucketToQuarters(monthMap) {
  const out = {};
  Object.keys(monthMap).forEach(k => {
    const ql = quarterOfMonthKey_(k);
    out[ql] = (out[ql] || 0) + monthMap[k];
  });
  return out;
}

/** The quarter label containing `date` (defaults to today). */
function currentQuarterLabel(date) {
  date = date || new Date();
  return quarterLabel_(date.getFullYear(), Math.floor(date.getMonth() / 3) + 1);
}

/** Sortable integer for a '2026 Q2' label. */
function quarterSortNum(label) {
  const p = label.split(' Q');
  return parseInt(p[0], 10) * 4 + parseInt(p[1], 10);
}

/**
 * Normalise an inventory-item string to its Xero product code so budgets and
 * actuals join on one key. Budget cells read "WW_25_TOI_002 - General
 * management"; Xero line items are already the code "WW_25_TOI_002".
 */
function itemCode_(value) {
  return String(value == null ? '' : value).split(' - ')[0].trim();
}

/**
 * Day-weighted distribution of a numeric field (`cost` or `income`) across the
 * months each line spans, into { monthKey: amount }.
 */
function distributeByMonth_(lines, field) {
  const byMonth = {};
  lines.forEach(l => {
    const amount = l[field];
    if (!l.start || !l.end || !amount) return;
    const totalDays = (l.end - l.start) / 86400000 + 1;
    DateMath.monthsBetween(l.start, l.end).forEach(m => {
      const days = DateMath.overlapDays(l.start, l.end, DateMath.firstOfMonth(m), DateMath.lastOfMonth(m));
      const key = DateMath.monthKey(m);
      byMonth[key] = (byMonth[key] || 0) + (amount * days) / totalDays;
    });
  });
  return byMonth;
}

/**
 * Top-level: compute the forecast bundle for one funding source's budget lines.
 * The funding-source `Budget` tab is keyed on milestone, not chart-of-accounts:
 *   expense = day-weighted `Cost`, income = day-weighted `Income`.
 * (`Contribution` = Income − Cost is the margin that funds General/overhead;
 *  it's already embedded in Income, so we don't double-count it here.)
 */
function computeFundingSourceForecast(lines) {
  const expenseByMonth = distributeByMonth_(lines, 'cost');
  const incomeByMonth = distributeByMonth_(lines, 'income');
  const sum = obj => Object.keys(obj).reduce((t, k) => t + obj[k], 0);
  return {
    expenseByMonth, incomeByMonth,
    totalBudgetExpense: sum(expenseByMonth),
    totalBudgetIncome: sum(incomeByMonth),
    totalContribution: sum(distributeByMonth_(lines, 'contribution'))
  };
}
