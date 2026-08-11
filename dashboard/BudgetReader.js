/**
 * BudgetReader.js
 * Walks the Budgets Drive tree and parses each funding source's `Budget` tab into
 * structured budget lines. Funding sources are discovered from folder structure:
 *
 *   BUDGETS_ROOT_FOLDER_ID
 *     |- <project folder>
 *          |- secured/   <funding-source Gsheet>  -> status 'secured'
 *          |- proposed/  <funding-source Gsheet>  -> status 'proposed'
 *
 * Funding sources whose name starts with ARCHIVE_PREFIX are skipped.
 *
 * The funding-source `Budget` tab is keyed on milestone, not chart-of-accounts.
 * Real columns: Description, Start, End, Cost, Income, Contribution, Comments,
 * Milestone, Xero Inventory Item. A budget line is therefore:
 *   { description, start, end, cost, income, contribution, milestone, item, project }
 */

/**
 * Everything this reader discards is reported rather than silently dropped: a
 * malformed line used to vanish behind a Logger.log nobody reads. Each entry
 * carries `issues: [{ check, row?, detail }]` where `check` is a health-check id
 * from dashboard/HEALTH_CHECKS.md.
 *
 * @return {Array<{name, status, projectFolder, lines, hasProjectColumn, metadata,
 *                 tabs, forecast, sheetUrl, issues}>}
 *   one entry per funding-source spreadsheet found.
 */
function readAllBudgets() {
  const root = DriveApp.getFolderById(CONFIG.BUDGETS_ROOT_FOLDER_ID);
  const sources = [];
  const projectFolders = root.getFolders();
  while (projectFolders.hasNext()) {
    const projectFolder = projectFolders.next();
    if (startsWithArchive_(projectFolder.getName())) continue;
    readStatusFolder_(projectFolder, CONFIG.SECURED_FOLDER_NAME, 'secured', sources);
    readStatusFolder_(projectFolder, CONFIG.PROPOSED_FOLDER_NAME, 'proposed', sources);
  }
  return sources;
}

function startsWithArchive_(name) {
  return name.indexOf(CONFIG.ARCHIVE_PREFIX) === 0;
}

function readStatusFolder_(projectFolder, subName, status, out) {
  const subs = projectFolder.getFoldersByName(subName);
  if (!subs.hasNext()) return;
  const files = subs.next().getFilesByType(MimeType.GOOGLE_SHEETS);
  while (files.hasNext()) {
    const file = files.next();
    if (startsWithArchive_(file.getName())) continue;

    const entry = { name: file.getName(), status: status,
      projectFolder: projectFolder.getName(), lines: [], hasProjectColumn: false,
      metadata: {}, tabs: [], sheetUrl: '',
      forecast: { cost: {}, income: {}, comments: {} }, issues: [] };
    try {
      // Open once and share. This file used to be opened twice - once for the
      // Budget tab, once for Forecast - and the crawl already approaches the
      // Apps Script 6-minute limit.
      const ss = SpreadsheetApp.openById(file.getId());
      entry.sheetUrl = ss.getUrl();
      entry.tabs = ss.getSheets().map(s => s.getName());

      const parsed = parseBudgetFile_(ss, projectFolder.getName());
      entry.lines = parsed.lines;
      entry.hasProjectColumn = parsed.hasProjectColumn;
      entry.issues = entry.issues.concat(parsed.issues);

      // Metadata: the Funding_info tab wins where present, falling back to a
      // `key | value` block above the Budget columns.
      entry.metadata = parsed.metadata || {};
      const info = parseFundingInfoTab_(ss);
      if (info) Object.keys(info).forEach(k => { entry.metadata[k] = info[k]; });

      const forecast = parseForecastTab_(ss);
      entry.forecast = forecast.data;
      entry.issues = entry.issues.concat(forecast.issues);
    } catch (e) {
      // A file that will not parse is invisible on the dashboard. Report it as a
      // finding instead of only logging, and still push the entry so the source
      // appears in Health rather than disappearing without trace.
      entry.issues.push({ check: 'A1', detail: e.message });
      Logger.log('Skipped ' + file.getName() + ': ' + e.message);
    }
    out.push(entry);
  }
}

function parseBudgetFile_(ss, projectFolderName) {
  const sheet = ss.getSheetByName(CONFIG.BUDGET_TAB);
  if (!sheet) throw new Error('no "' + CONFIG.BUDGET_TAB + '" tab');
  const data = sheet.getDataRange().getValues();

  // The header row is located, not assumed: a `key | value` metadata block may sit
  // above it (see BUDGET_SHEET_TEMPLATE.md).
  const headerRow = findHeaderRow_(data);
  if (headerRow === -1) {
    throw new Error('no column header row found in the first ' +
      (CONFIG.HEADER_SCAN_ROWS || 40) + ' rows (needs both "' +
      CONFIG.BUDGET_COLUMNS.start + '" and "' + CONFIG.BUDGET_COLUMNS.cost + '")');
  }
  const metadata = readMetadataBlock_(data, headerRow);

  const header = data[headerRow];
  const col = {};
  Object.keys(CONFIG.BUDGET_COLUMNS).forEach(k => {
    const expected = clean_(CONFIG.BUDGET_COLUMNS[k]).toLowerCase();
    col[k] = header.findIndex(h => clean_(h).toLowerCase() === expected);
  });
  CONFIG.BUDGET_REQUIRED_COLUMNS.forEach(req => {
    if (col[req] === -1) throw new Error('missing column ' + CONFIG.BUDGET_COLUMNS[req]);
  });
  const hasProjectColumn = col.project !== -1;

  const lines = [];
  const issues = [];
  let zeroRows = 0, missingItem = 0;

  for (let i = headerRow + 1; i < data.length; i++) {
    const row = data[i];
    const rowNo = i + 1; // 1-based, as it appears in Sheets
    const cost = parseAmount_(col.cost !== -1 ? row[col.cost] : 0);
    const income = parseAmount_(col.income !== -1 ? row[col.income] : 0);
    if (cost === 0 && income === 0) { zeroRows++; continue; } // blank/summary row

    const start = parseSheetDate_(row[col.start]);
    const end = parseSheetDate_(row[col.end]);
    if (!start || !end) {
      issues.push({ check: 'B2', row: rowNo, detail: 'unparseable date (Start "' +
        clean_(row[col.start]) + '", End "' + clean_(row[col.end]) + '") - line ignored' });
      continue;
    }
    if (end < start) {
      // Day-weighting a reversed range yields nonsense. Two-digit years are the
      // usual cause: "30/Jun/01" parses as the year 2001.
      issues.push({ check: 'B3', row: rowNo, detail: 'End ' + isoDate_(end) +
        ' precedes Start ' + isoDate_(start) + ' - line ignored' });
      continue;
    }

    // Project attribution: a line's `Project` value wins when set; otherwise
    // (no column, or blank cell) the line belongs to its parent project folder.
    const project = hasProjectColumn && row[col.project]
      ? clean_(row[col.project])
      : (projectFolderName || CONFIG.DEFAULT_PROJECT);

    const item = col.item !== -1 ? clean_(row[col.item]) : '';
    if (!item) missingItem++;

    lines.push({
      description: col.description !== -1 ? clean_(row[col.description]) : '',
      start: start, end: end,
      cost: cost, income: income,
      contribution: col.contribution !== -1 ? parseAmount_(row[col.contribution]) : (income - cost),
      account: col.account !== -1 ? clean_(row[col.account]) : '',
      milestone: col.milestone !== -1 ? clean_(row[col.milestone]) : '',
      item: item,
      project: project
    });
  }

  if (zeroRows) {
    issues.push({ check: 'B1', detail: zeroRows +
      ' line(s) skipped because Cost and Income are both 0' });
  }
  if (missingItem) {
    issues.push({ check: 'B4', detail: missingItem +
      ' line(s) have no "' + CONFIG.BUDGET_COLUMNS.item +
      '", so they fall outside milestone tracking' });
  }
  if (!hasProjectColumn) {
    issues.push({ check: 'A4', detail: 'no "' + CONFIG.BUDGET_COLUMNS.project +
      '" column, so lines cannot be split across projects' });
  }

  return { lines: lines, hasProjectColumn: hasProjectColumn,
    metadata: metadata, issues: issues };
}

/**
 * Locate the column header row: the first row (within HEADER_SCAN_ROWS) carrying
 * both the Start and Cost header names. Requiring both means a metadata row such
 * as `Funding start | 10/Aug/25` cannot be mistaken for it.
 */
function findHeaderRow_(data) {
  const startName = clean_(CONFIG.BUDGET_COLUMNS.start).toLowerCase();
  const costName = clean_(CONFIG.BUDGET_COLUMNS.cost).toLowerCase();
  const limit = Math.min(data.length, CONFIG.HEADER_SCAN_ROWS || 40);
  for (let i = 0; i < limit; i++) {
    const row = (data[i] || []).map(h => clean_(h).toLowerCase());
    if (row.indexOf(startName) !== -1 && row.indexOf(costName) !== -1) return i;
  }
  return -1;
}

/**
 * Read the Funding_info tab as `key | value` pairs from columns A and B.
 * Returns null when the tab is absent, so the caller can fall back to a metadata
 * block above the Budget columns. Rows with no value in column B - the tab title,
 * any legend, section headings - are skipped.
 */
function parseFundingInfoTab_(ss) {
  const sheet = ss.getSheetByName(CONFIG.FUNDING_INFO_TAB);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  const meta = {};
  for (let i = 0; i < data.length; i++) {
    const row = data[i] || [];
    const key = clean_(row[0]);
    if (!key) continue;
    const raw = row.length > 1 ? row[1] : '';
    const val = (raw instanceof Date) ? raw : clean_(raw);
    if (val === '') continue;
    meta[key.toLowerCase()] = val;
  }
  return meta;
}

/** Read the `key | value` block above the header row into a lower-cased map. */
function readMetadataBlock_(data, headerRow) {
  const meta = {};
  for (let i = 0; i < headerRow; i++) {
    const row = data[i] || [];
    const key = clean_(row[0]);
    if (!key) continue;
    const raw = row.length > 1 ? row[1] : '';
    meta[key.toLowerCase()] = (raw instanceof Date) ? raw : clean_(raw);
  }
  return meta;
}

/** yyyy-MM-dd without depending on Utilities/Session, so it stays unit-testable. */
function isoDate_(d) {
  const p = n => (n < 10 ? '0' : '') + n;
  return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
}

/** Parse "$5,600" / 5600 / "" into a Number (0 when blank/unparseable). */
function parseAmount_(value) {
  if (typeof value === 'number') return value;
  const n = parseFloat(String(value == null ? '' : value).replace(/[^0-9.\-]+/g, ''));
  return isNaN(n) ? 0 : n;
}

/** Trim and strip stray zero-width / non-breaking spaces seen in some cells. */
function clean_(value) {
  return String(value == null ? '' : value)
    .replace(/[​-‍﻿ ]/g, '')
    .trim();
}

/** Accepts Date objects, DD/MMM/YY strings, and standard date strings. */
function parseSheetDate_(value) {
  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === 'string' && value.trim()) {
    const m = /^(\d{2})\/([A-Za-z]{3})\/(\d{2})$/.exec(value.trim());
    if (m) {
      const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                      'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
      const month = months.indexOf(m[2]);
      if (month === -1) return null;
      return new Date(2000 + parseInt(m[3], 10), month, parseInt(m[1], 10));
    }
    const d = new Date(value);
    if (!isNaN(d)) return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  return null;
}

// ---- Forecast tab reader --------------------------------------------------

/**
 * Read the "Forecast" tab from a funding source spreadsheet.
 * Returns { data: { cost, income, comments }, sheetUrl }.
 * The Forecast tab has Revenue and Expenses sections with item codes in column A
 * and quarter forecasts in subsequent columns. Stops at "Funding Source Details".
 */
function parseForecastTab_(ss) {
  const sheetUrl = ss.getUrl();
  const sheet = ss.getSheetByName(CONFIG.FORECAST_TAB);
  const issues = [];
  const empty = { data: { cost: {}, income: {}, comments: {} },
    sheetUrl: sheetUrl, issues: issues };
  // An absent or empty Forecast tab is a valid state, not a finding: a quarter
  // with no override falls back to the budget baseline.
  if (!sheet) return empty;

  const data = sheet.getDataRange().getValues();
  if (!data.length) return empty;

  const result = { cost: {}, income: {}, comments: {} };
  var section = null; // 'revenue' | 'expenses'
  var quarterCols = []; // [{ col, label }]
  var commentCol = -1;

  for (var i = 0; i < data.length; i++) {
    var cellA = clean_(String(data[i][0] || ''));

    // Stop at "Funding Source Details"
    if (cellA === 'Funding Source Details') break;

    // Detect section headers
    if (cellA === 'Revenue' || cellA === 'Expenses') {
      section = cellA.toLowerCase();
      quarterCols = [];
      commentCol = -1;
      for (var j = 1; j < data[i].length; j++) {
        var header = clean_(String(data[i][j] || ''));
        if (!header) continue;
        if (header.toLowerCase() === 'comments') { commentCol = j; continue; }
        var ql = parseQuarterHeader_(header);
        if (ql) { quarterCols.push({ col: j, label: ql }); continue; }
        // A header that does not match "MMM-MMM YY Forecast" is ignored, which
        // silently discards a whole quarter of forecast. Say so.
        issues.push({ check: 'A7', row: i + 1, detail: 'Forecast column header "' +
          header + '" does not match "MMM-MMM YY Forecast" (e.g. "Jul-Sep 26 ' +
          'Forecast") - that column is ignored' });
      }
      continue;
    }

    if (!section) continue;
    // Skip blank, Total, and summary rows
    if (!cellA || cellA.indexOf('Total') === 0) continue;

    // Extract item code from "CODE - Name" format
    var code = itemCode_(cellA);
    if (!code) {
      issues.push({ check: 'A8', row: i + 1, detail: 'Forecast row "' + cellA +
        '" is not a "CODE - Name" milestone, so it is ignored' });
      continue;
    }

    var bucket = (section === 'revenue') ? result.income : result.cost;
    quarterCols.forEach(function (qc) {
      var val = data[i][qc.col];
      if (val !== null && val !== '' && !isNaN(Number(val))) {
        var key = code + '||' + qc.label;
        bucket[key] = (bucket[key] || 0) + Number(val);
      }
    });

    if (commentCol >= 0) {
      var comment = clean_(String(data[i][commentCol] || ''));
      if (comment) result.comments[code] = comment;
    }
  }

  return { data: result, sheetUrl: sheetUrl, issues: issues };
}

/**
 * Parse a forecast tab header like "Jul-Sep 26 Forecast" into an FY quarter
 * label like "26/27 Q2". Returns null if the header doesn't match.
 */
function parseQuarterHeader_(header) {
  var m = /^([A-Za-z]{3})-[A-Za-z]{3}\s+(\d{2})\s+Forecast$/i.exec(header);
  if (!m) return null;
  var months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  var monthIdx = months.indexOf(m[1].toLowerCase());
  if (monthIdx === -1) return null;
  var year = 2000 + parseInt(m[2], 10);
  return labelOfQi_(qiOfDate_(new Date(year, monthIdx, 1)));
}
