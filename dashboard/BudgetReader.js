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
 * @return {Array<{name, status, projectFolder, lines, hasProjectColumn}>}
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
    try {
      const parsed = parseBudgetFile_(file, projectFolder.getName());
      out.push({ name: file.getName(), status: status,
        projectFolder: projectFolder.getName(), lines: parsed.lines,
        hasProjectColumn: parsed.hasProjectColumn });
    } catch (e) {
      Logger.log('Skipped ' + file.getName() + ': ' + e.message);
    }
  }
}

function parseBudgetFile_(file, projectFolderName) {
  const sheet = SpreadsheetApp.openById(file.getId()).getSheetByName(CONFIG.BUDGET_TAB);
  if (!sheet) throw new Error('no "' + CONFIG.BUDGET_TAB + '" tab');
  const data = sheet.getDataRange().getValues();
  const header = data[0];
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
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const cost = parseAmount_(col.cost !== -1 ? row[col.cost] : 0);
    const income = parseAmount_(col.income !== -1 ? row[col.income] : 0);
    if (cost === 0 && income === 0) continue; // skip blank/summary rows
    const start = parseSheetDate_(row[col.start]);
    const end = parseSheetDate_(row[col.end]);
    if (!start || !end) continue;

    // Project attribution: a line's `Project` value wins when set; otherwise
    // (no column, or blank cell) the line belongs to its parent project folder.
    const project = hasProjectColumn && row[col.project]
      ? clean_(row[col.project])
      : (projectFolderName || CONFIG.DEFAULT_PROJECT);

    lines.push({
      description: col.description !== -1 ? clean_(row[col.description]) : '',
      start: start, end: end,
      cost: cost, income: income,
      contribution: col.contribution !== -1 ? parseAmount_(row[col.contribution]) : (income - cost),
      account: col.account !== -1 ? clean_(row[col.account]) : '',
      milestone: col.milestone !== -1 ? clean_(row[col.milestone]) : '',
      item: col.item !== -1 ? clean_(row[col.item]) : '',
      project: project
    });
  }
  return { lines: lines, hasProjectColumn: hasProjectColumn };
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
