/**
 * ForecastStore.js
 * Read/write access to the single central "Forecast" Google Sheet that replaces
 * the per-funding-source "Budget, Actual, Forecast Tracking" tab.
 *
 * One row per (funding source, item, quarter) holding the GM's forward forecast
 * of *spend* for that milestone in that quarter. Sparse: only quarters the GM has
 * explicitly set appear; everything else defaults to the frozen baseline budget.
 *
 * Columns (see CONFIG.FORECAST.HEADER):
 *   Funding Source | Milestone | Item | Quarter | Forecast Cost | Updated By | Updated At
 */

function getForecastSheet_() {
  const ss = openOrCreateForecastSpreadsheet_();
  let sheet = ss.getSheetByName(CONFIG.FORECAST.TAB);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.FORECAST.TAB);
    sheet.appendRow(CONFIG.FORECAST.HEADER);
    sheet.setFrozenRows(1);
  } else {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headers[5] !== 'Forecast Income') {
      sheet.insertColumnsBefore(6, 2);
      sheet.getRange(1, 1, 1, CONFIG.FORECAST.HEADER.length).setValues([CONFIG.FORECAST.HEADER]);
    }
  }
  return sheet;
}

/**
 * Open the configured Forecast spreadsheet, or auto-create one inside the Budgets
 * root folder if none is configured/found, persisting its id for next time.
 */
function openOrCreateForecastSpreadsheet_() {
  const id = getForecastSheetId();
  if (id && id !== 'PUT_FORECAST_SPREADSHEET_ID_HERE') {
    try { return SpreadsheetApp.openById(id); }
    catch (e) { /* id stale/inaccessible - fall through and recreate */ }
  }
  return createForecastSpreadsheet_();
}

function createForecastSpreadsheet_() {
  const ss = SpreadsheetApp.create(CONFIG.FORECAST.FILE_NAME);
  // Move it from My Drive into the Budgets root folder so it lives alongside the
  // budgets it forecasts.
  try {
    const folder = DriveApp.getFolderById(CONFIG.BUDGETS_ROOT_FOLDER_ID);
    DriveApp.getFileById(ss.getId()).moveTo(folder);
  } catch (e) {
    Logger.log('Forecast sheet created but could not move to Budgets folder: ' + e.message);
  }
  setSecret('COCKPIT_FORECAST_SHEET_ID', ss.getId());
  Logger.log('Created Forecast sheet ' + ss.getId() + ' ("' + CONFIG.FORECAST.FILE_NAME + '").');
  return ss;
}

function forecastKey_(source, item, quarter) {
  return source + '||' + item + '||' + quarter;
}

function commentKey_(source, item) { return source + '||' + item; }

// Column indices (0-based) within a Forecast sheet row.
const FC_COST = 4, FC_INCOME = 5, FC_COMMENT = 6, FC_USER = 7, FC_AT = 8;

/**
 * @return {{ amounts: { 'source||item||quarter': { cost, income } },
 *            comments: { 'source||item': comment } }}
 * Amount rows have a Quarter; comment rows have a blank Quarter.
 */
function getForecastMap() {
  const sheet = getForecastSheet_();
  const data = sheet.getDataRange().getValues();
  const amounts = {};
  const comments = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const source = String(row[0] || '').trim();
    const item = String(row[2] || '').trim();
    const quarter = String(row[3] || '').trim();
    const comment = String(row[FC_COMMENT] || '').trim();
    if (!source) continue;
    if (quarter) {
      amounts[forecastKey_(source, item, quarter)] = {
        cost: numOrNull_(row[FC_COST]), income: numOrNull_(row[FC_INCOME]) };
    } else if (comment) {
      comments[commentKey_(source, item)] = comment;
    }
  }
  return { amounts: amounts, comments: comments };
}

function numOrNull_(v) {
  if (v === '' || v === null || typeof v === 'undefined') return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

/** Find all 1-based sheet rows for a (source, item, quarter) match. */
function findForecastRows_(data, source, item, quarter) {
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === source &&
        String(data[i][2]).trim() === item &&
        String(data[i][3]).trim() === quarter) rows.push(i + 1);
  }
  return rows;
}

/**
 * Insert or update one forecast amount cell for a given measure ('cost' or
 * 'income'). A null/'' value clears that measure; the row is removed only when
 * both cost and income are then empty so the cell reverts to baseline.
 */
function upsertForecast(source, milestone, item, quarter, measure, value, user) {
  const sheet = getForecastSheet_();
  const data = sheet.getDataRange().getValues();
  user = user || (Session.getActiveUser().getEmail() || 'unknown');
  
  const rowIndices = findForecastRows_(data, source, item, quarter);
  const rowIndex = rowIndices.length > 0 ? rowIndices[0] : -1;
  const existing = rowIndex !== -1 ? data[rowIndex - 1] : null;

  let cost = existing ? numOrNull_(existing[FC_COST]) : null;
  let income = existing ? numOrNull_(existing[FC_INCOME]) : null;
  const cleared = (value === null || value === '' || typeof value === 'undefined');
  const num = cleared ? null : (Number(value) || 0);
  if (measure === 'income') income = num; else cost = num;

  for (let i = rowIndices.length - 1; i > 0; i--) {
    sheet.deleteRow(rowIndices[i]);
  }

  if (cost === null && income === null) { // nothing left -> revert to baseline
    if (rowIndex !== -1) sheet.deleteRow(rowIndex);
    return { source, item, quarter, cost: null, income: null };
  }
  const rowValues = [source, milestone, item, quarter,
    cost === null ? '' : cost, income === null ? '' : income, '', user, new Date()];
  if (rowIndex === -1) sheet.appendRow(rowValues);
  else sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  return { source, item, quarter, cost: cost, income: income };
}

/**
 * Insert/update a milestone's Comment (stored on a row with a blank Quarter).
 * An empty comment clears the row.
 */
function upsertForecastComment(source, milestone, item, comment, user) {
  const sheet = getForecastSheet_();
  const data = sheet.getDataRange().getValues();
  user = user || (Session.getActiveUser().getEmail() || 'unknown');
  
  const rowIndices = findForecastRows_(data, source, item, ''); // blank quarter = comment row
  const rowIndex = rowIndices.length > 0 ? rowIndices[0] : -1;

  for (let i = rowIndices.length - 1; i > 0; i--) {
    sheet.deleteRow(rowIndices[i]);
  }

  const text = String(comment == null ? '' : comment).trim();
  if (!text) {
    if (rowIndex !== -1) sheet.deleteRow(rowIndex);
    return { source, item, comment: '' };
  }
  const rowValues = [source, milestone, item, '', '', '', text, user, new Date()];
  if (rowIndex === -1) sheet.appendRow(rowValues);
  else sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  return { source, item, comment: text };
}
