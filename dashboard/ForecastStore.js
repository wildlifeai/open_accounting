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

/**
 * @return { 'source||item||quarter': { cost, milestone } } for every stored row.
 */
function getForecastMap() {
  const sheet = getForecastSheet_();
  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const source = String(row[0] || '').trim();
    const milestone = String(row[1] || '').trim();
    const item = String(row[2] || '').trim();
    const quarter = String(row[3] || '').trim();
    if (!source || !quarter) continue;
    map[forecastKey_(source, item, quarter)] = {
      cost: Number(row[4]) || 0, milestone: milestone
    };
  }
  return map;
}

/**
 * Insert or update one forecast cell. Returns the saved value.
 * `cost` of null/'' clears the override (row removed) so the cell reverts to
 * baseline.
 */
function upsertForecast(source, milestone, item, quarter, cost, user) {
  const sheet = getForecastSheet_();
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  user = user || (Session.getActiveUser().getEmail() || 'unknown');

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === source &&
        String(data[i][2]).trim() === item &&
        String(data[i][3]).trim() === quarter) { rowIndex = i + 1; break; }
  }

  const clearing = (cost === null || cost === '' || typeof cost === 'undefined');
  if (clearing) {
    if (rowIndex !== -1) sheet.deleteRow(rowIndex);
    return { source, item, quarter, cost: null };
  }

  const value = Number(cost) || 0;
  const rowValues = [source, milestone, item, quarter, value, user, now];
  if (rowIndex === -1) {
    sheet.appendRow(rowValues);
  } else {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  }
  return { source, item, quarter, cost: value };
}
