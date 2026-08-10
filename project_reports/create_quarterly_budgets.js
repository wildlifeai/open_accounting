// Quarterly Budgets Integration for Google Apps Script.
//
// This runs as an ordinary Apps Script file. There is no remote loader: the code
// lives in the project and is deployed with clasp, and credentials live in Script
// Properties rather than in source.
//
// Setup
//   1. Add the OAuth2 library as `OAuth2` (Script ID
//      1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF).
//   2. Project Settings > Script Properties, add:
//        XERO_CLIENT_ID      - from your Xero app
//        XERO_CLIENT_SECRET  - from your Xero app
//   3. Run logXeroRedirectUri() and register the logged URI in the Xero app.
//   4. Run showAuthorizationUrl(), open the URL, approve the organisation.
//   5. Run generateAndUploadQuarterlyBudgets().
//
// Internal helpers carry a trailing underscore. Apps Script treats those as
// private, which keeps the editor's Run menu to the real entry points and avoids
// clashing with same-named helpers in sibling files (funding-aggregator.js also
// defines a getFolderByName).

const BUDGETS_CONFIG = {
  BUDGETS_FOLDER_ID: '10105co6S5qHFSVVg0pb0fkoPidN3ScJZ',

  // Project folder name -> abbreviation used in generated sheet names.
  // These MUST stay as SPY/GEN/WAI/WW: they match project_reports/README.md and,
  // more importantly, the funding-source file names already in the Budgets Drive
  // (WAI_25_OMV, WW_25_TOI, WW_26_SALES, SPY_26_UOA) and the corresponding Xero
  // "Funding source" tracking values. Renaming them decouples generated budgets
  // from the real data.
  PROJECTS: {
    'Spyfish Aotearoa': 'SPY',
    'General': 'GEN',
    'Wild About AI': 'WAI',
    'Wildlife Watcher': 'WW'
  },

  TARGET_SHEET_NAME: 'Budget, Actual, Forecast Tracking',
  XERO_SHEET_ID: '1VHtZsZRzJJ29tt3SRDXnIP6ebKChoDtlHqyV5Ua-3Yg',
  ACCOUNT_COL_INDEX: 9,               // Column J in the Xero chart of accounts sheet
  COLUMNS_TO_EXTRACT: [6, 8, 9, 10, 11], // 0-indexed: G, I, J, K, L

  // Google Sheets tab-name constraints.
  MAX_TAB_NAME_LENGTH: 31,
  INVALID_TAB_CHARS: /[*?:\\/\[\]]/g
};


// ===================== SECRETS =====================

/** Read a secret from Script Properties (returns '' if unset). */
function getSecret_(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}


// ===================== OAUTH2 & XERO API =====================

function getXeroService_() {
  const clientId = getSecret_('XERO_CLIENT_ID');
  const clientSecret = getSecret_('XERO_CLIENT_SECRET');
  if (!clientId || !clientSecret) {
    throw new Error('XERO_CLIENT_ID / XERO_CLIENT_SECRET are not set. Add them in ' +
      'Project Settings > Script Properties.');
  }
  return OAuth2.createService('xero')
    .setAuthorizationBaseUrl('https://login.xero.com/identity/connect/authorize')
    .setTokenUrl('https://identity.xero.com/connect/token')
    .setClientId(clientId)
    .setClientSecret(clientSecret)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setScope('offline_access accounting.settings.read accounting.reports.read accounting.transactions.read')
    .setParam('response_type', 'code')
    .setTokenHeaders({
      Authorization: 'Basic ' + Utilities.base64Encode(clientId + ':' + clientSecret)
    });
}

/** Run once and register the logged URI as a redirect URI in the Xero app. */
function logXeroRedirectUri() {
  Logger.log('Register this redirect URI in your Xero app:');
  Logger.log(getXeroService_().getRedirectUri());
}

/** 1. Run this first, then open the logged URL to authorise Xero. */
function showAuthorizationUrl() {
  Logger.log('Open this URL to authorize Xero: ');
  Logger.log(getXeroService_().getAuthorizationUrl());
}

/** OAuth2 redirect handler. Must stay a bare global - OAuth2 calls it by name. */
function authCallback(request) {
  const isAuthorized = getXeroService_().handleCallback(request);
  if (isAuthorized) return HtmlService.createHtmlOutput('Success! You can close this tab.');
  return HtmlService.createHtmlOutput('Denied.');
}

function getTenantId_(service) {
  const res = UrlFetchApp.fetch('https://api.xero.com/connections', {
    headers: { Authorization: 'Bearer ' + service.getAccessToken() },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    throw new Error('Xero connections API ' + res.getResponseCode() + ': ' + res.getContentText());
  }
  const data = JSON.parse(res.getContentText());
  if (!data.length) throw new Error('No Xero tenant found');
  return data[0].tenantId;
}

function fetchXeroBudgets_(service, tenantId) {
  Logger.log('Fetching current budgets from Xero API...');
  const url = 'https://api.xero.com/api.xro/2.0/Budgets';
  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + service.getAccessToken(),
      'xero-tenant-id': tenantId
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() >= 400) {
    Logger.log('Failed to fetch budgets: ' + response.getContentText());
    return [];
  }

  const data = JSON.parse(response.getContentText());
  return data.Budgets || [];
}


// ===================== MAIN SCRIPT =====================

/** 2. Generate the CSV budgets and diff them against Xero. */
function generateAndUploadQuarterlyBudgets() {
  Logger.log('Starting Quarterly Budgets Generation...');

  const service = getXeroService_();
  if (!service.hasAccess()) {
    Logger.log('⚠️ Xero not authorized. Please run showAuthorizationUrl() first to authorize Xero API access.');
    Logger.log('The script will generate CSVs, but will skip checking what needs to be updated.');
  }

  let xeroBudgetMap = null;
  if (service.hasAccess()) {
    try {
      const tenantId = getTenantId_(service);
      const xeroBudgets = fetchXeroBudgets_(service, tenantId);

      xeroBudgetMap = {};
      xeroBudgets.forEach(b => {
        const desc = b.Description;
        xeroBudgetMap[desc] = {};
        (b.BudgetLines || []).forEach(line => {
          const code = line.AccountCode;
          xeroBudgetMap[desc][code] = {};
          (line.BudgetBalances || []).forEach(bal => {
            xeroBudgetMap[desc][code][bal.Period] = bal.Amount;
          });
        });
      });
      Logger.log(`Successfully mapped ${xeroBudgets.length} existing Xero budgets for diffing.`);
    } catch (e) {
      Logger.log(`⚠️ Error connecting to Xero API: ${e.message}. Will skip diffing.`);
    }
  }

  const validAccounts = fetchValidXeroAccounts_();
  if (!validAccounts || validAccounts.length === 0) {
    Logger.log('No valid accounts found. Exiting.');
    return;
  }

  const budgetsFolder = DriveApp.getFolderById(BUDGETS_CONFIG.BUDGETS_FOLDER_ID);
  const budgetsToUpdate = [];
  const skippedFiles = [];

  for (const projectName of Object.keys(BUDGETS_CONFIG.PROJECTS)) {
    const projectAbbr = BUDGETS_CONFIG.PROJECTS[projectName];
    Logger.log(`\n=== Processing Project: ${projectName} (${projectAbbr}) ===`);

    const projectFolder = getFolderByName_(budgetsFolder, projectName);
    if (!projectFolder) {
      Logger.log(`Project folder not found: ${projectName}`);
      continue;
    }

    const securedFolder = getFolderByName_(projectFolder, 'secured');
    if (!securedFolder) {
      Logger.log(`'secured' folder not found in ${projectName}`);
      continue;
    }

    const projectData = {};
    let earliestMonth = null;
    let latestMonth = null;

    const budgetFiles = securedFolder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (budgetFiles.hasNext()) {
      const file = budgetFiles.next();
      const fundingSource = file.getName();
      Logger.log(`  - Found funding source: ${fundingSource}`);

      // One unreadable or malformed file must not abort the whole run.
      try {
        const spreadsheet = SpreadsheetApp.openById(file.getId());
        const sheet = spreadsheet.getSheetByName(BUDGETS_CONFIG.TARGET_SHEET_NAME);

        if (!sheet) {
          Logger.log(`    Missing target sheet '${BUDGETS_CONFIG.TARGET_SHEET_NAME}'. Skipping.`);
          continue;
        }

        const data = sheet.getDataRange().getValues();
        if (data.length < 4) {
          Logger.log('    Sheet too small to contain required data. Skipping.');
          continue;
        }

        let headerRowIndex = -1;
        let accountColIndex = -1;

        for (let i = 0; i < data.length; i++) {
          for (let j = 0; j < data[i].length; j++) {
            const cellStr = String(data[i][j]).trim();
            if (cellStr === '*Account' || cellStr.includes('*Account') || cellStr === 'Account') {
              headerRowIndex = i;
              accountColIndex = j;
              break;
            }
          }
          if (headerRowIndex !== -1) break;
        }

        if (headerRowIndex === -1) {
          Logger.log("    Could not find '*Account' header in any cell. Skipping.");
          continue;
        }

        const dateRow = data[3];

        if (!projectData[fundingSource]) {
          projectData[fundingSource] = {};
        }

        BUDGETS_CONFIG.COLUMNS_TO_EXTRACT.forEach(colIndex => {
          const dateVal = dateRow[colIndex];
          if (dateVal && dateVal !== '') {
            const quarterInfo = getQuarterInfo_(dateVal);
            if (quarterInfo) {
              if (!earliestMonth || quarterInfo.startMonth < earliestMonth) {
                earliestMonth = quarterInfo.startMonth;
              }
              if (!latestMonth || quarterInfo.endMonth > latestMonth) {
                latestMonth = quarterInfo.endMonth;
              }
            }
          }
        });

        for (let i = headerRowIndex + 1; i < data.length; i++) {
          const row = data[i];
          const accountName = row[accountColIndex];

          if (!accountName || accountName === '' || !validAccounts.includes(accountName)) {
            continue;
          }

          BUDGETS_CONFIG.COLUMNS_TO_EXTRACT.forEach(colIndex => {
            const dateVal = dateRow[colIndex];
            const val = parseFloat(String(row[colIndex]).replace(/[^0-9.-]+/g, ''));

            if (dateVal && !isNaN(val) && val !== 0) {
              const quarterInfo = getQuarterInfo_(dateVal);
              if (quarterInfo) {
                if (!projectData[fundingSource][accountName]) {
                  projectData[fundingSource][accountName] = {};
                }
                const timeKey = quarterInfo.middleMonth.getTime();
                if (!projectData[fundingSource][accountName][timeKey]) {
                  projectData[fundingSource][accountName][timeKey] = 0;
                }
                projectData[fundingSource][accountName][timeKey] += val;
              }
            }
          });
        }
      } catch (e) {
        Logger.log(`    ⚠️ Failed to process '${fundingSource}': ${e.message}. Skipping this file.`);
        skippedFiles.push(`${projectName} / ${fundingSource}: ${e.message}`);
        continue;
      }
    }

    if (!earliestMonth || !latestMonth) {
      Logger.log('  No valid dates found in columns G, I, J, K, L. Cannot create summary sheet.');
      continue;
    }

    const timelineMonths = generateTimeline_(earliestMonth, latestMonth);
    const currQuarterInfo = getQuarterInfo_(new Date());
    const spreadsheetName = `${currQuarterInfo.yy}${currQuarterInfo.q}_${projectAbbr}`;

    Logger.log(`\n  Consolidating into Summary Spreadsheet: ${spreadsheetName}`);
    const summarySpreadsheet = getOrCreateSummarySpreadsheet_(projectFolder, spreadsheetName);

    const overallProjectData = {};

    for (const fundingSource of Object.keys(projectData)) {
      const fData = projectData[fundingSource];

      updateFundingSourceTabContinuous_(summarySpreadsheet, fundingSource, validAccounts, fData, timelineMonths);

      const budgetDescription = `${spreadsheetName}_${fundingSource}`;
      const csvFileName = `${budgetDescription}_XERO_IMPORT.csv`;
      generateXeroBudgetCSV_(projectFolder, csvFileName, fData, validAccounts, timelineMonths);

      if (xeroBudgetMap) {
        if (compareBudget_(budgetDescription, fData, validAccounts, timelineMonths, xeroBudgetMap)) {
          budgetsToUpdate.push(csvFileName);
        }
      }

      for (const acc of Object.keys(fData)) {
        if (!overallProjectData[acc]) {
          overallProjectData[acc] = {};
        }
        for (const timeKey of Object.keys(fData[acc])) {
          if (!overallProjectData[acc][timeKey]) {
            overallProjectData[acc][timeKey] = 0;
          }
          overallProjectData[acc][timeKey] += fData[acc][timeKey];
        }
      }
    }

    const overallDescription = `${spreadsheetName}_OVERALL`;
    const overallCsvFileName = `${overallDescription}_XERO_IMPORT.csv`;

    updateFundingSourceTabContinuous_(summarySpreadsheet, 'OVERALL_PROJECT', validAccounts, overallProjectData, timelineMonths);
    generateXeroBudgetCSV_(projectFolder, overallCsvFileName, overallProjectData, validAccounts, timelineMonths);

    if (xeroBudgetMap) {
      if (compareBudget_(overallDescription, overallProjectData, validAccounts, timelineMonths, xeroBudgetMap)) {
        budgetsToUpdate.push(overallCsvFileName);
      }
    }
  }

  Logger.log('\n=======================================================');
  Logger.log('                BUDGET UPLOAD CHECKLIST                 ');
  Logger.log('=======================================================');
  if (!service.hasAccess() || xeroBudgetMap === null) {
    Logger.log('⚠️ Xero API not connected. Could not perform budget diffing.');
    Logger.log('Please setup Xero OAuth to enable automated checklist.');
  } else if (budgetsToUpdate.length === 0) {
    Logger.log('All Xero budgets are perfectly up-to-date! No CSV uploads required.');
  } else {
    Logger.log(`Found ${budgetsToUpdate.length} budgets that differ from Xero. Please upload the following CSVs:`);
    budgetsToUpdate.forEach(b => Logger.log(`  [UPDATE REQUIRED] -> ${b}`));
  }
  if (skippedFiles.length) {
    Logger.log('-------------------------------------------------------');
    Logger.log(`⚠️ ${skippedFiles.length} file(s) were skipped due to errors:`);
    skippedFiles.forEach(s => Logger.log(`  [SKIPPED] ${s}`));
  }
  Logger.log('=======================================================');
}


// ===================== HELPER FUNCTIONS =====================

function compareBudget_(budgetName, accountData, validAccounts, timelineMonths, xeroBudgetMap) {
  if (!xeroBudgetMap[budgetName]) return true;

  const xeroBudget = xeroBudgetMap[budgetName];
  let needsUpdate = false;

  for (let i = 0; i < validAccounts.length; i++) {
    const acc = validAccounts[i];
    const accData = accountData[acc] || {};

    const hasValues = timelineMonths.some(m => (accData[m.getTime()] || 0) !== 0);
    if (!hasValues) continue;

    // Xero budgets are keyed on the short account code. Without a code we cannot
    // compare, and defaulting to '' would silently report every period as a
    // difference - so skip and say so rather than emit a false positive.
    const match = String(acc).match(/^(.*?)\s*\((\d+)\)$/);
    if (!match) {
      Logger.log(`    ⚠️ Account "${acc}" has no code in parentheses. Skipping comparison.`);
      continue;
    }
    const accountCode = match[2];

    for (let j = 0; j < timelineMonths.length; j++) {
      const m = timelineMonths[j];
      const generatedAmount = accData[m.getTime()] || 0;
      const periodStr = Utilities.formatDate(m, 'UTC', 'yyyy-MM');

      const xeroAmount = (xeroBudget[accountCode] && xeroBudget[accountCode][periodStr])
        ? xeroBudget[accountCode][periodStr] : 0;

      if (Math.abs(generatedAmount - xeroAmount) > 0.01) {
        needsUpdate = true;
        break;
      }
    }
    if (needsUpdate) break;
  }

  return needsUpdate;
}

/**
 * Quote a value for CSV per RFC 4180: double any embedded quote, and wrap the
 * field when it contains a comma, quote, CR or LF. Wrapping alone (the previous
 * behaviour) produced files Xero could not parse whenever an account name
 * contained a quote character.
 */
function csvCell_(value) {
  const s = String(value == null ? '' : value);
  if (/[",\r\n]/.test(s)) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function generateXeroBudgetCSV_(parentFolder, fileName, accountData, validAccounts, timelineMonths) {
  Logger.log(`\n    --- Generating Xero CSV Budget: ${fileName} ---`);

  const tz = Session.getScriptTimeZone();
  const headers = ['*Account'];
  timelineMonths.forEach(m => {
    headers.push(Utilities.formatDate(m, tz, 'MMM-yyyy'));
  });

  const csvLines = [];
  csvLines.push(headers.map(csvCell_).join(','));

  let rowsAdded = 0;

  validAccounts.forEach(acc => {
    const accData = accountData[acc] || {};
    const hasValues = timelineMonths.some(m => (accData[m.getTime()] || 0) !== 0);

    if (hasValues) {
      const row = [csvCell_(acc)];
      timelineMonths.forEach(m => {
        row.push(accData[m.getTime()] || 0);
      });
      csvLines.push(row.join(','));
      rowsAdded++;
    }
  });

  if (rowsAdded > 0) {
    const csvContent = csvLines.join('\n');
    const existingFiles = parentFolder.getFilesByName(fileName);
    if (existingFiles.hasNext()) {
      const file = existingFiles.next();
      file.setContent(csvContent);
      Logger.log(`    Updated existing CSV file: ${fileName}`);
    } else {
      parentFolder.createFile(fileName, csvContent, MimeType.CSV);
      Logger.log(`    Created new CSV file: ${fileName}`);
    }
  } else {
    Logger.log(`    No non-zero values found. CSV generation skipped for ${fileName}`);
  }
}

/**
 * Google Sheets tab names are limited to 31 characters and cannot contain
 * * ? : \ / [ ]. Funding-source file names can exceed both, which made
 * insertSheet throw and aborted the run.
 */
function sanitizeTabName_(tabName) {
  const cleaned = String(tabName == null ? '' : tabName)
    .replace(BUDGETS_CONFIG.INVALID_TAB_CHARS, '')
    .trim()
    .substring(0, BUDGETS_CONFIG.MAX_TAB_NAME_LENGTH);
  return cleaned || 'UNNAMED';
}

function updateFundingSourceTabContinuous_(spreadsheet, tabName, validAccounts, accountData, timelineMonths) {
  const safeName = sanitizeTabName_(tabName);
  if (safeName !== tabName) {
    Logger.log(`    Tab name '${tabName}' sanitized to '${safeName}' (31 char / character limits).`);
  }

  let sheet = spreadsheet.getSheetByName(safeName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(safeName);
  }
  sheet.clearContents();
  const tz = Session.getScriptTimeZone();
  const headers = ['*Account'];
  timelineMonths.forEach(m => {
    headers.push(Utilities.formatDate(m, tz, 'MMM yyyy'));
  });
  const rows = [headers];

  validAccounts.forEach(acc => {
    const row = [acc];
    const accData = accountData[acc] || {};
    timelineMonths.forEach(m => {
      row.push(accData[m.getTime()] || 0);
    });
    rows.push(row);
  });

  sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
  Logger.log(`    Updated tab: ${safeName} with ${rows.length - 1} accounts across ${timelineMonths.length} months.`);
}

function generateTimeline_(startMonth, endMonth) {
  const months = [];
  const current = new Date(startMonth.getFullYear(), startMonth.getMonth(), 15);
  const end = new Date(endMonth.getFullYear(), endMonth.getMonth(), 15);
  while (current <= end) {
    months.push(new Date(current));
    current.setMonth(current.getMonth() + 1);
  }
  return months;
}

function getFolderByName_(parentFolder, name) {
  const folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return null;
}

function getOrCreateSummarySpreadsheet_(parentFolder, spreadsheetName) {
  const files = parentFolder.getFilesByName(spreadsheetName);
  if (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
      return SpreadsheetApp.openById(file.getId());
    }
  }
  const newSpreadsheet = SpreadsheetApp.create(spreadsheetName);
  DriveApp.getFileById(newSpreadsheet.getId()).moveTo(parentFolder);
  return newSpreadsheet;
}

function getQuarterInfo_(dateValue) {
  let date;
  if (dateValue instanceof Date) date = dateValue;
  else if (typeof dateValue === 'string') date = new Date(dateValue);
  if (!date || isNaN(date)) return null;

  const month = date.getMonth();
  const year = date.getFullYear();

  let q = '';
  let fyYear = year;
  let startMonthOffset = 0;

  if (month >= 3 && month <= 5) { q = 'Q1'; startMonthOffset = 3; }
  else if (month >= 6 && month <= 8) { q = 'Q2'; startMonthOffset = 6; }
  else if (month >= 9 && month <= 11) { q = 'Q3'; startMonthOffset = 9; }
  else if (month >= 0 && month <= 2) {
    q = 'Q4';
    fyYear = year - 1;
    startMonthOffset = 0;
  }

  const yy = fyYear.toString().slice(-2);
  const startMonthDate = new Date(year, startMonthOffset, 15);
  const middleMonthDate = new Date(year, startMonthOffset + 1, 15);
  const endMonthDate = new Date(year, startMonthOffset + 2, 15);

  return {
    q: q,
    yy: yy,
    startMonth: startMonthDate,
    middleMonth: middleMonthDate,
    endMonth: endMonthDate
  };
}

function fetchValidXeroAccounts_() {
  try {
    // getSheets()[0] rather than getActiveSheet(): "active" depends on whichever
    // tab a user last left selected, so getActiveSheet() can silently read the
    // wrong sheet and find no accounts.
    const xeroSheet = SpreadsheetApp.openById(BUDGETS_CONFIG.XERO_SHEET_ID).getSheets()[0];
    const xeroData = xeroSheet.getDataRange().getValues();
    const validAccounts = xeroData
      .map(row => row[BUDGETS_CONFIG.ACCOUNT_COL_INDEX])
      // Coerce before string tests: a numeric cell would throw on startsWith.
      .map(account => (account == null ? '' : String(account)))
      .filter(account => account &&
        account !== '*Account' &&
        account !== 'Account Name' &&
        !account.startsWith('Less Accumulated Depreciation'));
    return validAccounts.filter(a => a.trim().length > 0);
  } catch (e) {
    Logger.log(`Error fetching valid Xero accounts: ${e.message}`);
    return [];
  }
}
