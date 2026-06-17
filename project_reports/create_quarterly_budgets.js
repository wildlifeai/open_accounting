// Quarterly Budgets Integration module for Google Sheets
// This script runs via a loader to keep credentials private.

function QuarterlyBudgetsIntegration(PRIVATE_CONFIG) {

  // Configuration object
  const CONFIG = {
    BUDGETS_FOLDER_ID: "10105co6S5qHFSVVg0pb0fkoPidN3ScJZ",
    PROJECTS: {
      "Spyfish Aotearoa": "SPY",
      "General": "GEN",
      "Wild About AI": "WAI",
      "Wildlife Watcher": "WW"
    },
    TARGET_SHEET_NAME: "Budget, Actual, Forecast Tracking",
    XERO_SHEET_ID: "1VHtZsZRzJJ29tt3SRDXnIP6ebKChoDtlHqyV5Ua-3Yg",
    ACCOUNT_COL_INDEX: 9, // Column J in the Xero chart of accounts sheet
    COLUMNS_TO_EXTRACT: [6, 8, 9, 10, 11] // 0-indexed: G(6), I(8), J(9), K(10), L(11)
  };

  // ===================== OAUTH2 & XERO API =====================
  function getXeroService() {
    const props = PropertiesService.getDocumentProperties() || PropertiesService.getScriptProperties();
    return OAuth2.createService('xero')
      .setAuthorizationBaseUrl('https://login.xero.com/identity/connect/authorize')
      .setTokenUrl('https://identity.xero.com/connect/token')
      .setClientId(PRIVATE_CONFIG.CLIENT_ID)
      .setClientSecret(PRIVATE_CONFIG.CLIENT_SECRET)
      .setCallbackFunction('authCallback')
      .setPropertyStore(props)
      .setScope('offline_access accounting.settings.read accounting.reports.read accounting.transactions.read')
      .setParam('response_type', 'code')
      .setTokenHeaders({
        Authorization: 'Basic ' + Utilities.base64Encode(PRIVATE_CONFIG.CLIENT_ID + ':' + PRIVATE_CONFIG.CLIENT_SECRET)
      });
  }

  function showAuthorizationUrl() {
    Logger.log("Open this URL to authorize Xero: ");
    Logger.log(getXeroService().getAuthorizationUrl());
  }

  function authCallback(request) {
    const isAuthorized = getXeroService().handleCallback(request);
    if (isAuthorized) return HtmlService.createHtmlOutput('Success! You can close this tab.');
    return HtmlService.createHtmlOutput('Denied.');
  }

  function getTenantId(service) {
    const res = UrlFetchApp.fetch('https://api.xero.com/connections', {
      headers: { Authorization: 'Bearer ' + service.getAccessToken() }
    });
    const data = JSON.parse(res.getContentText());
    if (!data.length) throw new Error('No Xero tenant found');
    return data[0].tenantId;
  }

  function fetchXeroBudgets(service, tenantId) {
    Logger.log("Fetching current budgets from Xero API...");
    const url = 'https://api.xero.com/api.xro/2.0/Budgets';
    const response = UrlFetchApp.fetch(url, {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken(),
        'xero-tenant-id': tenantId
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() >= 400) {
      Logger.log("Failed to fetch budgets: " + response.getContentText());
      return [];
    }
    
    const data = JSON.parse(response.getContentText());
    return data.Budgets || [];
  }

  // ===================== MAIN SCRIPT =====================
  function generateAndUploadQuarterlyBudgets() {
    Logger.log("Starting Quarterly Budgets Generation...");
    
    const service = getXeroService();
    if (!service.hasAccess()) {
      Logger.log("⚠️ Xero not authorized. Please run showAuthorizationUrl() first to authorize Xero API access.");
      Logger.log("The script will generate CSVs, but will skip checking what needs to be updated.");
    }
    
    let xeroBudgetMap = null;
    if (service.hasAccess()) {
      try {
        const tenantId = getTenantId(service);
        const xeroBudgets = fetchXeroBudgets(service, tenantId);
        
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
      } catch(e) {
        Logger.log(`⚠️ Error connecting to Xero API: ${e.message}. Will skip diffing.`);
      }
    }

    const validAccounts = fetchValidXeroAccounts();
    if (!validAccounts || validAccounts.length === 0) {
      Logger.log("No valid accounts found. Exiting.");
      return;
    }

    const budgetsFolder = DriveApp.getFolderById(CONFIG.BUDGETS_FOLDER_ID);
    const budgetsToUpdate = [];

    for (const projectName of Object.keys(CONFIG.PROJECTS)) {
      const projectAbbr = CONFIG.PROJECTS[projectName];
      Logger.log(`\n=== Processing Project: ${projectName} (${projectAbbr}) ===`);

      const projectFolder = getFolderByName(budgetsFolder, projectName);
      if (!projectFolder) {
        Logger.log(`Project folder not found: ${projectName}`);
        continue;
      }

      const securedFolder = getFolderByName(projectFolder, "secured");
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

        const spreadsheet = SpreadsheetApp.openById(file.getId());
        const sheet = spreadsheet.getSheetByName(CONFIG.TARGET_SHEET_NAME);

        if (!sheet) {
          Logger.log(`    Missing target sheet '${CONFIG.TARGET_SHEET_NAME}'. Skipping.`);
          continue;
        }

        const data = sheet.getDataRange().getValues();
        if (data.length < 4) {
          Logger.log(`    Sheet too small to contain required data. Skipping.`);
          continue;
        }

        let headerRowIndex = -1;
        let accountColIndex = -1;

        for (let i = 0; i < data.length; i++) {
          for (let j = 0; j < data[i].length; j++) {
            const cellStr = String(data[i][j]).trim();
            if (cellStr === "*Account" || cellStr.includes("*Account") || cellStr === "Account") {
              headerRowIndex = i;
              accountColIndex = j;
              break;
            }
          }
          if (headerRowIndex !== -1) break;
        }

        if (headerRowIndex === -1) {
          Logger.log(`    Could not find '*Account' header in any cell. Skipping.`);
          continue;
        }

        const dateRow = data[3];

        if (!projectData[fundingSource]) {
          projectData[fundingSource] = {};
        }

        CONFIG.COLUMNS_TO_EXTRACT.forEach(colIndex => {
          const dateVal = dateRow[colIndex];
          if (dateVal && dateVal !== "") {
            const quarterInfo = getQuarterInfo(dateVal);
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

          if (!accountName || accountName === "" || !validAccounts.includes(accountName)) {
            continue;
          }

          CONFIG.COLUMNS_TO_EXTRACT.forEach(colIndex => {
            const dateVal = dateRow[colIndex];
            const val = parseFloat(String(row[colIndex]).replace(/[^0-9.-]+/g, ""));

            if (dateVal && !isNaN(val) && val !== 0) {
              const quarterInfo = getQuarterInfo(dateVal);
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
      } 

      if (!earliestMonth || !latestMonth) {
        Logger.log(`  No valid dates found in columns G, I, J, K, L. Cannot create summary sheet.`);
        continue;
      }

      const timelineMonths = generateTimeline(earliestMonth, latestMonth);
      const currQuarterInfo = getQuarterInfo(new Date());
      const spreadsheetName = `${currQuarterInfo.yy}${currQuarterInfo.q}_${projectAbbr}`;

      Logger.log(`\n  Consolidating into Summary Spreadsheet: ${spreadsheetName}`);
      const summarySpreadsheet = getOrCreateSummarySpreadsheet(projectFolder, spreadsheetName);

      const overallProjectData = {};

      for (const fundingSource of Object.keys(projectData)) {
        const fData = projectData[fundingSource];
        
        updateFundingSourceTabContinuous(summarySpreadsheet, fundingSource, validAccounts, fData, timelineMonths);

        const budgetDescription = `${spreadsheetName}_${fundingSource}`;
        const csvFileName = `${budgetDescription}_XERO_IMPORT.csv`;
        generateXeroBudgetCSV(projectFolder, csvFileName, fData, validAccounts, timelineMonths);
        
        if (xeroBudgetMap) {
          if (compareBudget(budgetDescription, fData, validAccounts, timelineMonths, xeroBudgetMap)) {
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
      
      updateFundingSourceTabContinuous(summarySpreadsheet, "OVERALL_PROJECT", validAccounts, overallProjectData, timelineMonths);
      generateXeroBudgetCSV(projectFolder, overallCsvFileName, overallProjectData, validAccounts, timelineMonths);
      
      if (xeroBudgetMap) {
        if (compareBudget(overallDescription, overallProjectData, validAccounts, timelineMonths, xeroBudgetMap)) {
          budgetsToUpdate.push(overallCsvFileName);
        }
      }
    }

    Logger.log("\n=======================================================");
    Logger.log("                BUDGET UPLOAD CHECKLIST                 ");
    Logger.log("=======================================================");
    if (!service.hasAccess() || xeroBudgetMap === null) {
      Logger.log("⚠️ Xero API not connected. Could not perform budget diffing.");
      Logger.log("Please setup Xero OAuth to enable automated checklist.");
    } else if (budgetsToUpdate.length === 0) {
      Logger.log("All Xero budgets are perfectly up-to-date! No CSV uploads required.");
    } else {
      Logger.log(`Found ${budgetsToUpdate.length} budgets that differ from Xero. Please upload the following CSVs:`);
      budgetsToUpdate.forEach(b => Logger.log(`  [UPDATE REQUIRED] -> ${b}`));
    }
    Logger.log("=======================================================");
  }

  // ===================== HELPER FUNCTIONS =====================
  function compareBudget(budgetName, accountData, validAccounts, timelineMonths, xeroBudgetMap) {
    if (!xeroBudgetMap[budgetName]) return true; 
    
    const xeroBudget = xeroBudgetMap[budgetName];
    let needsUpdate = false;
    
    for (let i = 0; i < validAccounts.length; i++) {
      const acc = validAccounts[i];
      const accData = accountData[acc] || {};
      
      const hasValues = timelineMonths.some(m => (accData[m.getTime()] || 0) !== 0);
      if (!hasValues) continue; 
      
      const match = acc.match(/^(.*?)\s*\((\d+)\)$/);
      const accountCode = match ? match[2] : "";
      
      for (let j = 0; j < timelineMonths.length; j++) {
        const m = timelineMonths[j];
        const generatedAmount = accData[m.getTime()] || 0;
        const periodStr = Utilities.formatDate(m, "UTC", "yyyy-MM");
        
        const xeroAmount = (xeroBudget[accountCode] && xeroBudget[accountCode][periodStr]) ? xeroBudget[accountCode][periodStr] : 0;
        
        if (Math.abs(generatedAmount - xeroAmount) > 0.01) {
          needsUpdate = true;
          break;
        }
      }
      if (needsUpdate) break;
    }
    
    return needsUpdate;
  }

  function generateXeroBudgetCSV(parentFolder, fileName, accountData, validAccounts, timelineMonths) {
    Logger.log(`\n    --- Generating Xero CSV Budget: ${fileName} ---`);
    
    const tz = Session.getScriptTimeZone();
    const headers = ["*Account"];
    timelineMonths.forEach(m => {
      headers.push(Utilities.formatDate(m, tz, "MMM-yyyy"));
    });
    
    const csvLines = [];
    csvLines.push(headers.join(","));
    
    let rowsAdded = 0;
    
    validAccounts.forEach(acc => {
      const accData = accountData[acc] || {};
      const hasValues = timelineMonths.some(m => (accData[m.getTime()] || 0) !== 0);
      
      if (hasValues) {
        let accountName = acc;
        if (accountName.includes(",")) {
          accountName = `"${accountName}"`;
        }
        const row = [accountName];
        timelineMonths.forEach(m => {
          row.push(accData[m.getTime()] || 0);
        });
        csvLines.push(row.join(","));
        rowsAdded++;
      }
    });
    
    if (rowsAdded > 0) {
      const csvContent = csvLines.join("\n");
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

  function updateFundingSourceTabContinuous(spreadsheet, tabName, validAccounts, accountData, timelineMonths) {
    let sheet = spreadsheet.getSheetByName(tabName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(tabName);
    }
    sheet.clearContents();
    const tz = Session.getScriptTimeZone();
    const headers = ["*Account"];
    timelineMonths.forEach(m => {
      headers.push(Utilities.formatDate(m, tz, "MMM yyyy"));
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
    Logger.log(`    Updated tab: ${tabName} with ${rows.length - 1} accounts across ${timelineMonths.length} months.`);
  }

  function generateTimeline(startMonth, endMonth) {
    const months = [];
    const current = new Date(startMonth.getFullYear(), startMonth.getMonth(), 15);
    const end = new Date(endMonth.getFullYear(), endMonth.getMonth(), 15);
    while (current <= end) {
      months.push(new Date(current));
      current.setMonth(current.getMonth() + 1);
    }
    return months;
  }

  function getFolderByName(parentFolder, name) {
    const folders = parentFolder.getFoldersByName(name);
    if (folders.hasNext()) return folders.next();
    return null;
  }

  function getOrCreateSummarySpreadsheet(parentFolder, spreadsheetName) {
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

  function getQuarterInfo(dateValue) {
    let date;
    if (dateValue instanceof Date) date = dateValue;
    else if (typeof dateValue === 'string') date = new Date(dateValue);
    if (!date || isNaN(date)) return null;
    
    const month = date.getMonth();
    const year = date.getFullYear();
    
    let q = "";
    let fyYear = year;
    let startMonthOffset = 0;

    if (month >= 3 && month <= 5) { q = "Q1"; startMonthOffset = 3; }
    else if (month >= 6 && month <= 8) { q = "Q2"; startMonthOffset = 6; }
    else if (month >= 9 && month <= 11) { q = "Q3"; startMonthOffset = 9; }
    else if (month >= 0 && month <= 2) {
      q = "Q4";
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

  function fetchValidXeroAccounts() {
    try {
      const xeroSheet = SpreadsheetApp.openById(CONFIG.XERO_SHEET_ID).getActiveSheet();
      const xeroData = xeroSheet.getDataRange().getValues();
      const validAccounts = xeroData
        .map(row => row[CONFIG.ACCOUNT_COL_INDEX])
        .filter(account => account && account !== "*Account" && !account.startsWith("Less Accumulated Depreciation") && account !== "Account Name");
      return validAccounts.filter(a => a.trim().length > 0);
    } catch (e) {
      Logger.log(`Error fetching valid Xero accounts: ${e.message}`);
      return [];
    }
  }

  // Export functions to the loader
  return {
    generateAndUploadQuarterlyBudgets,
    showAuthorizationUrl,
    authCallback
  };
}
