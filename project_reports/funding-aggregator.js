/**
 * Script to be stored in GitHub public repository
 * Aggregates funding data from PROJECT_YEAR_FUNDING sheets
 */

function aggregateFundingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const overviewSheet = ss.getSheetByName('Overview') || ss.getActiveSheet();
  
  // Get the parent folder of the current spreadsheet
  const file = DriveApp.getFileById(ss.getId());
  const parentFolder = file.getParents().next();
  
  // Get secured and proposed folders
  const securedFolder = getFolderByName(parentFolder, 'secured');
  const proposedFolder = getFolderByName(parentFolder, 'proposed');
  
  if (!securedFolder && !proposedFolder) {
    throw new Error('Neither "secured" nor "proposed" folders found in the same directory as this spreadsheet');
  }
  
  // Clear existing data (starting from row 2 to preserve headers)
  const lastRow = overviewSheet.getLastRow();
  if (lastRow > 1) {
    overviewSheet.getRange(2, 1, lastRow - 1, overviewSheet.getLastColumn()).clearContent();
  }
  
  // Set up headers if they don't exist
  const headers = ['Funding Source', 'Secured', 'Milestone', 'Cost', 'Income', 'Total Actual to Date'];
  overviewSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  
  // Collect all data
  const allData = [];
  
  // Process secured folder
  if (securedFolder) {
    const securedData = processFundingFolder(securedFolder, 'Yes');
    allData.push(...securedData);
  }
  
  // Process proposed folder
  if (proposedFolder) {
    const proposedData = processFundingFolder(proposedFolder, 'No');
    allData.push(...proposedData);
  }
  
  // Write all data to sheet
  if (allData.length > 0) {
    overviewSheet.getRange(2, 1, allData.length, headers.length).setValues(allData);
  }
  
  // Auto-resize columns
  for (let i = 1; i <= headers.length; i++) {
    overviewSheet.autoResizeColumn(i);
  }
  
  Logger.log(`Successfully processed ${allData.length} funding entries`);
}

/**
 * Get a folder by name within a parent folder
 */
function getFolderByName(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : null;
}

/**
 * Process all spreadsheets in a funding folder
 */
function processFundingFolder(folder, securedStatus) {
  const data = [];
  const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  
  while (files.hasNext()) {
    const file = files.next();
    const fundingSource = file.getName();
    
    try {
      const spreadsheet = SpreadsheetApp.openById(file.getId());
      const fundingData = extractFundingData(spreadsheet, fundingSource, securedStatus);
      data.push(...fundingData);
      
      Logger.log(`Processed: ${fundingSource} (${securedStatus})`);
    } catch (error) {
      Logger.log(`Error processing ${fundingSource}: ${error.toString()}`);
    }
  }
  
  return data;
}

/**
 * Extract funding data from a PROJECT_YEAR_FUNDING spreadsheet
 */
function extractFundingData(spreadsheet, fundingSource, securedStatus) {
  const data = [];
  
  // Try to get Budget sheet
  const budgetSheet = spreadsheet.getSheetByName('Budget');
  const trackingSheet = spreadsheet.getSheetByName('Budget, Actual, Forecast Tracking');
  
  if (!budgetSheet && !trackingSheet) {
    Logger.log(`Warning: No Budget or Tracking sheets found in ${fundingSource}`);
    return data;
  }
  
  // Extract data from Budget sheet
  if (budgetSheet) {
    const budgetData = extractFromBudgetSheet(budgetSheet, fundingSource, securedStatus, trackingSheet);
    data.push(...budgetData);
  }
  
  return data;
}

/**
 * Extract data from Budget sheet and match with Tracking sheet
 */
function extractFromBudgetSheet(budgetSheet, fundingSource, securedStatus, trackingSheet) {
  const data = [];
  const budgetValues = budgetSheet.getDataRange().getValues();
  
  // Find header row indices in Budget sheet
  let headerRow = -1;
  let milestoneCol = -1;
  let costCol = -1;
  let incomeCol = -1;
  let xeroInventoryCol = -1;
  
  for (let i = 0; i < budgetValues.length; i++) {
    const row = budgetValues[i];
    for (let j = 0; j < row.length; j++) {
      const cell = row[j].toString().trim();
      if (cell === 'Milestone') milestoneCol = j;
      if (cell === 'Cost') costCol = j;
      if (cell === 'Income') incomeCol = j;
      if (cell === 'Xero Inventory Item') xeroInventoryCol = j;
    }
    
    if (milestoneCol !== -1 && costCol !== -1 && incomeCol !== -1 && xeroInventoryCol !== -1) {
      headerRow = i;
      break;
    }
  }
  
  if (headerRow === -1) {
    Logger.log(`Warning: Could not find required headers in Budget sheet of ${fundingSource}`);
    return data;
  }
  
  // Build milestone to xero item mapping from Budget sheet
  const milestoneToXeroItem = {};
  for (let i = headerRow + 1; i < budgetValues.length; i++) {
    const row = budgetValues[i];
    const milestone = row[milestoneCol] ? row[milestoneCol].toString().trim() : '';
    const xeroItem = row[xeroInventoryCol] ? row[xeroInventoryCol].toString().trim() : '';
    
    if (milestone && xeroItem) {
      milestoneToXeroItem[milestone] = xeroItem;
    }
  }
  
  // Get tracking data using Xero Inventory Items
  const xeroItemActuals = trackingSheet ? getTrackingDataByXeroItem(trackingSheet) : {};
  
  // Extract data rows from Budget sheet
  for (let i = headerRow + 1; i < budgetValues.length; i++) {
    const row = budgetValues[i];
    const milestone = row[milestoneCol] ? row[milestoneCol].toString().trim() : '';
    const cost = row[costCol] || 0;
    const income = row[incomeCol] || 0;
    
    // Skip empty rows
    if (!milestone && !cost && !income) continue;
    
    // Get actual to date using the Xero Inventory Item
    const xeroItem = milestoneToXeroItem[milestone] || '';
    const actualToDate = xeroItem ? (xeroItemActuals[xeroItem] || 0) : 0;
    
    data.push([
      fundingSource,
      securedStatus,
      milestone,
      cost,
      income,
      actualToDate
    ]);
  }
  
  return data;
}

/**
 * Extract "Total Actual to date" from Tracking sheet for Xero Inventory Items
 * between "Expenses" and "Total Expenses" rows
 */
function getTrackingDataByXeroItem(trackingSheet) {
  const xeroItemMap = {};
  const values = trackingSheet.getDataRange().getValues();
  
  // Find the "Expenses" and "Total Expenses" rows, and column indices
  let expensesRow = -1;
  let totalExpensesRow = -1;
  let colA = 0; // Column A (Description/Xero Inventory Item names)
  let actualCol = -1; // Column E (Total Actual to date)
  
  // Find "Total Actual to date" column (should be column E)
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    for (let j = 0; j < row.length; j++) {
      const cell = row[j].toString().trim();
      if (cell === 'Total Actual to date' || cell.includes('Total Actual')) {
        actualCol = j;
      }
    }
    
    // Find "Expenses" row in column A
    const cellA = row[colA].toString().trim();
    if (cellA === 'Expenses') {
      expensesRow = i;
    }
    if (cellA === 'Total Expenses') {
      totalExpensesRow = i;
    }
  }
  
  if (expensesRow === -1 || totalExpensesRow === -1 || actualCol === -1) {
    Logger.log('Warning: Could not find Expenses section or Total Actual to date column in Tracking sheet');
    return xeroItemMap;
  }
  
  // Extract Xero Inventory Items and their actual values between Expenses and Total Expenses
  for (let i = expensesRow + 1; i < totalExpensesRow; i++) {
    const row = values[i];
    const xeroItem = row[colA] ? row[colA].toString().trim() : '';
    const actual = row[actualCol] || 0;
    
    if (xeroItem) {
      xeroItemMap[xeroItem] = actual;
      Logger.log(`Mapped Xero Item: "${xeroItem}" -> ${actual}`);
    }
  }
  
  return xeroItemMap;
}