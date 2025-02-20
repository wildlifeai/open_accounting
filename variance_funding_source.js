function variance_funding_source() {
  function updateAccountValidation() {
    Logger.log("Starting account validation update process...");
  
    // Get valid accounts from Xero chart of accounts
    const xeroSheetId = "1VHtZsZRzJJ29tt3SRDXnIP6ebKChoDtlHqyV5Ua-3Yg";
    Logger.log("Fetching accounts from Xero sheet: " + xeroSheetId);
    
    const xeroSheet = SpreadsheetApp.openById(xeroSheetId).getActiveSheet();
    const xeroData = xeroSheet.getDataRange().getValues();
    const accountColIndexXero = 9; // Column J (0-based)
    
    // Get valid accounts (excluding header)
    const validAccounts = xeroData
      .map(row => row[accountColIndexXero])
      .filter(account => account && account !== "*Account");
    Logger.log(`Found ${validAccounts.length} valid accounts`);
  
    // Get the current spreadsheet and find Budget sheet
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const budgetSheet = spreadsheet.getSheetByName("Budget");
  
    if (!budgetSheet) {
      Logger.log("No Budget sheet found");
      return;
    }
  
    // Find the *Account column
    const headerRow = budgetSheet.getRange(1, 1, 1, budgetSheet.getLastColumn()).getValues()[0];
    const accountColIndex = headerRow.indexOf("*Account");
  
    if (accountColIndex === -1) {
      Logger.log("No *Account column found in Budget sheet");
      return;
    }
  
    Logger.log(`Found *Account column at index ${accountColIndex + 1}`);
  
    // Get the data range for the *Account column
    const lastRow = budgetSheet.getLastRow();
    const accountColumn = budgetSheet.getRange(1, accountColIndex + 1, lastRow);
  
    // Remove existing data validation
    Logger.log("Removing existing data validation...");
    accountColumn.clearDataValidations();
  
    // Create new data validation rule
    Logger.log("Creating new data validation rule...");
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(validAccounts, true)
      .setAllowInvalid(false)
      .build();
  
    // Apply new validation to all cells in the column except header
    Logger.log("Applying new data validation...");
    if (lastRow > 1) {
      const dataRange = budgetSheet.getRange(2, accountColIndex + 1, lastRow - 1);
      dataRange.setDataValidation(rule);
    }
  
    Logger.log("Account validation update completed successfully");
  }
  
  updateAccountValidation()
  
  Logger.log('Starting to retrieve the actuals for the budget...');  
  
  try {
    // Get the current spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentFileName = ss.getName();
    
    // Get the parent folder of the parent folder of the current spreadsheet
    const fileId = ss.getId();
    const initialParentFolder = DriveApp.getFileById(fileId).getParents().next();
    const parentOfParentFolder = initialParentFolder.getParents().next();
    const parentOfParentOfParentFolder = parentOfParentFolder.getParents().next();
    
    if (!parentOfParentOfParentFolder) {
      throw new Error('Cannot access grandparent folder of the current spreadsheet.');
    }

    // Find the "income_expenses_overview" folder
    let incomeExpensesFolder;
    const parentFolders = parentOfParentOfParentFolder.getFolders();
    while (parentFolders.hasNext()) {
      const folder = parentFolders.next();
      if (folder.getName() === 'income_expenses_overview') {
        incomeExpensesFolder = folder;
        break;
      }
    }

    if (!incomeExpensesFolder) {
       // Log all folder names for debugging
      const parentFolders = parentOfParentOfParentFolder.getFolders();
      let folderNames = [];
      while (parentFolders.hasNext()) {
        const folder = parentFolders.next();
        folderNames.push(folder.getName());
      }
      
      throw new Error(`Cannot find income_expenses_overview folder. Available folders: ${folderNames.join(', ')}`);
    }

    // Find the "previous_quarters" subfolder
    const previousQuartersFolder = incomeExpensesFolder.getFoldersByName('previous_quarters').next();

    if (!previousQuartersFolder) {
      throw new Error('Cannot find previous_quarters folder');
    }

    // Collect reconciliation sheets
    const reconciliationSheets = [];
    const foldersToSearch = [previousQuartersFolder, incomeExpensesFolder];

    foldersToSearch.forEach(folder => {
      const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
      while (files.hasNext()) {
        const file = files.next();
        if (file.getName().includes('reconciliation_expense_income')) {
          const spreadsheet = SpreadsheetApp.openById(file.getId());
          const sheet = spreadsheet.getSheetByName('Reconciliation');
          reconciliationSheets.push({
            spreadsheet: spreadsheet,
            sheet: sheet,
            fileName: file.getName()
          });
        }
      }
    });

    // Output sheet
    let outputSheet = ss.getSheetByName("Budget");
    if (!outputSheet) {
      throw new Error('Budget sheet not found in this Gsheet');
    }

    // Get Budget sheet headers
    const budgetHeaders = outputSheet.getRange(1, 1, 1, outputSheet.getLastColumn()).getValues()[0];
    const expenseIncomeIdx = budgetHeaders.indexOf('Expense/Income');
    const amountIdx = budgetHeaders.indexOf('Amount');
    const actualIdx = budgetHeaders.indexOf('Actual');
    const forecastIdx = budgetHeaders.indexOf('Forecast');

    if (expenseIncomeIdx === -1 || actualIdx === -1 || amountIdx === -1 || forecastIdx === -1) {
      throw new Error('Required columns not found in Budget sheet');
    }

    // Get list of matching Expense/Income combinations
    const budgetData = outputSheet.getRange(2, 1, outputSheet.getLastRow() - 1, outputSheet.getLastColumn()).getValues();
    const currentSheetCategories = new Set(budgetData.map(row => row[expenseIncomeIdx]));
                                        
    // Process each reconciliation sheet matching current file
    const varianceResults = new Map();

    reconciliationSheets.forEach(({sheet, fileName}) => {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const reconciledExpIncIdx = headers.indexOf('Reconciled Expense/Income');
      const fundingSourceIdx = headers.indexOf('Funding Source');
      const debitIdx = headers.indexOf('Debit (NZD)');
      const creditIdx = headers.indexOf('Credit (NZD)');

      if (reconciledExpIncIdx === -1 || fundingSourceIdx === -1 || debitIdx === -1 || creditIdx === -1) {
        throw new Error('Required Reconciled Expense/Income, Funding Source, Debit (NZD), or Credit (NZD) columns not found in' + sheet.name);
      }

      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

      data.forEach(row => {
        const reconciledExpInc = row[reconciledExpIncIdx];
        const fundingSource = row[fundingSourceIdx];
        const debit = Number(row[debitIdx]) || 0;
        const credit = Number(row[creditIdx]) || 0;

        // Check if the Funding Source is the current sheet's filename
        if (fundingSource !== currentFileName) {
          return;
        }
        
        // Check if the Expense/Income doesn't match any of the Expense/Income of the current sheet
        if (!currentSheetCategories.has(reconciledExpInc)) {
          return;
        }

        const categoryKey = reconciledExpInc;
        const currentTotal = varianceResults.get(categoryKey) || { debitSum: 0, creditSum: 0, sources: [] };
        const updatedTotal = {
          debitSum: currentTotal.debitSum + debit,
          creditSum: currentTotal.creditSum + credit,
          sources: [...currentTotal.sources, { 
            fileName: fileName, 
            debit: debit, 
            credit: credit, 
            reconciledExpInc: reconciledExpInc
          }]
        };
        varianceResults.set(categoryKey, updatedTotal);
      });
    });

    // Detailed logging of variance results
    Logger.log('=== Detailed Variance Breakdown ===');
    varianceResults.forEach((variance, expenseIncome) => {
      Logger.log(`Expense/Income: ${expenseIncome}`);
      Logger.log(`Total Debit: $${variance.debitSum.toFixed(2)}`);
      Logger.log(`Total Credit: $${variance.creditSum.toFixed(2)}`);
      Logger.log('Sources:');
      variance.sources.forEach(source => {
        Logger.log(`- File: ${source.fileName}`);
        Logger.log(`  Debit: $${source.debit.toFixed(2)}`);
        Logger.log(`  Credit: $${source.credit.toFixed(2)}`);
      });
      Logger.log('---');
    });

    // Update only the Actual column in the Budget sheet
    const actualRange = outputSheet.getRange(2, actualIdx + 1, budgetData.length, 1);
    const actualValues = budgetData.map(row => {
      const expenseIncome = row[expenseIncomeIdx];
      const variance = varianceResults.get(expenseIncome);
      
      return [(variance ? Math.abs(variance.debitSum - variance.creditSum) : row[actualIdx])];
    });

    // Write back only the Actual column
    actualRange.setValues(actualValues);

    // Update all forecast formulas
    for (let i = 0; i < budgetData.length; i++) {
      const rowNumber = i + 2; // +2 because we start from row 2
      const forecastCell = outputSheet.getRange(rowNumber, forecastIdx + 1);
      const formula = `=${getA1Notation(amountIdx + 1)}${rowNumber}-${getA1Notation(actualIdx + 1)}${rowNumber}`;
      forecastCell.setFormula(formula);
    }
    
    Logger.log('Budget actuals and forecasts updated successfully');
    
  } catch (error) {
    Logger.log('Critical error in income and expenses variance generation: ' + error.toString());
    throw error;
  }
}

// Helper function to convert column index to A1 notation
function getA1Notation(column) {
  let temp = '';
  let letter = '';
  
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  
  return letter;
}