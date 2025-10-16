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

  // Get the current spreadsheet and find Budget and Submitted_budget sheets
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const budgetSheet = spreadsheet.getSheetByName("Budget");  
  const submittedBudgetSheet = spreadsheet.getSheetByName("Submitted_budget");

  // Copy the valid account values
  if (submittedBudgetSheet) {
    Logger.log("Processing 'Submitted_budget' sheet...");

    // Add the new total account value to the end of the list
    const accountsForSubmittedBudget = [...validAccounts, "Total expenses without GST"];
    Logger.log("Added 'Total expenses without GST' to the list.");

    // Prepare data for writing (needs to be a 2D array)
    const valuesToWrite = accountsForSubmittedBudget.map(account => [account]);

    // Clear existing values from A3 downwards
    const clearRange = submittedBudgetSheet.getRange("A3:A");
    Logger.log("Clearing existing data from A3 downwards in 'Submitted_budget'.");
    clearRange.clearContent();

    // Write the new values starting from A3
    if (valuesToWrite.length > 0) {
      const targetRange = submittedBudgetSheet.getRange(3, 1, valuesToWrite.length, 1);
      Logger.log(`Writing ${valuesToWrite.length} new values to A3 downwards.`);
      targetRange.setValues(valuesToWrite);
    }
  } else {
    Logger.log("Sheet 'Submitted_budget' not found. Skipping.");
  }
  
  // Update budgetsheet data range
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