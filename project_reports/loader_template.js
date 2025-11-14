/**
 * Main script to be placed in PROJECT_overview Google Sheet
 * This fetches and runs the script from GitHub
 */

function loadAndRunGitHubScript() {
  // Replace with your actual GitHub raw URL
  const GITHUB_SCRIPT_URL = 'https://raw.githubusercontent.com/wildlifeai/open_accounting/refs/heads/main/project_reports/funding-aggregator.js';
  
  try {
    // Fetch the script from GitHub
    const response = UrlFetchApp.fetch(GITHUB_SCRIPT_URL);
    const scriptCode = response.getContentText();
    
    // Execute the fetched script
    eval(scriptCode);
    
    // Call the main function from the GitHub script
    aggregateFundingData();
    
    SpreadsheetApp.getUi().alert('Success!', 'Funding data has been updated.', SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', 'Failed to load or run script: ' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('Error: ' + error.toString());
  }
}

/**
 * Creates a custom menu when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Funding Tools')
      .addItem('Update Funding Data', 'loadAndRunGitHubScript')
      .addToUi();
}