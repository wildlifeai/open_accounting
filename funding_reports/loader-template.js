// ==========================================
// XERO INTEGRATION LOADER SCRIPT - TEMPLATE
// ==========================================
// Copy this file to your Google Apps Script project
// Rename it to something like "Config" or "MyPrivateScript"
// DO NOT commit this file to GitHub after adding your credentials!

// ==========================================
// CONFIGURATION - ADD YOUR CREDENTIALS HERE
// ==========================================
const PRIVATE_CONFIG = {
  // 1. Get these from your Xero app: https://developer.xero.com/app/manage
  CLIENT_ID: 'YOUR_CLIENT_ID_HERE',
  CLIENT_SECRET: 'YOUR_CLIENT_SECRET_HERE',
  
  // 2. Format: https://script.google.com/macros/d/{YOUR_SCRIPT_ID}/usercallback
  //    Get Script ID from: Apps Script > Project Settings (gear icon)
  REDIRECT_URI: 'https://script.google.com/macros/d/YOUR_SCRIPT_ID_HERE/usercallback',
  
  // 3. URL to the main script file in your GitHub repo
  //    Format: https://raw.githubusercontent.com/USERNAME/REPO/BRANCH/xero-integration.js
  //    Example: https://raw.githubusercontent.com/john/xero-sync/main/xero-integration.js
  GITHUB_SCRIPT_URL: 'https://raw.githubusercontent.com/YOUR_USERNAME/YOUR_REPO/main/xero-integration.js'
};

// ==========================================
// DO NOT EDIT BELOW THIS LINE
// ==========================================

/**
 * Loads and evaluates the main script from GitHub
 */
function loadXeroScript() {
  try {
    const response = UrlFetchApp.fetch(PRIVATE_CONFIG.GITHUB_SCRIPT_URL);
    const scriptCode = response.getContentText();
    
    // Evaluate the script to load all functions into this context
    eval(scriptCode);
    
    // Set the configuration with private credentials
    setConfig(
      PRIVATE_CONFIG.CLIENT_ID,
      PRIVATE_CONFIG.CLIENT_SECRET,
      PRIVATE_CONFIG.REDIRECT_URI
    );
    
    return true;
  } catch (error) {
    Logger.log('Error loading script from GitHub: ' + error.toString());
    SpreadsheetApp.getUi().alert('Error loading Xero script: ' + error.toString());
    return false;
  }
}

/**
 * Creates custom menu - automatically runs when spreadsheet opens
 */
function onOpen() {
  if (loadXeroScript()) {
    // Call the onOpen from the loaded script
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Xero Sync')
      .addItem('1. Authorize Xero', 'showAuthorizationUrlWrapper')
      .addItem('2. Update Transactions', 'updateXeroTransactionsWrapper')
      .addItem('Clear Authorization', 'clearAuthorizationWrapper')
      .addToUi();
  }
}

/**
 * Wrapper functions that load the script before calling main functions
 */
function showAuthorizationUrlWrapper() {
  if (loadXeroScript()) {
    showAuthorizationUrl();
  }
}

function updateXeroTransactionsWrapper() {
  if (loadXeroScript()) {
    updateXeroTransactions();
  }
}

function clearAuthorizationWrapper() {
  if (loadXeroScript()) {
    clearAuthorization();
  }
}

/**
 * OAuth callback handler - loads script before handling callback
 */
function authCallback(request) {
  if (loadXeroScript()) {
    return authCallback(request);
  }
  return HtmlService.createHtmlOutput('Error: Could not load Xero script.');
}