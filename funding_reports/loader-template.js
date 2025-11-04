// ==========================================
// XERO INTEGRATION LOADER SCRIPT - WORKING VERSION
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
  GITHUB_SCRIPT_URL: 'https://raw.githubusercontent.com/wildlifeai/open_accounting/refs/heads/main/funding_reports/xero-integration.js'
};

// ==========================================
// SIMPLE APPROACH: FETCH ON DEMAND
// ==========================================

/**
 * Fetches the script from GitHub and returns it as text
 */
function getScriptCode() {
  try {
    // Try cache first
    let scriptCode = PropertiesService.getUserProperties().getProperty('CACHED_XERO_SCRIPT');
    
    if (!scriptCode) {
      // Fetch from GitHub
      const response = UrlFetchApp.fetch(PRIVATE_CONFIG.GITHUB_SCRIPT_URL);
      scriptCode = response.getContentText();
      
      // Cache it
      PropertiesService.getUserProperties().setProperty('CACHED_XERO_SCRIPT', scriptCode);
    }
    
    return scriptCode;
  } catch (error) {
    throw new Error('Cannot load script: ' + error.toString());
  }
}

/**
 * Updates the cached script from GitHub
 */
function updateScriptFromGitHub() {
  try {
    const response = UrlFetchApp.fetch(PRIVATE_CONFIG.GITHUB_SCRIPT_URL);
    const scriptCode = response.getContentText();
    PropertiesService.getUserProperties().setProperty('CACHED_XERO_SCRIPT', scriptCode);
    SpreadsheetApp.getUi().alert('Script updated successfully from GitHub!');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error updating script: ' + error.toString());
  }
}

/**
 * Initializes on first run
 */
function initialize() {
  SpreadsheetApp.getActiveSpreadsheet();
  UrlFetchApp.fetch('https://www.google.com');
  PropertiesService.getUserProperties();
  
  updateScriptFromGitHub();
  
  SpreadsheetApp.getUi().alert(
    'Initialization complete!\n\n' +
    'Reload your spreadsheet to see the "Xero Sync" menu.'
  );
}

/**
 * Creates menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const hasCachedScript = PropertiesService.getUserProperties().getProperty('CACHED_XERO_SCRIPT') !== null;
  
  if (hasCachedScript) {
    ui.createMenu('Xero Sync')
      .addItem('1. Authorize Xero', 'showAuthorizationUrl')
      .addItem('2. Update Transactions', 'updateXeroTransactions')
      .addSeparator()
      .addItem('Update Script from GitHub', 'updateScriptFromGitHub')
      .addItem('Clear Authorization', 'clearAuthorization')
      .addToUi();
  } else {
    ui.createMenu('Xero Sync')
      .addItem('⚠️ Setup Required - Run Initialize', 'initialize')
      .addToUi();
  }
}

// ==========================================
// XERO FUNCTIONS - These execute the loaded script
// ==========================================

function showAuthorizationUrl() {
  const code = getScriptCode();
  eval(code);
  eval(`setConfig('${PRIVATE_CONFIG.CLIENT_ID}', '${PRIVATE_CONFIG.CLIENT_SECRET}', '${PRIVATE_CONFIG.REDIRECT_URI}')`);
  eval('showAuthorizationUrl()');
}

function updateXeroTransactions() {
  const code = getScriptCode();
  eval(code);
  eval(`setConfig('${PRIVATE_CONFIG.CLIENT_ID}', '${PRIVATE_CONFIG.CLIENT_SECRET}', '${PRIVATE_CONFIG.REDIRECT_URI}')`);
  eval('updateXeroTransactions()');
}

function clearAuthorization() {
  const code = getScriptCode();
  eval(code);
  eval(`setConfig('${PRIVATE_CONFIG.CLIENT_ID}', '${PRIVATE_CONFIG.CLIENT_SECRET}', '${PRIVATE_CONFIG.REDIRECT_URI}')`);
  eval('clearAuthorization()');
}

function authCallback(request) {
  const code = getScriptCode();
  eval(code);
  eval(`setConfig('${PRIVATE_CONFIG.CLIENT_ID}', '${PRIVATE_CONFIG.CLIENT_SECRET}', '${PRIVATE_CONFIG.REDIRECT_URI}')`);
  return eval('authCallback(request)');
}