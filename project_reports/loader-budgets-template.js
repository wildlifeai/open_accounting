// ==========================================
// QUARTERLY BUDGETS LOADER SCRIPT
// ==========================================
// Copy this file to your Google Apps Script project
// Rename it to something like "Config" or "MyPrivateScript"
// DO NOT commit this file to GitHub after adding your credentials!

const PRIVATE_CONFIG = {
  CLIENT_ID: 'YOUR_CLIENT_ID',
  CLIENT_SECRET: 'YOUR_CLIENT_SECRET',
  REDIRECT_URI: 'https://script.google.com/macros/d/YOUR_SCRIPT_ID/usercallback',

  GITHUB_SCRIPT_URL: 'https://raw.githubusercontent.com/wildlifeai/open_accounting/refs/heads/feature/budget_xero/project_reports/create_quarterly_budgets.js'
};

// Cache key for the script code
const CACHE_KEY = 'BUDGETS_SCRIPT_CACHE';

// ================= LOAD SCRIPT =================
function loadBudgetsModule() {
  const props = PropertiesService.getDocumentProperties() || PropertiesService.getScriptProperties();
  let code = props.getProperty(CACHE_KEY);

  if (!code) {
    const res = UrlFetchApp.fetch(PRIVATE_CONFIG.GITHUB_SCRIPT_URL);
    code = res.getContentText();
    props.setProperty(CACHE_KEY, code);
  }

  const moduleFactory = new Function(`
    ${code}
    return QuarterlyBudgetsIntegration;
  `);

  const QuarterlyBudgetsIntegration = moduleFactory();
  return QuarterlyBudgetsIntegration(PRIVATE_CONFIG);
}

// ================= EXPOSED ENTRY POINTS =================

/**
 * 1. Run this first to get the URL to authorize Xero
 */
function showAuthorizationUrl() {
  const module = loadBudgetsModule();
  module.showAuthorizationUrl();
}

/**
 * Required callback function for OAuth2
 */
function authCallback(request) {
  const module = loadBudgetsModule();
  return module.authCallback(request);
}

/**
 * 2. Run this to generate the CSV budgets and diff them against Xero
 */
function generateAndUploadQuarterlyBudgets() {
  const module = loadBudgetsModule();
  module.generateAndUploadQuarterlyBudgets();
}

/**
 * Helper to force a code update from GitHub by clearing the cache
 */
function updateScriptFromGitHub() {
  const props = PropertiesService.getDocumentProperties() || PropertiesService.getScriptProperties();
  props.deleteProperty(CACHE_KEY);
  Logger.log("Cache cleared. Next run will fetch latest code from GitHub.");
}

/**
 * Quick test to verify which PropertiesService is available in the current environment
 */
function testLoader() {
  Logger.log(PropertiesService.getDocumentProperties());
  Logger.log(PropertiesService.getScriptProperties());
}

/**
 * Dummy function to force Apps Script to request necessary permissions.
 * You do not need to run this function. The Apps Script analyzer will see
 * these calls and prompt you for Drive and Spreadsheet permissions.
 */
function _forcePermissions() {
  if (false) {
    SpreadsheetApp.openById('');
    SpreadsheetApp.create('');
    DriveApp.getFolderById('');
  }
}
