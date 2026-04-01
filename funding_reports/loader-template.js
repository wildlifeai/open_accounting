// ==========================================
// XERO INTEGRATION LOADER SCRIPT
// ==========================================
// Copy this file to your Google Apps Script project
// Rename it to something like "Config" or "MyPrivateScript"
// DO NOT commit this file to GitHub after adding your credentials!

const PRIVATE_CONFIG = {
  CLIENT_ID: 'YOUR_CLIENT_ID',
  CLIENT_SECRET: 'YOUR_CLIENT_SECRET',
  REDIRECT_URI: 'https://script.google.com/macros/d/YOUR_SCRIPT_ID/usercallback',

  GITHUB_SCRIPT_URL: 'https://raw.githubusercontent.com/wildlifeai/open_accounting/refs/heads/main/funding_reports/xero-integration.js'
};

// ✅ Per-sheet storage (NOT shared globally)
const CACHE_KEY = 'XERO_SCRIPT_CACHE';

// ================= LOAD SCRIPT =================
function loadXeroModule() {
  const props = PropertiesService.getDocumentProperties();

  let code = props.getProperty(CACHE_KEY);

  if (!code) {
    const res = UrlFetchApp.fetch(PRIVATE_CONFIG.GITHUB_SCRIPT_URL);
    code = res.getContentText();

    props.setProperty(CACHE_KEY, code);
  }

  const moduleFactory = new Function(`
    ${code}
    return XeroIntegration;
  `);

  const XeroIntegration = moduleFactory();

  return XeroIntegration(PRIVATE_CONFIG);
}

// ================= MENU =================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Xero Sync')
    .addItem('1. Authorize Xero', 'authorize')
    .addItem('2. Update Transactions', 'runUpdate')
    .addItem('System Check', 'systemCheck')
    .addItem('Refresh Script Cache', 'refreshScript')
    .addItem('Clear Auth', 'clearAuth')
    .addToUi();
}

// ================= ACTIONS =================
function authorize() {
  const xero = loadXeroModule();
  const url = xero.showAuthorizationUrl();

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(`<a href="${url}" target="_blank">Authorize Xero</a>`),
    'Authorize Xero'
  );
}

function runUpdate() {
  const ui = SpreadsheetApp.getUi();

  try {
    const xero = loadXeroModule();
    const result = xero.updateXeroTransactions();

    ui.alert('✅ ' + result);

  } catch (e) {
    if (e.message === 'NOT_AUTHORIZED') {
      ui.alert('Please authorize first.');
    } else {
      ui.alert('❌ Error: ' + e.message);
    }
  }
}

function clearAuth() {
  loadXeroModule().clearAuthorization();
  SpreadsheetApp.getUi().alert('Auth cleared');
}

// ================= SYSTEM CHECK =================
function systemCheck() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();

  try {
    const xero = loadXeroModule();

    const cache = props.getProperty(CACHE_KEY);
    if (!cache) throw new Error('Script cache missing');

    const lastSync = props.getProperty('XERO_LAST_SYNC');

    ui.alert(
      '✅ System OK\n' +
      'Cache: OK\n' +
      'Last Sync: ' + (lastSync || 'Never')
    );

  } catch (e) {
    ui.alert('❌ System Error: ' + e.message);
  }
}

// ================= CALLBACK =================
function authCallback(request) {
  const success = loadXeroModule().handleCallback(request);

  return HtmlService.createHtmlOutput(
    success ? 'Success! You can close this tab.' : 'Auth failed'
  );
}

// ================= CACHE CONTROL =================
function refreshScript() {
  const res = UrlFetchApp.fetch(PRIVATE_CONFIG.GITHUB_SCRIPT_URL);

  PropertiesService.getDocumentProperties().setProperty(
    CACHE_KEY,
    res.getContentText()
  );

  SpreadsheetApp.getUi().alert('Script updated');
}

// ================= AUTO SYNC =================
function setupAutoSync() {
  const triggers = ScriptApp.getProjectTriggers();
  if (!triggers.some(t => t.getHandlerFunction() === 'runUpdate')) {
    ScriptApp.newTrigger('runUpdate')
      .timeBased()
      .everyHours(1)
      .create();
  }
}