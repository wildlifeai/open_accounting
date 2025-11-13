// Xero API Integration for Google Sheets
// This script fetches journal/transaction data from Xero and updates the "Xero Transactions" sheet
// GitHub: [Your Repository URL Here]

// CONFIGURATION - These are set via setConfig() function in your private script
let CONFIG = {
  CLIENT_ID: '',
  CLIENT_SECRET: '',
  REDIRECT_URI: '',
  SHEET_NAME: 'Xero Transactions',
  TRACKING_CATEGORY_NAME: 'Funding source',
  DATE_SHEET_NAME: 'Budget, Actual, Forecast Tracking',
  DATE_CELL: 'B3',
  EXCLUDED_ACCOUNT_CODES: ['835', '600', '610', '800', '820', '877']
};

/**
 * Sets the configuration with sensitive credentials
 * This should be called from your private config script
 */
function setConfig(clientId, clientSecret, redirectUri) {
  CONFIG.CLIENT_ID = clientId;
  CONFIG.CLIENT_SECRET = clientSecret;
  CONFIG.REDIRECT_URI = redirectUri;
}

/**
 * Gets the current configuration
 */
function getConfig() {
  return CONFIG;
}

// OAuth 2.0 endpoints
const XERO_AUTH_URL = 'https://login.xero.com/identity/connect/authorize';
const XERO_TOKEN_URL = 'https://identity.xero.com/connect/token';
const XERO_API_BASE = 'https://api.xero.com/api.xro/2.0';
const XERO_IDENTITY_URL = 'https://api.xero.com/connections';

/**
 * Creates a custom menu in Google Sheets when the document opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Xero Sync')
    .addItem('1. Authorize Xero', 'showAuthorizationUrl')
    .addItem('2. Update Transactions', 'updateXeroTransactions')
    .addItem('Clear Authorization', 'clearAuthorization')
    .addToUi();
}

/**
 * Step 1: Display authorization URL for user to connect to Xero
 */
function showAuthorizationUrl() {
  const service = getXeroService();
  
  if (service.hasAccess()) {
    SpreadsheetApp.getUi().alert('Already authorized! You can now update transactions.');
    return;
  }
  
  const authorizationUrl = service.getAuthorizationUrl();
  const template = HtmlService.createHtmlOutput(
    '<p>Click the link below to authorize access to Xero:</p>' +
    '<p><a href="' + authorizationUrl + '" target="_blank">Authorize Xero Access</a></p>' +
    '<p>After authorizing, close this window and run "Update Transactions".</p>'
  );
  
  SpreadsheetApp.getUi().showModalDialog(template, 'Xero Authorization');
}

/**
 * OAuth2 callback handler
 */
function authCallback(request) {
  const service = getXeroService();
  const isAuthorized = service.handleCallback(request);
  
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab and return to your spreadsheet.');
  } else {
    return HtmlService.createHtmlOutput('Authorization failed. Please try again.');
  }
}

/**
 * Clears stored authorization
 */
function clearAuthorization() {
  const service = getXeroService();
  service.reset();
  SpreadsheetApp.getUi().alert('Authorization cleared. Run "Authorize Xero" to reconnect.');
}

/**
 * Creates and configures the OAuth2 service for Xero
 */
function getXeroService() {
  return OAuth2.createService('xero')
    .setAuthorizationBaseUrl(XERO_AUTH_URL)
    .setTokenUrl(XERO_TOKEN_URL)
    .setClientId(CONFIG.CLIENT_ID)
    .setClientSecret(CONFIG.CLIENT_SECRET)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('offline_access accounting.transactions.read accounting.reports.read accounting.journals.read')
    .setParam('response_type', 'code')
    .setTokenHeaders({
      'Authorization': 'Basic ' + Utilities.base64Encode(CONFIG.CLIENT_ID + ':' + CONFIG.CLIENT_SECRET)
    });
}

/**
 * Main function: Fetches journal data from Xero and updates the sheet
 */
function updateXeroTransactions() {
  const ui = SpreadsheetApp.getUi();
  const service = getXeroService();
  
  // Check authorization
  if (!service.hasAccess()) {
    ui.alert('Not authorized. Please run "Authorize Xero" first.');
    return;
  }
  
  try {
    ui.alert('Fetching data from Xero... This may take a moment.');
    
    // Get the spreadsheet filename to use as tracking category value
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const trackingCategoryValue = ss.getName();
    
    // Get start date from the specified sheet and cell
    const startDate = getStartDateFromSheet();
    
    if (!startDate) {
      ui.alert('Error: Could not read start date from cell ' + CONFIG.DATE_CELL + 
               ' in sheet "' + CONFIG.DATE_SHEET_NAME + '". Please ensure the cell contains a valid date.');
      return;
    }
    
    // Get tenant ID
    const tenantId = getXeroTenantId(service);
    
    if (!tenantId) {
      ui.alert('Error: Could not retrieve Xero organization. Please re-authorize.');
      return;
    }
    
    // Fetch journals (account transactions)
    const journals = fetchXeroJournals(service, tenantId, startDate);
    
    // Filter for specific tracking category
    const filteredTransactions = filterTransactionsByTracking(journals, 
      CONFIG.TRACKING_CATEGORY_NAME, 
      trackingCategoryValue);
    
    // Update the sheet
    updateSheet(filteredTransactions);
    
    ui.alert('Success! Updated ' + filteredTransactions.length + ' transactions.');
    
  } catch (error) {
    Logger.log('Error: ' + error.toString());
    ui.alert('Error: ' + error.toString());
  }
}

/**
 * Retrieves the start date from the specified sheet and cell
 */
function getStartDateFromSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.DATE_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('Sheet "' + CONFIG.DATE_SHEET_NAME + '" not found.');
      return null;
    }
    
    const dateValue = sheet.getRange(CONFIG.DATE_CELL).getValue();
    
    if (!dateValue) {
      Logger.log('Cell ' + CONFIG.DATE_CELL + ' is empty.');
      return null;
    }
    
    // Convert to date if it's not already
    const date = new Date(dateValue);
    
    if (isNaN(date.getTime())) {
      Logger.log('Invalid date in cell ' + CONFIG.DATE_CELL + ': ' + dateValue);
      return null;
    }
    
    return date;
  } catch (error) {
    Logger.log('Error reading start date: ' + error.toString());
    return null;
  }
}

/**
 * Retrieves the Xero tenant (organization) ID
 */
function getXeroTenantId(service) {
  const response = UrlFetchApp.fetch(XERO_IDENTITY_URL, {
    headers: {
      'Authorization': 'Bearer ' + service.getAccessToken(),
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() === 200) {
    const connections = JSON.parse(response.getContentText());
    if (connections && connections.length > 0) {
      return connections[0].tenantId;
    }
  }
  
  return null;
}

/**
 * Fetches journals from Xero API with pagination
 */
function fetchXeroJournals(service, tenantId, startDate) {
  const allJournals = [];
  let offset = 0;
  const pageSize = 100;
  let hasMore = true;
  
  // Format start date for Xero API
  const fromDateStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  Logger.log('Fetching journals from ' + fromDateStr + ' to today');
  
  while (hasMore) {
    const url = `${XERO_API_BASE}/Journals?offset=${offset}&if-modified-since=${fromDateStr}`;
    
    const response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken(),
        'xero-tenant-id': tenantId,
        'Accept': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    if (response.getResponseCode() !== 200) {
      throw new Error('Failed to fetch journals: ' + response.getContentText());
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (data.Journals && data.Journals.length > 0) {
      allJournals.push(...data.Journals);
      offset += pageSize;
      
      // If we got fewer than pageSize results, we've reached the end
      if (data.Journals.length < pageSize) {
        hasMore = false;
      }
    } else {
      hasMore = false;
    }
    
    // Avoid hitting rate limits
    Utilities.sleep(300);
  }
  
  Logger.log('Fetched ' + allJournals.length + ' journals total');
  
  return allJournals;
}

/**
 * Parses Xero date format to JavaScript Date object
 * Xero dates come in format like "/Date(1609459200000+0000)/"
 */
function parseXeroDate(xeroDateString) {
  if (!xeroDateString) return new Date();
  
  // If it's already a proper date string
  if (xeroDateString.indexOf('/Date(') === -1) {
    return new Date(xeroDateString);
  }
  
  // Extract timestamp from Xero's /Date(timestamp)/ format
  const timestamp = parseInt(xeroDateString.match(/\d+/)[0]);
  return new Date(timestamp);
}

/**
 * Filters journal entries for specific tracking category and value
 */
function filterTransactionsByTracking(journals, trackingCategoryName, trackingCategoryValue) {
  const transactions = [];
  
  Logger.log('Filtering for tracking category: ' + trackingCategoryName + ' = ' + trackingCategoryValue);
  Logger.log('Excluding account codes: ' + CONFIG.EXCLUDED_ACCOUNT_CODES.join(', '));
  
  journals.forEach(journal => {
    if (!journal.JournalLines) return;
    
    journal.JournalLines.forEach(line => {
      // Skip excluded account codes
      if (CONFIG.EXCLUDED_ACCOUNT_CODES.includes(line.AccountCode)) {
        return;
      }
      
      // Check if this line has the required tracking category
      let hasMatchingTracking = false;
      
      if (line.TrackingCategories && line.TrackingCategories.length > 0) {
        hasMatchingTracking = line.TrackingCategories.some(tracking => 
          tracking.Name === trackingCategoryName && tracking.Option === trackingCategoryValue
        );
      }
      
      // Only include lines with matching tracking category
      if (hasMatchingTracking) {
        transactions.push({
          date: parseXeroDate(journal.JournalDate),
          journalNumber: journal.JournalNumber,
          reference: journal.Reference || '',
          sourceType: journal.SourceType || '',
          sourceID: journal.SourceID || '',
          accountCode: line.AccountCode,
          accountName: line.AccountName || '',
          description: line.Description || '',
          debit: line.NetAmount > 0 ? line.NetAmount : 0,
          credit: line.NetAmount < 0 ? Math.abs(line.NetAmount) : 0,
          netAmount: line.NetAmount,
          taxAmount: line.TaxAmount || 0,
          grossAmount: line.GrossAmount || 0,
          trackingCategory1: line.TrackingCategories && line.TrackingCategories[0] ? 
            line.TrackingCategories[0].Name + ': ' + line.TrackingCategories[0].Option : '',
          trackingCategory2: line.TrackingCategories && line.TrackingCategories[1] ? 
            line.TrackingCategories[1].Name + ': ' + line.TrackingCategories[1].Option : ''
        });
      }
    });
  });
  
  // Sort by date (newest first)
  transactions.sort((a, b) => new Date(b.date) - new Date(a.date));
  
  Logger.log('Filtered to ' + transactions.length + ' transactions');
  
  return transactions;
}

/**
 * Updates the Google Sheet with transaction data
 */
function updateSheet(transactions) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  // Create sheet if it doesn't exist
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }
  
  // Clear existing data
  sheet.clear();
  
  // Set headers
  const headers = [
    'Date', 
    'Journal #', 
    'Reference', 
    'Source Type',
    'Source ID',
    'Account Code', 
    'Account Name', 
    'Description', 
    'Debit', 
    'Credit', 
    'Net Amount',
    'Tax Amount',
    'Gross Amount',
    'Tracking 1',
    'Tracking 2'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  
  // Add data
  if (transactions.length > 0) {
    const data = transactions.map(t => [
      t.date, // Already a Date object now
      t.journalNumber,
      t.reference,
      t.sourceType,
      t.sourceID,
      t.accountCode,
      t.accountName,
      t.description,
      t.debit,
      t.credit,
      t.netAmount,
      t.taxAmount,
      t.grossAmount,
      t.trackingCategory1,
      t.trackingCategory2
    ]);
    
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    
    // Format date column
    sheet.getRange(2, 1, data.length, 1).setNumberFormat('yyyy-mm-dd');
    
    // Format currency columns (Debit, Credit, Net Amount, Tax Amount, Gross Amount)
    sheet.getRange(2, 9, data.length, 5).setNumberFormat('$#,##0.00');
    
    // Auto-resize columns
    for (let i = 1; i <= headers.length; i++) {
      sheet.autoResizeColumn(i);
    }
  }
}

/**
 * SETUP INSTRUCTIONS:
 * 
 * 1. Install OAuth2 Library:
 *    - In Apps Script editor, click "+" next to Libraries
 *    - Enter Script ID: 1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF
 *    - Select the latest version and click "Add"
 * 
 * 2. Create Xero OAuth App:
 *    - Go to https://developer.xero.com/app/manage
 *    - Create a new app
 *    - Set redirect URI to: https://script.google.com/macros/d/{YOUR_SCRIPT_ID}/usercallback
 *    - Copy Client ID and Client Secret to CONFIG above
 * 
 * 3. Deploy as Web App:
 *    - Click "Deploy" > "New deployment"
 *    - Select type "Web app"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 *    - Click "Deploy" and authorize
 * 
 * 4. Get Script ID:
 *    - In Apps Script, go to Project Settings
 *    - Copy the Script ID
 *    - Update REDIRECT_URI in CONFIG above
 * 
 * 5. Run the script:
 *    - Reload your spreadsheet
 *    - Use the "Xero Sync" menu
 *    - First run "Authorize Xero"
 *    - Then run "Update Transactions"
 */