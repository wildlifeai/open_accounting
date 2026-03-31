// Xero API Integration module for Google Sheets
// This script fetches journal/transaction data from Xero and updates the "Xero Transactions" sheet

function XeroIntegration(config) {

  const CONFIG = {
    SHEET_NAME: 'Xero Transactions',
    TRACKING_CATEGORY_NAME: 'Funding source',

    DATE_SHEET_NAME: 'Budget, Actual, Forecast Tracking',
    DATE_CELL: 'B3',

    EXCLUDED_ACCOUNT_CODES: ['835', '600', '610', '800', '820', '877'],

    TEST_MODE: false
  };

  const XERO_API_BASE = 'https://api.xero.com/api.xro/2.0';
  const XERO_IDENTITY_URL = 'https://api.xero.com/connections';

  // Cache for invoice item lookups (keyed by SourceID)
  const itemCache = {};

  // ================= LOGGER =================
  function log(message, data = null) {
    Logger.log(JSON.stringify({
      time: new Date().toISOString(),
      message,
      data
    }));
  }

  // ================= RETRY =================
  function fetchWithRetry(url, options, retries = 3) {
    for (let i = 0; i < retries; i++) {
      try {
        const res = UrlFetchApp.fetch(url, options);

        if (res.getResponseCode() < 300) return res;

      } catch (e) {
        if (i === retries - 1) throw e;
        Utilities.sleep(1000 * (i + 1));
      }
    }
  }

  // ================= ITEM ENRICHMENT =================
  function getItemCodes(service, tenantId, sourceType, sourceID) {
    if (itemCache[sourceID]) return itemCache[sourceID];

    // Only invoices (ACCPAY = bills, ACCREC = sales invoices) carry ItemCodes
    if (sourceType !== 'ACCPAY' && sourceType !== 'ACCREC') {
      itemCache[sourceID] = [''];
      return itemCache[sourceID];
    }

    try {
      const url = `${XERO_API_BASE}/Invoices/${sourceID}`;

      const res = fetchWithRetry(url, {
        headers: {
          Authorization: 'Bearer ' + service.getAccessToken(),
          'xero-tenant-id': tenantId
        }
      });

      const data = JSON.parse(res.getContentText());
      const invoice = data.Invoices[0];

      const codes = (invoice.LineItems || [])
        .map(li => li.ItemCode || '')
        .filter(Boolean);

      itemCache[sourceID] = codes.length ? codes : [''];
    } catch (e) {
      log('Failed to fetch item codes for ' + sourceID, e.message);
      itemCache[sourceID] = [''];
    }

    return itemCache[sourceID];
  }

  // ================= AUTH =================
  function getXeroService() {
    return OAuth2.createService('xero')
      .setAuthorizationBaseUrl('https://login.xero.com/identity/connect/authorize')
      .setTokenUrl('https://identity.xero.com/connect/token')
      .setClientId(config.CLIENT_ID)
      .setClientSecret(config.CLIENT_SECRET)
      .setCallbackFunction('authCallback')
      .setPropertyStore(PropertiesService.getDocumentProperties())
      .setScope('offline_access accounting.transactions.read accounting.journals.read')
      .setParam('response_type', 'code')
      .setTokenHeaders({
        Authorization: 'Basic ' + Utilities.base64Encode(config.CLIENT_ID + ':' + config.CLIENT_SECRET)
      });
  }

  // ================= MAIN =================
  function updateXeroTransactions() {
    const service = getXeroService();

    if (!service.hasAccess()) {
      throw new Error('NOT_AUTHORIZED');
    }

    log('Starting sync');

    const props = PropertiesService.getDocumentProperties();

    try {
      const tenantId = getTenantId(service);
      const startDate = getStartDateFromSheet();

      log('Fetching from', startDate);

      const journals = fetchJournals(service, tenantId, startDate);

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const trackingValue = ss.getName();

      const filtered = filterTransactionsByTracking(
        service,
        tenantId,
        journals,
        CONFIG.TRACKING_CATEGORY_NAME,
        trackingValue,
        startDate
      );

      log('Filtered transactions', filtered.length);

      if (CONFIG.TEST_MODE) {
        log('TEST MODE - skipping write');
        return 'TEST SUCCESS';
      }

      updateSheet(filtered);

      
      return `SUCCESS: ${filtered.length} transactions`;

    } catch (e) {
      log('Sync failed', e.message);
      throw e;
    }
  }

  // ================= TENANT =================
  function getTenantId(service) {
    const res = fetchWithRetry(XERO_IDENTITY_URL, {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      }
    });

    const data = JSON.parse(res.getContentText());

    if (!data.length) throw new Error('No Xero tenant found');

    return data[0].tenantId;
  }

  // ================= FETCH =================
  function fetchJournals(service, tenantId, startDate) {
    let all = [];
    let offset = 0;
    const pageSize = 100;

    const fromDate = Utilities.formatDate(startDate, 'GMT', 'yyyy-MM-dd');

    while (true) {
      const url = `${XERO_API_BASE}/Journals?offset=${offset}`;

      const res = fetchWithRetry(url, {
        headers: {
          Authorization: 'Bearer ' + service.getAccessToken(),
          'xero-tenant-id': tenantId
        }
      });

      const data = JSON.parse(res.getContentText());

      if (!data.Journals || data.Journals.length === 0) break;

      all.push(...data.Journals);

      if (data.Journals.length < pageSize) break;

      offset += pageSize;
      Utilities.sleep(300);
    }

    log('Fetched journals', all.length);
    return all;
  }

  // ================= SYNC DATE =================
  function getStartDateFromSheet() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CONFIG.DATE_SHEET_NAME);

    const val = sheet.getRange(CONFIG.DATE_CELL).getValue();
    return new Date(val);
  }

  // ================= FILTER =================
  function filterTransactionsByTracking(service, tenantId, journals, name, value, startDate) {
    const rows = [];

    journals.forEach(j => {
      const journalNumber = j.JournalNumber || '';
      const sourceType = j.SourceType || '';
      const reference = j.Reference || '';
      const sourceID = j.SourceID || '';
      const date = parseXeroDate(j.JournalDate);
      if (date < startDate) return;

      (j.JournalLines || []).forEach(l => {
        if (CONFIG.EXCLUDED_ACCOUNT_CODES.includes(l.AccountCode)) return;

        // ✅ tracking match
        const tracking = l.TrackingCategories || [];

        const match = tracking.some(
          t => t.Name === name && t.Option === value
        );

        if (!match) return;

        const net = l.NetAmount || 0;
        const tax = l.TaxAmount || 0;
        const gross = net + tax;

        const debit = net > 0 ? net : 0;
        const credit = net < 0 ? Math.abs(net) : 0;

        // Format tracking nicely
        const tracking1 = tracking[0]
          ? `${tracking[0].Name}: ${tracking[0].Option}`
          : '';

        const tracking2 = tracking[1]
          ? `${tracking[1].Name}: ${tracking[1].Option}`
          : '';

        // Enrich with Product/Service (ItemCode) from source invoice
        const itemCodes = getItemCodes(service, tenantId, sourceType, sourceID);

        // Duplicate row per item code when multiple items exist on the source document
        itemCodes.forEach(code => {
          rows.push([
            date,
            journalNumber,
            reference,
            sourceType,
            sourceID,
            l.AccountCode,
            l.AccountName || '',
            code,              // Product/Service (ItemCode)
            l.Description || '',
            debit,
            credit,
            net,
            tax,
            gross,
            tracking1,
            tracking2
          ]);
        });
      });
    });

    return rows;
  }

  // ================= DATE =================
  function parseXeroDate(x) {
    if (!x.includes('/Date(')) return new Date(x);
    return new Date(parseInt(x.match(/\d+/)[0]));
  }

  // ================= SHEET =================
  function updateSheet(rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);

  sheet.clear();

  const headers = [[
    'Date',
    'Journal #',
    'Reference',
    'Source Type',
    'Source ID',
    'Account Code',
    'Account Name',
    'Product/Service',
    'Description',
    'Debit',
    'Credit',
    'Net Amount',
    'Tax Amount',
    'Gross Amount',
    'Tracking 1',
    'Tracking 2'
  ]];

  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers[0].length)
      .setValues(rows);
  }
}

  // ================= AUTH HELPERS =================
  function showAuthorizationUrl() {
    return getXeroService().getAuthorizationUrl();
  }

  function handleCallback(request) {
    return getXeroService().handleCallback(request);
  }

  function clearAuthorization() {
    getXeroService().reset();
  }

  return {
    updateXeroTransactions,
    showAuthorizationUrl,
    handleCallback,
    clearAuthorization
  };
}