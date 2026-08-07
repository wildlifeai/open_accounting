/**
 * XeroClient.js
 * OAuth2 connection to Xero and retrieval of actuals as normalised transaction lines.
 *
 * Reuses the apps-script-oauth2 library (declared in appsscript.json as `OAuth2`),
 * the same library the funding_reports script already depends on.
 *
 * A "normalised line" is the atomic unit the rest of the app reasons about:
 *   { date: Date, account: String, project: String, fundingSource: String,
 *     item: String, amount: Number, kind: 'expense' | 'income' }
 * `amount` is always positive; direction is carried by `kind`.
 */

/** Build (or fetch) the OAuth2 service for Xero. */
function getXeroService() {
  return OAuth2.createService('xero')
    .setAuthorizationBaseUrl(CONFIG.XERO.AUTH_URL)
    .setTokenUrl(CONFIG.XERO.TOKEN_URL)
    .setClientId(getSecret('XERO_CLIENT_ID'))
    .setClientSecret(getSecret('XERO_CLIENT_SECRET'))
    .setCallbackFunction('xeroAuthCallback')
    .setPropertyStore(PropertiesService.getScriptProperties())
    .setScope(CONFIG.XERO.SCOPE)
    .setParam('response_type', 'code')
    // Xero requires a fixed redirect URI registered on the app; getRedirectUri()
    // returns the script's callback URL — register exactly that value in Xero.
    .setTokenHeaders({
      'Authorization': 'Basic ' + Utilities.base64Encode(
        getSecret('XERO_CLIENT_ID') + ':' + getSecret('XERO_CLIENT_SECRET'))
    });
}

/** OAuth2 redirect handler (named in setCallbackFunction above). */
function xeroAuthCallback(request) {
  const ok = getXeroService().handleCallback(request);
  return HtmlService.createHtmlOutput(
    ok ? 'Xero connected. You can close this tab.' : 'Xero authorisation denied.');
}

/** The redirect URI to register in the Xero developer app. Run + read the log. */
function logXeroRedirectUri() {
  Logger.log('Register this redirect URI in your Xero app:\n' +
    getXeroService().getRedirectUri());
}

/** Log whether Xero is connected right now (run this from the editor to check). */
function checkXeroConnection() {
  const connected = isXeroConnected();
  Logger.log(connected
    ? 'Xero IS connected.'
    : 'Xero is NOT connected - run logXeroAuthUrl and open the URL to authorise.');
  return connected;
}

/** Log the Xero authorisation URL. Open it (as the deploying user) to connect. */
function logXeroAuthUrl() {
  const service = getXeroService();
  if (service.hasAccess()) { Logger.log('Already connected.'); return; }
  if (!getSecret('XERO_CLIENT_ID') || !getSecret('XERO_CLIENT_SECRET')) {
    Logger.log('Set XERO_CLIENT_ID and XERO_CLIENT_SECRET in Script Properties first.');
    return;
  }
  Logger.log('Open this URL (signed in as the account that DEPLOYED the web app) ' +
    'to connect Xero:\n' + service.getAuthorizationUrl());
}

/** Clear the stored Xero token so you can re-authorise from scratch. */
function resetXeroConnection() {
  getXeroService().reset();
  setSecret('XERO_TENANT_ID', '');
  Logger.log('Xero token cleared. Run logXeroAuthUrl to reconnect.');
}

/** True once a valid Xero token exists. */
function isXeroConnected() {
  return getXeroService().hasAccess();
}

/** Resolve and cache the Xero tenant (organisation) id. */
function getXeroTenantId() {
  let tenantId = getSecret('XERO_TENANT_ID');
  if (tenantId) return tenantId;

  const resp = UrlFetchApp.fetch(CONFIG.XERO.CONNECTIONS_URL, {
    headers: { Authorization: 'Bearer ' + getXeroService().getAccessToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error('Xero connections API ' + resp.getResponseCode() + ': ' + resp.getContentText());
  }
  const connections = JSON.parse(resp.getContentText());
  if (!connections.length) throw new Error('No Xero organisations connected to this app.');
  tenantId = connections[0].tenantId;
  setSecret('XERO_TENANT_ID', tenantId);
  return tenantId;
}

/**
 * Log every Xero organisation this app is connected to, with its tenant id.
 * Run from the editor AFTER a successful connection. Copy the tenantId into the
 * XERO_TENANT_ID Script Property only if you want to pin a specific org (e.g. if
 * the app is connected to more than one).
 */
function logXeroConnections() {
  if (!isXeroConnected()) {
    Logger.log('Not connected yet - authorise Xero first (logXeroAuthUrl).');
    return;
  }
  const resp = UrlFetchApp.fetch(CONFIG.XERO.CONNECTIONS_URL, {
    headers: { Authorization: 'Bearer ' + getXeroService().getAccessToken() },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 300) {
    Logger.log('Error fetching Xero connections (' + resp.getResponseCode() + '): ' + resp.getContentText());
    return;
  }
  const connections = JSON.parse(resp.getContentText());
  if (!connections.length) { Logger.log('No organisations connected.'); return; }
  connections.forEach(c => Logger.log(
    'Org: ' + c.tenantName + '  |  tenantId: ' + c.tenantId + '  |  type: ' + c.tenantType));
}

/** Low-level GET against the Xero Accounting API, returns parsed JSON. */
function xeroGet_(path, params) {
  const query = params
    ? '?' + Object.keys(params).map(k => k + '=' + encodeURIComponent(params[k])).join('&')
    : '';
  const resp = UrlFetchApp.fetch(CONFIG.XERO.API_BASE + path + query, {
    headers: {
      Authorization: 'Bearer ' + getXeroService().getAccessToken(),
      'Xero-tenant-id': getXeroTenantId(),
      Accept: 'application/json'
    },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() >= 300) {
    throw new Error('Xero API ' + resp.getResponseCode() + ': ' + resp.getContentText());
  }
  return JSON.parse(resp.getContentText());
}

/**
 * Pull all actuals since `sinceDate` and return normalised lines.
 * Sources: bank transactions (cash) + invoices (accrual). Extend with
 * ManualJournals if you book accruals/payroll via journals.
 */
function fetchXeroActuals(sinceDate) {
  const modifiedHeader = sinceDate ? Utilities.formatDate(
    sinceDate, 'UTC', "yyyy-MM-dd'T'HH:mm:ss") : null;
  const lines = [];

  lines.push.apply(lines, fetchBankTransactionLines_(modifiedHeader));
  lines.push.apply(lines, fetchInvoiceLines_(modifiedHeader));
  return lines;
}

/** Paginate a Xero endpoint and flatten line items via `mapper`. */
function paginate_(path, collectionKey, mapper, modifiedAfter) {
  const out = [];
  let page = 1;
  while (true) {
    const params = { page: page };
    if (modifiedAfter) params['where'] = 'UpdatedDateUTC>=DateTime(' +
      modifiedAfter.substring(0, 10).split('-').join(',') + ')';
    const data = xeroGet_(path, params);
    const rows = data[collectionKey] || [];
    if (!rows.length) break;
    rows.forEach(r => mapper(r, out));
    if (rows.length < 100) break; // Xero pages at 100
    page++;
  }
  return out;
}

function trackingValue_(tracking, categoryName) {
  if (!tracking) return '';
  const hit = tracking.filter(t => t.Name === categoryName)[0];
  return hit ? hit.Option : '';
}

function parseXeroDate_(value) {
  // Xero returns "/Date(1612137600000+0000)/"
  const m = /\/Date\((\d+)/.exec(value || '');
  return m ? new Date(parseInt(m[1], 10)) : null;
}

function fetchBankTransactionLines_(modifiedAfter) {
  return paginate_('/BankTransactions', 'BankTransactions', (tx, out) => {
    const date = parseXeroDate_(tx.DateString ? null : tx.Date) || new Date(tx.DateString);
    const kind = tx.Type === 'RECEIVE' ? 'income' : 'expense';
    (tx.LineItems || []).forEach(li => out.push(normaliseLine_(li, date, kind)));
  }, modifiedAfter);
}

function fetchInvoiceLines_(modifiedAfter) {
  return paginate_('/Invoices', 'Invoices', (inv, out) => {
    if (inv.Status === 'DELETED' || inv.Status === 'VOIDED') return;
    const date = new Date(inv.DateString || parseXeroDate_(inv.Date));
    const kind = inv.Type === 'ACCREC' ? 'income' : 'expense';
    (inv.LineItems || []).forEach(li => out.push(normaliseLine_(li, date, kind)));
  }, modifiedAfter);
}

/** Normalise a Xero line item into the app's actual-line shape. */
function normaliseLine_(li, date, kind) {
  // Prefer the product/service code; if absent, recover a {SOURCE}_{NNN} code
  // from the line description (some lines, e.g. bank fees, carry the code only
  // in the description). This keeps such amounts attached to their milestone.
  const code = li.Item ? li.Item.Code : codeFromDescription_(li.Description);
  return {
    date: date,
    account: accountLabelFromCode_(li.AccountCode),
    project: trackingValue_(li.Tracking, CONFIG.XERO.PROJECT_TRACKING_CATEGORY),
    fundingSource: trackingValue_(li.Tracking, CONFIG.XERO.FUNDING_TRACKING_CATEGORY),
    item: code,
    itemName: li.Item ? (li.Item.Name || '') : '',
    amount: Number(li.LineAmount) || 0,
    kind: kind
  };
}

/** Extract a leading {SOURCE}_{NNN}-style code from a description, or ''. */
function codeFromDescription_(desc) {
  const c = String(desc == null ? '' : desc).split(' - ')[0].trim();
  return /^[A-Za-z0-9]+_\d{2}_[A-Za-z0-9]+_\d{3}$/.test(c) ? c : '';
}

/**
 * Map a Xero account *code* (e.g. "500") to the "Name (code)" label used in budgets.
 * Cached per run from the Accounts endpoint so budgets and actuals share one key.
 */
let _accountLabelCache = null;
function accountLabelFromCode_(code) {
  if (code == null || code === '') return '';
  if (!_accountLabelCache) {
    _accountLabelCache = {};
    const data = xeroGet_('/Accounts');
    (data.Accounts || []).forEach(a => {
      _accountLabelCache[a.Code] = a.Name + ' (' + a.Code + ')';
    });
  }
  return _accountLabelCache[code] || ('(' + code + ')');
}
