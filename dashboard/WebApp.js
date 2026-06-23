/**
 * WebApp.js
 * Web-app entry point and the server-side functions the client calls via
 * google.script.run. Also a spreadsheet menu for setup/admin if this script is
 * bound to a Gsheet (optional — it runs standalone too).
 */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Wildlife.ai Funding Cockpit')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Allow Index.html to pull in Stylesheet.html / JavaScript.html partials. */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** Client API: the cached snapshot (fast). */
function apiGetSnapshot() {
  return getSnapshot();
}

/** Client API: force a rebuild ("Refresh now" button). */
function apiRefresh() {
  return refreshSnapshot();
}

/**
 * Diagnostic (run from the editor): rebuild the snapshot and log how many
 * breakdown rows / projects / sources it produced, plus a sample row. Confirms
 * the data side independently of the deployed web app + cache.
 */
function diagBreakdown() {
  const s = refreshSnapshot();
  Logger.log('breakdownRows: ' + (s.breakdownRows ? s.breakdownRows.length : 'MISSING'));
  Logger.log('projects: ' + (s.projects ? s.projects.length : 0) +
    ' | fundingSources: ' + (s.fundingSources ? s.fundingSources.length : 0));
  if (s.breakdownRows && s.breakdownRows.length) {
    Logger.log('sample row: ' + JSON.stringify(s.breakdownRows[0]));
  }
  return s.breakdownRows ? s.breakdownRows.length : -1;
}

/** Client API: is Xero connected, and the auth URL if not. */
function apiXeroStatus() {
  const service = getXeroService();
  return { connected: service.hasAccess(),
    authUrl: service.hasAccess() ? null : service.getAuthorizationUrl() };
}

/** Client API: the list of funding sources for the tracking selector. */
function apiListSources() {
  const snap = getSnapshot();
  return (snap.tracking || []).map(t => ({
    source: t.source, status: t.status, project: t.project }));
}

/**
 * Client API: the quarterly tracking grid for one funding source, with the live
 * forecast layered on (so GM edits show without a full refresh).
 */
function apiGetTracking(sourceName) {
  const snap = getSnapshot();
  const trackingSource = (snap.tracking || []).filter(t => t.source === sourceName)[0];
  if (!trackingSource) throw new Error('Unknown funding source: ' + sourceName);
  return composeTrackingForSource(
    trackingSource, getForecastMap(), snap.currentQuarter || currentQuarterLabel());
}

/**
 * Client API: save one forecast cell. `cost` of null clears the override (reverts
 * to baseline). Returns the recomposed grid for the source so the UI can refresh
 * totals without another round trip.
 */
function apiSaveForecast(sourceName, milestone, item, quarter, cost) {
  upsertForecast(sourceName, milestone, item, quarter, cost);
  return apiGetTracking(sourceName);
}

// ---- Admin menu (only appears when bound to a spreadsheet) -----------------

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Cockpit')
      .addItem('Connect Xero', 'menuConnectXero')
      .addItem('Show Xero redirect URI', 'menuShowRedirectUri')
      .addItem('Refresh snapshot now', 'refreshSnapshot')
      .addItem('Install auto-refresh trigger', 'installRefreshTrigger')
      .addToUi();
  } catch (e) { /* not bound to a sheet — ignore */ }
}

function menuConnectXero() {
  const url = getXeroService().getAuthorizationUrl();
  SpreadsheetApp.getUi().alert('Open this URL to connect Xero:\n\n' + url);
}

function menuShowRedirectUri() {
  SpreadsheetApp.getUi().alert('Register this redirect URI in your Xero app:\n\n' +
    getXeroService().getRedirectUri());
}
