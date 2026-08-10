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

/**
 * Client API: the entities selectable in the tracking dropdown.
 * The General project (aggregated across funding sources) is listed first,
 * followed by each funding source. Each entry: { id, label, type }.
 */
function apiListSources() {
  const snap = getSnapshot();
  const entries = [];
  const hasGeneral = (snap.tracking || []).some(t =>
    (t.milestones || []).some(m => m.project === CONFIG.GENERAL_PROJECT));
  if (hasGeneral) {
    entries.push({ id: 'project:' + CONFIG.GENERAL_PROJECT,
      label: CONFIG.GENERAL_PROJECT + ' (project)', type: 'project' });
  }
  (snap.tracking || []).forEach(t => entries.push({
    id: t.source, label: t.source + ' (' + t.status + ' · ' + t.project + ')',
    type: 'source' }));
  return entries;
}

/**
 * Client API: the quarterly tracking grid for one entity id (a funding source
 * name, or 'project:<Name>'), with the live forecast layered on. Or an array of ids.
 */
function apiGetTracking(ids, measure) {
  const snap = getSnapshot();
  const currentQi = quarterSortNum(snap.currentQuarter || currentQuarterLabel());
  return composeTracking(resolveEntity_(snap, ids), getForecastMap(), currentQi, measure);
}

/** Build the entity (source or aggregated project) the tracking grid renders. */
function resolveEntity_(snap, ids) {
  if (!Array.isArray(ids)) ids = [ids];
  const tracking = snap.tracking || [];
  
  if (ids.length === 1 && ids[0].indexOf('project:') === 0) {
    const projectName = ids[0].substring('project:'.length);
    const milestones = [];
    tracking.forEach(t => (t.milestones || []).forEach(m => {
      if (m.project === projectName) milestones.push(m);
    }));
    if (!milestones.length) throw new Error('No milestones for project: ' + projectName);
    return { id: ids[0], label: projectName + ' (project)', type: 'project',
      project: projectName, milestones: milestones };
  }
  
  if (ids.length === 1) {
    const src = tracking.filter(t => t.source === ids[0])[0];
    if (!src) throw new Error('Unknown funding source: ' + ids[0]);
    return { id: ids[0], label: ids[0], type: 'source', source: ids[0], status: src.status,
      project: src.project, milestones: src.milestones };
  }
  
  // Multiple sources selected
  const milestones = [];
  ids.forEach(id => {
    const src = tracking.filter(t => t.source === id)[0];
    if (src && src.milestones) {
      milestones.push(...src.milestones);
    }
  });
  
  return { id: ids.join(','), label: 'Multiple sources selected', type: 'composite',
    project: 'Multiple', milestones: milestones };
}

/**
 * Client API: save one forecast amount cell for a measure ('cost' | 'income').
 * A null value clears that measure. Returns the recomposed grid (same measure).
 */
function apiSaveForecast(entityIds, source, milestone, item, quarter, measure, value) {
  upsertForecast(source, milestone, item, quarter, measure, value);
  return apiGetTracking(entityIds, measure);
}

/** Client API: save a milestone's forecast Comment. Returns the recomposed grid. */
function apiSaveComment(entityIds, source, milestone, item, comment, measure) {
  upsertForecastComment(source, milestone, item, comment);
  return apiGetTracking(entityIds, measure);
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
