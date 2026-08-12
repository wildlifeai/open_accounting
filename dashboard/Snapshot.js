/**
 * Snapshot.js
 * Caches the aggregated snapshot so the dashboard loads instantly and a single
 * Xero/Drive crawl is shared by every viewer. The crawl is the slow part and can
 * approach Apps Script's 6-minute limit, so it runs on a time trigger, not on
 * page load.
 *
 * Storage: a single JSON file in Drive (no size ceiling, unlike Script
 * Properties which cap at ~500KB total). The file id is kept in a Script
 * Property so it's found again across executions.
 */

// ---- Refresh progress ------------------------------------------------------
// A refresh takes about a minute, almost all of it opening one spreadsheet at a time
// (SpreadsheetApp.openById costs 1-3 seconds each) and paging Xero. Apps Script cannot
// stream to a client, so the current phase is written to CacheService here and the
// browser polls apiRefreshProgress for it.
//
// Progress is never worth failing a refresh for, so every call below swallows errors.

/**
 * Measured phase durations for a 9-source refresh, which is where the bands in
 * PROGRESS_BANDS come from: folder walk ~13s, reading ~14s, Xero ~30s. Each call site
 * reports its own band so the bar only ever moves forward. Without that, a phase with
 * no countable total (Xero pages) reset the bar to zero and looked like a regression.
 *
 * @param {string} phase   human-readable, shown verbatim in the UI
 * @param {number=} done   units completed, so the bar is real rather than an animation
 * @param {number=} total  0 when the total is not yet known
 * @param {number=} pct    overall completion, 0-100. Preferred by the client over done/total.
 */
function setRefreshProgress_(phase, done, total, pct) {
  try {
    CacheService.getScriptCache().put(
      CONFIG.PROGRESS_CACHE_KEY,
      JSON.stringify({ phase: phase, done: done || 0, total: total || 0,
                       pct: pct === undefined ? null : pct,
                       at: new Date().getTime() }),
      CONFIG.PROGRESS_CACHE_TTL_SECONDS);
  } catch (e) { /* ignore */ }
}

function clearRefreshProgress_() {
  try { CacheService.getScriptCache().remove(CONFIG.PROGRESS_CACHE_KEY); } catch (e) { /* ignore */ }
}

/** The current phase, or null when nothing is running. */
function readRefreshProgress_() {
  try {
    const raw = CacheService.getScriptCache().get(CONFIG.PROGRESS_CACHE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}

/** Rebuild the snapshot from source and persist it. Run by the time trigger. */
function refreshSnapshot() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // a refresh is already running
  try {
    setRefreshProgress_('Starting', 0, 0, 1);
    const snapshot = buildSnapshot();
    setRefreshProgress_('Saving snapshot', 0, 0, 96);
    storeSnapshot_(snapshot);
    return snapshot;
  } finally {
    // Cleared even on failure, so a dead run cannot leave a bar stuck on screen.
    clearRefreshProgress_();
    lock.releaseLock();
  }
}

/** Return the cached snapshot, rebuilding once if none exists yet. */
function getSnapshot() {
  const cached = readSnapshot_();
  return cached || refreshSnapshot();
}

function snapshotFile_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(CONFIG.SNAPSHOT_FILE_ID_PROPERTY);
  if (id) {
    try { return DriveApp.getFileById(id); } catch (e) { /* recreate below */ }
  }
  const file = DriveApp.createFile(CONFIG.SNAPSHOT_FILE_NAME, '{}', 'application/json');
  props.setProperty(CONFIG.SNAPSHOT_FILE_ID_PROPERTY, file.getId());
  return file;
}

function storeSnapshot_(snapshot) {
  snapshotFile_().setContent(JSON.stringify(snapshot));
}

function readSnapshot_() {
  const content = snapshotFile_().getBlob().getDataAsString();
  if (!content || content === '{}') return null;
  try {
    const obj = JSON.parse(content);
    return obj && obj.generatedAt ? obj : null;
  } catch (e) {
    return null;
  }
}

/** Install the recurring refresh trigger (idempotent). */
function installRefreshTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'refreshSnapshot') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('refreshSnapshot')
    .timeBased().everyHours(CONFIG.REFRESH_TRIGGER_HOURS).create();
}
