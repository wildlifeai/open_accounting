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

/** Rebuild the snapshot from source and persist it. Run by the time trigger. */
function refreshSnapshot() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return; // a refresh is already running
  try {
    const snapshot = buildSnapshot();
    storeSnapshot_(snapshot);
    return snapshot;
  } finally {
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
