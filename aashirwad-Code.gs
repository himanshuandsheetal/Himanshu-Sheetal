/**
 * Aashirwad wall for the Himanshu & Sheetal wedding website.
 * Free: runs in your own Google account and saves every wish into a Google Sheet.
 *
 * ONE-TIME SETUP (about 10 minutes)
 *  1. Open https://sheets.new (a new Google Sheet). Name it "Aashirwad".
 *  2. In the Sheet: Extensions > Apps Script. Delete what is there, paste this whole file, click Save.
 *  3. Click Deploy > New deployment > the gear icon > Web app.
 *       Execute as: Me
 *       Who has access: Anyone
 *     Click Deploy, then Authorize access and allow it with your Google account
 *     (if you see "Google hasn't verified this app": Advanced > Go to project > Allow).
 *  4. Copy the Web app URL (it ends with /exec).
 *  5. In index.html find  const AASHIRWAD_URL='';  and paste it between the quotes.
 *  6. Upload index.html to Netlify again. Done.
 *
 * DAILY USE
 *  - Every wish appears as a new row in the "Aashirwad" tab.
 *  - To hide a wish from the website, untick its "Show" box (or delete the row).
 *  - If you change this code later: Deploy > Manage deployments > Edit > Version: New version > Deploy
 *    (the URL stays the same).
 */

const SHEET_NAME = 'Aashirwad';
const MAX_WISHES = 300;

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Time', 'Name', 'Relation', 'Message', 'Show']);
    sh.setFrozenRows(1);
    sh.getRange('A1:E1').setFontWeight('bold');
    sh.setColumnWidth(4, 420);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Trim, squash spaces, cap length, and stop spreadsheet formulas
function clean_(value, max) {
  let s = String(value || '').replace(/[<>]/g, '').replace(/\s+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

// Website reads the wall: newest first, only rows with Show ticked
function doGet() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return json_({ ok: true, wishes: [] });
  const rows = sh.getRange(2, 1, last - 1, 5).getValues();
  const wishes = rows
    .filter(r => r[3] && r[4] !== false && String(r[4]).toUpperCase() !== 'FALSE')
    .map(r => ({ t: new Date(r[0]).getTime(), name: String(r[1]), rel: String(r[2]), msg: String(r[3]) }))
    .reverse()
    .slice(0, MAX_WISHES);
  return json_({ ok: true, wishes: wishes });
}

// A guest sends a wish
function doPost(e) {
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (d.website) return json_({ ok: true });             // hidden spam trap was filled: ignore quietly
    const name = clean_(d.name, 40), rel = clean_(d.rel, 30), msg = clean_(d.msg, 300);
    if (!name || !msg) return json_({ ok: false, error: 'missing' });

    // Ignore the same wish sent twice within 10 minutes (double taps)
    const cache = CacheService.getScriptCache();
    const key = 'w' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, name + '|' + msg, Utilities.Charset.UTF_8));
    if (cache.get(key)) return json_({ ok: true, duplicate: true });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sh = sheet_();
      sh.appendRow([new Date(), name, rel, msg, true]);
      sh.getRange(sh.getLastRow(), 5).insertCheckboxes().setValue(true);
    } finally {
      lock.releaseLock();
    }
    cache.put(key, '1', 600);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  }
}
