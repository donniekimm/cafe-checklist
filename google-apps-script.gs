/**
 * Receives the weekly report from the checklist app. Each week (Monday to Sunday) gets its own tab.
 *
 * Setup (one time):
 *   1. Create a new Google Sheet (e.g. "Cafe Sign-offs").
 *   2. In the sheet: Extensions → Apps Script. Delete what's there and paste this whole file. Save.
 *   3. Deploy → New deployment → type "Web app".
 *        Execute as: Me
 *        Who has access: Anyone
 *      Click Deploy and allow access when Google asks.
 *   4. Copy the Web app URL (ends in /exec) and paste it into SHEET_URL in index.html.
 *
 * If you change this script later, use Deploy → Manage deployments → Edit → New version,
 * so the URL stays the same.
 */

const HEADER = ['Date', 'Day', 'Shift', 'Task', 'Step', 'Employee', 'Time signed', 'Closed at'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const { tab, rows } = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Newest week first. A week sent again after an owner's "Undo finish" replaces its tab.
    const sheet = ss.getSheetByName(tab) || ss.insertSheet(tab, 0);
    sheet.clearContents();
    sheet.getRange(1, 1, 1, HEADER.length).setValues([HEADER]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    if (rows.length) {
      // Plain text so Sheets doesn't turn dates and times into its own formats.
      sheet.getRange(2, 1, rows.length, HEADER.length)
        .setNumberFormat('@')
        .setValues(rows);
    }
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
