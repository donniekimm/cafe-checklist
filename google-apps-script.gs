/**
 * Receives each closed day from the checklist app and adds it to a Google Sheet.
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

const SHEET_NAME = 'Sign-offs';
const HEADER = ['Date', 'Day', 'Shift', 'Task', 'Step', 'Employee', 'Time signed', 'Closed at'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const rows = JSON.parse(e.postData.contents).rows;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADER);
      sheet.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    // A day closed again after an owner's "Undo close" replaces that day's earlier rows.
    const date = rows.length && rows[0][0];
    if (date && sheet.getLastRow() > 1) {
      const dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
      // Delete from the bottom up, one block of matching rows at a time.
      for (let end = dates.length - 1; end >= 0; end--) {
        if (String(dates[end][0]) !== date) continue;
        let start = end;
        while (start > 0 && String(dates[start - 1][0]) === date) start--;
        sheet.deleteRows(start + 2, end - start + 1);
        end = start;
      }
    }
    if (rows.length) {
      // Plain text so Sheets doesn't turn dates and times into its own formats.
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADER.length)
        .setNumberFormat('@')
        .setValues(rows);
    }
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}
