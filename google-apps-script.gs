/**
 * Google Apps Script — receives form submissions from the GitHub Pages
 * registration site and appends each one as a row in a Google Sheet.
 *
 * SETUP (one time):
 *  1. Create a new Google Sheet. In the first row, add these headers
 *     (left to right):
 *       Timestamp | First Name | Last Name | Email |
 *       Certified Bodyshop Name | State | City | Zip |
 *       Network Member Confirmed | Consent Given
 *  2. In that Sheet: Extensions ▸ Apps Script. Delete any code and paste
 *     this whole file in. Save.
 *  3. Click Deploy ▸ New deployment ▸ type "Web app".
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     Click Deploy and authorise. Copy the "Web app URL".
 *  4. Paste that URL into index.html where it says SHEET_ENDPOINT.
 *
 * To get future submissions in a different sheet/tab, change SHEET_NAME.
 */

var SHEET_NAME = ''; // leave blank to use the first/active sheet

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Basic spam guard: if the hidden honeypot field is filled, ignore it.
    if (data.company) {
      return ok({ ignored: true });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];

    sheet.appendRow([
      new Date(),
      data.firstName || '',
      data.lastName  || '',
      data.email     || '',
      data.bodyshop  || '',
      data.state     || '',
      data.city      || '',
      data.zip       || '',
      data.confirm ? 'Yes' : 'No',
      data.consent ? 'Yes' : 'No'
    ]);

    return ok({ saved: true });
  } catch (err) {
    return ok({ error: String(err) });
  }
}

function ok(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
