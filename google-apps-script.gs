/**
 * Google Apps Script — receives form submissions from the GitHub Pages
 * registration site and appends each one as a row in a Google Sheet.
 * Headers are styled for readability and incoming rows are auto-shaded
 * with alternating light-grey / white banding.
 *
 * SETUP (one time):
 *  1. Open the Google Sheet, then: Extensions ▸ Apps Script.
 *     Delete any code and paste this whole file in. Save.
 *  2. Restrict permissions to THIS sheet only (recommended):
 *       - Project Settings (gear icon) ▸ tick
 *         "Show appsscript.json manifest file in editor".
 *       - Open appsscript.json and replace its contents with the
 *         appsscript.json from this repo. Save.
 *     This makes the consent screen grant access to ONLY this
 *     spreadsheet (scope: spreadsheets.currentonly) instead of all
 *     of your spreadsheets.
 *  3. (Optional) Select "setupSheet" in the function dropdown and
 *     click Run once to format the header immediately.
 *  4. Click Deploy ▸ New deployment ▸ type "Web app".
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     (Anyone is required so the public form can post. The endpoint
 *     only runs doPost, which can ONLY append a row — it cannot read
 *     or delete anything.)
 *     Click Deploy and authorise. Copy the "Web app URL".
 *  5. Paste that URL into index.html where it says SHEET_ENDPOINT.
 *
 * To use a different sheet/tab, change SHEET_NAME.
 */

var SHEET_NAME = '';            // blank = first/active sheet
var NUM_COLS   = 9;             // number of columns in the table

var HEADER_BG   = '#202124';    // dark header background
var HEADER_TEXT = '#ffffff';    // white header text
var ROW_LIGHT   = '#f1f3f4';    // light grey row
var ROW_WHITE   = '#ffffff';    // white row
var ROW_TEXT    = '#202124';    // dark row text

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Basic spam guard: if the hidden honeypot field is filled, ignore it.
    if (data.company) {
      return ok({ ignored: true });
    }

    // Lightweight validation: require a plausible email and cap field
    // sizes so a bad actor can't bloat the sheet with huge payloads.
    var email = clean_(data.email, 120);
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
      return ok({ rejected: 'invalid email' });
    }

    var sheet = getSheet_();
    formatHeader_(sheet);

    sheet.appendRow([
      new Date(),
      clean_(data.bodyshop,  120),
      clean_(data.firstName, 80),
      clean_(data.lastName,  80),
      email,
      clean_(data.state,     40),
      clean_(data.oem,       400),   // OEM certifications (comma-separated)
      clean_(data.drp,       400),   // DRP programs (comma-separated)
      data.confirm ? 'Yes' : 'No'
    ]);

    styleRow_(sheet, sheet.getLastRow());

    return ok({ saved: true });
  } catch (err) {
    return ok({ error: String(err) });
  }
}

/** Lets you test the endpoint by opening its URL in a browser.
 *  If you see {"status":"ok"} the web app is publicly reachable.
 *  If you see a Google sign-in page instead, the deployment is
 *  restricted to your organisation — re-deploy with access "Anyone". */
function doGet() {
  return ok({ status: 'ok' });
}

/** Run this once from the editor to name + format the tab straight away. */
function setupSheet() {
  formatHeader_(getSheet_());
}

/** Trim a value to a string and cap its length. */
function clean_(v, max) {
  return String(v == null ? '' : v).trim().slice(0, max);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (SHEET_NAME) return ss.getSheetByName(SHEET_NAME);
  var sheet = ss.getSheetByName('Submissions');
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName('Submissions');
  }
  return sheet;
}

/** Bold, dark, frozen, easy-to-read header row. */
function formatHeader_(sheet) {
  var header = sheet.getRange(1, 1, 1, NUM_COLS);
  // Write the column labels (only if the header row is still empty).
  if (!String(sheet.getRange(1, 1).getValue()).trim()) {
    header.setValues([[
      'Submitted', 'Bodyshop', 'First Name', 'Last Name', 'Email',
      'State', 'OEM Certifications', 'DRP Programs', 'Confirmed'
    ]]);
  }
  header
    .setBackground(HEADER_BG)
    .setFontColor(HEADER_TEXT)
    .setFontWeight('bold')
    .setFontSize(11)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left')
    .setWrap(true);
  sheet.setRowHeight(1, 38);
  sheet.setFrozenRows(1);

  // Reasonable starting column widths for legibility.
  var widths = [150, 220, 110, 110, 230, 110, 260, 260, 110];
  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }
}

/** Alternating light-grey / white shading for each incoming data row. */
function styleRow_(sheet, row) {
  if (row < 2) return;
  var isGrey = (row % 2 === 0); // row 2 grey, row 3 white, row 4 grey ...
  sheet.getRange(row, 1, 1, NUM_COLS)
    .setBackground(isGrey ? ROW_LIGHT : ROW_WHITE)
    .setFontColor(ROW_TEXT)
    .setVerticalAlignment('middle');
}

function ok(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
