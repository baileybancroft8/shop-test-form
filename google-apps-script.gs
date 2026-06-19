/**
 * Google Apps Script — receives form submissions from the GitHub Pages
 * registration site and appends each one as a row in a Google Sheet.
 * Headers are styled for readability and incoming rows are auto-shaded
 * with alternating light-grey / white banding.
 *
 * SETUP (one time):
 *  1. Open the Google Sheet, then: Extensions ▸ Apps Script.
 *     Delete any code and paste this whole file in. Save.
 *  2. (Optional) In the editor, select "setupSheet" in the function
 *     dropdown and click Run once to format the header immediately.
 *  3. Click Deploy ▸ New deployment ▸ type "Web app".
 *       - Execute as: Me
 *       - Who has access: Anyone
 *     Click Deploy and authorise. Copy the "Web app URL".
 *  4. Paste that URL into index.html where it says SHEET_ENDPOINT.
 *
 * To use a different sheet/tab, change SHEET_NAME.
 */

var SHEET_NAME = '';            // blank = first/active sheet
var NUM_COLS   = 10;            // number of columns in the table

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

    var sheet = getSheet_();
    formatHeader_(sheet);

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

    styleRow_(sheet, sheet.getLastRow());

    return ok({ saved: true });
  } catch (err) {
    return ok({ error: String(err) });
  }
}

/** Run this once from the editor to format the header straight away. */
function setupSheet() {
  formatHeader_(getSheet_());
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getSheets()[0];
}

/** Bold, dark, frozen, easy-to-read header row. */
function formatHeader_(sheet) {
  var header = sheet.getRange(1, 1, 1, NUM_COLS);
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
  var widths = [150, 110, 110, 220, 220, 120, 130, 90, 150, 130];
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
