/**
 * ScaleIn Growth - Lead form backend
 * Saves form submissions (Timestamp, Name, Email, Phone) to the "Responses" tab.
 *
 * SETUP (required every time you change this code):
 * 1. Open your Google Sheet > Extensions > Apps Script
 * 2. Replace everything in Code.gs with this file and click Save
 * 3. Deploy > Manage deployments > pencil icon (Edit)
 *      - Version: New version
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Click Deploy. (Editing the code alone does NOT update the live web app.)
 * 4. Test the URL in a private/incognito window.
 */

var SHEET_NAME = 'Responses';
var HEADERS = ['Timestamp', 'Name', 'Email', 'Phone'];

/**
 * Returns the "Responses" sheet, creating it with bold headers if missing.
 */
function getOrCreateResponsesSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange.setValues([HEADERS]);
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * Reads name/email/phone from the parameters and appends one row.
 * Returns true if saved, false if every field was empty.
 */
function saveResponse_(params) {
  var name = String(params.name || '').trim();
  var email = String(params.email || '').trim();
  var phone = String(params.phone || '').trim();

  if (!name && !email && !phone) {
    return false;
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = getOrCreateResponsesSheet_();
    sheet.appendRow([new Date(), name, email, phone]);
  } finally {
    lock.releaseLock();
  }

  return true;
}

/**
 * Returns a JSON response: {"success": true} or {"success": false}.
 */
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * GET: used by the website form.
 * Example: SCRIPT_URL?name=Asha&email=a@b.com&phone=%2B919876543210
 */
function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    return jsonResponse_({ success: saveResponse_(params) });
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err) });
  }
}

/**
 * POST: accepts form-encoded fields or a JSON body.
 */
function doPost(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};

    if (e && e.postData && e.postData.contents && !params.name) {
      try {
        params = JSON.parse(e.postData.contents) || params;
      } catch (parseErr) {
        // Not JSON, keep the form-encoded parameters
      }
    }

    return jsonResponse_({ success: saveResponse_(params) });
  } catch (err) {
    return jsonResponse_({ success: false, error: String(err) });
  }
}
