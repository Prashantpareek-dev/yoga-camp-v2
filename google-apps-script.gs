/**
 * Veda Yoga Ashram — Registration form -> Google Sheet
 *
 * Setup:
 * 1. Create (or open) the Google Sheet you want registrations saved to.
 * 2. Extensions > Apps Script, delete any boilerplate code, and paste this file's contents.
 * 3. Deploy > New deployment > type "Web app".
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy the deployment URL and paste it into GOOGLE_SHEETS_URL in index.html.
 * 5. Re-deploy (New deployment, not just Save) any time you edit this file, otherwise
 *    the live URL keeps running the old code.
 */

// Must match the <input name="..."> / <select name="..."> attributes on #regForm in index.html.
var FIELDS = [
  { key: "fullName", header: "Full Name" },
  { key: "mobile", header: "Mobile Number" },
  { key: "stateCity", header: "State / City" },
  { key: "age", header: "Age" },
  { key: "gender", header: "Gender" },
  { key: "campDate", header: "Camp Date" },
  { key: "participants", header: "Participants" },
  { key: "message", header: "Message / Preferences" },
];

function doPost(e) {
  try {
    var params = (e && e.parameter) || {};
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    ensureHeaderRow_(sheet);

    var row = [new Date()];
    FIELDS.forEach(function (field) {
      row.push(params[field.key] || "");
    });
    sheet.appendRow(row);

    return jsonResponse_({ result: "success" });
  } catch (err) {
    return jsonResponse_({ result: "error", message: err.message });
  }
}

function doGet() {
  return jsonResponse_({ result: "error", message: "This endpoint only accepts POST requests." });
}

function ensureHeaderRow_(sheet) {
  if (sheet.getLastRow() > 0) return;
  var headers = ["Timestamp"].concat(FIELDS.map(function (field) { return field.header; }));
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  sheet.setFrozenRows(1);
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
