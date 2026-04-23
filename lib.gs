// Shared helpers — sheet access, column map, date/hex utilities
// Keep this file thin: only things used by 2+ other files belong here.

var SHEET_NAME = "Sales Tracker";

// 1-indexed column numbers matching the Sales Tracker layout.
// When row data is read via getValues(), use COL.X - 1 to index the array.
var COL = {
  DATE:           1,
  NO:             2,
  REDEEM_TYPE:    3,
  PACKAGE:        4,
  TRIAL:          5,
  PRODUCT:        6,
  AMOUNT:         7,
  PAYMENT_METHOD: 8,
  SALES_PERSON:   9,
  REMARK:         10,
  CREATED_BY:     11
};
var NUM_COLS = 11;

function getSalesSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SHEET_NAME) || ss.getActiveSheet();
}

function getTz_() { return Session.getScriptTimeZone(); }

function todayStr_() {
  return Utilities.formatDate(new Date(), getTz_(), "yyyy-MM-dd");
}

function rowDateStr_(raw) {
  return raw ? Utilities.formatDate(new Date(raw), getTz_(), "yyyy-MM-dd") : "";
}

// Iterate today's rows on Sales Tracker.
// cb(row, rowIndex1Based, sheet) — return true from cb to stop iteration.
function forEachTodayRow_(cb) {
  var sheet = getSalesSheet_();
  var data = sheet.getDataRange().getValues();
  var today = todayStr_();
  for (var i = 1; i < data.length; i++) {
    if (rowDateStr_(data[i][COL.DATE - 1]) !== today) continue;
    if (cb(data[i], i + 1, sheet) === true) return;
  }
}

// Signed-byte array → lowercase hex string.
// Works for both positive numbers (random bytes) and signed bytes (HMAC output).
function bytesToHex_(bytes) {
  return bytes.map(function(b) {
    var v = b < 0 ? b + 256 : b;
    return ("0" + v.toString(16)).slice(-2);
  }).join("");
}

function setDropdown(sheet, range, options, allowInvalid) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(allowInvalid === true)
    .build();
  sheet.getRange(range).setDataValidation(rule);
}
