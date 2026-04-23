// PIN management — stored in a hidden "Users" sheet.
// Layout: A1:B1 headers "Name" | "PIN Hash"  ·  A2:B… rows  ·  D1 = HMAC key

var USERS_SHEET_NAME = "Users";

function getUsersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    sheet.getRange("A1:B1")
      .setValues([["Name", "PIN Hash"]])
      .setFontWeight("bold").setBackground("#4a86e8").setFontColor("#ffffff");
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 420);
    sheet.hideSheet();
  }
  return sheet;
}

function getEncryptionKey_() {
  var cell = getUsersSheet_().getRange("D1");
  var key = cell.getValue();
  if (!key) {
    var bytes = [];
    for (var i = 0; i < 32; i++) bytes.push(Math.floor(Math.random() * 256));
    key = bytesToHex_(bytes);
    cell.setValue(key);
  }
  return String(key);
}

function hashPin_(pin, key) {
  return bytesToHex_(Utilities.computeHmacSha256Signature(String(pin), String(key)));
}

// Returns 1-based row index of the first row whose Name column matches (case-insensitive),
// or -1 if not found.
function findUserRowIndex_(sheet, name) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var data = sheet.getRange(2, 1, last - 1, 1).getValues();
  var target = String(name).toLowerCase();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === target) return i + 2;
  }
  return -1;
}

function findUserByPin(pin) {
  if (!pin) return null;
  var sheet = getUsersSheet_();
  var target = hashPin_(pin, getEncryptionKey_());
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var data = sheet.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] && String(data[i][1]) === target) return String(data[i][0]);
  }
  return null;
}

function upsertUserPin(name, pin) {
  name = String(name || "").trim();
  pin  = String(pin  || "").trim();
  if (!name) throw new Error("Name is required");
  if (!/^\d{4}$/.test(pin)) throw new Error("PIN must be 4 digits");

  var sheet = getUsersSheet_();
  var hash = hashPin_(pin, getEncryptionKey_());
  var rowIdx = findUserRowIndex_(sheet, name);

  if (rowIdx !== -1) {
    sheet.getRange(rowIdx, 1, 1, 2).setValues([[name, hash]]);
    return "updated";
  }
  sheet.appendRow([name, hash]);
  return "created";
}

function removeUserPin(name) {
  name = String(name || "").trim();
  if (!name) return false;
  var sheet = getUsersSheet_();
  var rowIdx = findUserRowIndex_(sheet, name);
  if (rowIdx === -1) return false;
  sheet.deleteRow(rowIdx);
  return true;
}

function listUsersWithPin() {
  var sheet = getUsersSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, 1).getValues()
    .map(function(r) { return String(r[0]); })
    .filter(Boolean);
}
