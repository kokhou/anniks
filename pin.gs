// PIN management — stored in hidden "Users" sheet
// A1:B1 headers "Name" | "PIN Hash"  · A2:B... data · D1 holds encryption key

var USERS_SHEET_NAME = "Users";

function getUsersSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    sheet.getRange("A1:B1").setValues([["Name", "PIN Hash"]]);
    sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#4a86e8").setFontColor("#ffffff");
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 420);
    sheet.hideSheet();
  }
  return sheet;
}

function getEncryptionKey_() {
  var sheet = getUsersSheet_();
  var cell = sheet.getRange("D1");
  var key = cell.getValue();
  if (!key) {
    var bytes = [];
    for (var i = 0; i < 32; i++) bytes.push(Math.floor(Math.random() * 256));
    key = bytes.map(function(b) { return ("0" + b.toString(16)).slice(-2); }).join("");
    cell.setValue(key);
  }
  return String(key);
}

function hashPin_(pin, key) {
  var sig = Utilities.computeHmacSha256Signature(String(pin), String(key));
  return sig.map(function(b) {
    var v = b < 0 ? b + 256 : b;
    return ("0" + v.toString(16)).slice(-2);
  }).join("");
}

function findUserByPin(pin) {
  if (!pin) return null;
  var sheet = getUsersSheet_();
  var key = getEncryptionKey_();
  var target = hashPin_(pin, key);
  var last = sheet.getLastRow();
  if (last < 2) return null;
  var data = sheet.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][1] && String(data[i][1]) === target) {
      return String(data[i][0]);
    }
  }
  return null;
}

function upsertUserPin(name, pin) {
  name = String(name || "").trim();
  pin = String(pin || "").trim();
  if (!name) throw new Error("Name is required");
  if (!/^\d{4}$/.test(pin)) throw new Error("PIN must be 4 digits");

  var sheet = getUsersSheet_();
  var key = getEncryptionKey_();
  var hash = hashPin_(pin, key);

  var last = sheet.getLastRow();
  if (last >= 2) {
    var data = sheet.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).toLowerCase() === name.toLowerCase()) {
        sheet.getRange(i + 2, 1, 1, 2).setValues([[name, hash]]);
        return "updated";
      }
    }
  }
  sheet.appendRow([name, hash]);
  return "created";
}

function removeUserPin(name) {
  name = String(name || "").trim();
  if (!name) return false;
  var sheet = getUsersSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return false;
  var data = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === name.toLowerCase()) {
      sheet.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

function listUsersWithPin() {
  var sheet = getUsersSheet_();
  var last = sheet.getLastRow();
  if (last < 2) return [];
  var data = sheet.getRange(2, 1, last - 1, 1).getValues();
  return data.map(function(r) { return String(r[0]); }).filter(Boolean);
}
