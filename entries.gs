// Reading + writing redeem entries

function getDialogData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sales Tracker") || ss.getActiveSheet();
  var salesPersons = getSalesPersons();

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var rowDate = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd") : "";
    if (rowDate === today) count++;
  }

  return { salesPersons: salesPersons, nextNo: count + 1 };
}

function parseEntryDate_(raw) {
  if (!raw) return new Date();
  var datePart = raw, timePart = "00:00";
  if (raw.indexOf("T") !== -1) {
    var parts = raw.split("T");
    datePart = parts[0];
    timePart = parts[1] || "00:00";
  }
  var d = datePart.split("-");
  var t = timePart.split(":");
  return new Date(
    parseInt(d[0], 10),
    parseInt(d[1], 10) - 1,
    parseInt(d[2], 10),
    parseInt(t[0], 10) || 0,
    parseInt(t[1], 10) || 0
  );
}

// Returns { name, entries: [{no, date, redeemType, package, trial, product, amount, paymentMethod, salesPerson, remark}] }
// — filtered to today + created by the PIN's user.
function getMyTodayEntries(pin) {
  var name = findUserByPin(pin);
  if (!name) throw new Error("Invalid PIN");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sales Tracker") || ss.getActiveSheet();
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var rowDate = Utilities.formatDate(new Date(row[0]), tz, "yyyy-MM-dd");
    if (rowDate !== today) continue;
    if (String(row[10] || "") !== name) continue;
    out.push({
      no:            row[1],
      date:          Utilities.formatDate(new Date(row[0]), tz, "yyyy-MM-dd'T'HH:mm"),
      redeemType:    row[2],
      package:       row[3],
      trial:         row[4],
      product:       row[5],
      amount:        row[6] === "" ? "" : row[6],
      paymentMethod: row[7],
      salesPerson:   row[8],
      remark:        row[9]
    });
  }
  return { name: name, entries: out };
}

// Finds today's row matching entry.no, updates columns except Date + No.
// Overwrites Created By with the editor's name.
function updateEntry(entry) {
  var editor = findUserByPin(entry.pin);
  if (!editor) throw new Error("Invalid PIN");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sales Tracker") || ss.getActiveSheet();
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    var rowDate = Utilities.formatDate(new Date(row[0]), tz, "yyyy-MM-dd");
    if (rowDate !== today) continue;
    if (parseInt(row[1], 10) !== parseInt(entry.no, 10)) continue;

    // Update columns C..K (3..11), leave A (date) + B (no) alone
    sheet.getRange(i + 1, 3, 1, 9).setValues([[
      entry.redeemType,
      entry.package,
      entry.trial,
      entry.product,
      entry.amount === "" ? "" : parseFloat(entry.amount),
      entry.paymentMethod,
      entry.salesPerson,
      entry.remark,
      editor
    ]]);
    return editor;
  }
  throw new Error("Entry No. " + entry.no + " not found for today");
}

function addEntry(entry) {
  var createdBy = findUserByPin(entry.pin);
  if (!createdBy) throw new Error("Invalid PIN");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sales Tracker") || ss.getActiveSheet();
  var date = parseEntryDate_(entry.date);
  sheet.appendRow([
    date,
    entry.no,
    entry.redeemType,
    entry.package,
    entry.trial,
    entry.product,
    entry.amount === "" ? "" : parseFloat(entry.amount),
    entry.paymentMethod,
    entry.salesPerson,
    entry.remark,
    createdBy
  ]);

  // Force standard format on the new date cell (appendRow auto-formats based on locale)
  sheet.getRange(sheet.getLastRow(), 1).setNumberFormat("yyyy-mm-dd HH:mm");

  return createdBy;
}
