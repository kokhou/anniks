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

  notifyWhatsApp_(entry, createdBy);

  return createdBy;
}
