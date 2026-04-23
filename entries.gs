// Reading + writing redeem entries

function getDialogData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
    entry.remark
  ]);
}
