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

function addEntry(entry) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var date = new Date(entry.date.indexOf("T") === -1 ? entry.date + "T00:00:00" : entry.date);
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
