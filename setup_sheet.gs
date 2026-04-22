// Runs on every sheet open — creates the custom menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚙️ Manage")
    .addItem("➕ New Redeem Entry", "showRedeemDialog")
    .addSeparator()
    .addItem("Add Sales Person", "showAddSalesPersonDialog")
    .addItem("Remove Sales Person", "showRemoveSalesPersonDialog")
    .addToUi();
}

// ── Redeem Entry Dialog (desktop menu) ──

function showRedeemDialog() {
  var html = HtmlService.createHtmlOutputFromFile("dialog")
    .setWidth(420)
    .setHeight(580);
  SpreadsheetApp.getUi().showModalDialog(html, "New Redeem Entry");
}

// ── Web App entry point (mobile) ──

function doGet() {
  return HtmlService.createHtmlOutputFromFile("dialog")
    .setTitle("New Redeem Entry")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDialogData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var salesPersons = getSalesPersons();

  // Auto-calculate next No for today
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
  var date = new Date(entry.date + "T00:00:00");
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

function setupSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setName("Sales Tracker");

  // ── Headers ──
  var headers = ["Date", "No", "Redeem Type", "Package", "Trial", "Product", "Amount", "Payment Method", "Sales Person", "Remark"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Style headers
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#4a86e8");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");

  // ── Dropdowns (rows 2–1000) ──
  setDropdown(sheet, "C2:C1000", ["New", "Existing"]);

  setDropdown(sheet, "D2:D1000", [
    "P6880 脸部塑型",
    "P6880 开肩",
    "P6880 体态",
    "P6880 祈龄",
    "P6880 局部",
    "P4880 高级波肽",
    "Gold 脸部塑型",
    "Gold 开肩",
    "T2388 小腿"
  ]);

  setDropdown(sheet, "E2:E1000", ["Yes", "No"]);

  setDropdown(sheet, "F2:F1000", [
    "T388 脸部塑型",
    "T388 祈龄魔法",
    "T298 体态",
    "Firming Cream"
  ]);

  setDropdown(sheet, "H2:H1000", ["Cash", "Card", "Online Transfer", "Debit Card", "Credit Card", "QR", "Transfer"]);

  setDropdown(sheet, "I2:I1000", getSalesPersons(), true);

  // ── Column widths ──
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 50);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 60);
  sheet.setColumnWidth(6, 160);
  sheet.setColumnWidth(7, 90);
  sheet.setColumnWidth(8, 130);
  sheet.setColumnWidth(9, 110);
  sheet.setColumnWidth(10, 180);

  sheet.setFrozenRows(1);

  // Save default sales persons to script properties
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty("salesPersons")) {
    props.setProperty("salesPersons", JSON.stringify(["Florence", "Annika", "Celine", "Jane", "KitKit", "Tracy"]));
  }

  SpreadsheetApp.getUi().alert("Sheet setup complete!");
}

// ── Sales Person helpers ──

function getSalesPersons() {
  var props = PropertiesService.getScriptProperties();
  var stored = props.getProperty("salesPersons");
  return stored ? JSON.parse(stored) : ["Florence", "Annika", "Celine", "Jane", "KitKit", "Tracy"];
}

function saveSalesPersons(list) {
  PropertiesService.getScriptProperties().setProperty("salesPersons", JSON.stringify(list));
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  setDropdown(sheet, "I2:I1000", list, true);
}

// ── Dialogs ──

function showAddSalesPersonDialog() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    "Add Sales Person",
    "Enter the new sales person's name:",
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  var name = result.getResponseText().trim();
  if (!name) {
    ui.alert("Name cannot be empty.");
    return;
  }

  var list = getSalesPersons();
  if (list.indexOf(name) !== -1) {
    ui.alert('"' + name + '" is already in the list.');
    return;
  }

  list.push(name);
  list.sort();
  saveSalesPersons(list);
  ui.alert('"' + name + '" has been added successfully!');
}

function showRemoveSalesPersonDialog() {
  var ui = SpreadsheetApp.getUi();
  var list = getSalesPersons();

  var result = ui.prompt(
    "Remove Sales Person",
    "Current list:\n" + list.join(", ") + "\n\nEnter the name to remove:",
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  var name = result.getResponseText().trim();
  var idx = list.indexOf(name);
  if (idx === -1) {
    ui.alert('"' + name + '" was not found in the list.');
    return;
  }

  list.splice(idx, 1);
  saveSalesPersons(list);
  ui.alert('"' + name + '" has been removed.');
}

// ── Utility ──

function setDropdown(sheet, range, options, allowInvalid) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(allowInvalid === true)
    .build();
  sheet.getRange(range).setDataValidation(rule);
}
