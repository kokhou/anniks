// One-time sheet setup — run setupSheet() from Apps Script editor after schema changes

function setupSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setName("Sales Tracker");

  // ── Headers ──
  var headers = ["Date", "No", "Redeem Type", "Package", "Trial", "Product", "Amount", "Payment Method", "Sales Person", "Remark", "Created By"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#4a86e8");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");

  // ── Dropdowns (rows 2–1000) · allowInvalid so custom entries from web app don't error ──
  setDropdown(sheet, "C2:C1000", ["New", "Existing"], true);

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
  ], true);

  setDropdown(sheet, "E2:E1000", ["Yes", "No"], true);

  setDropdown(sheet, "F2:F1000", [
    "T388 脸部塑型",
    "T388 祈龄魔法",
    "T298 体态",
    "Firming Cream"
  ], true);

  setDropdown(sheet, "H2:H1000", ["Cash", "Debit Card", "Credit Card", "Online Transfer", "QR Pay"], true);

  setDropdown(sheet, "I2:I1000", getSalesPersons(), true);

  // ── Column widths ──
  sheet.setColumnWidth(1, 150);
  sheet.getRange("A2:A1000").setNumberFormat("yyyy-mm-dd HH:mm");
  sheet.setColumnWidth(2, 50);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 60);
  sheet.setColumnWidth(6, 160);
  sheet.setColumnWidth(7, 90);
  sheet.setColumnWidth(8, 130);
  sheet.setColumnWidth(9, 110);
  sheet.setColumnWidth(10, 180);
  sheet.setColumnWidth(11, 110);

  sheet.setFrozenRows(1);

  // Provision hidden Users sheet + encryption key
  getUsersSheet_();
  getEncryptionKey_();

  // Seed default sales persons if not set
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty("salesPersons")) {
    props.setProperty("salesPersons", JSON.stringify(["Florence", "Annika", "Celine", "Jane", "KitKit", "Tracy"]));
  }

  SpreadsheetApp.getUi().alert("Sheet setup complete!");
}

function setDropdown(sheet, range, options, allowInvalid) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(allowInvalid === true)
    .build();
  sheet.getRange(range).setDataValidation(rule);
}
