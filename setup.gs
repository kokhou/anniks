// One-time sheet setup — run setupSheet() from the Apps Script editor after schema changes.

function setupSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setName(SHEET_NAME);

  // ── Headers ──
  var headers = [
    "Date", "No", "Redeem Type", "Package", "Trial", "Product",
    "Amount", "Payment Method", "Sales Person", "Remark", "Created By"
  ];
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground("#4a86e8")
    .setFontColor("#ffffff")
    .setFontWeight("bold");

  // ── Dropdowns (rows 2–1000). allowInvalid so custom web-app entries don't error ──
  setDropdown(sheet, "C2:C1000", ["New", "Existing"], true);
  setDropdown(sheet, "D2:D1000", [
    "P6880 脸部塑型", "P6880 开肩", "P6880 体态", "P6880 祈龄", "P6880 局部",
    "P4880 高级波肽", "Gold 脸部塑型", "Gold 开肩", "T2388 小腿"
  ], true);
  setDropdown(sheet, "E2:E1000", ["Yes", "No"], true);
  setDropdown(sheet, "F2:F1000", [
    "T388 脸部塑型", "T388 祈龄魔法", "T298 体态", "Firming Cream"
  ], true);
  setDropdown(sheet, "H2:H1000", [
    "Cash", "Debit Card", "Credit Card", "Online Transfer", "QR Pay"
  ], true);
  setDropdown(sheet, "I2:I1000", getSalesPersons(), true);

  // ── Column widths ──
  var widths = [150, 50, 110, 160, 60, 160, 90, 130, 110, 180, 110];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });
  sheet.getRange("A2:A1000").setNumberFormat("yyyy-mm-dd HH:mm");
  sheet.setFrozenRows(1);

  // Provision hidden Users sheet + encryption key on first run
  getUsersSheet_();
  getEncryptionKey_();

  // Seed default sales persons on first run
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty(SALES_PERSONS_KEY)) {
    props.setProperty(SALES_PERSONS_KEY, JSON.stringify(DEFAULT_SALES_PERSONS));
  }

  SpreadsheetApp.getUi().alert("Sheet setup complete!");
}
