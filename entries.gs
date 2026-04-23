// Reading + writing redeem entries on the Sales Tracker sheet

function getDialogData() {
  var count = 0;
  forEachTodayRow_(function() { count++; });
  return { salesPersons: getSalesPersons(), nextNo: count + 1 };
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

function rowToEntry_(row) {
  var tz = getTz_();
  return {
    no:            row[COL.NO - 1],
    date:          Utilities.formatDate(new Date(row[COL.DATE - 1]), tz, "yyyy-MM-dd'T'HH:mm"),
    redeemType:    row[COL.REDEEM_TYPE - 1],
    package:       row[COL.PACKAGE - 1],
    trial:         row[COL.TRIAL - 1],
    product:       row[COL.PRODUCT - 1],
    amount:        row[COL.AMOUNT - 1] === "" ? "" : row[COL.AMOUNT - 1],
    paymentMethod: row[COL.PAYMENT_METHOD - 1],
    salesPerson:   row[COL.SALES_PERSON - 1],
    remark:        row[COL.REMARK - 1]
  };
}

// Returns { name, entries: [...] } — today's rows created by the PIN's owner.
function getMyTodayEntries(pin) {
  var name = findUserByPin(pin);
  if (!name) throw new Error("Invalid PIN");

  var out = [];
  forEachTodayRow_(function(row) {
    if (String(row[COL.CREATED_BY - 1] || "") === name) out.push(rowToEntry_(row));
  });
  return { name: name, entries: out };
}

// Finds today's row matching entry.no, updates columns C..K (leaves Date + No).
// Stamps Created By with the editor's name.
function updateEntry(entry) {
  var editor = findUserByPin(entry.pin);
  if (!editor) throw new Error("Invalid PIN");

  var targetNo = parseInt(entry.no, 10);
  var written = false;
  forEachTodayRow_(function(row, rowIdx, sheet) {
    if (parseInt(row[COL.NO - 1], 10) !== targetNo) return;
    sheet.getRange(rowIdx, COL.REDEEM_TYPE, 1, NUM_COLS - (COL.REDEEM_TYPE - 1)).setValues([[
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
    written = true;
    return true; // stop iteration
  });
  if (!written) throw new Error("Entry No. " + entry.no + " not found for today");
  return editor;
}

function addEntry(entry) {
  var createdBy = findUserByPin(entry.pin);
  if (!createdBy) throw new Error("Invalid PIN");

  var sheet = getSalesSheet_();
  sheet.appendRow([
    parseEntryDate_(entry.date),
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

  // appendRow auto-formats date cells based on locale — force our standard format.
  sheet.getRange(sheet.getLastRow(), COL.DATE).setNumberFormat("yyyy-mm-dd HH:mm");
  return createdBy;
}
