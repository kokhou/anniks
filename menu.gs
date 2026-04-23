// Custom "⚙️ Manage" menu (desktop sheet only) + its dialogs

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚙️ Manage")
    .addItem("➕ New Redeem Entry", "showRedeemDialog")
    .addSeparator()
    .addItem("Add Sales Person", "showAddSalesPersonDialog")
    .addItem("Remove Sales Person", "showRemoveSalesPersonDialog")
    .addSeparator()
    .addItem("🔑 Create / Update PIN", "showCreatePinDialog")
    .addItem("🗑️ Remove PIN", "showRemovePinDialog")
    .addItem("📋 List Users with PIN", "showListUsersDialog")
    .addSeparator()
    .addItem("💬 WhatsApp URL by No.", "showWhatsAppUrlDialog")
    .addToUi();
}

function showCreatePinDialog() {
  var ui = SpreadsheetApp.getUi();

  var r1 = ui.prompt("Create / Update PIN", "Enter user's name:", ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var name = r1.getResponseText().trim();
  if (!name) { ui.alert("Name cannot be empty."); return; }

  var r2 = ui.prompt("Create / Update PIN", 'Enter 4-digit PIN for "' + name + '":', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var pin = r2.getResponseText().trim();
  if (!/^\d{4}$/.test(pin)) { ui.alert("PIN must be exactly 4 digits."); return; }

  var r3 = ui.prompt("Create / Update PIN", "Confirm PIN:", ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;
  if (r3.getResponseText().trim() !== pin) { ui.alert("PINs do not match."); return; }

  try {
    var action = upsertUserPin(name, pin);
    ui.alert('"' + name + '" PIN ' + action + ' successfully.');
  } catch (e) {
    ui.alert("Error: " + e.message);
  }
}

function showRemovePinDialog() {
  var ui = SpreadsheetApp.getUi();
  var list = listUsersWithPin();
  if (!list.length) { ui.alert("No users with PIN yet."); return; }

  var r = ui.prompt("Remove PIN",
    "Current users:\n" + list.join(", ") + "\n\nEnter the name to remove:",
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  var name = r.getResponseText().trim();
  if (removeUserPin(name)) {
    ui.alert('"' + name + '" PIN removed.');
  } else {
    ui.alert('"' + name + '" was not found.');
  }
}

function showListUsersDialog() {
  var ui = SpreadsheetApp.getUi();
  var list = listUsersWithPin();
  ui.alert("Users with PIN", list.length ? list.join("\n") : "(none)", ui.ButtonSet.OK);
}

function showWhatsAppUrlDialog() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt("WhatsApp URL by No.",
    "Enter the entry No. (for today):",
    ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  var no = parseInt(r.getResponseText().trim(), 10);
  if (!no) { ui.alert("Invalid No."); return; }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sales Tracker") || ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var tz = Session.getScriptTimeZone();
  var today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  var found = null;
  for (var i = data.length - 1; i >= 1; i--) {
    var rowDate = data[i][0] ? Utilities.formatDate(new Date(data[i][0]), tz, "yyyy-MM-dd") : "";
    if (rowDate === today && parseInt(data[i][1], 10) === no) { found = data[i]; break; }
  }
  if (!found) { ui.alert("No entry found for today with No. " + no); return; }

  var entry = {
    salesPerson:   found[8],
    redeemType:    found[2],
    package:       found[3],
    product:       found[5],
    amount:        found[6],
    paymentMethod: found[7],
    remark:        found[9]
  };
  var text = buildWhatsAppText_(entry);
  var url  = 'https://wa.me/?text=' + encodeURIComponent(text);
  var escapedText = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  var html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;padding:16px;">' +
      '<p style="margin:0 0 12px;font-size:13px;color:#555;">Entry No. ' + no + '</p>' +
      '<a href="' + url + '" target="_blank" ' +
         'style="display:inline-block;padding:12px 20px;background:#25d366;color:#fff;' +
         'text-decoration:none;border-radius:8px;font-weight:600;">Open WhatsApp</a>' +

      '<p style="margin:16px 0 4px;font-size:11px;color:#888;">Message text (tap to select):</p>' +
      '<textarea readonly onclick="this.select()" ' +
         'style="width:100%;height:120px;font-size:12px;padding:8px;border:1px solid #ddd;' +
         'border-radius:4px;box-sizing:border-box;font-family:monospace;">' + escapedText + '</textarea>' +

      '<p style="margin:12px 0 4px;font-size:11px;color:#888;">URL (tap to select):</p>' +
      '<textarea readonly onclick="this.select()" ' +
         'style="width:100%;height:70px;font-size:11px;padding:8px;border:1px solid #ddd;' +
         'border-radius:4px;box-sizing:border-box;">' + url + '</textarea>' +
    '</div>'
  ).setWidth(440).setHeight(400);
  ui.showModalDialog(html, "WhatsApp Share");
}

function buildWhatsAppText_(entry) {
  var lines = ['*Sales ' + (entry.salesPerson || '') + '*'];
  lines.push('Package: '  + (entry.package       || ''));
  lines.push('Product: '  + (entry.product       || ''));
  lines.push('Amount: '   + (entry.amount ? 'RM ' + entry.amount : ''));
  lines.push('Pay: '      + (entry.paymentMethod || ''));
  lines.push('Redeem: '   + (entry.redeemType    || ''));
  if (entry.remark) {
    lines.push('');
    lines.push('_' + entry.remark + '_');
  }
  return lines.join('\n');
}

function showRedeemDialog() {
  var html = HtmlService.createHtmlOutputFromFile("index")
    .setWidth(420)
    .setHeight(580);
  SpreadsheetApp.getUi().showModalDialog(html, "New Redeem Entry");
}

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
