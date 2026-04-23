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
