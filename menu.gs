// Custom "⚙️ Manage" menu (desktop sheet only) + its dialogs

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚙️ Manage")
    .addItem("➕ New Redeem Entry", "showRedeemDialog")
    .addSeparator()
    .addItem("Add Sales Person", "showAddSalesPersonDialog")
    .addItem("Remove Sales Person", "showRemoveSalesPersonDialog")
    .addToUi();
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
