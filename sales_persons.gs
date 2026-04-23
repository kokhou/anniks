// Sales person list — stored in Script Properties, mirrored to column I dropdown

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
