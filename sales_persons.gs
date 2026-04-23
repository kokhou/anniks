// Sales person list — stored in Script Properties, mirrored to column I dropdown

var DEFAULT_SALES_PERSONS = ["Florence", "Annika", "Celine", "Jane", "KitKit", "Tracy"];
var SALES_PERSONS_KEY = "salesPersons";

function getSalesPersons() {
  var stored = PropertiesService.getScriptProperties().getProperty(SALES_PERSONS_KEY);
  return stored ? JSON.parse(stored) : DEFAULT_SALES_PERSONS.slice();
}

function saveSalesPersons(list) {
  PropertiesService.getScriptProperties().setProperty(SALES_PERSONS_KEY, JSON.stringify(list));
  // Always target the Sales Tracker sheet — not whatever tab happens to be active.
  setDropdown(getSalesSheet_(), "I2:I1000", list, true);
}
