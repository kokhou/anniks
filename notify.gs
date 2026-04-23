// WhatsApp notification via CallMeBot
// Setup:
//   1. Add @ApiWhatsBot as a contact on WhatsApp
//   2. Add that contact to your group
//   3. In the group, send: "I allow callmebot to send me messages"
//   4. The bot replies with your API key + the group's phone identifier
//   5. Use ⚙️ Manage → Configure WhatsApp Notify to save them

var CMB_PHONE_KEY  = "cmb_phone";
var CMB_APIKEY_KEY = "cmb_apikey";

function notifyWhatsApp_(entry, createdBy) {
  var props = PropertiesService.getScriptProperties();
  var phone  = props.getProperty(CMB_PHONE_KEY);
  var apikey = props.getProperty(CMB_APIKEY_KEY);
  if (!phone || !apikey) return; // not configured — skip silently

  var lines = [
    "🧾 New Redeem Entry",
    "By: " + createdBy,
    "Sales: " + (entry.salesPerson || "-"),
    "Type: " + (entry.redeemType || "-")
  ];
  if (entry.package)       lines.push("Package: " + entry.package);
  if (entry.product)       lines.push("Product: " + entry.product);
  if (entry.amount)        lines.push("Amount: RM " + entry.amount);
  if (entry.paymentMethod) lines.push("Pay: " + entry.paymentMethod);
  if (entry.remark)        lines.push("Note: " + entry.remark);

  var url = "https://api.callmebot.com/whatsapp.php"
    + "?phone="  + encodeURIComponent(phone)
    + "&apikey=" + encodeURIComponent(apikey)
    + "&text="   + encodeURIComponent(lines.join("\n"));

  try {
    UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (e) {
    // fail silently so submission isn't blocked
    console.log("WhatsApp notify failed: " + e.message);
  }
}

function showConfigureWhatsAppDialog() {
  var ui = SpreadsheetApp.getUi();
  var props = PropertiesService.getScriptProperties();

  var r1 = ui.prompt("WhatsApp Notify Setup",
    "Enter CallMeBot phone/group identifier (e.g. +60123456789):\n\n" +
    "Current: " + (props.getProperty(CMB_PHONE_KEY) || "(not set)"),
    ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  var phone = r1.getResponseText().trim();
  if (!phone) { ui.alert("Phone cannot be empty."); return; }

  var r2 = ui.prompt("WhatsApp Notify Setup",
    "Enter CallMeBot API key:\n\n" +
    "Current: " + (props.getProperty(CMB_APIKEY_KEY) ? "(set)" : "(not set)"),
    ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  var apikey = r2.getResponseText().trim();
  if (!apikey) { ui.alert("API key cannot be empty."); return; }

  props.setProperty(CMB_PHONE_KEY, phone);
  props.setProperty(CMB_APIKEY_KEY, apikey);
  ui.alert("Saved. Test it by submitting an entry.");
}

function showDisableWhatsAppDialog() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.alert("Disable WhatsApp Notify?",
    "Clear the stored phone and API key?",
    ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(CMB_PHONE_KEY);
  props.deleteProperty(CMB_APIKEY_KEY);
  ui.alert("WhatsApp notifications disabled.");
}
