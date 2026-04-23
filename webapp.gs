// Web App JSON API — called from GitHub Pages frontend via fetch()

function doGet(e) {
  return jsonResponse_(getDialogData());
}

function doPost(e) {
  try {
    var entry = JSON.parse(e.postData.contents);
    addEntry(entry);
    return jsonResponse_({ ok: true });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
