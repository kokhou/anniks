// Web App JSON API — called from GitHub Pages frontend via fetch()

function doGet(e) {
  return jsonResponse_(getDialogData());
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    if (payload.action === "listMine") {
      var res = getMyTodayEntries(payload.pin);
      return jsonResponse_({ ok: true, name: res.name, entries: res.entries });
    }
    if (payload.action === "update") {
      var editor = updateEntry(payload);
      return jsonResponse_({ ok: true, createdBy: editor });
    }
    // default: add new entry
    var createdBy = addEntry(payload);
    return jsonResponse_({ ok: true, createdBy: createdBy });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err.message || err) });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
