// Teacher Portal backend. Paste into Extensions > Apps Script of a Google Sheet.
function sh_(n, head) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), s = ss.getSheetByName(n);
  if (!s) { s = ss.insertSheet(n); s.appendRow(head); }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() { return out_({ ok: true, msg: 'Portal endpoint is running' }); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var p = JSON.parse((e.postData && e.postData.contents) || '{}');
    var pin = String(PropertiesService.getScriptProperties().getProperty('PIN') || '');
    var authed = pin !== '' && String(p.pin || '') === pin;
    var R = sh_('Responses', ['id', 'createdAt', 'name', 'json']);
    var V = sh_('Reviews', ['id', 'status', 'action', 'at']);

    if (p.action === 'submit' || p.action === 'add') {
      if (p.action === 'add' && !authed) return out_({ ok: false, error: 'pin' });
      var d = p.doc;
      if (!d || !d.id || !d.a || !d.a.name) return out_({ ok: false, error: 'bad data' });
      var j = JSON.stringify(d);
      if (j.length > 49000) return out_({ ok: false, error: 'too long' });
      var n = R.getLastRow();
      var ids = n > 1 ? R.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return String(r[0]); }) : [];
      if (ids.indexOf(String(d.id)) < 0) {
        var safeName = String(d.a.name).replace(/^[=+\-@]/, "'$&");
        R.appendRow([String(d.id), d.createdAt || Date.now(), safeName, j]);
      }
      return out_({ ok: true });
    }

    if (!authed) return out_({ ok: false, error: 'pin' });

    if (p.action === 'list') {
      var rows = R.getLastRow() > 1 ? R.getRange(2, 1, R.getLastRow() - 1, 4).getValues() : [];
      var entries = [];
      rows.forEach(function (r) { try { entries.push(JSON.parse(r[3])); } catch (x) {} });
      var vr = V.getLastRow() > 1 ? V.getRange(2, 1, V.getLastRow() - 1, 3).getValues() : [];
      var reviews = {};
      vr.forEach(function (r) { reviews[String(r[0])] = { status: r[1], action: r[2] }; });
      return out_({ ok: true, entries: entries, reviews: reviews });
    }

    if (p.action === 'review') {
      var id = String(p.id || '');
      if (!id) return out_({ ok: false, error: 'bad data' });
      var m = V.getLastRow() > 1 ? V.getRange(2, 1, V.getLastRow() - 1, 1).getValues() : [];
      var row = -1;
      for (var i = 0; i < m.length; i++) if (String(m[i][0]) === id) { row = i + 2; break; }
      var vals = [id, String(p.status || 'New'), String(p.note || '').slice(0, 500), Date.now()];
      if (row < 0) V.appendRow(vals); else V.getRange(row, 1, 1, 4).setValues([vals]);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}
