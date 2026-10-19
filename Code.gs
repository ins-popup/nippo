/**
 * T8 日報・売上・在庫管理：保存用の Google Apps Script
 * このスクリプトを「保存先にするスプレッドシート」の Apps Script に貼り付けて、ウェブアプリとして公開します。
 * スクリプトプロパティに PASSCODE（スタッフ用の合言葉）と ADMIN_CODE（管理者コード）を設定してください。
 */
var SHEET = 'docs';
var ADMIN_ONLY = { skus: true, config: true };
var LOG_LIMIT = 300;

function doGet() { return ContentService.createTextOutput('T8 backend OK'); }

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  var props = PropertiesService.getScriptProperties();
  var pass = props.getProperty('PASSCODE'), adminCode = props.getProperty('ADMIN_CODE');
  if (!pass || !adminCode) return out_({ ok: false, error: 'not_configured' });
  var isAdmin = !!req.admin && req.admin === adminCode;
  if (!isAdmin && !(req.code && (req.code === pass || req.code === adminCode))) return out_({ ok: false, error: 'unauthorized' });
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    switch (req.action) {
      case 'list': return out_({ ok: true, data: list_() });
      case 'write': return out_(write_(req.ops || [], isAdmin));
      case 'seed': if (!isAdmin) return out_({ ok: false, error: 'admin_required' }); return out_(seed_(req.payload || {}));
      case 'upload': return out_(upload_(req.b64, req.mime));
      case 'trash': return out_(trash_(req.id));
      default: return out_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return out_({ ok: false, error: 'server:' + err });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(['key', 'collection', 'id', 'json', 'updated']); sh.setFrozenRows(1); }
  return sh;
}

function list_() {
  var sh = sheet_(), n = sh.getLastRow(), res = {};
  if (n < 2) return res;
  var rows = sh.getRange(2, 1, n - 1, 5).getValues(), logs = [];
  rows.forEach(function (r) {
    var c = r[1], id = r[2], d;
    try { d = JSON.parse(r[3]); } catch (e) { return; }
    if (c === 'log') { logs.push([id, d]); return; }
    (res[c] = res[c] || {})[id] = d;
  });
  logs.sort(function (a, b) { return (b[1].at || 0) - (a[1].at || 0); });
  res.log = {};
  logs.slice(0, LOG_LIMIT).forEach(function (x) { res.log[x[0]] = x[1]; });
  return res;
}

function index_(sh) {
  var n = sh.getLastRow(), m = {};
  if (n < 2) return m;
  sh.getRange(2, 1, n - 1, 1).getValues().forEach(function (r, i) { m[r[0]] = i + 2; });
  return m;
}

function write_(ops, isAdmin) {
  var sh = sheet_(), now = new Date().toISOString();
  ops.forEach(function (o) {
    if (!o || !o.collection || !o.id) throw 'bad_op';
    if (ADMIN_ONLY[o.collection] && !isAdmin) throw 'admin_required';
    var key = o.collection + '/' + o.id, idx = index_(sh), row = idx[key];
    if (o.op === 'delete') { if (row) sh.deleteRow(row); return; }
    var data = o.data || {};
    if (o.op === 'update' && row) {
      var cur = {};
      try { cur = JSON.parse(sh.getRange(row, 4).getValue()); } catch (e) {}
      Object.keys(data).forEach(function (k) { cur[k] = data[k]; });
      data = cur;
    }
    var vals = [[key, o.collection, o.id, JSON.stringify(data), now]];
    if (row) sh.getRange(row, 1, 1, 5).setValues(vals); else sh.getRange(sh.getLastRow() + 1, 1, 1, 5).setValues(vals);
  });
  return { ok: true };
}

function seed_(p) {
  var sh = sheet_(), now = new Date().toISOString(), idx = index_(sh), add = [];
  function put(c, id, d) {
    var key = c + '/' + id, row = idx[key], vals = [key, c, id, JSON.stringify(d), now];
    if (row) sh.getRange(row, 1, 1, 5).setValues([vals]); else add.push(vals);
  }
  if (p.config) put('config', 'main', p.config);
  (p.skus || []).forEach(function (s) { var id = s.id; var d = {}; Object.keys(s).forEach(function (k) { if (k !== 'id') d[k] = s[k]; }); put('skus', id, d); });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 5).setValues(add);
  return { ok: true, added: add.length };
}

function folder_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('PHOTO_FOLDER_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  var f = DriveApp.createFolder('T8日報_売上レポート写真');
  props.setProperty('PHOTO_FOLDER_ID', f.getId());
  return f;
}

function upload_(b64, mime) {
  if (!b64) throw 'no_data';
  var blob = Utilities.newBlob(Utilities.base64Decode(b64), mime || 'image/jpeg', 't8_' + Date.now() + '.jpg');
  var file = folder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, id: file.getId() };
}

function trash_(id) {
  if (!id) throw 'no_id';
  var f = DriveApp.getFileById(id), inFolder = false, it = f.getParents();
  var fid = folder_().getId();
  while (it.hasNext()) { if (it.next().getId() === fid) inFolder = true; }
  if (!inFolder) throw 'not_allowed';
  f.setTrashed(true);
  return { ok: true };
}
