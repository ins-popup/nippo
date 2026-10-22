/**
 * T8 日報・売上・在庫管理：保存用の Google Apps Script
 * このスクリプトを「保存先にするスプレッドシート」の Apps Script に貼り付けて、ウェブアプリとして公開します。
 * スクリプトプロパティに PASSCODE（社内スタッフ用の合言葉）と ADMIN_CODE（管理者コード）を設定してください。
 *
 * 入室できる範囲（誰が何を見られるか）
 *   ADMIN_CODE      … すべてのイベントを見られる＋商品マスタ・イベント設定を変更できる（管理者）
 *   PASSCODE        … すべてのイベントを見られる（社内スタッフ用。イベント切替ができる）
 *   イベント別の合言葉 … そのイベントだけ見られる（社外メンバー用）。管理者がサイトの
 *                      「管理者用：イベントの設定・追加」から設定します（保存先：スクリプトプロパティ EVENT_CODES）。
 *                      他のイベントのデータはサーバーが渡さないので、画面を細工しても見られません。
 */
var SHEET = 'docs';
var ADMIN_ONLY = { skus: true, config: true, imports: true, sqseen: true };
var LOG_LIMIT = 300;

function doGet() { return ContentService.createTextOutput('T8 backend OK'); }

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_request' }); }
  var props = PropertiesService.getScriptProperties();
  var pass = props.getProperty('PASSCODE'), adminCode = props.getProperty('ADMIN_CODE');
  if (!pass || !adminCode) return out_({ ok: false, error: 'not_configured' });
  var isAdmin = !!req.admin && req.admin === adminCode;
  var scope = { role: 'all', event: '' };
  if (isAdmin || (req.code && req.code === adminCode)) { scope = { role: 'admin', event: '' }; }
  else if (req.code && req.code === pass) { scope = { role: 'all', event: '' }; }
  else {
    var ev = req.code ? eventOfCode_(props, req.code) : '';
    if (!ev) return out_({ ok: false, error: 'unauthorized' });
    scope = { role: 'event', event: ev };
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    switch (req.action) {
      case 'list': return out_({ ok: true, data: list_(scope), scope: scope });
      case 'write': return out_(write_(req.ops || [], isAdmin, scope));
      case 'getcodes': if (!isAdmin) return out_({ ok: false, error: 'admin_required' }); return out_({ ok: true, codes: codes_(props) });
      case 'setcode': if (!isAdmin) return out_({ ok: false, error: 'admin_required' }); return out_(setCode_(props, req.event, req.newCode, pass, adminCode));
      case 'squarenow': if (!isAdmin) return out_({ ok: false, error: 'admin_required' }); return out_(squareScanNow_());
      case 'squareassign': if (!isAdmin) return out_({ ok: false, error: 'admin_required' }); return out_(squareAssign_(req.id, req.event));
      case 'squaredismiss': if (!isAdmin) return out_({ ok: false, error: 'admin_required' }); return out_(squareDismiss_(req.id));
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

/* イベント別の合言葉（スクリプトプロパティ EVENT_CODES に {イベントID: 合言葉} で保存） */
function codes_(props) {
  try { return JSON.parse(props.getProperty('EVENT_CODES') || '{}') || {}; } catch (e) { return {}; }
}
function eventOfCode_(props, code) {
  var m = codes_(props);
  for (var k in m) { if (m[k] && m[k] === code) return k; }
  return '';
}
function setCode_(props, ev, code, pass, adminCode) {
  if (!ev || !/^[A-Za-z0-9_-]{1,40}$/.test(ev)) throw 'bad_event';
  var m = codes_(props);
  code = String(code || '').trim();
  if (!code) { delete m[ev]; }
  else {
    if (code.length < 4) throw 'code_too_short';
    if (code === pass || code === adminCode) throw 'code_in_use';
    for (var k in m) { if (k !== ev && m[k] === code) throw 'code_in_use'; }
    m[ev] = code;
  }
  props.setProperty('EVENT_CODES', JSON.stringify(m));
  return { ok: true, set: !!code };
}
/* ドキュメントIDがどのイベントのものか：先頭が「ev~ID~」ならそのイベント、「~」を含まなければ最初のイベント（t8） */
function evOfId_(id) {
  id = String(id);
  if (id.indexOf('~') < 0) return 't8';
  var m = /^ev~([^~]+)~/.exec(id);
  return m ? m[1] : null;
}
function inScope_(scope, collection, id) {
  if (scope.role !== 'event') return true;
  if (collection === 'config' && id === 'events') return false;
  return evOfId_(id) === scope.event;
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function sheet_() {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(['key', 'collection', 'id', 'json', 'updated']); sh.setFrozenRows(1); }
  return sh;
}

function list_(scope) {
  scope = scope || { role: 'all' };
  var sh = sheet_(), n = sh.getLastRow(), res = {};
  if (n < 2) return res;
  var rows = sh.getRange(2, 1, n - 1, 5).getValues(), logs = [];
  rows.forEach(function (r) {
    var c = r[1], id = r[2], d;
    try { d = JSON.parse(r[3]); } catch (e) { return; }
    if (c === 'sqseen') return;
    if (scope.role === 'event') {
      if (c === 'config' && id === 'events') {
        /* 自分のイベントの1件だけを見せる（他のイベント名は渡さない） */
        var only = (d.list || []).filter(function (x) { return x.id === scope.event; });
        (res[c] = res[c] || {})[id] = { list: only };
        return;
      }
      if (!inScope_(scope, c, id)) return;
    }
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

function write_(ops, isAdmin, scope) {
  var sh = sheet_(), now = new Date().toISOString();
  ops.forEach(function (o) {
    if (!o || !o.collection || !o.id) throw 'bad_op';
    if (ADMIN_ONLY[o.collection] && !isAdmin) throw 'admin_required';
    if (scope && scope.role === 'event' && !inScope_(scope, o.collection, o.id)) throw 'not_allowed';
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

/* =====================================================================
 * Square 売上レポート（メール）の自動取り込み
 *  - Gmail に届いた「Square売上レポート」を定期的に読み取り、イベントごとの売上として登録します。
 *  - どのイベントか：①イベント設定の「Squareの拠点名」→②会期→③商品の一致、の順で自動判定します。
 *    決められないものは「未振り分け」（imports の pending~…）に溜めて、管理者が画面で選びます。
 *  - 同じメール・同じ日の再取り込みでも二重になりません（IDが日付＋拠点＋商品で決まるため）。
 *  - 初回だけ setupSquareTrigger を手動で1回実行（Gmailの許可が必要）。以後は30分ごとに自動で動きます。
 * ===================================================================== */
/* Square からは自動で届かないため、担当者が転送・送信したものも読む（送信元は MAIL_ALLOW で許可したものだけ反映） */
var SQ_QUERY = 'subject:(Square売上レポート) newer_than:14d';

function setupSquareTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'squareScan') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('squareScan').timeBased().everyMinutes(30).create();
  return squareScanNow_();
}

/* 時間トリガーから呼ばれる */
function squareScan() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return squareScanCore_(); } finally { try { lock.releaseLock(); } catch (x) {} }
}
/* 画面の「今すぐ取り込む」（doPost 内でロック済み） */
function squareScanNow_() { return squareScanCore_(); }

function sqDecode_(t) {
  return String(t).replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;|&apos;/g, "'")
    .replace(/&#(\d+);/g, function (m, n) { return String.fromCharCode(parseInt(n, 10)); })
    .replace(/&#x([0-9a-f]+);/gi, function (m, n) { return String.fromCharCode(parseInt(n, 16)); });
}
function sqText_(h) { return sqDecode_(String(h).replace(/<!--[\s\S]*?-->/g, ' ').replace(/<[^>]*>/g, ' ')).replace(/\s+/g, ' ').trim(); }
function sqYen_(t) {
  var m = /(\()?\s*[¥￥]\s*([\d,]+)/.exec(String(t));
  if (!m) return null;
  var n = parseInt(m[2].replace(/,/g, ''), 10);
  return m[1] ? -n : n;
}
function sqKey_(t) { return String(t == null ? '' : t).normalize('NFKC').toUpperCase().replace(/[^0-9A-Z\u3040-\u30ff\u4e00-\u9fff]/g, ''); }
function sqPad_(n) { return (n < 10 ? '0' : '') + n; }
function sqRow_(html, label) {
  var re = new RegExp('<td[^>]*>\\s*' + label + '\\s*</td>\\s*<td[^>]*>([\\s\\S]*?)</td>');
  var m = re.exec(html);
  return m ? sqText_(m[1]) : null;
}

/* Square の売上レポート（HTMLメール）を読み取る。形が違うときは ok:false */
function parseSquareReport_(html) {
  var r = { ok: false, lines: [], totals: {} };
  html = String(html || '');
  var h1 = /class="h1"[^>]*>([\s\S]*?)<\/td>/.exec(html);
  if (!h1 || sqText_(h1[1]).indexOf('売上レポート') < 0) { r.reason = '売上レポートの形式ではありません'; return r; }
  var after = html.slice(h1.index + h1[0].length);
  var loc = /class="p"[^>]*>([\s\S]*?)<\/td>/.exec(after);
  r.loc = loc ? sqText_(loc[1]) : '';
  var dr = /id="date-range"[^>]*>([\s\S]*?)<\/td>/.exec(html), dtxt = dr ? sqText_(dr[1]) : '';
  var ds = [], re = /(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日/g, m;
  while ((m = re.exec(dtxt))) { var d = m[1] + '-' + sqPad_(+m[2]) + '-' + sqPad_(+m[3]); if (ds.indexOf(d) < 0) ds.push(d); }
  if (!ds.length) { r.reason = '日付を読み取れませんでした'; return r; }
  r.date = ds[0]; r.multi = ds.length > 1;
  var t = r.totals;
  t.customers = parseInt(String(sqRow_(html, '客数') || '').replace(/[^\d-]/g, ''), 10);
  if (isNaN(t.customers)) t.customers = null;
  t.gross = sqYen_(sqRow_(html, '総売上高')); t.ret = sqYen_(sqRow_(html, '返品')); t.disc = sqYen_(sqRow_(html, 'ディスカウントと無料提供'));
  t.net = sqYen_(sqRow_(html, '純売上高')); t.tax = sqYen_(sqRow_(html, '税金')); t.total = sqYen_(sqRow_(html, '合計'));
  var i = html.search(/>\s*商品別売上\s*</), items = [];
  if (i >= 0) {
    var seg = html.slice(i), cre = /<td valign="top" width="300"[^>]*>([\s\S]*?)<\/td>\s*<td valign="top" width="80"[^>]*>([\s\S]*?)<\/td>/g, c;
    while ((c = cre.exec(seg))) {
      var before = seg.slice(0, c.index), k = before.lastIndexOf('<td style="padding: 0 2px');
      var variant = k >= 0 && before.slice(k, k + 60).indexOf('0 12px') >= 0;
      var qm = /×(?:&nbsp;|\s)*(-?[\d,]+)/.exec(c[1]);
      var name = sqText_(String(c[1]).replace(/<span[\s\S]*?<\/span>/g, ' '));
      var row = { name: name, qty: qm ? parseInt(qm[1].replace(/,/g, ''), 10) : 1, amount: sqYen_(c[2]) };
      if (variant && items.length) items[items.length - 1].variants.push(row); else items.push({ name: name, qty: row.qty, amount: row.amount, variants: [] });
    }
  }
  items.forEach(function (it) {
    if (it.variants.length) it.variants.forEach(function (v) { r.lines.push({ name: it.name, variant: v.name, qty: v.qty, amount: v.amount }); });
    else r.lines.push({ name: it.name, variant: '', qty: it.qty, amount: it.amount });
  });
  r.ok = true;
  return r;
}

/* ---- 商品マスタとの突き合わせ ---- */
function sqSkuIndex_(docs, P) {
  var ix = { full: {}, byName: {}, n: 0 };
  Object.keys(docs).forEach(function (k) {
    if (k.indexOf('skus/') !== 0) return;
    var raw = k.slice(5), d = docs[k];
    if (P === '' ? raw.indexOf('~') >= 0 : raw.indexOf(P) !== 0) return;
    var id = P === '' ? raw : raw.slice(P.length), co = sqKey_(d.color), sz = sqKey_(d.size);
    ix.n++;
    [sqKey_(d.name), sqKey_(d.hin), sqKey_(String(d.hin || '') + String(d.name || ''))].forEach(function (nm) {
      if (!nm) return;
      ix.full[nm + '|' + co + '|' + sz] = id; ix.full[nm + '|' + sz + '|' + co] = id;
      (ix.byName[nm] = ix.byName[nm] || []).push({ id: id, co: co, sz: sz });
    });
  });
  return ix;
}
function sqUniq_(a) { var u = {}, o = []; a.forEach(function (x) { if (!u[x.id]) { u[x.id] = 1; o.push(x); } }); return o; }
function sqMatch_(ix, line) {
  var nm = sqKey_(line.name), v = String(line.variant || '').split('/').map(function (x) { return x.trim(); }).filter(Boolean), cand, id;
  if (v.length >= 2) {
    var last = sqKey_(v[v.length - 1]), rest = sqKey_(v.slice(0, v.length - 1).join('/')), first = sqKey_(v[0]), rest2 = sqKey_(v.slice(1).join('/'));
    id = ix.full[nm + '|' + rest + '|' + last] || ix.full[nm + '|' + first + '|' + rest2];
    if (id) return id;
  }
  cand = ix.byName[nm] || [];
  if (v.length === 1) {
    var t = sqKey_(v[0]), hit = cand.filter(function (c) { return c.co === t || c.sz === t; });
    return sqUniq_(hit).length === 1 ? hit[0].id : null;
  }
  if (!v.length) return sqUniq_(cand).length === 1 ? cand[0].id : null;
  return null;
}

/* ---- 全データの読み込み（キー→データ）と書き込み ---- */
function sqDocs_() {
  var sh = sheet_(), n = sh.getLastRow(), m = {};
  if (n < 2) return m;
  sh.getRange(2, 1, n - 1, 4).getValues().forEach(function (r) {
    if (!/^(skus|config|days|imports|sqseen)$/.test(r[1])) return;
    try { m[r[0]] = JSON.parse(r[3]); } catch (e) {}
  });
  return m;
}
/* ops: {op:'set'|'delete', collection, id, data} をまとめて書く（行の検索は1回だけ） */
function sqWrite_(ops) {
  var sh = sheet_(), now = new Date().toISOString(), idx = index_(sh), del = [], add = [];
  ops.forEach(function (o) { if (o.op === 'delete' && idx[o.collection + '/' + o.id]) del.push(idx[o.collection + '/' + o.id]); });
  del.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  if (del.length) idx = index_(sh);
  ops.forEach(function (o) {
    if (o.op !== 'set') return;
    var key = o.collection + '/' + o.id, vals = [key, o.collection, o.id, JSON.stringify(o.data), now];
    if (idx[key]) sh.getRange(idx[key], 1, 1, 5).setValues([vals]); else add.push(vals);
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, 5).setValues(add);
}

function sqEvents_(docs) {
  var reg = docs['config/events'], list = reg && reg.list && reg.list.length ? reg.list : [{ id: 't8' }];
  return list.map(function (e) {
    var P = e.id === 't8' ? '' : 'ev~' + e.id + '~', cfg = docs['config/' + P + 'main'] || {};
    var names = String(cfg.squareNames || '').split(/[,、\n]/).map(sqKey_).filter(Boolean);
    return { id: e.id, name: cfg.name || e.name || e.id, P: P, cfg: cfg, names: names, start: cfg.start || e.start || '', end: cfg.end || e.end || '' };
  });
}
function sqInRange_(e, d) { return !!(e.start && e.end && d >= e.start && d <= e.end); }

function sqRoute_(rep, events, docs) {
  var loc = sqKey_(rep.loc), byName = events.filter(function (e) { return loc && e.names.indexOf(loc) >= 0; }), pool = byName.slice(), how = '拠点名';
  if (!pool.length) { pool = events.filter(function (e) { return sqInRange_(e, rep.date); }); how = '会期'; }
  if (!pool.length) return { reason: 'この日付が会期に入っているイベントがありません', cands: [] };
  if (byName.length) {
    var dated = pool.filter(function (e) { return sqInRange_(e, rep.date); });
    if (!dated.length) return { reason: '「' + pool[0].name + '」の会期（' + pool[0].start + '〜' + pool[0].end + '）に入っていない日付です', cands: pool.map(function (e) { return e.id; }) };
    pool = dated;
  }
  var scored = pool.map(function (e) {
    var ix = sqSkuIndex_(docs, e.P), n = 0;
    rep.lines.forEach(function (l) { if (l.qty > 0 && sqMatch_(ix, l)) n++; });
    return { e: e, n: n };
  });
  var best = scored.slice().sort(function (a, b) { return b.n - a.n; });
  if (best.length === 1) {
    if (rep.lines.length && best[0].n === 0) return { reason: '「' + best[0].e.name + '」の商品マスタに、このレポートの商品が見つかりません', cands: [best[0].e.id] };
    return { event: best[0].e, how: how };
  }
  if (rep.lines.length && best[0].n > 0 && best[0].n > best[1].n) return { event: best[0].e, how: how + '＋商品' };
  return { reason: '複数のイベントに当てはまります（拠点名をイベント設定に入れると自動で決まります）', cands: best.map(function (x) { return x.e.id; }) };
}

/* 振り分け先が決まったレポートを、売上として登録する */
function sqApply_(rep, ev, msgId, docs) {
  var P = ev.P, ix = sqSkuIndex_(docs, P), tax = ev.cfg.tax != null ? Number(ev.cfg.tax) : 0.1, L = sqKey_(rep.loc).slice(0, 24) || 'x';
  var kind = rep.kind || 'sq', label = rep.label || 'Square', srcName = kind === 'ar' ? 'airregi' : kind === 'sc' ? 'squarecsv' : 'square';
  var ops = [], matched = [], unmatched = [], i = 0, pre = kind === 'ar' ? P + 'ar-' + rep.date + '-' : P + kind + '-' + rep.date + '-' + L + '-', day = P + rep.date;
  Object.keys(docs).forEach(function (k) { if (k.indexOf('tx/' + pre) === 0) ops.push({ op: 'delete', collection: 'tx', id: k.slice(3) }); });
  var at0 = Date.parse(rep.date + 'T23:30:00+09:00') || Date.now(), acc = {};
  rep.lines.forEach(function (l) {
    var id = sqMatch_(ix, l);
    if (l.qty <= 0 || l.amount == null || l.amount < 0) { unmatched.push({ name: l.name, variant: l.variant, qty: l.qty, amount: l.amount, why: '返品・マイナス（手入力してください）' }); return; }
    if (!id) { unmatched.push({ name: l.name, variant: l.variant, qty: l.qty, amount: l.amount, why: l.name === '任意の金額' ? '金額だけの入力（商品なし）' : '商品マスタに見つかりません' }); return; }
    var a = acc[id] = acc[id] || { qty: 0, amount: 0 }; a.qty += l.qty; a.amount += (l.incl ? l.amount / (1 + tax) : l.amount);
  });
  Object.keys(acc).forEach(function (id) {
    var sku = docs['skus/' + P + id] || {}, a = acc[id], std = (sku.sale != null ? Number(sku.sale) : 0) * a.qty, rec = { date: rep.date, skuId: id, kind: 'sale', qty: a.qty, at: at0 + (i++), by: label + '自動', memo: label + '自動取込', src: srcName };
    if (Math.abs(std - a.amount) > a.qty) rec.amount = Math.round(a.amount * (1 + tax));
    ops.push({ op: 'set', collection: 'tx', id: pre + id, data: rec });
    matched.push({ id: id, qty: a.qty, amount: a.amount });
  });
  if (rep.totals && rep.totals.total != null) {
    var cur = docs['days/' + day] || {}, reps = {}, tot = 0, cust = 0, disc = 0;
    Object.keys(cur.sqReports || {}).forEach(function (k) { reps[k] = cur.sqReports[k]; });
    reps[L] = { total: rep.totals.total, cust: rep.totals.customers, disc: Math.round(Math.abs(rep.totals.disc || 0) * (1 + tax)), msg: msgId };
    Object.keys(reps).forEach(function (k) { tot += reps[k].total || 0; cust += reps[k].cust || 0; disc += reps[k].disc || 0; });
    var dd = {}; Object.keys(cur).forEach(function (k) { dd[k] = cur[k]; });
    dd.sqReports = reps; dd.reportTotal = tot; dd.disc = disc; if (cust > 0) dd.customers = cust;
    ops.push({ op: 'set', collection: 'days', id: day, data: dd });
  }
  var log = { msgId: msgId, loc: rep.loc, date: rep.date, event: ev.id, status: unmatched.length ? 'partial' : 'done', matched: matched.length, pcs: matched.reduce(function (s, x) { return s + x.qty; }, 0), unmatched: unmatched, totals: rep.totals || {}, label: label, at: Date.now() };
  ops.push({ op: 'set', collection: 'imports', id: P + msgId, data: log });
  ops.push({ op: 'delete', collection: 'imports', id: 'pending~' + msgId });
  ops.push({ op: 'set', collection: 'sqseen', id: 'seen~' + msgId, data: { at: Date.now(), event: ev.id } });
  sqWrite_(ops);
  /* 次の処理でも最新が見えるよう、手元の読み込み結果も更新する */
  ops.forEach(function (o) { var k = o.collection + '/' + o.id; if (o.op === 'set') docs[k] = o.data; else delete docs[k]; });
  return log;
}

function sqPending_(rep, msg, reason, cands, docs) {
  var d = { msgId: msg.id, subject: msg.subject, loc: rep.loc || '', date: rep.date || '', reason: reason, cands: cands || [], status: 'pending', parsed: { loc: rep.loc, date: rep.date, totals: rep.totals, lines: rep.lines, kind: rep.kind || 'sq', label: rep.label || 'Square', days: rep.days || null }, at: Date.now() };
  var ops = [{ op: 'set', collection: 'imports', id: 'pending~' + msg.id, data: d }, { op: 'set', collection: 'sqseen', id: 'seen~' + msg.id, data: { at: Date.now(), pending: true } }];
  sqWrite_(ops);
  ops.forEach(function (o) { docs[o.collection + '/' + o.id] = o.data; });
  return d;
}

/* メッセージ1通ぶん。{msgId, status} を返す */
function sqHandle_(msg, docs, events) {
  if (docs['sqseen/seen~' + msg.id]) return null;
  var rep = parseSquareReport_(msg.body);
  if (!rep.ok) { sqPending_({ totals: {}, lines: [] }, msg, rep.reason || '読み取れませんでした', [], docs); return { id: msg.id, status: 'pending' }; }
  if (rep.multi) { sqPending_(rep, msg, '複数日のレポートです（1日ごとのレポートに設定してください）', [], docs); return { id: msg.id, status: 'pending' }; }
  return sqHandleRep_(rep, msg, docs, events);
}
function sqHandleRep_(rep, msg, docs, events) {
  var rt = sqRoute_(rep, events, docs);
  if (!rt.event) { sqPending_(rep, msg, rt.reason, rt.cands, docs); return { id: msg.id, status: 'pending' }; }
  var log = sqApply_(rep, rt.event, msg.id, docs);
  return { id: msg.id, status: log.status, event: rt.event.id };
}

function squareScanCore_() {
  var docs = sqDocs_(), events = sqEvents_(docs), res = { done: 0, partial: 0, pending: 0, skipped: 0 }, threads = [];
  try { threads = GmailApp.search(SQ_QUERY, 0, 50); } catch (e) { return { ok: false, error: 'gmail:' + e }; }
  var msgs = [];
  threads.forEach(function (t) { t.getMessages().forEach(function (m) { msgs.push({ id: m.getId(), subject: m.getSubject(), from: m.getFrom(), body: null, m: m }); }); });
  msgs.sort(function (a, b) { return a.m.getDate() - b.m.getDate(); });
  msgs.forEach(function (x) {
    if (docs['sqseen/seen~' + x.id]) { res.skipped++; return; }
    if (String(x.subject).indexOf('Square売上レポート') < 0) return;
    x.body = x.m.getBody();
    if (!mailAllowed_(x.from)) { sqPending_({ totals: {}, lines: [] }, x, '送信元（' + x.from + '）が許可されていません', [], docs); res.pending++; return; }
    try { var r = sqHandle_(x, docs, events); if (r) res[r.status === 'pending' ? 'pending' : r.status === 'partial' ? 'partial' : 'done']++; }
    catch (e) { Logger.log('square import failed ' + x.id + ': ' + e); }
  });
  try { csvMailScan_(docs, events, res); } catch (e) { Logger.log('csv mail scan failed: ' + e); res.csvError = String(e); }
  res.ok = true;
  return res;
}

/* =====================================================================
 * メールに添付されたCSV（エアレジ・Square）の自動取り込み
 *  - エアレジ：「バリエーション別売上_YYYYMMDD-YYYYMMDD.csv」（1日ぶん・ファイル名のまま送る）／「売上集計_….csv」
 *  - Square：ダッシュボードの商品別売上CSV（日付・アイテム・数量・総売上・割引・純売上 の列）
 *  - 送信元は MAIL_ALLOW（スクリプトプロパティ・カンマ区切りのドメインまたはアドレス。未設定なら ins.fm と squareup.com）のみ自動で反映。
 *    それ以外の送信元は「未振り分け」に溜めて、管理者が確認してから反映します。
 * ===================================================================== */
var SQ_QUERY_CSV = 'has:attachment filename:csv newer_than:14d';

function arKey_(t) { return sqKey_(t); }
function arNum_(x) { var n = parseFloat(String(x == null ? '' : x).replace(/[,\s¥￥]/g, '')); return isNaN(n) ? 0 : n; }
function arMoney_(x) {
  var s = String(x == null ? '' : x).trim(), neg = /^\(.*\)$/.test(s) || /^[-−△▲]/.test(s);
  var n = Math.abs(arNum_(s.replace(/^[(（△▲−-]+|[)）]+$/g, '')));
  return neg ? -n : n;
}
function arDate_(v) { var m = /(\d{4})\D+(\d{1,2})\D+(\d{1,2})/.exec(String(v || '')); return m ? m[1] + '-' + sqPad_(+m[2]) + '-' + sqPad_(+m[3]) : ''; }
function mailAllowed_(from) {
  var props = PropertiesService.getScriptProperties(), list = String(props.getProperty('MAIL_ALLOW') || 'ins.fm,squareup.com').toLowerCase().split(/[,\s;]+/).filter(Boolean);
  var m = /<([^>]+)>/.exec(String(from || '')), addr = (m ? m[1] : String(from || '')).trim().toLowerCase(), dom = addr.indexOf('@') >= 0 ? addr.split('@').pop() : '';
  return list.some(function (a) { return a === addr || (a.indexOf('@') < 0 && (dom === a || dom.slice(-(a.length + 1)) === '.' + a)); });
}
function csvText_(blob) {
  var t = blob.getDataAsString('UTF-8');
  if (t.indexOf('\uFFFD') >= 0) { try { t = blob.getDataAsString('Shift_JIS'); } catch (e) {} }
  return String(t).replace(/^\uFEFF/, '');
}
/* CSV を解釈する。戻り値：{kind:'air_items'|'air_sum'|'sq_items', ...} か null */
function csvParse_(name, text) {
  var rows = Utilities.parseCsv(text).filter(function (r) { return r.some(function (x) { return String(x).trim() !== ''; }); });
  if (!rows.length) return null;
  var hd = rows[0].map(function (x) { return String(x).trim(); });
  function ix(n) { return hd.indexOf(n); }
  function g(r, n) { var k = ix(n); return k < 0 ? '' : String(r[k] == null ? '' : r[k]).trim(); }
  if (ix('商品名') >= 0 && ix('販売商品数') >= 0) {
    var m = /(\d{8})\D+(\d{8})/.exec(name), out = { kind: 'air_items', from: m ? m[1] : '', to: m ? m[2] : '', items: [] };
    rows.slice(1).forEach(function (r) {
      var nm = g(r, '商品名'); if (!nm || nm === '合計') return;
      out.items.push({ name: nm, v1: g(r, '種別1'), v2: g(r, '種別2'), tax: g(r, '税区分'), amount: arNum_(g(r, '販売総売上')), sold: arNum_(g(r, '販売商品数')), ret: arNum_(g(r, '返品商品数')) });
    });
    return out;
  }
  if (ix('集計期間') >= 0 && ix('会計数') >= 0) {
    var o2 = { kind: 'air_sum', days: [] };
    rows.slice(1).forEach(function (r) {
      var q = g(r, '集計期間');
      if (/^\d{8}$/.test(q)) o2.days.push({ date: q.slice(0, 4) + '-' + q.slice(4, 6) + '-' + q.slice(6, 8), total: arNum_(g(r, '売上')), cust: arNum_(g(r, '客数')), disc: Math.abs(arNum_(g(r, '割引額'))) });
    });
    return o2;
  }
  function col(al) { for (var a = 0; a < al.length; a++) { for (var k = 0; k < hd.length; k++) if (arKey_(hd[k]) === arKey_(al[a])) return k; } return -1; }
  var cD = col(['日付', 'Date']), cI = col(['アイテム', '商品', '商品名', 'Item']), cQ = col(['数量', 'Qty', 'Quantity']), cG = col(['総売上', '総売上高', 'Gross Sales']), cN = col(['純売上', '純売上高', 'Net Sales']),
    cV = col(['価格ポイント名', '価格ポイント', 'バリエーション', 'バリエーション名', 'Price Point Name', 'Variation', 'Variation Name']), cX = col(['割引', 'ディスカウント', 'Discounts']),
    cL = col(['場所', '店舗', 'ロケーション', '拠点', 'Location']), cE = col(['イベントタイプ', 'Event Type']);
  if (cD >= 0 && cI >= 0 && cQ >= 0 && (cG >= 0 || cN >= 0)) {
    var agg = {}, list = [];
    function cell(r, k) { return k < 0 ? '' : String(r[k] == null ? '' : r[k]).trim(); }
    rows.slice(1).forEach(function (r) {
      var d = arDate_(cell(r, cD)), nm = cell(r, cI); if (!d || !nm) return;
      var q = arNum_(cell(r, cQ)), gross = cG >= 0 ? arMoney_(cell(r, cG)) : 0, disc = cX >= 0 ? Math.abs(arMoney_(cell(r, cX))) : 0;
      var net = cN >= 0 ? arMoney_(cell(r, cN)) : gross - disc;
      var vr = cell(r, cV); if (/^(regular|レギュラー|通常|標準)$/i.test(vr)) vr = '';
      var loc = cell(r, cL), ev = cell(r, cE), isRet = q < 0 || net < 0 || /refund|返金|返品|払い戻し/i.test(ev);
      var key = [d, loc, nm, vr].join('\u0001'), a = agg[key];
      if (!a) { a = agg[key] = { date: d, loc: loc, name: nm, variant: vr, amount: 0, sold: 0, ret: 0 }; list.push(a); }
      if (isRet) a.ret += Math.abs(q); else { a.sold += q; a.amount += net; }
    });
    return { kind: 'sq_items', items: list };
  }
  return null;
}
/* 1つのCSVを、取り込み単位（rep）の一覧に直す */
function csvReps_(name, parsed) {
  var reps = [], notes = [];
  if (parsed.kind === 'air_items') {
    if (!parsed.from) { notes.push('ファイル名から日付を読み取れませんでした（「バリエーション別売上_YYYYMMDD-YYYYMMDD.csv」のままお送りください）'); return { reps: [], notes: notes }; }
    if (parsed.from !== parsed.to) { notes.push('複数日のCSVです（商品別は1日ずつダウンロードしてください）'); return { reps: [], notes: notes }; }
    var lines = [];
    parsed.items.forEach(function (it) {
      var v = [it.v1, it.v2].filter(Boolean).join(' / ');
      if (it.sold > 0) lines.push({ name: it.name, variant: v, qty: it.sold, amount: it.amount, incl: it.tax !== '外税' });
      if (it.ret > 0) lines.push({ name: it.name, variant: v, qty: -it.ret, amount: null });
    });
    reps.push({ ok: true, kind: 'ar', label: 'エアレジ', loc: '', date: parsed.from.slice(0, 4) + '-' + parsed.from.slice(4, 6) + '-' + parsed.from.slice(6, 8), totals: {}, lines: lines });
  } else if (parsed.kind === 'sq_items') {
    var by = {}, order = [];
    parsed.items.forEach(function (it) {
      var k = it.date + '\u0001' + it.loc, r = by[k];
      if (!r) { r = by[k] = { ok: true, kind: 'sc', label: 'Square', loc: it.loc, date: it.date, totals: {}, lines: [] }; order.push(k); }
      if (it.sold > 0) r.lines.push({ name: it.name, variant: it.variant, qty: it.sold, amount: it.amount });
      if (it.ret > 0) r.lines.push({ name: it.name, variant: it.variant, qty: -it.ret, amount: null });
    });
    order.forEach(function (k) { reps.push(by[k]); });
  }
  return { reps: reps, notes: notes };
}
/* エアレジの売上集計（日別）を、会期の合う日だけ days に反映する。戻り値：反映した日数 */
function arSumApply_(days, ev, docs) {
  var ops = [], n = 0;
  days.forEach(function (d) {
    var key = 'days/' + ev.P + d.date, cur = docs[key] || {}, dd = {};
    Object.keys(cur).forEach(function (k) { dd[k] = cur[k]; });
    dd.reportTotal = Math.round(d.total); dd.disc = Math.round(d.disc || 0); if (d.cust > 0) dd.customers = Math.round(d.cust);
    ops.push({ op: 'set', collection: 'days', id: ev.P + d.date, data: dd }); n++;
  });
  if (ops.length) { sqWrite_(ops); ops.forEach(function (o) { docs[o.collection + '/' + o.id] = o.data; }); }
  return n;
}
function arSumHandle_(parsed, msg, docs, events) {
  var byEv = {}, pend = [];
  parsed.days.forEach(function (d) {
    var pool = events.filter(function (e) { return sqInRange_(e, d.date); });
    if (pool.length === 1) { (byEv[pool[0].id] = byEv[pool[0].id] || { e: pool[0], days: [] }).days.push(d); } else pend.push(d);
  });
  var done = 0, pending = 0;
  Object.keys(byEv).forEach(function (k) { done += arSumApply_(byEv[k].days, byEv[k].e, docs); });
  if (pend.length) {
    var rep = { loc: '', date: pend[0].date, totals: {}, lines: [], kind: 'air_sum', label: 'エアレジ', days: pend };
    sqPending_(rep, { id: msg.id, subject: msg.subject }, pend.length + ' 日ぶんが、会期に入っているイベントがない（または複数ある）ため振り分けできません', [], docs);
    pending = 1;
  }
  return { done: done, pending: pending };
}

function csvMailScan_(docs, events, res) {
  var threads = GmailApp.search(SQ_QUERY_CSV, 0, 50), msgs = [];
  threads.forEach(function (t) { t.getMessages().forEach(function (m) { msgs.push(m); }); });
  msgs.sort(function (a, b) { return a.getDate() - b.getDate(); });
  msgs.forEach(function (m) {
    var mid = m.getId();
    if (docs['sqseen/seen~' + mid]) return;
    var atts = [];
    try { atts = m.getAttachments(); } catch (e) { return; }
    var csvs = atts.filter(function (a) { return /\.csv$/i.test(String(a.getName())); });
    if (!csvs.length) return;
    var allowed = mailAllowed_(m.getFrom()), recognized = 0;
    csvs.forEach(function (a, i) {
      var name = String(a.getName()), parsed = null;
      try { parsed = csvParse_(name, csvText_(a)); } catch (e) { Logger.log('csv parse failed ' + name + ': ' + e); }
      if (!parsed) return;
      recognized++;
      var pm = { id: mid + '#' + i, subject: name };
      if (parsed.kind === 'air_sum') {
        if (!parsed.days.length) { sqPending_({ loc: '', date: '', totals: {}, lines: [], kind: 'air_sum', label: 'エアレジ', days: [] }, pm, '日別の行がありません', [], docs); res.pending++; return; }
        if (!allowed) { sqPending_({ loc: '', date: parsed.days[0].date, totals: {}, lines: [], kind: 'air_sum', label: 'エアレジ', days: parsed.days }, pm, '送信元（' + m.getFrom() + '）が許可されていません', [], docs); res.pending++; return; }
        var r0 = arSumHandle_(parsed, pm, docs, events); res.done += r0.done ? 1 : 0; res.pending += r0.pending; return;
      }
      var cr = csvReps_(name, parsed);
      if (!cr.reps.length) { sqPending_({ loc: '', date: parsed.from ? parsed.from.slice(0, 4) + '-' + parsed.from.slice(4, 6) + '-' + parsed.from.slice(6, 8) : '', totals: {}, lines: [], kind: 'ar', label: 'エアレジ' }, pm, cr.notes.join(' / ') || '読み取れませんでした', [], docs); res.pending++; return; }
      cr.reps.forEach(function (rep, j) {
        var rm = { id: mid + '#' + i + '.' + j, subject: name };
        if (!allowed) { sqPending_(rep, rm, '送信元（' + m.getFrom() + '）が許可されていません', [], docs); res.pending++; return; }
        var r = sqHandleRep_(rep, rm, docs, events);
        res[r.status === 'pending' ? 'pending' : r.status === 'partial' ? 'partial' : 'done']++;
      });
    });
    var seen = { op: 'set', collection: 'sqseen', id: 'seen~' + mid, data: { at: Date.now(), csv: recognized } };
    sqWrite_([seen]); docs['sqseen/seen~' + mid] = seen.data;
  });
}

function squareAssign_(pid, evId) {
  var docs = sqDocs_(), events = sqEvents_(docs), key = 'imports/' + pid, p = docs[key];
  if (!p || !p.parsed) throw 'no_pending';
  var ev = events.filter(function (e) { return e.id === evId; })[0];
  if (!ev) throw 'no_event';
  if (p.parsed.kind === 'air_sum') {
    var n = arSumApply_(p.parsed.days || [], ev, docs);
    sqWrite_([{ op: 'delete', collection: 'imports', id: pid }]);
    return { ok: true, status: 'done', matched: n, unmatched: 0 };
  }
  var rep = { ok: true, loc: p.parsed.loc, date: p.parsed.date, totals: p.parsed.totals || {}, lines: p.parsed.lines || [], kind: p.parsed.kind || 'sq', label: p.parsed.label || 'Square' };
  var log = sqApply_(rep, ev, p.msgId, docs);
  return { ok: true, status: log.status, matched: log.matched, unmatched: log.unmatched.length };
}
function squareDismiss_(pid) {
  var key = 'imports/' + pid;
  sqWrite_([{ op: 'delete', collection: 'imports', id: pid }]);
  return { ok: true };
}
