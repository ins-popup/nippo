/* データ連携：Google Apps Script（ウェブアプリ）経由で Google スプレッドシートに保存します。
   画面側（index.html）からは「コレクション／ドキュメント」の形で読み書きできるようにしてあります。 */
(function () {
  "use strict";
  var cfg = window.T8_CONFIG || {};
  var COLS = ["skus", "tx", "days", "counts", "photos", "log", "config", "imports"];
  var store = {}, subs = [], code = "", inflight = 0, timer = null, scope = { role: "all", event: "" };
  /* イベント切替：イベントごとに、ドキュメントIDの先頭に「ev~イベントID~」を付けて同じ保存先に分けて入れる。
     最初のイベント（京都T8＝t8）は付けない（これまでのデータをそのまま使うため） */
  var LEGACY = "t8", evId = LEGACY, pfx = "";
  function mine(id) { return pfx === "" ? String(id).indexOf("~") < 0 : String(id).indexOf(pfx) === 0; }
  function strip(id) { return pfx === "" ? id : String(id).slice(pfx.length); }
  COLS.forEach(function (c) { store[c] = {}; });

  function configured() { return !!cfg.url && /^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec/.test(cfg.url); }
  function fail(msg, c) { var e = new Error(msg); e.code = c || "server"; return e; }
  function post(body) {
    return fetch(cfg.url, { method: "POST", redirect: "follow", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(Object.assign({ code: code }, body)) })
      .then(function (r) { return r.json(); })
      .then(function (j) { if (!j.ok) throw fail(j.error || "error", j.error === "unauthorized" ? "unauthorized" : "server"); return j; });
  }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function notify(name) { subs.slice().forEach(function (s) { if (s.name === name) s.fire(); }); }

  function applyList(data) {
    COLS.forEach(function (c) {
      var inc = data[c] || {};
      if (JSON.stringify(inc) !== JSON.stringify(store[c])) { store[c] = inc; notify(c); }
    });
  }
  function refresh() {
    if (inflight > 0) return Promise.resolve();
    return post({ action: "list" }).then(function (j) { if (inflight === 0) applyList(j.data || {}); });
  }
  function connect() {
    return post({ action: "list" }).then(function (j) {
      /* 入った合言葉の範囲（admin＝管理者／all＝全イベント／event＝そのイベントだけ）。古い保存先では返らないので all 扱い */
      scope = j.scope && j.scope.role ? { role: j.scope.role, event: j.scope.event || "" } : { role: "all", event: "" };
      applyList(j.data || {});
      clearInterval(timer);
      timer = setInterval(function () { if (!document.hidden) refresh().catch(function () {}); }, 20000);
      document.addEventListener("visibilitychange", function () { if (!document.hidden) refresh().catch(function () {}); });
    });
  }
  /* 書き込みは1件ずつ順番に送る（続けて保存したとき、後の内容が先に届いて前の内容で上書きされるのを防ぐ） */
  var sendChain = Promise.resolve();
  function send(ops) {
    inflight++;
    var pr = sendChain.then(function () { return post({ action: "write", ops: ops }); });
    sendChain = pr.catch(function () {});
    return pr.then(function () { inflight--; return refresh(); }, function (e) { inflight--; refresh().catch(function () {}); throw e; });
  }
  function rid() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }

  function Query(name, ord, lim) {
    return {
      orderBy: function (f, d) { return Query(name, { f: f, d: d || "asc" }, lim); },
      limit: function (n) { return Query(name, ord, n); },
      onSnapshot: function (cb) {
        var s = { name: name, fire: function () {
          var o = store[name] || {}, docs = Object.keys(o).filter(mine).map(function (raw) { return { id: strip(raw), exists: true, data: function () { return o[raw]; } }; });
          if (ord) docs.sort(function (a, b) { var x = a.data()[ord.f], y = b.data()[ord.f]; return (x < y ? -1 : x > y ? 1 : 0) * (ord.d === "desc" ? -1 : 1); });
          if (lim) docs = docs.slice(0, lim);
          cb({ docs: docs, size: docs.length });
        } };
        subs.push(s); setTimeout(function () { s.fire(); }, 0);
        return function () { subs = subs.filter(function (x) { return x !== s; }); };
      }
    };
  }
  function docRef(name, id0) {
    var id = pfx + id0;
    return {
      id: id0,
      set: function (d) { store[name][id] = clone(d); notify(name); return send([{ op: "set", collection: name, id: id, data: d }]); },
      update: function (d) { store[name][id] = Object.assign({}, store[name][id] || {}, clone(d)); notify(name); return send([{ op: "update", collection: name, id: id, data: d }]); },
      delete: function () { delete store[name][id]; notify(name); return send([{ op: "delete", collection: name, id: id }]); },
      onSnapshot: function (cb) {
        var s = { name: name, fire: function () { var v = store[name][id]; cb({ exists: !!v, data: function () { return v; } }); } };
        subs.push(s); setTimeout(function () { s.fire(); }, 0);
        return function () { subs = subs.filter(function (x) { return x !== s; }); };
      }
    };
  }
  var db = {
    collection: function (name) {
      var q = Query(name);
      q.doc = function (id) { return docRef(name, id || rid()); };
      q.add = function (d) { var r = docRef(name, rid()); return r.set(d).then(function () { return r; }); };
      return q;
    },
    doc: function (path) { var p = path.split("/"); return docRef(p[0], p[1]); }
  };

  function toB64(blob) {
    return new Promise(function (res, rej) {
      var r = new FileReader();
      r.onload = function () { res(String(r.result).split(",")[1]); };
      r.onerror = function () { rej(fail("read", "server")); };
      r.readAsDataURL(blob);
    });
  }
  var assets = {
    upload: function (blob, opts) {
      return toB64(blob).then(function (b64) { return post({ action: "upload", mime: (opts && opts.type) || blob.type || "image/jpeg", b64: b64 }); })
        .then(function (j) { return { id: j.id }; });
    },
    delete: function (id) { return post({ action: "trash", id: id }).then(function () { return { deleted: true }; }); }
  };
  var downloads = {
    save: function (req) {
      try {
        var blob = new Blob([req.data], { type: "text/csv;charset=utf-8" }), a = document.createElement("a");
        a.href = URL.createObjectURL(blob); a.download = req.filename; document.body.appendChild(a); a.click();
        setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
        return Promise.resolve({ status: "saved" });
      } catch (e) { return Promise.reject({ code: "unavailable" }); }
    }
  };

  /* 保存先が他の人の処理で混み合って「ロックのタイムアウト」になったときは、少し待ってやり直す */
  function postRetry(body, tries) {
    return post(body).catch(function (e) {
      if (tries > 0 && /ロック|lock/i.test(String(e && e.message))) {
        return new Promise(function (res) { setTimeout(res, 5000); }).then(function () { return postRetry(body, tries - 1); });
      }
      throw e;
    });
  }
  /* 管理者コード付きの書き込み（商品マスタ・イベント設定。多いときは分けて送る） */
  function adminWrite(admin, ops, onProgress) {
    var chunks = [], i;
    for (i = 0; i < ops.length; i += 20) chunks.push(ops.slice(i, i + 20));
    inflight++;
    var p = Promise.resolve();
    var done = 0;
    chunks.forEach(function (c) { p = p.then(function () { return postRetry({ action: "write", admin: admin, ops: c }, 8); }).then(function (r) { done += c.length; if (onProgress) onProgress(done, ops.length); return r; }); });
    return p.then(function () { inflight--; return refresh(); }, function (e) { inflight--; refresh().catch(function () {}); throw e; });
  }
  window.T8Data = {
    configured: configured,
    setCode: function (c) { code = c; },
    scope: function () { return scope; },
    /* Square自動取り込み：管理者用（未振り分けと取り込み履歴は、イベントに関係なく生のIDのまま返す） */
    rawImports: function () { var o = store.imports || {}; return Object.keys(o).map(function (k) { return { id: k, data: o[k] }; }); },
    squareNow: function (admin) { return post({ action: "squarenow", admin: admin }).then(function (j) { return refresh().then(function () { return j; }); }); },
    squareAssign: function (admin, id, ev) { return post({ action: "squareassign", admin: admin, id: id, event: ev }).then(function (j) { return refresh().then(function () { return j; }); }); },
    squareDismiss: function (admin, id) { return post({ action: "squaredismiss", admin: admin, id: id }).then(function (j) { return refresh().then(function () { return j; }); }); },
    /* 通常の書き込みをまとめて送る（今のイベントのIDに自動で接頭辞を付ける。エアレジCSVの取り込み用） */
    batch: function (ops) {
      var p = Promise.resolve(), i;
      for (i = 0; i < ops.length; i += 20) (function (c) {
        p = p.then(function () { return send(c.map(function (o) { return { op: o.op, collection: o.collection, id: pfx + o.id, data: o.data }; })); });
      })(ops.slice(i, i + 20));
      return p;
    },
    /* 管理者コード付きの書き込み（今のイベントのIDに自動で接頭辞を付ける。エアレジCSVの取り込み用） */
    writeAdmin: function (admin, ops, onProgress) {
      return adminWrite(admin, ops.map(function (o) { return { op: o.op, collection: o.collection, id: pfx + o.id, data: o.data }; }), onProgress);
    },
    getEventCodes: function (admin) { return post({ action: "getcodes", admin: admin }).then(function (j) { return j.codes || {}; }); },
    setEventCode: function (admin, ev, newCode) { return post({ action: "setcode", admin: admin, event: ev, newCode: newCode }); },
    connect: connect, refresh: refresh, db: db, assets: assets, downloads: downloads,
    photoUrl: function (id) { return "https://drive.google.com/thumbnail?id=" + encodeURIComponent(id) + "&sz=w1600"; },
    seed: function (admin, payload, onProgress) {
      if (pfx === "") return post({ action: "seed", admin: admin, payload: payload }).then(function () { return refresh(); });
      var cur = store.skus || {};
      var ops = (payload.skus || []).map(function (s) { var d = {}; Object.keys(s).forEach(function (k) { if (k !== "id") d[k] = s[k]; }); return { op: "set", collection: "skus", id: pfx + s.id, data: d }; })
        .filter(function (o) { return JSON.stringify(cur[o.id]) !== JSON.stringify(o.data); });
      if (!ops.length) return refresh();
      return adminWrite(admin, ops, onProgress);
    },
    /* イベントの一覧と切替 */
    LEGACY: LEGACY,
    setEvent: function (id) { evId = id || LEGACY; pfx = evId === LEGACY ? "" : "ev~" + evId + "~"; },
    eventId: function () { return evId; },
    events: function () { var d = (store.config || {}).events; return d && d.list ? d.list : null; },
    saveEvent: function (admin, ev, cfgData, makeIndex) {
      var list = (window.T8Data.events() || makeIndex || []).slice(), i = -1;
      list.forEach(function (x, k) { if (x.id === ev.id) i = k; });
      if (i >= 0) list[i] = ev; else list.push(ev);
      var p = ev.id === LEGACY ? "" : "ev~" + ev.id + "~";
      return adminWrite(admin, [{ op: "set", collection: "config", id: "events", data: { list: list } }, { op: "update", collection: "config", id: p + "main", data: cfgData }]);
    }
  };
})();
