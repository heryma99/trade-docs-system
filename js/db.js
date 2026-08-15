/* L4 主数据存储：IndexedDB封装（DB_VER升版迁移 + 内存兜底） */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else { root.TD = root.TD || {}; root.TD.db = factory(); }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var DB_NAME = 'trade_docs';
  var DB_VER = 1; // ⚠ 新增/修改 store 必须升版本号
  var STORES = ['orders', 'packings', 'parties', 'declare_reqs', 'templates', 'documents', 'config'];
  var KEYS = { declare_reqs: 'sku', config: 'key' }; // 其余用 id

  var _db = null;
  var _useMem = false;             // IndexedDB不可用时内存兜底
  var _mem = {};

  function _memStore(name) { if (!_mem[name]) _mem[name] = {}; return _mem[name]; }
  function keyOf(store) { return KEYS[store] || 'id'; }

  function open() {
    if (_db) return Promise.resolve(_db);
    if (_useMem || typeof indexedDB === 'undefined') { _useMem = true; return Promise.resolve(null); }
    return new Promise(function (resolve) {
      var req;
      try { req = indexedDB.open(DB_NAME, DB_VER); }
      catch (e) { _useMem = true; return resolve(null); }
      req.onupgradeneeded = function (ev) {
        var db = ev.target.result;
        STORES.forEach(function (s) {
          if (!db.objectStoreNames.contains(s)) db.createObjectStore(s, { keyPath: keyOf(s) });
        });
      };
      req.onsuccess = function (ev) { _db = ev.target.result; resolve(_db); };
      req.onerror = function () { _useMem = true; resolve(null); };
    });
  }

  function _tx(store, mode, fn) {
    return open().then(function (db) {
      if (_useMem || !db) return fn(null);
      return new Promise(function (resolve, reject) {
        var tx = db.transaction(store, mode);
        var os = tx.objectStore(store);
        var out = fn(os);
        tx.oncomplete = function () { resolve(out && out._v !== undefined ? out._v : out); };
        tx.onerror = function () { reject(tx.error); };
      });
    });
  }

  function genId() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 8); }

  function put(store, obj) {
    var k = keyOf(store);
    if (!obj[k]) obj[k] = genId();
    if (obj.createdAt === undefined) obj.createdAt = Date.now();
    obj.updatedAt = Date.now();
    if (_useMem) { _memStore(store)[obj[k]] = obj; return Promise.resolve(obj); }
    return _tx(store, 'readwrite', function (os) {
      if (!os) { _memStore(store)[obj[k]] = obj; return obj; }
      os.put(obj); return obj;
    });
  }

  function bulkPut(store, list) {
    var k = keyOf(store);
    list.forEach(function (o) { if (!o[k]) o[k] = genId(); if (o.createdAt === undefined) o.createdAt = Date.now(); o.updatedAt = Date.now(); });
    if (_useMem) { list.forEach(function (o) { _memStore(store)[o[k]] = o; }); return Promise.resolve(list.length); }
    return _tx(store, 'readwrite', function (os) {
      if (!os) { list.forEach(function (o) { _memStore(store)[o[k]] = o; }); return list.length; }
      list.forEach(function (o) { os.put(o); });
      return list.length;
    });
  }

  function get(store, key) {
    if (_useMem) return Promise.resolve(_memStore(store)[key] || null);
    return open().then(function (db) {
      if (!db) return _memStore(store)[key] || null;
      return new Promise(function (resolve, reject) {
        var req = db.transaction(store).objectStore(store).get(key);
        req.onsuccess = function () { resolve(req.result || null); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function all(store) {
    if (_useMem) return Promise.resolve(Object.values(_memStore(store)));
    return open().then(function (db) {
      if (!db) return Object.values(_memStore(store));
      return new Promise(function (resolve, reject) {
        var req = db.transaction(store).objectStore(store).getAll();
        req.onsuccess = function () { resolve(req.result || []); };
        req.onerror = function () { reject(req.error); };
      });
    });
  }

  function del(store, key) {
    if (_useMem) { delete _memStore(store)[key]; return Promise.resolve(true); }
    return _tx(store, 'readwrite', function (os) {
      if (!os) { delete _memStore(store)[key]; return true; }
      os.delete(key); return true;
    });
  }

  function clear(store) {
    if (_useMem) { _mem[store] = {}; return Promise.resolve(true); }
    return _tx(store, 'readwrite', function (os) {
      if (!os) { _mem[store] = {}; return true; }
      os.clear(); return true;
    });
  }

  function forceMemMode() { _useMem = true; } // 测试用

  return {
    DB_NAME: DB_NAME, DB_VER: DB_VER, STORES: STORES,
    open: open, put: put, bulkPut: bulkPut, get: get, all: all, del: del, clear: clear,
    genId: genId, keyOf: keyOf, forceMemMode: forceMemMode,
    isMem: function () { return _useMem; }
  };
});
