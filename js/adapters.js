/* L2 数据源适配器：Local(IndexedDB默认) / RemoteHttp(有服务器时配URL) / JstApi(占位) */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else { root.TD = root.TD || {}; root.TD.adapters = factory(); }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /**
   * 统一接口 DataSourceAdapter：
   *   list(store)           → Promise<rows[]>
   *   pull(store, since)    → Promise<rows[]>   增量拉取
   *   available()           → Promise<bool>
   */

  function LocalAdapter(db) {
    this.name = 'local';
    this.db = db;
  }
  LocalAdapter.prototype.available = function () { return Promise.resolve(true); };
  LocalAdapter.prototype.list = function (store) { return this.db.all(store); };
  LocalAdapter.prototype.pull = function (store, since) {
    return this.db.all(store).then(function (rows) {
      return since ? rows.filter(function (r) { return (r.updatedAt || 0) > since; }) : rows;
    });
  };

  /**
   * RemoteHttpAdapter：约定 GET {baseURL}/{store}.json → [{...}]
   * 可注入 fetchFn 便于测试。失败抛错，由 sync 层降级本地。
   */
  function RemoteHttpAdapter(config, fetchFn) {
    this.name = 'remote';
    this.baseURL = (config && config.baseURL || '').replace(/\/+$/, '');
    this.token = config && config.token || '';
    this.fetchFn = fetchFn || (typeof fetch !== 'undefined' ? fetch.bind(typeof self !== 'undefined' ? self : this) : null);
    this.timeoutMs = (config && config.timeoutMs) || 8000;
  }
  RemoteHttpAdapter.prototype._get = function (path) {
    var self = this;
    if (!self.fetchFn) return Promise.reject(new Error('fetch不可用'));
    if (!self.baseURL) return Promise.reject(new Error('未配置远程数据源URL'));
    var headers = {};
    if (self.token) headers['Authorization'] = 'Bearer ' + self.token;
    var timeout = new Promise(function (_, rej) { setTimeout(function () { rej(new Error('远程请求超时')); }, self.timeoutMs); });
    var req = self.fetchFn(self.baseURL + path, { headers: headers }).then(function (res) {
      if (!res.ok) throw new Error('远程返回 ' + res.status);
      return res.json();
    });
    return Promise.race([req, timeout]);
  };
  RemoteHttpAdapter.prototype.available = function () {
    return this._get('/ping.json').then(function () { return true; }).catch(function () { return false; });
  };
  RemoteHttpAdapter.prototype.list = function (store) { return this._get('/' + store + '.json'); };
  RemoteHttpAdapter.prototype.pull = function (store, since) {
    return this.list(store).then(function (rows) {
      if (!Array.isArray(rows)) throw new Error('远程数据格式错误（应为数组）');
      return since ? rows.filter(function (r) { return (r.updatedAt || 0) > since; }) : rows;
    });
  };

  /** JstApiAdapter：聚水潭开放平台占位（拿到app_key/secret后实现签名与拉单） */
  function JstApiAdapter(config) {
    this.name = 'jst';
    this.config = config || {};
  }
  JstApiAdapter.prototype.available = function () { return Promise.resolve(false); };
  JstApiAdapter.prototype.list = function () { return Promise.reject(new Error('聚水潭API适配器未启用：请在设置中配置开放平台 app_key/app_secret 后开通')); };
  JstApiAdapter.prototype.pull = JstApiAdapter.prototype.list;

  /**
   * 远程同步（含降级）：从remote拉store合并进本地db。
   * 合并规则：远程按主键覆盖本地（远程为权威源），本地独有保留。
   * 返回 {ok, source, merged, error}
   */
  function syncFromRemote(db, remote, store) {
    return remote.pull(store).then(function (rows) {
      if (!rows || !rows.length) return { ok: true, source: 'remote', merged: 0 };
      var key = db.keyOf(store);
      var valid = rows.filter(function (r) { return r && r[key]; });
      return db.bulkPut(store, valid).then(function (n) {
        return { ok: true, source: 'remote', merged: n };
      });
    }).catch(function (e) {
      return { ok: false, source: 'local', merged: 0, error: e.message + '（已降级使用本地数据）' };
    });
  }

  return {
    LocalAdapter: LocalAdapter,
    RemoteHttpAdapter: RemoteHttpAdapter,
    JstApiAdapter: JstApiAdapter,
    syncFromRemote: syncFromRemote
  };
});
