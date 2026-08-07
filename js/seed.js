/* 首次种子数据：示例收发货人 + 示例申报信息 + 内置模板（骨架，等真实样例替换） */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else { root.TD = root.TD || {}; root.TD.seed = factory(); }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var SEED_VER = 5;

  var PARTIES = [
    { id: 'p_shipper_demo', type: 'shipper', isSeed: true, name: '示例发货公司 DEMO TRADING CO., LTD.', address: 'ROOM 101, DEMO BUILDING, SHENZHEN, CHINA', tel: '0755-00000000', contact: 'DEMO', country: 'CN', remark: '示例数据，请在"收发货人"页维护真实信息' },
    { id: 'p_consignee_demo', type: 'consignee', isSeed: true, name: 'DEMO IMPORT LLC', address: '100 DEMO STREET, LOS ANGELES, CA, USA', tel: '+1-000-000-0000', contact: 'DEMO', country: 'US', remark: '示例数据' },
    { id: 'p_notify_demo', type: 'notify', isSeed: true, name: 'SAME AS CONSIGNEE', address: '', tel: '', contact: '', country: '', remark: '' }
  ];

  var DECLARES = [
    { sku: 'DEMO-SKU-001', nameCn: '示例商品A', nameEn: 'DEMO PRODUCT A', hsCode: '9503008900', declarePrice: 2.5, currency: 'USD', material: 'PLASTIC 塑料', usage: 'TOY 玩具', brand: 'NO BRAND', model: 'DEMO-SKU-001', nw: 0.2, gw: 0.25, unit: 'PCS', origin: 'CN', ver: 1 },
    { sku: 'DEMO-SKU-002', nameCn: '示例商品B', nameEn: 'DEMO PRODUCT B', hsCode: '6307900000', declarePrice: 1.8, currency: 'USD', material: 'POLYESTER 涤纶', usage: 'HOME USE 家用', brand: 'NO BRAND', model: 'DEMO-SKU-002', nw: 0.15, gw: 0.18, unit: 'PCS', origin: 'CN', ver: 1 }
  ];

  /** 幂等seed：按 config.seedVer 防重复。engine/ExcelJS 用于生成内置模板。 */
  function run(db, engine, ExcelJS) {
    var realTplsEarly = (typeof window !== 'undefined' && window.TD && window.TD.realTemplates) ? window.TD.realTemplates : [];
    var byIdEarly = {};
    realTplsEarly.forEach(function (rt) { byIdEarly[rt.id] = rt; });
    // 先做幂等迁移：给缺 logo / 缺 mapping.scanned 的内置模板补字段（不破坏用户手填数据）。
    var migrateJobs = db.all('templates').then(function (existing) {
      var ups = [];
      existing.forEach(function (t) {
        var src = byIdEarly[t.id];
        if (!src) return;
        var dirty = false;
        if (t.builtin && !t.logo && src.logo) { t.logo = src.logo; dirty = true; }
        // 方案 B：给缺 previewBuf 的内置模板补回源文件原样（预览直接显示 LOGO + 样张）
        if (t.builtin && src.previewBufB64 && !t.previewBuf) {
          try {
            var pbin = src.previewBufB64;
            var pbinary = (typeof atob !== 'undefined') ? atob(pbin) : Buffer.from(pbin, 'base64').toString('binary');
            var pab2 = new ArrayBuffer(pbinary.length);
            var pvu2 = new Uint8Array(pab2);
            for (var pj = 0; pj < pbinary.length; pj++) pvu2[pj] = pbinary.charCodeAt(pj);
            t.previewBuf = pab2;
            dirty = true;
          } catch (e) { /* ignore */ }
        }
        if (t.builtin && src.fileBufB64 && t.mapping && !t.mapping.scanned) {
          // 映射扫描结果也补上（v1.4.14 引入的扫描缓存）
          try {
            var bin = src.fileBufB64;
            var binary = (typeof atob !== 'undefined') ? atob(bin) : Buffer.from(bin, 'base64').toString('binary');
            var ab2 = new ArrayBuffer(binary.length);
            var vu2 = new Uint8Array(ab2);
            for (var i2 = 0; i2 < binary.length; i2++) vu2[i2] = binary.charCodeAt(i2);
            var wb2 = new ExcelJS.Workbook();
            wb2.xlsx.load(ab2).then(function (wb) {
              t.mapping.scanned = engine.scanTemplate(wb);
              db.put('templates', t);
            }).catch(function () { /* ignore */ });
          } catch (e) { /* ignore */ }
        }
        if (dirty) ups.push(db.put('templates', t));
      });
      return Promise.all(ups);
    });
    return migrateJobs.then(function () {
      // 申报要素主数据：从本地《商品申报信息》表整表重建（window.TD.declareData，build_declare_from_xlsx.py 生成）。
      // 用独立的 declareSeedVer 触发「清空 declare_reqs + 整表重建」；仅触发一次，不碰收发货人/模板，亦不覆盖首次 seed 逻辑。
      var declareJob = db.get('config', 'declareSeedVer').then(function (dcfg) {
        var DECLARE_VER = 3;
        if (dcfg && dcfg.value >= DECLARE_VER) return null;
        var rd = (typeof window !== 'undefined' && window.TD && window.TD.declareData) ? window.TD.declareData : [];
        return db.clear('declare_reqs').then(function () {
          return rd.length ? db.bulkPut('declare_reqs', rd) : null;
        }).then(function () {
          return db.put('config', { key: 'declareSeedVer', value: DECLARE_VER });
        });
      });
      return db.get('config', 'seedVer').then(function (cfg) {
        if (cfg && cfg.value >= SEED_VER) return declareJob.then(function () { return { seeded: false }; });
        var jobs = [
          db.bulkPut('parties', PARTIES.slice())
        ];
        // v1.4.59：不再内置/内嵌任何模板。模板唯一真源改为 GitHub userdata.json（stores.templates），
        // 启动时 pullShared 对 templates 做全量同步（清掉历史内置/嵌入/老浏览器缓存的残留模板）。
        jobs.push(declareJob);
        return Promise.all(jobs).then(function () {
          return db.put('config', { key: 'seedVer', value: SEED_VER }).then(function () { return { seeded: true }; });
        });
      });
    });
  }

  return { SEED_VER: SEED_VER, PARTIES: PARTIES, DECLARES: DECLARES, run: run };
});
