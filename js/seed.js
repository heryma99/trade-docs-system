/* 首次种子数据：示例收发货人 + 示例申报信息 + 内置模板（骨架，等真实样例替换） */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else { root.TD = root.TD || {}; root.TD.seed = factory(); }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var SEED_VER = 5;

  var PARTIES = [
    { id: 'p_shipper_demo', type: 'shipper', name: '示例发货公司 DEMO TRADING CO., LTD.', address: 'ROOM 101, DEMO BUILDING, SHENZHEN, CHINA', tel: '0755-00000000', contact: 'DEMO', country: 'CN', remark: '示例数据，请在"收发货人"页维护真实信息' },
    { id: 'p_consignee_demo', type: 'consignee', name: 'DEMO IMPORT LLC', address: '100 DEMO STREET, LOS ANGELES, CA, USA', tel: '+1-000-000-0000', contact: 'DEMO', country: 'US', remark: '示例数据' },
    { id: 'p_notify_demo', type: 'notify', name: 'SAME AS CONSIGNEE', address: '', tel: '', contact: '', country: '', remark: '' }
  ];

  var DECLARES = [
    { sku: 'DEMO-SKU-001', nameCn: '示例商品A', nameEn: 'DEMO PRODUCT A', hsCode: '9503008900', declarePrice: 2.5, currency: 'USD', material: 'PLASTIC 塑料', usage: 'TOY 玩具', brand: 'NO BRAND', model: 'DEMO-SKU-001', nw: 0.2, gw: 0.25, unit: 'PCS', origin: 'CN', ver: 1 },
    { sku: 'DEMO-SKU-002', nameCn: '示例商品B', nameEn: 'DEMO PRODUCT B', hsCode: '6307900000', declarePrice: 1.8, currency: 'USD', material: 'POLYESTER 涤纶', usage: 'HOME USE 家用', brand: 'NO BRAND', model: 'DEMO-SKU-002', nw: 0.15, gw: 0.18, unit: 'PCS', origin: 'CN', ver: 1 }
  ];

  /** 幂等seed：按 config.seedVer 防重复。engine/ExcelJS 用于生成内置模板。 */
  function run(db, engine, ExcelJS) {
    return db.get('config', 'seedVer').then(function (cfg) {
      if (cfg && cfg.value >= SEED_VER) return { seeded: false };
      var jobs = [
        db.bulkPut('parties', PARTIES.slice()),
        db.bulkPut('declare_reqs', DECLARES.slice())
      ];
      // 飞书《申报信息》双表合并镜像（tests/build_declare_data.js 生成，js/declare_data.js 提供）。
      // 仅新增本地缺失的 SKU，不覆盖用户已手填/已存在的申报要素。
      var realDeclares = (typeof window !== 'undefined' && window.TD && window.TD.declareData) ? window.TD.declareData : [];
      if (realDeclares.length) {
        jobs.push(db.all('declare_reqs').then(function (existing) {
          var have = {};
          existing.forEach(function (e) { have[e.sku] = true; });
          var toAdd = realDeclares.filter(function (r) { return !have[r.sku]; });
          return toAdd.length ? db.bulkPut('declare_reqs', toAdd) : null;
        }));
      }
      // 内置模板
      var inv = engine.makeBuiltinInvoiceTemplate(ExcelJS);
      var bok = engine.makeBuiltinBookingTemplate(ExcelJS);
      jobs.push(inv.xlsx.writeBuffer().then(function (buf) {
        return db.put('templates', {
          id: 'tpl_builtin_invoice', name: '内置·通用商业发票模板', kind: 'invoice', carrier: '通用',
          status: 'active', builtin: true, fileBuf: buf,
          mapping: { required: engine.REQUIRED_FIELDS.invoice }
        });
      }));
      jobs.push(bok.xlsx.writeBuffer().then(function (buf) {
        return db.put('templates', {
          id: 'tpl_builtin_booking', name: '内置·通用订舱单模板(BOOKING FORM)', kind: 'booking', carrier: '通用',
          status: 'active', builtin: true, fileBuf: buf,
          mapping: { required: engine.REQUIRED_FIELDS.booking }
        });
      }));
      // 真实业务模板（tests/build_real_templates.js 生成，js/real_templates.js 提供 base64）
      var realTpls = (typeof window !== 'undefined' && window.TD && window.TD.realTemplates) ? window.TD.realTemplates : [];
      realTpls.forEach(function (rt) {
        var bin = rt.fileBufB64 || '';
        var binary = (typeof atob !== 'undefined') ? atob(bin) : Buffer.from(bin, 'base64').toString('binary');
        var ab = new ArrayBuffer(binary.length);
        var vu = new Uint8Array(ab);
        for (var i = 0; i < binary.length; i++) vu[i] = binary.charCodeAt(i);
        var mapping = { required: engine.REQUIRED_FIELDS[rt.kind] || [] };
        // 首次 seed 时扫描模板占位符并缓存，供后续 boxMode 等逻辑使用
        var scanJob = new ExcelJS.Workbook().xlsx.load(ab).then(function (wb) {
          mapping.scanned = engine.scanTemplate(wb);
        }).catch(function (e) {
          mapping.scanned = { fields: [], itemFields: [], itemsRow: -1, itemHeaderMap: {}, sheetName: '' };
        });
        jobs.push(scanJob.then(function () {
          return db.put('templates', {
            id: rt.id, name: rt.name, kind: rt.kind, carrier: rt.carrier || '通用',
            status: 'active', builtin: true, fileBuf: ab,
            logo: rt.logo || null,
            mapping: mapping
          });
        }));
      });
      return Promise.all(jobs).then(function () {
        return db.put('config', { key: 'seedVer', value: SEED_VER }).then(function () { return { seeded: true }; });
      });
    });
  }

  return { SEED_VER: SEED_VER, PARTIES: PARTIES, DECLARES: DECLARES, run: run };
});
