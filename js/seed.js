/* 首次种子数据：示例收发货人 + 示例申报信息 + 内置模板（骨架，等真实样例替换） */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else { root.TD = root.TD || {}; root.TD.seed = factory(); }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var SEED_VER = 6;

  var PARTIES = [
    { id: 'p_shipper_demo', type: 'shipper', isSeed: true, name: '示例发货公司 DEMO TRADING CO., LTD.', address: 'ROOM 101, DEMO BUILDING, SHENZHEN, CHINA', tel: '0755-00000000', contact: 'DEMO', country: 'CN', remark: '示例数据，请在"收发货人"页维护真实信息' },
    { id: 'p_consignee_demo', type: 'consignee', isSeed: true, name: 'DEMO IMPORT LLC', address: '100 DEMO STREET, LOS ANGELES, CA, USA', tel: '+1-000-000-0000', contact: 'DEMO', country: 'US', remark: '示例数据' },
    { id: 'p_notify_demo', type: 'notify', isSeed: true, name: 'SAME AS CONSIGNEE', address: '', tel: '', contact: '', country: '', remark: '' },
    { id: 'p_cons_macy_s_dc_pd', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC PD)', company: 'MACY\'S LOGISTICS DC', address: '1155 VAUGHN PKWY', city: 'PORTLAND', state: 'TN', zip: '37148', country: 'US', tel: '+1 615-745-2000', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC PD' },
    { id: 'p_cons_macy_s_dc_ok', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC OK)', company: 'MACY\'S LOGISTICS DC', address: '7120 E 76TH ST N', city: 'OWASSA', state: 'OK', zip: '74055', country: 'US', tel: '+1 918-401-2828', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC OK' },
    { id: 'p_cons_macy_s_dc_mb', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC MB)', company: 'MACY\'S LOGISTICS DC', address: '333 CAPERTON BLVD', city: 'MARTINSBURG', state: 'WV', zip: '25403', country: 'US', tel: '+1 304-901-3100', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC MB' },
    { id: 'p_cons_macy_s_dc_az', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC AZ)', company: 'MACY\'S LOGISTICS DC', address: '16575 W COMMERCE LANE', city: 'GOODYEAR', state: 'AZ', zip: '85338', country: 'US', tel: '+1 623-925-3600', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC AZ' },
    { id: 'p_cons_macy_s_dc_ci', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC CI)', company: 'MACY\'S LOGISTICS DC', address: '15541 EAST GATE AVE', city: 'CITY OF INDUSTRY', state: 'CA', zip: '91745', country: 'US', tel: '+1 626-855-4142', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC CI' },
    { id: 'p_cons_macy_s_dc_cl', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC CL)', company: 'MACY\'S LOGISTICS DC', address: '601 MIDPOINT RD', city: 'MINOOKA', state: 'IL', zip: '60447', country: 'US', tel: '+1 815-521-3533', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC CL' },
    { id: 'p_cons_macy_s_dc_ha', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC HA)', company: 'MACY\'S LOGISTICS DC', address: '28701 HALL RD', city: 'HAYWARD', state: 'CA', zip: '94545', country: 'US', tel: '+1 510-887-7333', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC HA' },
    { id: 'p_cons_macy_s_dc_jp', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC JP)', company: 'MACY\'S LOGISTICS DC', address: '3300 FASHION WAY', city: 'JOPPA', state: 'MD', zip: '21085', country: 'US', tel: '+1 410-612-8189', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC JP' },
    { id: 'p_cons_macy_s_dc_sc', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC SC)', company: 'MACY\'S LOGISTICS DC', address: '500 MEADOWLANDS PKWY', city: 'SECAUCUS', state: 'NJ', zip: '07094', country: 'US', tel: '+1 201-863-3250', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC SC' },
    { id: 'p_cons_macy_s_dc_st', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC ST)', company: 'MACY\'S LOGISTICS DC', address: '4401 SARR PKWY', city: 'STONE MOUNTAIN', state: 'GA', zip: '30083', country: 'US', tel: '+1 770-491-2211', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC ST' },
    { id: 'p_cons_macy_s_dc_sw', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC SW)', company: 'MACY\'S LOGISTICS DC', address: '301 GOVERNORS HWY', city: 'S WINDSOR', state: 'CT', zip: '06074', country: 'US', tel: '+1 860-282-3019', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC SW' },
    { id: 'p_cons_macy_s_dc_tm', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC TM)', company: 'MACY\'S LOGISTICS DC', address: '19201 HAMISH ROAD', city: 'TOMBALL', state: 'TX', zip: '77377', country: 'US', tel: '+1 281-803-6603', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC TM' },
    { id: 'p_cons_macy_s_dc_tu', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC TU)', company: 'MACY\'S LOGISTICS DC', address: '17000 SOUTHCENTER PKWY', city: 'TUKWILA', state: 'WA', zip: '98188', country: 'US', tel: '+1 206-575-2060', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC TU' },
    { id: 'p_cons_macy_s_dc_cg', type: 'consignee', isSeed: true, name: 'MACY\'S (DC CG)', company: 'MACY\'S', address: '1305 LIBERTY RIDGE RD', city: 'CHINA GROVE', state: 'NC', zip: '28023', country: 'US', tel: '', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC CG' },
    { id: 'p_cons_macy_s_dc_dv', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC DV)', company: 'MACY\'S LOGISTICS DC', address: '510 E 51ST AVE', city: 'DENVER', state: 'CO', zip: '80216', country: 'US', tel: '', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC DV' },
    { id: 'p_cons_macy_s_dc_bn', type: 'consignee', isSeed: true, name: 'MACY\'S LOGISTICS DC (DC BN)', company: 'MACY\'S LOGISTICS DC', address: '270 DANIELS WAY', city: 'BURLINGTON', state: 'NJ', zip: '08016', country: 'US', tel: '', contact: '', email: '', remark: '品牌=MACY\'S | 代码=DC BN' },
    { id: 'p_cons_nordstrom_ship_to_569', type: 'consignee', isSeed: true, name: 'NORDSTROM FC (SHIP TO 569)', company: 'NORDSTROM FC', address: '30 DISTRIBUTION DR', city: 'ELIZABETHTOWN', state: 'PA', zip: '17022', country: 'US', tel: '+1 717-366-1300', contact: '', email: '', remark: '品牌=NORDSTROM线上 | 代码=SHIP TO 569' },
    { id: 'p_cons_nordstrom_ship_to_584', type: 'consignee', isSeed: true, name: 'NORDSTROM INVENTORY (SHIP TO 584)', company: 'NORDSTROM INVENTORY', address: '490 COLUMBIA AVE', city: 'RIVERSIDE', state: 'CA', zip: '92507', country: 'US', tel: '+1 951-892-7140', contact: '', email: '', remark: '品牌=NORDSTROM线上 | 代码=SHIP TO 584' },
    { id: 'p_cons_nordstrom_ship_to_599', type: 'consignee', isSeed: true, name: 'NORDSTROM FC (SHIP TO 599)', company: 'NORDSTROM FC', address: '7700 18TH ST SW', city: 'CEDAR RAPIDS', state: 'IA', zip: '52404', country: 'US', tel: '+1 319-846-4000', contact: '', email: '', remark: '品牌=NORDSTROM线上 | 代码=SHIP TO 599' },
    { id: 'p_cons_nordstrom_ship_to_89', type: 'consignee', isSeed: true, name: 'NORDSTROM DC (SHIP TO 89)', company: 'NORDSTROM DC', address: '5703 N MARINE DR', city: 'PORTLAND', state: 'OR', zip: '97203', country: 'US', tel: '+1 503-240-6071', contact: '', email: '', remark: '品牌=NORDSTROM线下 | 代码=SHIP TO 89' },
    { id: 'p_cons_nordstrom_ship_to_299', type: 'consignee', isSeed: true, name: 'NORDSTROM DC (SHIP TO 299)', company: 'NORDSTROM DC', address: '5050 CHAVENELLE DR', city: 'DUBUQUE', state: 'IA', zip: '52002', country: 'US', tel: '+1 563-556-4050', contact: '', email: '', remark: '品牌=NORDSTROM线下 | 代码=SHIP TO 299' },
    { id: 'p_cons_nordstrom_ship_to_399', type: 'consignee', isSeed: true, name: 'NORDSTROM DC (SHIP TO 399)', company: 'NORDSTROM DC', address: '1600 S MILLIKEN AVE', city: 'ONTARIO', state: 'CA', zip: '91761', country: 'US', tel: '+1 909-390-1040', contact: '', email: '', remark: '品牌=NORDSTROM线下 | 代码=SHIP TO 399' },
    { id: 'p_cons_nordstrom_ship_to_499', type: 'consignee', isSeed: true, name: 'NORDSTROM DC (SHIP TO 499)', company: 'NORDSTROM DC', address: '37599 FILBERT ST', city: 'NEWARK', state: 'CA', zip: '94560', country: 'US', tel: '+1 510-794-5440', contact: '', email: '', remark: '品牌=NORDSTROM线下 | 代码=SHIP TO 499' },
    { id: 'p_cons_nordstrom_ship_to_699', type: 'consignee', isSeed: true, name: 'NORDSTROM DC (SHIP TO 699)', company: 'NORDSTROM DC', address: '839 COMMERCE DR', city: 'UPPER MARLBORO', state: 'MD', zip: '20774', country: 'US', tel: '+1 301-390-7800', contact: '', email: '', remark: '品牌=NORDSTROM线下 | 代码=SHIP TO 699' },
    { id: 'p_cons_nordstrom_ship_to_799', type: 'consignee', isSeed: true, name: 'NORDSTROM DC (SHIP TO 799)', company: 'NORDSTROM DC', address: '5497 NE 49TH TERRACE', city: 'GAINESVILLE', state: 'FL', zip: '32609', country: 'US', tel: '+1 352-384-2111', contact: '', email: '', remark: '品牌=NORDSTROM线下 | 代码=SHIP TO 799' },
    { id: 'p_cons_anthropologies_wedding_kansas_city_kansas_fulfillment_center_kc1', type: 'consignee', isSeed: true, name: 'Kansas City, Kansas Fulfillment Center (KC1)', company: 'Kansas City, Kansas Fulfillment Center (KC1)', address: '11681 State Avenue', city: 'Kansas City', state: 'KS', zip: '66111', country: 'US', tel: '+1 913-980-7216', contact: '', email: '', remark: '品牌=Anthropologies&Wedding | 代码=Kansas City, Kansas Fulfillment Center (KC1)' },
    { id: 'p_cons_anthropologies_wedding_gap_direct_fc_gfc', type: 'consignee', isSeed: true, name: 'Gap Direct FC (GFC)', company: 'Gap Direct FC (GFC)', address: '766 Brackbill Rd.', city: 'Gap', state: 'PA', zip: '17527', country: 'US', tel: '+1 717-442-1000', contact: '', email: '', remark: '品牌=Anthropologies&Wedding | 代码=Gap Direct FC (GFC)' },
    { id: 'p_cons_anthropologies_wedding_reno_direct_fc_rno', type: 'consignee', isSeed: true, name: 'Reno Direct FC (RNO)', company: 'Reno Direct FC (RNO)', address: '12055 Moya Blvd.', city: 'Reno', state: 'NV', zip: '89506', country: 'US', tel: '+1 775-971-1362', contact: '', email: '', remark: '品牌=Anthropologies&Wedding | 代码=Reno Direct FC (RNO)' },
    { id: 'p_cons_anthropologies_wedding_reno_retail_dc_ren', type: 'consignee', isSeed: true, name: 'Reno Retail DC (REN)', company: 'Reno Retail DC (REN)', address: '6640 Echo Ave.', city: 'Reno', state: 'NV', zip: '89506', country: 'US', tel: '+1 775-971-1316', contact: '', email: '', remark: '品牌=Anthropologies&Wedding | 代码=Reno Retail DC (REN)' },
    { id: 'p_cons_anthropologies_wedding_gap_retail_dc_gap', type: 'consignee', isSeed: true, name: 'Gap Retail DC (GAP)', company: 'Gap Retail DC (GAP)', address: '755 Brackbill Rd.', city: 'Gap', state: 'PA', zip: '17527', country: 'US', tel: '+1 717-442-1218', contact: '', email: '', remark: '品牌=Anthropologies&Wedding | 代码=Gap Retail DC (GAP)' },
    { id: 'p_cons_nordstrom_rack_ship_to_0562', type: 'consignee', isSeed: true, name: 'NORDSTROM FC (SHIP TO 0562)', company: 'NORDSTROM FC', address: '30 DISTRIBUTION DR', city: 'ELIZABETHTOWN', state: 'PA', zip: '17022', country: 'US', tel: '+1 717-366-1300', contact: '', email: '', remark: '品牌=Nordstrom Rack | 代码=SHIP TO 0562' },
    { id: 'p_cons_nordstrom_rack_ship_to_5629', type: 'consignee', isSeed: true, name: 'NORDSTROM FC (SHIP TO 5629)', company: 'NORDSTROM FC', address: '490 COLUMBIA AVE', city: 'RIVERSIDE', state: 'CA', zip: '92507', country: 'US', tel: '', contact: '', email: '', remark: '品牌=Nordstrom Rack | 代码=SHIP TO 5629' }
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
