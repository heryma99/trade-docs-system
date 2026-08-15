/* L6 校验引擎：单号匹配、SKU数量勾稽、字段完整性、申报要素反查、状态机 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else { root.TD = root.TD || {}; root.TD.validator = factory(); }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /** 1. 单号匹配：所选订单号集合 vs 装箱清单单号集合，必须完全一致 */
  function matchOrderNos(selectedOrderNos, packingOrderNos) {
    var sel = {}, pk = {};
    (selectedOrderNos || []).forEach(function (n) { if (n) sel[String(n).trim()] = 1; });
    (packingOrderNos || []).forEach(function (n) { if (n) pk[String(n).trim()] = 1; });
    var missingInPacking = Object.keys(sel).filter(function (n) { return !pk[n]; }); // 选了但箱单没有
    var extraInPacking = Object.keys(pk).filter(function (n) { return !sel[n]; });   // 箱单有但没选
    return {
      ok: missingInPacking.length === 0 && extraInPacking.length === 0,
      missingInPacking: missingInPacking,
      extraInPacking: extraInPacking
    };
  }

  /** 2. SKU+数量勾稽：订单聚合 vs 装箱清单聚合，逐SKU比对 */
  function checkSkuQty(orders, packing) {
    var oAgg = {}, pAgg = {};
    (orders || []).forEach(function (o) {
      (o.items || []).forEach(function (it) {
        var k = String(it.sku).trim();
        oAgg[k] = (oAgg[k] || 0) + (Number(it.qty) || 0);
      });
    });
    ((packing && packing.boxes) || []).forEach(function (b) {
      var k = String(b.sku).trim();
      pAgg[k] = (pAgg[k] || 0) + (Number(b.qty) || 0);
    });
    var diffs = [];
    var all = {};
    Object.keys(oAgg).forEach(function (k) { all[k] = 1; });
    Object.keys(pAgg).forEach(function (k) { all[k] = 1; });
    Object.keys(all).sort().forEach(function (sku) {
      var oq = oAgg[sku] || 0, pq = pAgg[sku] || 0;
      if (oq !== pq) diffs.push({ sku: sku, orderQty: oq, packingQty: pq, diff: pq - oq });
    });
    return { ok: diffs.length === 0, diffs: diffs };
  }

  /** 3. 字段完整性：按路径取值，空即缺失。required: [{path,label}] */
  function getPath(obj, path) {
    var parts = String(path).split('.');
    var cur = obj;
    for (var i = 0; i < parts.length; i++) {
      if (cur === null || cur === undefined) return undefined;
      cur = cur[parts[i]];
    }
    return cur;
  }
  function requiredFieldCheck(data, required) {
    var missing = [];
    (required || []).forEach(function (f) {
      var v = getPath(data, f.path);
      if (v === undefined || v === null || String(v).trim() === '') missing.push(f);
    });
    return { ok: missing.length === 0, missing: missing };
  }

  /** 4. 申报要素反查：按模板实际用到的 items.xxx 字段去申报信息中检查。
   *  templateItemFields: 模板中扫描到的明细占位符数组，如 ['items.sku','items.qty','items.hsCode']。
   *  若未提供或为空，兼容旧逻辑（检查 HS编码/申报价/材质）。 */
  var ITEM_FIELD_TO_DECLARE = {
    'items.hsCode': 'hsCode',
    'items.material': 'material',
    'items.declarePrice': 'declarePrice',
    'items.nameEn': 'nameEn',
    'items.nameCn': 'nameCn',
    'items.brand': 'brand',
    'items.usage': 'usage',
    'items.model': 'model',
    'items.unitNw': 'nw',
    'items.unitGw': 'gw',
    'items.origin': 'origin'
  };
  function resolveDeclare(items, declareMap, templateItemFields) {
    var enriched = [], missing = [];
    var needFields = [];
    if (templateItemFields && templateItemFields.length) {
      templateItemFields.forEach(function (f) {
        var dk = ITEM_FIELD_TO_DECLARE[f];
        if (dk && needFields.indexOf(dk) < 0) needFields.push(dk);
      });
    } else {
      needFields = ['hsCode', 'declarePrice', 'material'];
    }
    (items || []).forEach(function (it) {
      var d = declareMap[String(it.sku).trim()] || null;
      var e = Object.assign({}, it);
      if (d) {
        e.nameEn = e.nameEn || d.nameEn || '';
        e.nameCn = e.nameCn || d.nameCn || e.name || '';
        e.hsCode = d.hsCode || '';
        e.declarePrice = (e.price !== undefined && e.price !== null && e.price !== 0) ? e.price : (d.declarePrice || 0);
        e.material = d.material || '';
        e.usage = d.usage || '';
        e.brand = d.brand || '';
        e.model = d.model || String(it.sku);
        e.unitNw = d.nw || 0;
        e.unitGw = d.gw || 0;
        e.origin = d.origin || 'CN';
        e._declareSource = 'master'; // 绿色：主数据反查
      } else {
        e._declareSource = 'none';
      }
      var lacks = [];
      if (!d) {
        if (needFields.length) lacks.push('申报要素缺失');
      } else {
        needFields.forEach(function (dk) {
          var v = e[dk] !== undefined ? e[dk] : d[dk];
          if (v === undefined || v === null || String(v).trim() === '' || (typeof v === 'number' && v === 0 && dk === 'declarePrice')) {
            var label = { hsCode: 'HS编码', declarePrice: '申报价', material: '材质', nameEn: '英文品名', nameCn: '中文品名', brand: '品牌', usage: '用途', model: '型号', nw: '净重', gw: '毛重', origin: '原产地' }[dk] || dk;
            lacks.push(label);
          }
        });
      }
      if (lacks.length) missing.push({ sku: it.sku, lacks: lacks });
      enriched.push(e);
    });
    return { ok: missing.length === 0, items: enriched, missing: missing, needFields: needFields };
  }

  /** 5. 单证状态机：draft → validated → confirmed → exported */
  var FLOW = { draft: ['validated'], validated: ['confirmed', 'draft'], confirmed: ['exported', 'draft'], exported: [] };
  function canTransition(from, to) {
    return (FLOW[from] || []).indexOf(to) >= 0;
  }

  /** 综合校验（发票/订舱单生成前）：返回分级报告 */
  function validateDocument(opts) {
    // opts: {kind, orders, packing, data, requiredFields, declareMap, skipPacking, templateItemFields}
    var report = { blocks: [], warns: [], ok: true };
    if (!opts.skipPacking && opts.packing) {
      var m = matchOrderNos(opts.orders.map(function (o) { return o.orderNo; }), opts.packing.orderNos);
      if (!m.ok) {
        if (m.missingInPacking.length) report.blocks.push({ type: 'orderNo', msg: '所选订单在装箱清单中缺失: ' + m.missingInPacking.join(', ') });
        if (m.extraInPacking.length) report.blocks.push({ type: 'orderNo', msg: '装箱清单含未勾选的单号: ' + m.extraInPacking.join(', ') });
      }
      var q = checkSkuQty(opts.orders, opts.packing);
      if (!q.ok) report.blocks.push({ type: 'skuQty', msg: 'SKU数量不一致', diffs: q.diffs });
    }
    if (opts.declareMap) {
      var allItems = [];
      opts.orders.forEach(function (o) { allItems = allItems.concat(o.items || []); });
      var d = resolveDeclare(allItems, opts.declareMap, opts.templateItemFields);
      if (!d.ok) report.blocks.push({ type: 'declare', msg: '申报要素缺失', missing: d.missing, needFields: d.needFields });
    }
    if (opts.requiredFields && opts.data) {
      var f = requiredFieldCheck(opts.data, opts.requiredFields);
      if (!f.ok) report.blocks.push({ type: 'field', msg: '必填字段缺失: ' + f.missing.map(function (x) { return x.label || x.path; }).join(', '), missing: f.missing });
    }
    report.ok = report.blocks.length === 0;
    return report;
  }

  return {
    matchOrderNos: matchOrderNos,
    checkSkuQty: checkSkuQty,
    getPath: getPath,
    requiredFieldCheck: requiredFieldCheck,
    resolveDeclare: resolveDeclare,
    canTransition: canTransition,
    validateDocument: validateDocument
  };
});
