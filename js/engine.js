/* L5 生成引擎：数据组装 + ExcelJS模板占位符填充（{{field}} / {{items.xxx}} 明细行复制） */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else { root.TD = root.TD || {}; root.TD.engine = factory(); }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var PH_RE = /\{\{\s*([A-Za-z0-9_.]+)\s*\}\}/g;

  /** 模板明细表头别名词典：把模板里真实列名映射到数据字段 */
  var HEADER_ALIASES = {
    no:      ['NO.', 'NO', '序号', '#', '项次'],
    boxNo:   ['CTNR NO', 'CTNR NO.', 'CTNRNO', 'CTNR', 'CARTON NO', 'CARTON#', '箱号', 'CARTON', 'CTN#'],
    boxCount:['CTNS', 'CARTONS', '箱数', '件数', 'TYPES OF PKG', 'TYPESOFPKG', 'NO. AND TYPES OF PKG'],
    sku:     ['SKU', '款号', '商品编码', '商家编码', '货号', '规格编码', 'MODEL NO', 'MODELNO'],
    model:   ['MODEL', '型号', 'MODEL NO.', 'MODEL NO'],
    nameEn:  ['DESCRIPTION', 'DESCRIPTION OF GOODS', 'GOODS DESCRIPTION', '英文品名', '品名(英文)', '品名'],
    nameCn:  ['中文品名', '品名(中文)', '商品名称', '中文名称'],
    hsCode:  ['HS CODE', 'HSCODE', 'HS编码', '海关编码', 'TARIFF CODE'],
    qty:     ['QTY', '数量', '件数', 'PCS', 'QUANTITY', 'TOTAL QUANTITY'],
    unit:    ['UNIT', '单位', 'UOM'],
    price:   ['UNIT PRICE', '单价', 'PRICE', 'DECLARED VALUE'],
    amount:  ['AMOUNT', '金额', 'TOTAL', 'TOTAL AMOUNT', 'TOTAL PRICE', '总申报价值', '申报价值'],
    nw:      ['N.W', 'N.W.', 'NW', 'NET WEIGHT', '净重'],
    gw:      ['G.W', 'G.W.', 'GW', 'GROSS WEIGHT', '毛重'],
    singleGw:['SINGLE GW', 'PER CTN G.W', 'WEIGHT/CTN', '单箱重量', '单箱毛重', 'WEIGHT PER CTN', '单箱重', '单箱重(KG)'],
    volume:  ['CBM', 'M3', 'VOL', 'MEAS', '体积', '尺码', 'MEAS\'T'],
    material:['MATERIAL', '材质', '质地'],
    usage:   ['用途', 'USE', 'USAGE', '用途说明', 'PURPOSE'],
    brand:   ['BRAND', '品牌'],
    origin:  ['ORIGIN', '原产地', '产地', 'COUNTRY OF ORIGIN'],
    // v1.4.43 海运类模板「长/宽/高/产品图片/产品性质/备注」等 per-item 字段
    // 注意：alias 不再放单字母 'L'/'W'/'H'（normalizeHeader 去空格后会被任意含该字母的词误命中）
    lengthCm:['LENGTH', 'L(CM)', '长', '长CM'],
    widthCm: ['WIDTH',  'W(CM)', '宽', '宽CM'],
    heightCm:['HEIGHT', 'H(CM)', '高', '高CM'],
    imageUrl:['IMAGE', 'PHOTO', 'PICTURE', '图片', '产品图片', '商品图片'],
    // v1.4.48：中运通达等模板中英双语表头「Purpose(用途)/产品电池类型/是否带电(Y/N)/是否带磁(Y/N)/
    //          产品销售价格/图片链接/单个产品净重/币种」此前因 alias 缺失而整列空（已被 scanTemplate 兜底选中
    //          为明细表头，但列未映射任何字段 → 不写值）。补齐下列中英别名。
    productNature:['NATURE', 'PRODUCT NATURE', '性质', '产品性质'],
    batteryType:['BATTERY TYPE', 'POWERED', '电池类型', '产品电池类型', '电池'],
    electrified:['ELECTRIFIED', '是否带电', '带电(Y/N)', '带电'],
    magnetic:['MAGNETIC', '是否带磁', '带磁(Y/N)', '带磁'],
    productPrice:['产品价格', '产品销售价格', '销售价格', '零售价', 'RETAIL PRICE', 'SALE PRICE', 'UNIT PRICE'],
    productLink:['产品链接', '图片链接', '商品链接', 'LINK', 'URL', 'PRODUCT LINK'],
    singleNw:['单个产品净重', '单品净重', 'NET WT/PC', 'NET WEIGHT/PC'],
    currency:['CURRENCY', '币种', '货币'],
    remark:  ['REMARK', 'NOTES', 'NOTE', '备注', '说明']
  };

  function normalizeHeader(s) {
    return String(s || '').replace(/\s+/g, '').replace(/[（(].*?[)）]/g, '').replace(/[*※]+$/, '').toUpperCase();
  }
  function matchHeaderAlias(s) {
    var ns = normalizeHeader(s);
    if (!ns) return null;
    var keys = Object.keys(HEADER_ALIASES);
    var best = null, bestLen = 0;
    for (var i = 0; i < keys.length; i++) {
      var al = HEADER_ALIASES[keys[i]];
      for (var j = 0; j < al.length; j++) {
        var a = normalizeHeader(al[j]);
        if (!a) continue;
        // 完全匹配优先级最高，立刻返回
        if (ns === a) return keys[i];
        // 子串匹配时，只接受以单词边界开头/结尾的，避免 "UNIT PRICE" 被 "UNIT" 截胡
        var hit = ns.indexOf(a) >= 0 || a.indexOf(ns) >= 0;
        if (hit && a.length > bestLen) { best = keys[i]; bestLen = a.length; }
      }
    }
    return best;
  }

  function getPath(obj, path) {
    var parts = String(path).split('.');
    var cur = obj;
    for (var i = 0; i < parts.length; i++) {
      if (cur === null || cur === undefined) return undefined;
      cur = cur[parts[i]];
    }
    return cur;
  }
  function round(n, d) { var p = Math.pow(10, d); return Math.round((Number(n) || 0) * p) / p; }

  /** 金额大写（美元英文） */
  function amountInWords(n, currency) {
    currency = currency || 'USD';
    var ones = ['', 'ONE', 'TWO', 'THREE', 'FOUR', 'FIVE', 'SIX', 'SEVEN', 'EIGHT', 'NINE', 'TEN', 'ELEVEN', 'TWELVE', 'THIRTEEN', 'FOURTEEN', 'FIFTEEN', 'SIXTEEN', 'SEVENTEEN', 'EIGHTEEN', 'NINETEEN'];
    var tens = ['', '', 'TWENTY', 'THIRTY', 'FORTY', 'FIFTY', 'SIXTY', 'SEVENTY', 'EIGHTY', 'NINETY'];
    function three(num) {
      var s = '';
      if (num >= 100) { s += ones[Math.floor(num / 100)] + ' HUNDRED'; num %= 100; if (num) s += ' AND '; }
      if (num >= 20) { s += tens[Math.floor(num / 10)]; num %= 10; if (num) s += '-' + ones[num]; }
      else if (num > 0) s += ones[num];
      return s;
    }
    var int = Math.floor(Math.abs(n)), cents = Math.round((Math.abs(n) - int) * 100);
    if (int === 0 && cents === 0) return 'ZERO ' + currency;
    var parts = [], units = ['', ' THOUSAND', ' MILLION', ' BILLION'], ui = 0;
    while (int > 0) { var seg = int % 1000; if (seg) parts.unshift(three(seg) + units[ui]); int = Math.floor(int / 1000); ui++; }
    var s = 'SAY ' + (currency === 'USD' ? 'US DOLLARS ' : currency + ' ') + (parts.join(' ') || 'ZERO');
    s += cents ? ' AND CENTS ' + three(cents) + ' ONLY' : ' ONLY';
    return s;
  }

  /**
   * 组装单证数据：orders(多单合并) + packing + meta(表头字段) + parties + declareMap
   * 返回统一 data：header字段 + items[] + totals
   */
  function buildDocData(opts) {
    var orders = opts.orders || [], packing = opts.packing || null;
    var meta = opts.meta || {}, shipper = opts.shipper || {}, consignee = opts.consignee || {}, notify = opts.notify || null;
    var declareMap = opts.declareMap || {};

    // 订单号列表
    var orderList = [];
    orders.forEach(function (o) { orderList.push(o.orderNo); });

    // 明细：方案A——模板需要箱号/尺寸(boxNo/length/width/height)且关联了装箱清单时，按「箱×SKU」逐行展开
    var boxMode = !!(opts.boxMode && packing && (packing.boxes || []).length);

    // 前置：按 SKU 聚合订单价格/币种/参考号
    var priceBySku = {}, currencyBySku = {}, orderSkuQty = {};
    orders.forEach(function (o) {
      (o.items || []).forEach(function (it) {
        var sku = String(it.sku).trim();
        var p = Number(it.price);
        if (p > 0 && priceBySku[sku] === undefined) priceBySku[sku] = p;
        if (it.currency && currencyBySku[sku] === undefined) currencyBySku[sku] = it.currency;
        if (!orderSkuQty[sku]) orderSkuQty[sku] = 0;
        orderSkuQty[sku] += Number(it.qty) || 0;
      });
    });
    var refId = meta.refId || (orderList.length ? orderList[0] : '');
    var destCountry = (consignee && consignee.country) || '';
    var tradeTerm = meta.incoterms || '';

    // 按 SKU 聚合装箱信息（箱数、总净毛重、体积、第一箱尺寸）
    var boxAgg = {};
    if (packing) {
      (packing.boxes || []).forEach(function (b) {
        var sku = String(b.sku == null ? '' : b.sku).trim();
        if (!sku) return;
        if (!boxAgg[sku]) boxAgg[sku] = { count: 0, qty: 0, nw: 0, gw: 0, volume: 0, volumeWeight: 0, firstDims: '', firstBoxSpec: '' };
        var L = Number(b.length) || 0, W = Number(b.width) || 0, H = Number(b.height) || 0;
        var vol = (L && W && H) ? round(L * W * H / 1000000, 6) : 0;
        var volW = (L && W && H) ? round(L * W * H / 6000, 3) : 0;
        var dims = (L && W && H) ? (L + '×' + W + '×' + H) : '';
        boxAgg[sku].count += 1;
        boxAgg[sku].qty += Number(b.qty) || 0;
        boxAgg[sku].nw += Number(b.nw) || 0;
        boxAgg[sku].gw += Number(b.gw) || 0;
        boxAgg[sku].volume += vol;
        boxAgg[sku].volumeWeight += volW;
        if (!boxAgg[sku].firstDims && dims) { boxAgg[sku].firstDims = dims; boxAgg[sku].firstBoxSpec = dims; }
      });
    }

    function makeItemCore(sku, d, itR) {
      itR = itR || {};
      var price = (priceBySku[sku] !== undefined) ? priceBySku[sku] : (d.declarePrice || 0);
      var currency = currencyBySku[sku] || d.currency || meta.currency || 'USD';
      // v1.4.43: 回退到订单 item 的字段（declareMap 缺字段时不再丢品名/材质等）
      //          同时新增 lengthCm/widthCm/heightCm/singleGw/imageUrl/remark 出口供明细 per-item 列填充
      function _pick(k) { return (d[k] !== undefined && d[k] !== '') ? d[k] : (itR[k] !== undefined && itR[k] !== '' ? itR[k] : ''); }
      return {
        sku: sku,
        model: _pick('model') || sku,
        nameEn: _pick('nameEn'), nameCn: _pick('nameCn'),
        hsCode: _pick('hsCode'), material: _pick('material'), usage: _pick('usage'), brand: _pick('brand'),
        lengthCm: _pick('lengthCm') || _pick('length'),
        widthCm:  _pick('widthCm')  || _pick('width'),
        heightCm: _pick('heightCm') || _pick('height'),
        singleGw: _pick('singleGw'),
        imageUrl: _pick('imageUrl') || _pick('image') || _pick('photo'),
        remark:   _pick('remark')   || _pick('note'),
        productNature: _pick('productNature') || _pick('batteryType') || _pick('nature'),
        productPrice: d.productPrice || d.retailPrice || d.salePrice || itR.productPrice || '',
        productLink: d.productLink || d.link || d.url || itR.productLink || '',
        singleNw: d.singleNw || d.netWeightPerPc || d.nwPerPc || itR.singleNw || '',
        price: price, unit: d.unit || itR.unit || 'PCS', origin: d.origin || itR.origin || 'CN',
        electrified: (d.electrified != null && d.electrified !== '') ? d.electrified : (itR.electrified != null ? itR.electrified : '否'),
        magnetic: (d.magnetic != null && d.magnetic !== '') ? d.magnetic : (itR.magnetic != null ? itR.magnetic : '否'),
        asin: d.asin || itR.asin || '', fnsku: d.fnsku || itR.fnsku || sku, note: d.note || itR.note || '', costPrice: d.costPrice || itR.costPrice || '',
        currency: currency, destCountry: destCountry, refId: refId,
        condition: d.condition || itR.condition || 'NEW', exportPrefer: d.exportPrefer || itR.exportPrefer || '', tradeTerm: tradeTerm,
        batteryType: d.batteryType || itR.batteryType || '',
        taxNo: d.taxNo || itR.taxNo || '',
        shippingMarks: '', poNo: refId, manufacturer: d.manufacturer || itR.manufacturer || ''
      };
    }

    var items = [];
    if (boxMode) {
      items = (packing.boxes || []).map(function (b) {
        var sku = String(b.sku == null ? '' : b.sku).trim();
        var d = declareMap[sku] || {};
        var L = Number(b.length) || 0, W = Number(b.width) || 0, H = Number(b.height) || 0;
        var dims = (L && W && H) ? (L + '×' + W + '×' + H) : (b.boxSpec || '');
        var volW = (L && W && H) ? round(L * W * H / 6000, 3) : (b.volumeWeight || 0);
        var vol = (L && W && H) ? round(L * W * H / 1000000, 6) : 0;
        var it = makeItemCore(sku, d);
        it.boxNo = b.boxNo || ''; it.boxSpec = dims; it.dims = dims;
        it.length = b.length || ''; it.width = b.width || ''; it.height = b.height || '';
        // v1.4.47：boxMode 直接把箱规同步给 per-item 长/宽/高/单箱重字段（模板列映射的是 lengthCm/... 而非 length）
        it.lengthCm = b.length || ''; it.widthCm = b.width || ''; it.heightCm = b.height || '';
        it.singleGw = Number(b.gw) || 0;
        it.volume = vol; it.volumeWeight = volW;
        it.boxCount = 1; it.ctns = 1;
        it.qty = Number(b.qty) || 0; it.amount = round(it.qty * it.price, 2);
        it.nw = Number(b.nw) || 0; it.gw = Number(b.gw) || 0; it.boxNw = it.nw;
        it._weightSource = 'packing';
        return it;
      });
      items.sort(function (a, b) { return (a.boxNo < b.boxNo ? -1 : a.boxNo > b.boxNo ? 1 : (a.sku < b.sku ? -1 : 1)); });
    } else {
      // 明细合并（跨订单同SKU同价合并）
      var agg = {};
      orders.forEach(function (o) {
        (o.items || []).forEach(function (it) {
          var sku = String(it.sku).trim();
          var d = declareMap[sku] || {};
          var price = (it.price !== undefined && it.price !== null && Number(it.price) > 0) ? Number(it.price) : (d.declarePrice || 0);
          var key = sku + '@' + price;
          if (!agg[key]) {
            agg[key] = makeItemCore(sku, d, it);
            agg[key].qty = 0; agg[key].amount = 0; agg[key].nw = 0; agg[key].gw = 0;
            agg[key]._priceSource = (it.price && Number(it.price) > 0) ? 'order' : (d.declarePrice ? 'master' : 'none');
          }
          agg[key].qty += Number(it.qty) || 0;
        });
      });
      items = Object.keys(agg).map(function (k) { return agg[k]; });
      items.sort(function (a, b) { return a.sku < b.sku ? -1 : 1; });
    }

    // 非箱模式：从装箱清单按SKU分摊净毛重，并补充箱数/体积/箱规
    var pkAgg = {};
    if (packing && !boxMode) {
      (packing.boxes || []).forEach(function (b) {
        var k = String(b.sku).trim();
        if (!pkAgg[k]) pkAgg[k] = { qty: 0, nw: 0, gw: 0 };
        pkAgg[k].qty += b.qty; pkAgg[k].nw += b.nw || 0; pkAgg[k].gw += b.gw || 0;
      });
    }
    var totals = { qty: 0, amount: 0, nw: 0, gw: 0, boxCount: 0, volume: 0, volumeWeight: 0 };
    items.forEach(function (it, idx) {
      it.no = idx + 1;
      it.amount = round(it.qty * it.price, 2);
      var sku = String(it.sku).trim();
      var ba = boxAgg[sku];
      if (!boxMode) {
        var pk = pkAgg[sku];
        if (pk && pk.qty > 0) {
          it.nw = round(pk.nw * (it.qty / pk.qty), 3);
          it.gw = round(pk.gw * (it.qty / pk.qty), 3);
          it._weightSource = 'packing';
        } else {
          var d = declareMap[sku] || {};
          it.nw = round((d.nw || 0) * it.qty, 3);
          it.gw = round((d.gw || 0) * it.qty, 3);
          it._weightSource = d.nw ? 'master' : 'none';
        }
        if (ba) {
          it.boxCount = ba.count; it.ctns = ba.count;
          it.dims = ba.firstDims; it.boxSpec = ba.firstBoxSpec;
          it.volume = round(ba.volume, 6); it.volumeWeight = round(ba.volumeWeight, 3);
          it.boxNw = ba.count ? round(ba.nw / ba.count, 3) : 0;
          // v1.4.47：非 boxMode（纯标签驱动模板未触发 boxMode）时，从第一箱尺寸兜底补 per-item 长/宽/高/单箱重，
          // 让「长(cm)/宽(cm)/高(cm)/单箱重量」等列有值可填（避免整列空白）
          if (!boxMode) {
            var _m = String(ba.firstDims || '').match(/(\d+(?:\.\d+)?)\s*[×xX*]\s*(\d+(?:\.\d+)?)\s*[×xX*]\s*(\d+(?:\.\d+)?)/);
            if (_m) { it.lengthCm = Number(_m[1]); it.widthCm = Number(_m[2]); it.heightCm = Number(_m[3]); }
            it.singleGw = ba.count ? round(ba.gw / ba.count, 3) : 0;
          }
        }
      } else {
        it._weightSource = 'packing';
      }
      totals.qty += it.qty; totals.amount += it.amount; totals.nw += it.nw; totals.gw += it.gw; totals.volumeWeight += (Number(it.volumeWeight) || 0);
    });
    totals.amount = round(totals.amount, 2); totals.nw = round(totals.nw, 3); totals.gw = round(totals.gw, 3); totals.volumeWeight = round(totals.volumeWeight, 3);
    if (packing) {
      totals.boxCount = packing.totals.boxCount;
      totals.volume = packing.totals.volume;
      // 装箱清单总重优先（权威源）
      if (packing.totals.nw) totals.nw = packing.totals.nw;
      if (packing.totals.gw) totals.gw = packing.totals.gw;
    }

    var currency = meta.currency || 'USD';
    // 给收发人补全常用字段默认值，避免模板占位符 unresolved
    function fillPartyDefaults(p) {
      p = p || {};
      ['name', 'company', 'warehouseCode', 'address', 'city', 'state', 'zip', 'country', 'tel', 'email', 'contact', 'taxNo', 'eori'].forEach(function (k) { if (p[k] === undefined) p[k] = ''; });
      return p;
    }
    shipper = fillPartyDefaults(shipper);
    consignee = fillPartyDefaults(consignee);
    notify = notify ? fillPartyDefaults(notify) : { name: 'SAME AS CONSIGNEE', address: '' };
    return {
      kind: opts.kind || 'invoice',
      invoiceNo: meta.invoiceNo || '',
      invoiceDate: meta.invoiceDate || new Date().toISOString().slice(0, 10),
      contractNo: meta.contractNo || '',
      orderNos: orderList.join(', '),
      refId: refId,
      shipper: shipper, consignee: consignee, notify: notify,
      incoterms: meta.incoterms || '',
      paymentTerms: meta.paymentTerms || '',
      transport: meta.transport || 'BY SEA',
      pol: meta.pol || '', pod: meta.pod || '',
      etd: meta.etd || '', vessel: meta.vessel || '',
      containerType: meta.containerType || '', containerQty: meta.containerQty || '',
      freightTerms: meta.freightTerms || '',
      shippingMarks: meta.shippingMarks || 'N/M',
      currency: currency,
      remark: meta.remark || '',
      dangerous: meta.dangerous || 'NON-DANGEROUS / GENERAL CARGO',
      customsType: meta.customsType || '',
      agent: meta.agent || '',
      goodsSummary: meta.goodsSummary || (items.length ? items.map(function (i) { return i.nameEn || i.nameCn; }).filter(function (v, i, a) { return v && a.indexOf(v) === i; }).slice(0, 3).join(', ') : ''),
      items: items,
      totals: totals,
      amountInWords: amountInWords(totals.amount, currency)
    };
  }

  // ---------- 模板扫描 ----------
  /** 扫描明细行上方的表头，返回 {colNumber -> dataField} */
  function scanItemHeader(ws, itemsRow, mergedMaps) {
    var map = {};
    if (!ws || itemsRow <= 1) return map;
    var sub = (mergedMaps && mergedMaps.subordinate) || {};
    // 向上扫最多 3 行，取每个列最近一行的非空表头
    for (var r = Math.max(1, itemsRow - 3); r < itemsRow; r++) {
      var row = ws.getRow(r);
      row.eachCell({ includeEmpty: false }, function (cell, colNumber) {
        if (map[colNumber]) return;
        if (sub[r + ',' + colNumber]) return; // 跳过合并从属格（B36:C36 合并的 C36 列等），避免重复映射字段
        var v = cell.value;
        var s = (v && v.richText) ? v.richText.map(function (t) { return t.text; }).join('') : (typeof v === 'string' ? v : '');
        var f = matchHeaderAlias(s);
        if (f) map[colNumber] = f;
      });
    }
    return map;
  }
  /** 扫描模板占位符，并尝试用表头识别兜底：返回 {fields, itemFields, itemsRow, itemHeaderMap, sheetName} */
  function scanTemplate(wb) {
    var result = { fields: [], itemFields: [], itemsRow: -1, itemHeaderMap: {}, sheetName: '' };
    var ws = wb.worksheets[0];
    if (!ws) return result;
    result.sheetName = ws.name;
    var seen = {}, seenItem = {};
    ws.eachRow({ includeEmpty: false }, function (row, rowNumber) {
      row.eachCell({ includeEmpty: false }, function (cell) {
        var v = cell.value;
        var s = (v && v.richText) ? v.richText.map(function (t) { return t.text; }).join('') : (typeof v === 'string' ? v : '');
        if (!s) return;
        var m; PH_RE.lastIndex = 0;
        while ((m = PH_RE.exec(s))) {
          var p = m[1];
          if (p.indexOf('items.') === 0) {
            if (result.itemsRow === -1) result.itemsRow = rowNumber;
            if (!seenItem[p]) { seenItem[p] = 1; result.itemFields.push(p); }
          } else if (!seen[p]) { seen[p] = 1; result.fields.push(p); }
        }
      });
    });
    // 兜底：没有占位符明细行时，识别真实表头区域（>=3 个可识别列头）
    if (result.itemsRow === -1) {
      var best = { row: -1, count: 0, map: {} };
      ws.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 30) return;
        var map = {}, count = 0;
        row.eachCell({ includeEmpty: false }, function (cell, colNumber) {
          var v = cell.value;
          var s = (v && v.richText) ? v.richText.map(function (t) { return t.text; }).join('') : (typeof v === 'string' ? v : '');
          var f = matchHeaderAlias(s);
          if (f && !map[colNumber]) { map[colNumber] = f; count++; }
        });
        if (count > best.count) best = { row: rowNumber, count: count, map: map };
      });
      if (best.count >= 3) {
        result.itemsRow = best.row + 1;
        result.itemHeaderMap = best.map;
      }
    } else {
      result.itemHeaderMap = scanItemHeader(ws, result.itemsRow);
    }
    return result;
  }

  /** 模板表头字符串读取（模块级，供表头识别填充使用） */
  function _cellStr(cell) {
    if (!cell) return '';
    var v = cell.value;
    if (!v) return '';
    if (v.richText) return v.richText.map(function (t) { return t.text; }).join('');
    if (typeof v === 'string') return v;
    if (v.text !== undefined) return v.text;
    if (v.formula !== undefined) return v.result !== undefined ? String(v.result) : '';
    if (v.sharedString !== undefined) return String(v.sharedString);
    return String(v);
  }

  /** 解析模板表头标签 -> {party, field, line}；非收发人字段标签返回 null */
  var PARTY_RE = [
    { re: /(发件人|发货人|托运人|SHIPPER|SELLER)/i, party: 'shipper' },
    { re: /(收货人|收货|收件人|收件|CONSIGNEE|BUYER)/i, party: 'consignee' },
    { re: /(通知人|NOTIFY)/i, party: 'notify' }
  ];
  function mapHeaderLabel(text) {
    if (!text || /\{\{/.test(text)) return null;
    var party = null;
    for (var i = 0; i < PARTY_RE.length; i++) if (PARTY_RE[i].re.test(text)) { party = PARTY_RE[i].party; break; }
    if (!party) return null;
    var field = null, line = 0;
    if (/(公司|COMPANY|企业)/i.test(text)) field = 'company';
    else if (/(邮箱|EMAIL|邮件|E-?MAIL)/i.test(text)) field = 'email';
    else if (/(电话|手机|TEL|PHONE|MOBILE)/i.test(text)) field = 'tel';
    else if (/(税号|TAX|EORI)/i.test(text)) field = 'taxNo';
    else if (/(国家|国别|COUNTRY)/i.test(text)) field = 'country';
    else if (/(邮编|ZIP|POSTAL?)/i.test(text)) field = 'zip';
    else if (/(省|州|STATE|PROVINCE)/i.test(text)) field = 'state';
    else if (/(城市|CITY)/i.test(text)) field = 'city';
    else if (/地址/i.test(text)) {
      field = 'address';
      var m = text.match(/地址([一二三四五六七八九十\d])/);
      if (m) { var cn = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10 }; line = cn[m[1]] || parseInt(m[1], 10) || 1; }
    }
    else if (/(联系人|CONTACT)/i.test(text)) field = 'contact';
    else if (/(姓名|NAME)/i.test(text)) field = 'name';
    if (!field) return null;
    return { party: party, field: field, line: line };
  }

  /** 通用（非收发人）字段标签 -> 数据路径。用于真实承运商模板里「运输方式 / 起运国 / 客户单号」等
   *  未被 mapHeaderLabel(只认收发人三类) 覆盖的表头字段。path='' 表示已知但本系统无对应字段(跳过)。 */
  var GENERAL_LABEL_RULES = [
    { re: /运输方式/, path: 'transport' },
    { re: /起运(国|港|地)/, path: 'pol' },
    { re: /目的(国|港|地区|地)/, path: 'pod' },
    { re: /(报关|清关)方式/, path: 'customsType' },
    { re: /成交方式/, path: 'incoterms' },
    { re: /(客户订单号|客户单号|订单号|PO\s*NUMBER|PO号|P\.O\.)/i, path: 'orderNos' },
    { re: /参考号/, path: 'refId' },
    { re: /(预计总件数|总包装件数|预计件数|箱数)/, path: 'totals.boxCount' },
    { re: /(预计重量|总毛重|总重)/, path: 'totals.gw' },
    { re: /(预计体积|总体积)/, path: 'totals.volume' },
    { re: /(申报总价值|总申报价值)/, path: 'totals.amount' },
    { re: /(货物品名|商品品名|^品名$)/, path: 'goodsSummary' },
    { re: /备注/, path: 'remark' },
    { re: /(VAT号|EORI)/, path: 'consignee.taxNo' },
    { re: /仓库代码|WAREHOUSE\s*CODE|海外仓代码|仓库编号|FBA代码/, path: 'consignee.warehouseCode' },
    { re: /(揽货渠道|客户渠道|服务渠道|服务$)/, path: '' } // 无对应字段，跳过
  ];

  /** 统一字段标签识别：先试收发人三类(mapHeaderLabel)，再试通用字段词典。
   *  返回 {party, field, path, line, kind} 或 null。 */
  function mapFieldLabel(text) {
    if (!text || /\{\{/.test(text)) return null;
    var p = mapHeaderLabel(text);
    if (p) return { party: p.party, field: p.field, path: p.party + '.' + p.field, line: p.line, kind: 'party' };
    var t = String(text).trim();
    for (var i = 0; i < GENERAL_LABEL_RULES.length; i++) {
      if (GENERAL_LABEL_RULES[i].re.test(t)) {
        var path = GENERAL_LABEL_RULES[i].path;
        if (!path) return null; // 跳过规则：模板有此标签但系统无对应字段
        return { party: null, field: null, path: path, line: 0, kind: 'general' };
      }
    }
    return null;
  }

  /** 扫描模板表头区，生成标签->字段映射(labelMap)，供填充与 UI 编辑。
   *  itemsRowNum：明细表头行号，其之上的表头区才是字段标签所在；传 -1 时默认扫前 25 行。 */
  function buildLabelMap(wb, itemsRowNum) {
    var ws = wb.worksheets[0];
    if (!ws) return [];
    var mergedMaps = buildMergedMaps(ws);
    var map = [];
    var end = (itemsRowNum && itemsRowNum !== -1) ? itemsRowNum - 1 : 25;
    end = Math.min(end, 60);
    var seen = {};
    for (var r = 1; r <= end; r++) {
      var row = ws.getRow(r);
      row.eachCell({ includeEmpty: false }, function (cell, c) {
        // 跳过合并从属格：同一合并块只在其主格处理一次，避免把值填进标签跨度破坏版式
        if (mergedMaps.masterOf[r + ',' + c]) return;
        var s = _cellStr(cell);
        if (!s || /\{\{/.test(s)) return;
        var info = mapFieldLabel(s);
        if (!info) return;
        var key = r + ',' + c;
        if (seen[key]) return;
        seen[key] = 1;
        map.push({ r: r, c: c, label: s, path: info.path, party: info.party, field: info.field, line: info.line, kind: info.kind, resolved: !!info.path });
      });
    }
    return map;
  }

  /** 按标签映射填充表头区字段：对每个 resolved 标签，找到其右侧值格(跳过合并标签跨度/下一个标签)写入。
   *  与 fillHeaderByLabels 互补——它只处理收发人三类，此处处理全部(含收发人，幂等不冲突)。 */
  function fillByFieldLabels(ws, data, labelMap, mergedMaps) {
    if (!labelMap || !labelMap.length) return;
    var labelCols = {};
    labelMap.forEach(function (e) { (labelCols[e.r] = labelCols[e.r] || {})[e.c] = true; });
    var PH_TEST = /\{\{\s*([A-Za-z0-9_.]+)\s*\}\}/;
    labelMap.forEach(function (e) {
      if (!e.resolved || !e.path) return;
      var val = getPath(data, e.path);
      if (val === undefined || val === null || val === '') {
        if (e.party && e.field === 'address' && e.line > 0 && data[e.party] && data[e.party].address) {
          var lines = String(data[e.party].address).split(/\r?\n/);
          val = lines[e.line - 1] || '';
        }
        if (val === undefined || val === null || val === '') return;
      }
      var spanEnd = e.c;
      var mr = mergedMaps.masterRange[e.r + ',' + e.c];
      if (mr) spanEnd = Math.max(spanEnd, mr.right);
      var vStart = spanEnd + 1;
      var target = null;
      for (var cc = vStart; cc <= vStart + 12; cc++) {
        if (labelCols[e.r] && labelCols[e.r][cc]) break; // 撞到下一个标签，停止
        var tc = ws.getRow(e.r).getCell(cc);
        var ts = _cellStr(tc);
        if (ts) { var m = ts.match(PH_TEST); if (m && m[1] === e.path) { target = tc; break; } }
      }
      if (!target && !(labelCols[e.r] && labelCols[e.r][vStart])) target = ws.getRow(e.r).getCell(vStart);
      if (!target) return;
      target.value = (typeof val === 'number') ? val : String(val);
    });
  }

  /** 模板表头识别填充：把发货人/收货人/地址/电话按模板自身表头标签写入对应单元格。
   *  两遍：先清理值列里的样本残留，再按标签写入（避免写入后被清理）。仅写入空单元格或同字段占位符。 */
  function fillHeaderByLabels(ws, data, headerEnd, itemsRowNumArg) {
    headerEnd = headerEnd || 25;
    // 值列：表头区（明细行之上）内包含 {{}} 占位符的列。
    // 关键：明细占位行(itemsRowNum)的 {{items.*}} 绝不能计入——否则其所在列(B/C/D…/I/K)
    // 被误判为「值列」，导致该列上方所有静态版式标签（承运商抬头/贸易术语/船名航次/货描表头…）
    // 在 PASS1 被误清，模板整体走样。仅扫描明细行之上区域。
    var valueCols = {};
    var valueEnd = (itemsRowNumArg && itemsRowNumArg !== -1) ? Math.min(headerEnd, itemsRowNumArg - 1) : headerEnd;
    for (var r = 1; r <= valueEnd; r++) {
      var row = ws.getRow(r);
      row.eachCell({ includeEmpty: true }, function (cell, col) {
        if (/\{\{/.test(_cellStr(cell))) valueCols[col] = true;
      });
    }
    // 合并单元格成员表：key = r + ',' + c；mergedMaster[r,c] = true 表示该单元格属于某个合并范围
    var mergedCell = {};
    (ws.model.merges || []).forEach(function (m) {
      var mm = (typeof m === 'string') ? m.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/) : null;
      if (!mm) return;
      var c1 = _colNum(mm[1]), r1 = parseInt(mm[2], 10), c2 = _colNum(mm[3]), r2 = parseInt(mm[4], 10);
      for (var rr = r1; rr <= r2; rr++) for (var cc = c1; cc <= c2; cc++) mergedCell[rr + ',' + cc] = true;
    });
    // 这些词出现在单元格里时视为「标签/静态文本」，保留不清理。
    // 注意：「地址/电话/邮箱/联系人」等由 mapHeaderLabel 识别保护，不在这里兜底，
    // 否则会把 JW PEI、Noul LLC 等模板样本地址残留也保留下来。
    var KEEP = /(编码|编号|号|库|液体|粉末|危险品|清关|交税|交货|派送|参考|备注|保价|投保|箱数|商品|申报|材质|海关|用途|品牌|型号|英文|中文|品名|数量|重量|长|宽|高|带电|带磁|图片|链接|销售|价格|名称|发件人|收件人|发货人|通知人|收货|装货|卸货|船名|航次|委托|贸易术语|SHIPPER|CONSIGNEE|NOTIFY|BOOKING|INSTRUCTION|PO|NO\.|承运人|运费|吨位|航班|日期|文件|通知|声明|托运|到达|始发|体积|进仓|运单|预留)/i;

    // PASS 1：清理值列里的样本残留（非占位符、非标签、非静态文本）。
    // 合并单元格成员只保留包含标签词的静态文本；具体样本数据（如旧地址、旧公司名）仍清空。
    for (var r1 = 1; r1 <= headerEnd; r1++) {
      var rowA = ws.getRow(r1);
      rowA.eachCell({ includeEmpty: true }, function (cell, col) {
        var s = _cellStr(cell);
        if (/\{\{/.test(s)) return;
        if (mapHeaderLabel(s)) return;
        if (mergedCell[r1 + ',' + col]) return; // 合并单元格（版式锚点文本）一律保留：块式模板(KEAS)的标题/承运商抬头/贸易术语/货描表头都是合并锚点，误清会整体走样；样本数据(如 Aramex 地址)是普通单元格非合并，仍会被下方 valueCols 规则清除。
        if (valueCols[col] && !KEEP.test(s)) cell.value = '';
      });
    }

    // PASS 2：按模板表头标签写入收发人数据
    for (var r2 = 1; r2 <= headerEnd; r2++) {
      var rowB = ws.getRow(r2);
      rowB.eachCell({ includeEmpty: true }, function (cell, col) {
        var info = mapHeaderLabel(_cellStr(cell));
        if (!info) return;
        var party = (data && data[info.party]) || {};
        var raw = party[info.field];
        var val = (info.field === 'address' && info.line > 0 && typeof raw === 'string')
          ? (raw.split(/\r?\n/)[info.line - 1] || '')
          : (raw === undefined || raw === null ? '' : raw);
        if (val === '') return;
        var target = null, isPh = false;
        for (var c = col + 1; c <= col + 12 && c <= ws.columnCount; c++) {
          var tc = rowB.getCell(c); var ts = _cellStr(tc);
          if (ts && /\{\{/.test(ts)) {
            var mp = ts.match(/\{\{\s*([A-Za-z0-9_.]+)\s*\}\}/);
            if (mp && mp[1] === info.party + '.' + info.field) { target = tc; isPh = true; break; }
            continue; // 占位符不匹配，继续找空单元格
          }
          if (!ts) { target = tc; isPh = false; break; }
        }
        if (!target) return;
        if (!isPh && _cellStr(target)) return; // 已有内容则不改写
        target.value = (typeof val === 'number') ? val : String(val);
      });
    }
  }

  /** 解析 Excel 范围字符串，如 A1:B2 -> {top,left,bottom,right,range} */
  function _colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
  function _colLet(n) { var s = ''; while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  function parseMergeRef(m) {
    var mm = (typeof m === 'string') ? m.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/) : null;
    if (!mm) return null;
    return {
      top: parseInt(mm[2], 10), left: _colNum(mm[1]),
      bottom: parseInt(mm[4], 10), right: _colNum(mm[3]),
      range: m
    };
  }
  function makeMergeRef(r) { return _colLet(r.left) + r.top + ':' + _colLet(r.right) + r.bottom; }

  /** 构建合并单元格映射：subordinate[rn,cn] 表示该格是合并区域的从属格；
   *  masterOf[rn,cn] 指向其主格坐标字符串。用于避免重复填充/识别从属格。 */
  function buildMergedMaps(ws) {
    var subordinate = {}, masterOf = {}, masterRange = {};
    (ws.model.merges || []).forEach(function (m) {
      var ref = parseMergeRef(m);
      if (!ref) return;
      var masterKey = ref.top + ',' + ref.left;
      masterRange[masterKey] = ref;
      for (var r = ref.top; r <= ref.bottom; r++) {
        for (var c = ref.left; c <= ref.right; c++) {
          if (r === ref.top && c === ref.left) continue;
          subordinate[r + ',' + c] = true;
          masterOf[r + ',' + c] = masterKey;
        }
      }
    });
    return { subordinate: subordinate, masterOf: masterOf, masterRange: masterRange };
  }

  /** 把源模板 logo 贴回 workbook（预览与导出都能显示）。ws 为目标工作表。
   *  logo = { dataB64, ext, from:{col,row}, to:{col,row} } */
  function addLogo(wb, ws, logo) {
    if (!logo || !logo.dataB64 || !logo.ext) return;
    try {
      var logoBuf;
      if (typeof Buffer !== 'undefined') {
        logoBuf = Buffer.from(logo.dataB64, 'base64');
      } else if (typeof atob !== 'undefined') {
        var binary = atob(logo.dataB64);
        var arr = new Uint8Array(binary.length);
        for (var ii = 0; ii < binary.length; ii++) arr[ii] = binary.charCodeAt(ii);
        logoBuf = arr;
      }
      if (!logoBuf) return;
      var imgId = wb.addImage({ buffer: logoBuf, extension: logo.ext });
      var tl = logo.from || { col: 0, row: 0 };
      var br = logo.to || { col: (logo.from ? logo.from.col + 2 : 2), row: (logo.from ? logo.from.row + 2 : 2) };
      ws.addImage(imgId, { tl: tl, br: br });
    } catch (e) {}
  }

  /** 模板填充：wb已加载的模板workbook，data为buildDocData输出。原地填充。
   *  支持两种模式：① {{items.xxx}} 占位符（老模板） ② 表头识别（物流商真实模板，无占位符也可填）
   *  options.logo = { dataB64, ext, from:{col,row}, to:{col,row} } 可选，用于把源模板 logo 贴回输出 workbook */
  function fillTemplate(wb, data, options) {
    options = options || {};
    var filled = { replaced: [], unresolved: [] };
    var ws = wb.worksheets[0];
    if (!ws) throw new Error('模板无工作表');

    // ① 抓取模板原始样式快照（边框/底色/字体/对齐），用于末尾 1:1 还原，
    //    避免 ExcelJS 写回时空值带样式的单元格丢边框。
    var origStyles = {};
    ws.eachRow({ includeEmpty: true }, function (row, rn) {
      row.eachCell({ includeEmpty: true }, function (cell, cn) {
        var st = cell.style;
        if (st && (st.border || st.fill || st.font || st.alignment)) {
          origStyles[rn] = origStyles[rn] || {};
          origStyles[rn][cn] = JSON.parse(JSON.stringify(st));
        }
      });
    });
    // 同时抓取原始合并单元格，插入明细行后需要按行位移重新计算并应用，
    // 否则下方合并区域会与新插入行重叠或错位，导致内容显示不全。
    var origMerges = (ws.model.merges || []).slice().map(parseMergeRef).filter(Boolean);
    // 同时构建合并从属格映射，避免对从属格重复识别表头/重复写值（如 B36:C36 合并的 C36）。
    var mergedMaps = buildMergedMaps(ws);

    // 行位移记录（仅第 2 段插入明细行会产生下移）；2.5 段改为清空值不删行，不产生位移。
    var splices = [];
    function mapRow(r) {
      var out = r;
      for (var i = 0; i < splices.length; i++) if (splices[i].at <= r) out += splices[i].delta;
      return out;
    }

    function replaceInString(s, ctx) {
      return s.replace(PH_RE, function (all, path) {
        var v = getPath(ctx, path);
        if (v === undefined || v === null) { filled.unresolved.push(path); return ''; }
        filled.replaced.push(path);
        return typeof v === 'number' ? String(v) : String(v);
      });
    }
    function isPureNumericPh(s, ctx) {
      PH_RE.lastIndex = 0;
      var m = PH_RE.exec(s);
      if (m && m[0] === s.trim()) {
        var v = getPath(ctx, m[1]);
        if (typeof v === 'number') return v;
      }
      return null;
    }
    function cellString(cell) {
      var v = cell && cell.value;
      return (v && v.richText) ? v.richText.map(function (t) { return t.text; }).join('') : (typeof v === 'string' ? v : '');
    }

    // 1) 定位明细行：优先占位符，其次表头识别
    var itemsRowNum = -1, itemHeaderMap = {};
    ws.eachRow({ includeEmpty: false }, function (row, rowNumber) {
      if (itemsRowNum !== -1) return;
      row.eachCell({ includeEmpty: false }, function (cell) {
        var s = cellString(cell);
        if (s && /\{\{\s*items\./.test(s)) itemsRowNum = rowNumber;
      });
    });
    if (itemsRowNum !== -1) {
      itemHeaderMap = scanItemHeader(ws, itemsRowNum, mergedMaps);
    } else {
      // 兜底：无占位符时识别真实表头区域
      var best = { row: -1, count: 0, map: {} };
      ws.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 30) return;
        var map = {}, count = 0;
        row.eachCell({ includeEmpty: false }, function (cell, colNumber) {
          if (map[colNumber]) return;
          var f = matchHeaderAlias(cellString(cell));
          if (f) { map[colNumber] = f; count++; }
        });
        if (count > best.count) best = { row: rowNumber, count: count, map: map };
      });
      if (best.count >= 3) { itemsRowNum = best.row + 1; itemHeaderMap = best.map; }
    }

    // 1.5) 模板表头识别填充（发货人/收货人/地址/电话等）：取模板自身表头标签填
    fillHeaderByLabels(ws, data, itemsRowNum === -1 ? 25 : itemsRowNum, itemsRowNum);

    // 1.6) 通用字段标签映射填充（运输方式/起运国/客户单号/总件数 等未被收发人三类覆盖的表头字段）
    //      优先用模板自带的 labelMap（可 UI 编辑），否则现场扫描；对纯 {{}} 占位符模板为空的，no-op。
    var _labelMap = (options && options.labelMap && options.labelMap.length)
      ? options.labelMap
      : buildLabelMap(wb, itemsRowNum);
    fillByFieldLabels(ws, data, _labelMap, mergedMaps);

    // 明细区下界（供 2.5 清理残留行与 3 跳过明细区使用，必须在前面算出）
    var itemEnd = itemsRowNum === -1 ? -1 : itemsRowNum + Math.max((data.items || []).length, 1) - 1;

    // 扫描模板实际明细槽位数：连续行「至少一格有 border」即视为同一明细槽
    //   v1.4.26 修复 KEAS 类「模板已预留 10 行只第 1 行有 items.* 又被插 9 行」bug
    //   - 仅靠 `{{items.*}}` 计数会漏掉「第 1 行有 items.* + 后续 N-1 行仅占位符+同款 border」的常见模板
    //   - 用 border 识别更稳：明细行必有边框（表头/页脚行一般无）
    function hasItemRowBorder(row) {
      var ok = false;
      row.eachCell({ includeEmpty: true }, function (cell) {
        if (cell.border && (cell.border.top || cell.border.bottom || cell.border.left || cell.border.right)) ok = true;
      });
      return ok;
    }
    var templateSlots = 1;
    if (itemsRowNum !== -1) {
      while (true) {
        var nextRow = ws.getRow(itemsRowNum + templateSlots);
        if (!hasItemRowBorder(nextRow)) break;
        templateSlots++;
        if (templateSlots > 50) break; // 防御
      }
    }

    // 2) 明细区展开：先复制模板行样式与占位内容 N-1 次
    var items = data.items || [];
    if (itemsRowNum !== -1 && items.length > 0) {
      var tplRow = ws.getRow(itemsRowNum);
      var tplCells = [];
      tplRow.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        tplCells.push({ col: colNumber, value: cell.value, style: cell.style });
      });
      var slotDelta = items.length - templateSlots; // 正=需插入；负=需删除
      if (slotDelta > 0) {
        // 模板槽位不足 → 插入 (items.length - templateSlots) 行
        var _insertArgs = [];
        for (var _k = 0; _k < slotDelta; _k++) _insertArgs.push([]);
        splices.push({ at: itemsRowNum + 1, delta: slotDelta });
        ws.spliceRows.apply(ws, [itemsRowNum + 1, 0].concat(_insertArgs));
        for (var i = 1; i < slotDelta + 1; i++) {
          var newRow = ws.getRow(itemsRowNum + i);
          newRow.height = tplRow.height;
          tplCells.forEach(function (tc) {
            var c = newRow.getCell(tc.col);
            c.value = tc.value;
            c.style = JSON.parse(JSON.stringify(tc.style || {}));
          });
        }
      } else if (slotDelta < 0) {
        // 模板槽位多于实际 items → 删除多余行（保留带边框整行→上移不影响：V1.4.15 实践已改为「保留行+清残留」，
        //   这里若直接 spliceRows 删 N 行会导致其下 VGM/footer 上移并破坏 2.5 段原样式，故暂保留「不减行」兜底）
        //   为避免插入/删除不对称导致版式错位，强制以 items.length 槽数为准（多删少插）—— 实测多数模板 items 数=模板槽数，无需走此分支。
        // 暂不实现：保持现状多槽位（多余槽位 2.5 段会清掉 {{items.*}} 占位符文本，整行仍带原样式）
      }
      // 逐行填充 items：占位符优先；无占位符但有表头映射的列，按表头写值
      for (var r = 0; r < items.length; r++) {
        var rowObj = ws.getRow(itemsRowNum + r);
        var itR = items[r];
        // v1.4.27 选项A：整行无货物内容（无 nameEn/nameCn/description/sku）时整行留空，
        //   含 B(boxCount)/I(S/O NO)/K(boxNo)/M(SEAL NO) 等均不写，仅清掉残留 {{items.*}} 占位符
        var rowEmpty = !itR || (!itR.nameEn && !itR.nameCn && !itR.description && !itR.sku);
        if (rowEmpty) {
          (function () {
            var rn = itemsRowNum + r;
            rowObj.eachCell({ includeEmpty: true }, function (cell) {
              if (mergedMaps.subordinate[rn + ',' + cell.col]) return;
              var s = cellString(cell);
              if (s && s.indexOf('{{items.') >= 0) cell.value = '';
            });
          })();
          continue;
        }
        var ctx = Object.assign({}, data, { items: itR });
        rowObj.eachCell({ includeEmpty: true }, function (cell) {
          if (mergedMaps.subordinate[(itemsRowNum + r) + ',' + cell.col]) return; // 合并从属格由主格统一显示，不单独写值
          var s = cellString(cell);
          var hasPh = s && s.indexOf('{{') >= 0;
          if (hasPh) {
            var num = isPureNumericPh(s, ctx);
            cell.value = (num !== null) ? num : replaceInString(s, ctx);
          }
          // 表头兜底：该列上方有识别出的表头，且该单元格无占位符（或占位符解析后为空），则按表头写值
          var field = itemHeaderMap[cell.col];
          if (field) {
            var v = getPath(ctx, 'items.' + field);
            if (v === undefined || v === null || v === '') v = getPath(ctx, field);
            if (v !== undefined && v !== null && v !== '') {
              // 如果占位符已经写了有效值，不覆盖；否则按表头写入
              var cur = cellString(cell);
              if (!hasPh || !cur || cur === '') cell.value = (typeof v === 'number') ? v : String(v);
            }
          }
        });
      }
      // 2.1.5) 兜底：模板预留下方明细行（如 KEAS 第 1 行有 items.* 后面 9 行仅 marks 占位符+border）
      //        rowObj.eachCell 不会遍历「无 value 仅有 style」的格，按表头映射显式补写
      if (Object.keys(itemHeaderMap).length > 0) {
        for (var r2 = 0; r2 < items.length; r2++) {
          var rowObj2 = ws.getRow(itemsRowNum + r2);
          var itR2 = items[r2];
          // v1.4.27 选项A：空内容行（无品名）整行留空，不写任何值（占位符已在上层清空）
          if (!itR2 || (!itR2.nameEn && !itR2.nameCn && !itR2.description && !itR2.sku)) continue;
          if (mergedMaps.subordinate[(itemsRowNum + r2) + ',' + 1]) continue; // 合并从属格主格统一显示
          var ctx2 = Object.assign({}, data, { items: itR2 });
          Object.keys(itemHeaderMap).forEach(function (colStr) {
            var col = parseInt(colStr, 10);
            if (mergedMaps.subordinate[(itemsRowNum + r2) + ',' + col]) return;
            var cell2 = rowObj2.getCell(col);
            var cur2 = cellString(cell2);
            if (cur2 && cur2 !== '') return; // 已有内容（占位符已填）跳过
            var field2 = itemHeaderMap[col];
            if (!field2) return;
            var v2 = getPath(ctx2, 'items.' + field2);
            if (v2 === undefined || v2 === null || v2 === '') v2 = getPath(ctx2, field2);
            if (v2 !== undefined && v2 !== null && v2 !== '') {
              cell2.value = (typeof v2 === 'number') ? v2 : String(v2);
            }
          });
        }
      }
      // 2.2) 列宽自适应：按 items 各字段最长值（中文字符按 2 倍宽估算）动态加宽
      //   itemHeaderMap 映射的列，避免长字符溢出覆盖相邻列；wrapText 关掉放末尾（避免 ④ 还原 alignment 被覆盖）。
      Object.keys(itemHeaderMap).forEach(function (colStr) {
        var col = parseInt(colStr, 10);
        var field = itemHeaderMap[col];
        if (!field) return;
        var sample = '';
        items.forEach(function (it) {
          var v = (it && it[field] !== undefined && it[field] !== null) ? it[field] : '';
          var s = String(v);
          var pxS = 0, pxSample = 0;
          for (var kk = 0; kk < s.length; kk++) {
            var ch = s.charCodeAt(kk);
            pxS += ((ch >= 0x4E00 && ch <= 0x9FFF) || (ch >= 0x3000 && ch <= 0x303F) || (ch >= 0xFF00 && ch <= 0xFFEF)) ? 14 : 7;
          }
          for (var kk2 = 0; kk2 < sample.length; kk2++) {
            var ch2 = sample.charCodeAt(kk2);
            pxSample += ((ch2 >= 0x4E00 && ch2 <= 0x9FFF) || (ch2 >= 0x3000 && ch2 <= 0x303F) || (ch2 >= 0xFF00 && ch2 <= 0xFFEF)) ? 14 : 7;
          }
          if (pxS > pxSample) sample = s;
        });
        if (!sample) return;
        var px = 0;
        for (var k = 0; k < sample.length; k++) {
          var ch3 = sample.charCodeAt(k);
          px += ((ch3 >= 0x4E00 && ch3 <= 0x9FFF) || (ch3 >= 0x3000 && ch3 <= 0x303F) || (ch3 >= 0xFF00 && ch3 <= 0xFFEF)) ? 14 : 7;
        }
        var wsCol = ws.getColumn(col);
        var oldW = wsCol.width || 8.43;
        var newW = Math.max(oldW, Math.ceil(px / 7) + 2);
        if (newW > oldW + 0.05) wsCol.width = newW;
      });
    }

    // 2.5) 清理模板多余的明细占位符：itemEnd 之后若仍有 {{items.*}} 残留（模板预设行数>实际条数），
    //      只清空占位符文本、保留带边框的整行（不再 spliceRows 删除行，避免底部行上移导致边框错位丢失）。
    if (itemsRowNum !== -1) {
      for (var rr = itemEnd + 1; rr <= ws.rowCount; rr++) {
        var rrow = ws.getRow(rr);
        rrow.eachCell({ includeEmpty: true }, function (cell) {
          var sv = cellString(cell);
          if (sv && /\{\{\s*items\./.test(sv)) cell.value = '';
        });
      }
    }

    // 3) 其余单元格占位符替换（{{items.*}} 已在第 2 段处理，此处跳过；
    //    但明细区内的非 items 占位符，如跨越明细行的 {{shippingMarks}} 合并主格，需要在此填充）
    ws.eachRow({ includeEmpty: false }, function (row, rowNumber) {
      row.eachCell({ includeEmpty: false }, function (cell) {
        var v = cell.value;
        var s = (v && v.richText) ? v.richText.map(function (t) { return t.text; }).join('') : (typeof v === 'string' ? v : null);
        if (s === null || s.indexOf('{{') < 0) return;
        if (/^\{\{\s*items\./.test(s)) return; // 明细占位符只在明细区展开，跳过
        var num = isPureNumericPh(s, data);
        cell.value = (num !== null) ? num : replaceInString(s, data);
      });
    });

    // ④ 1:1 还原模板样式：按行位移映射，把原始边框/底色/字体/对齐套回对应单元格，
    //    保证导出文件与模板「表格样式」完全一致（含底部边框行不被位移丢失）。
    Object.keys(origStyles).forEach(function (rn) {
      var tr = parseInt(rn, 10);
      var fr = mapRow(tr);
      Object.keys(origStyles[rn]).forEach(function (cn) {
        var cell = ws.getRow(fr).getCell(parseInt(cn, 10));
        cell.style = JSON.parse(JSON.stringify(origStyles[rn][cn]));
      });
    });

    // ⑤ 1:1 还原合并单元格：明细行插入导致下方合并区域需要整体下移；
    //    与明细行同处一行的合并需要复制到每一行插入行；纵向跨越明细行的合并需要扩展。
    if (itemsRowNum !== -1 && origMerges.length) {
      var delta = 0;
      for (var si = 0; si < splices.length; si++) if (splices[si].at === itemsRowNum + 1) { delta = splices[si].delta; break; }
      if (delta > 0) {
        var newMerges = [];
        // 货描表结构合并：以「首条明细行(itemsRowNum)的合并」为列模式
        var pattern = origMerges.filter(function (m) { return m.top === itemsRowNum; });
        origMerges.forEach(function (m) {
          if (m.bottom < itemsRowNum) {
            // 完全在明细区上方（表头/标签区）：原样保留
            newMerges.push({ top: m.top, left: m.left, bottom: m.bottom, right: m.right });
          } else if (m.top > itemEnd) {
            // 完全在明细区下方（页脚/免责声明/签字区）：整体下移 delta 行
            newMerges.push({ top: m.top + delta, left: m.left, bottom: m.bottom + delta, right: m.right });
          }
          // 与明细区相交的合并：不在此直接处理，统一由下方 pattern 复制到每一条明细行
        });
        // 把首条明细行的列合并模式复制到每一条明细行（单行合并），保证货描表每行格式一致；
        // 纵向 2 行合并（如 Aramex 的 I21:K22）也按单行复制，避免第二样例行等重复合并错位/悬空。
        for (var k = 0; k <= delta; k++) {
          pattern.forEach(function (p) {
            newMerges.push({ top: itemsRowNum + k, left: p.left, bottom: itemsRowNum + k, right: p.right });
          });
        }
        // 清空现有合并（ExcelJS 在 spliceRows 后不会自动更新 merge 范围）
        (ws.model.merges || []).slice().forEach(function (m) { try { ws.unMergeCells(m); } catch (e) {} });
        newMerges.forEach(function (m) {
          try { ws.mergeCells(makeMergeRef(m)); } catch (e) {}
        });
      }
    }
    // ⑤.5) 明细行 wrapText 关掉（放在 ④ 还原 origStyles 之后，避免被覆盖）：
    //   源模板 D/F/H 等列常 wrap=true，列宽刚好到边界时会强制把 "Handbag" 折成 "Handba\ng"
    //   显示错位；关 wrap 后长字符溢出右空白列更符合订舱单观感（且第 2.2 段已加宽列避免溢出到相邻内容列）。
    if (itemsRowNum !== -1) {
      for (var rrW = 0; rrW < items.length; rrW++) {
        var rrowW = ws.getRow(itemsRowNum + rrW);
        rrowW.eachCell({ includeEmpty: true }, function (cell) {
          if (cell.alignment && cell.alignment.wrapText) {
            cell.alignment = Object.assign({}, cell.alignment, { wrapText: false });
          }
        });
      }
    }
    // ⑥ 把源模板 logo 贴回输出 workbook（预览与导出都能显示）
    addLogo(wb, ws, options.logo);
    // ⑥.5) 合并从属格继承主格样式：ExcelJS 写回无值单元格时会丢弃其样式，导致合并块从属格
    //      （如 Aramex I21:K22 的 J21/K21/I22/J22/K22）在「写回→重读」后丢失对齐/字体等；
    //       把主格样式显式赋给从属格使其持久化。视觉不变（合并块本就按主格显示）。
    (ws.model.merges || []).forEach(function (mref) {
      var mm = parseMergeRef(mref); if (!mm) return;
      var master = ws.getCell(mm.top, mm.left);
      var ms = master.style; if (!ms) return;
      for (var rr = mm.top; rr <= mm.bottom; rr++) {
        for (var cc = mm.left; cc <= mm.right; cc++) {
          if (rr === mm.top && cc === mm.left) continue;
          var sc = ws.getCell(rr, cc);
          if (!sc.style || JSON.stringify(sc.style) !== JSON.stringify(ms)) {
            try { sc.style = JSON.parse(JSON.stringify(ms)); } catch (e) {}
          }
        }
      }
    });
    return filled;
  }

  // ---------- 内置模板（骨架版，等真实样例后替换） ----------
  function _hdr(ws, cell, text, size, bold) {
    var c = ws.getCell(cell);
    c.value = text;
    c.font = { size: size || 10, bold: !!bold, name: 'Arial' };
  }
  var thin = { style: 'thin' };
  function _box(ws, range) {
    // range like 'A1:H1'
    ws.getCell(range.split(':')[0]).border = { top: thin, left: thin, bottom: thin, right: thin };
  }

  /** 通用商业发票模板（Commercial Invoice + 装箱信息） */
  function makeBuiltinInvoiceTemplate(ExcelJS) {
    var wb = new ExcelJS.Workbook();
    var ws = wb.addWorksheet('INVOICE');
    ws.columns = [
      { width: 6 }, { width: 18 }, { width: 30 }, { width: 14 }, { width: 10 }, { width: 12 }, { width: 14 }, { width: 12 }, { width: 12 }
    ];
    ws.mergeCells('A1:I1'); _hdr(ws, 'A1', 'COMMERCIAL INVOICE', 16, true);
    ws.getCell('A1').alignment = { horizontal: 'center' };
    _hdr(ws, 'A3', 'INVOICE NO.:', 10, true); _hdr(ws, 'C3', '{{invoiceNo}}');
    _hdr(ws, 'F3', 'DATE:', 10, true); _hdr(ws, 'G3', '{{invoiceDate}}');
    _hdr(ws, 'A4', 'CONTRACT NO.:', 10, true); _hdr(ws, 'C4', '{{contractNo}}');
    _hdr(ws, 'F4', 'ORDER NO.:', 10, true); _hdr(ws, 'G4', '{{orderNos}}');
    _hdr(ws, 'A6', 'SHIPPER:', 10, true);
    ws.mergeCells('B6:E6'); _hdr(ws, 'B6', '{{shipper.name}}');
    ws.mergeCells('B7:E7'); _hdr(ws, 'B7', '{{shipper.address}}');
    ws.mergeCells('B8:E8'); _hdr(ws, 'B8', 'TEL: {{shipper.tel}}');
    _hdr(ws, 'F6', 'CONSIGNEE:', 10, true);
    ws.mergeCells('G6:I6'); _hdr(ws, 'G6', '{{consignee.name}}');
    ws.mergeCells('G7:I7'); _hdr(ws, 'G7', '{{consignee.address}}');
    ws.mergeCells('G8:I8'); _hdr(ws, 'G8', 'TEL: {{consignee.tel}}');
    _hdr(ws, 'A9', 'NOTIFY PARTY:', 10, true);
    ws.mergeCells('B9:E9'); _hdr(ws, 'B9', '{{notify.name}}');
    _hdr(ws, 'A11', 'FROM:', 10, true); _hdr(ws, 'B11', '{{pol}}');
    _hdr(ws, 'D11', 'TO:', 10, true); _hdr(ws, 'E11', '{{pod}}');
    _hdr(ws, 'F11', 'BY:', 10, true); _hdr(ws, 'G11', '{{transport}}');
    _hdr(ws, 'A12', 'PRICE TERMS:', 10, true); _hdr(ws, 'B12', '{{incoterms}}');
    _hdr(ws, 'D12', 'PAYMENT:', 10, true); _hdr(ws, 'E12', '{{paymentTerms}}');
    _hdr(ws, 'F12', 'MARKS:', 10, true); _hdr(ws, 'G12', '{{shippingMarks}}');
    // 表头
    var headRow = 14;
    var heads = ['NO.', 'SKU / MODEL', 'DESCRIPTION OF GOODS', 'HS CODE', 'QTY', 'UNIT', 'UNIT PRICE ({{currency}})', 'AMOUNT ({{currency}})', 'N.W.(KG)'];
    heads.forEach(function (h, i) {
      var c = ws.getRow(headRow).getCell(i + 1);
      c.value = h; c.font = { size: 9, bold: true, name: 'Arial' };
      c.border = { top: thin, left: thin, bottom: thin, right: thin };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    // 明细占位行
    var vals = ['{{items.no}}', '{{items.model}}', '{{items.nameEn}}', '{{items.hsCode}}', '{{items.qty}}', '{{items.unit}}', '{{items.price}}', '{{items.amount}}', '{{items.nw}}'];
    vals.forEach(function (v, i) {
      var c = ws.getRow(15).getCell(i + 1);
      c.value = v; c.font = { size: 9, name: 'Arial' };
      c.border = { top: thin, left: thin, bottom: thin, right: thin };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    // 合计
    _hdr(ws, 'A17', 'TOTAL:', 10, true);
    _hdr(ws, 'E17', '{{totals.qty}}', 10, true);
    _hdr(ws, 'H17', '{{totals.amount}}', 10, true);
    _hdr(ws, 'I17', '{{totals.nw}}', 10, true);
    _hdr(ws, 'A18', 'TOTAL PACKAGES:', 10, true); _hdr(ws, 'C18', '{{totals.boxCount}} CTNS');
    _hdr(ws, 'D18', 'G.W.:', 10, true); _hdr(ws, 'E18', '{{totals.gw}} KG');
    _hdr(ws, 'F18', 'VOLUME:', 10, true); _hdr(ws, 'G18', '{{totals.volume}} CBM');
    ws.mergeCells('A19:I19'); _hdr(ws, 'A19', '{{amountInWords}}', 10, true);
    _hdr(ws, 'A21', 'REMARK:', 10, true); ws.mergeCells('B21:I21'); _hdr(ws, 'B21', '{{remark}}');
    ws.mergeCells('F23:I23'); _hdr(ws, 'F23', 'SIGNATURE: ____________________', 10, true);
    return wb;
  }

  /** 通用订舱单模板（BOOKING FORM） */
  function makeBuiltinBookingTemplate(ExcelJS) {
    var wb = new ExcelJS.Workbook();
    var ws = wb.addWorksheet('BOOKING FORM');
    ws.columns = [{ width: 16 }, { width: 26 }, { width: 16 }, { width: 26 }, { width: 14 }, { width: 14 }];
    ws.mergeCells('A1:F1'); _hdr(ws, 'A1', 'SHIPPING BOOKING FORM 订舱委托书', 15, true);
    ws.getCell('A1').alignment = { horizontal: 'center' };
    _hdr(ws, 'A3', 'DATE 日期:', 10, true); _hdr(ws, 'B3', '{{invoiceDate}}');
    _hdr(ws, 'C3', 'REF NO. 委托号:', 10, true); _hdr(ws, 'D3', '{{invoiceNo}}');
    _hdr(ws, 'A5', 'SHIPPER 托运人:', 10, true);
    ws.mergeCells('B5:F5'); _hdr(ws, 'B5', '{{shipper.name}}');
    ws.mergeCells('B6:F6'); _hdr(ws, 'B6', '{{shipper.address}}  TEL: {{shipper.tel}}');
    _hdr(ws, 'A7', 'CONSIGNEE 收货人:', 10, true);
    ws.mergeCells('B7:F7'); _hdr(ws, 'B7', '{{consignee.name}}');
    ws.mergeCells('B8:F8'); _hdr(ws, 'B8', '{{consignee.address}}  TEL: {{consignee.tel}}');
    _hdr(ws, 'A9', 'NOTIFY 通知人:', 10, true);
    ws.mergeCells('B9:F9'); _hdr(ws, 'B9', '{{notify.name}}');
    _hdr(ws, 'A11', 'POL 起运港:', 10, true); _hdr(ws, 'B11', '{{pol}}');
    _hdr(ws, 'C11', 'POD 目的港:', 10, true); _hdr(ws, 'D11', '{{pod}}');
    _hdr(ws, 'A12', 'ETD 船期:', 10, true); _hdr(ws, 'B12', '{{etd}}');
    _hdr(ws, 'C12', 'VESSEL 船名航次:', 10, true); _hdr(ws, 'D12', '{{vessel}}');
    _hdr(ws, 'A13', 'CNTR 柜型柜量:', 10, true); _hdr(ws, 'B13', '{{containerType}} x {{containerQty}}');
    _hdr(ws, 'C13', 'TERMS 贸易条款:', 10, true); _hdr(ws, 'D13', '{{incoterms}}');
    _hdr(ws, 'A14', 'FREIGHT 运费条款:', 10, true); _hdr(ws, 'B14', '{{freightTerms}}');
    _hdr(ws, 'C14', 'AGENT 订舱代理:', 10, true); _hdr(ws, 'D14', '{{agent}}');
    _hdr(ws, 'A16', 'GOODS 品名概述:', 10, true);
    ws.mergeCells('B16:F16'); _hdr(ws, 'B16', '{{goodsSummary}}');
    _hdr(ws, 'A17', 'PACKAGES 总件数:', 10, true); _hdr(ws, 'B17', '{{totals.boxCount}} CTNS');
    _hdr(ws, 'C17', 'G.W. 总毛重:', 10, true); _hdr(ws, 'D17', '{{totals.gw}} KG');
    _hdr(ws, 'E17', 'CBM 总体积:', 10, true); _hdr(ws, 'F17', '{{totals.volume}}');
    _hdr(ws, 'A18', 'MARKS 唛头:', 10, true); ws.mergeCells('B18:F18'); _hdr(ws, 'B18', '{{shippingMarks}}');
    _hdr(ws, 'A19', 'DG 危险品声明:', 10, true); ws.mergeCells('B19:F19'); _hdr(ws, 'B19', '{{dangerous}}');
    _hdr(ws, 'A20', 'CUSTOMS 报关方式:', 10, true); _hdr(ws, 'B20', '{{customsType}}');
    _hdr(ws, 'A21', 'REMARK 备注:', 10, true); ws.mergeCells('B21:F21'); _hdr(ws, 'B21', '{{remark}}');
    ws.mergeCells('D23:F23'); _hdr(ws, 'D23', 'AUTHORIZED SIGNATURE: ____________', 10, true);
    return wb;
  }

  /** 各模板类型的默认必填字段（校验用） */
  var REQUIRED_FIELDS = {
    invoice: [
      { path: 'invoiceNo', label: '发票号' },
      { path: 'invoiceDate', label: '发票日期' },
      { path: 'shipper.name', label: 'SHIPPER名称' },
      { path: 'consignee.name', label: 'CONSIGNEE名称' },
      { path: 'incoterms', label: '贸易条款' },
      { path: 'pol', label: '起运港' },
      { path: 'pod', label: '目的港' },
      { path: 'currency', label: '币种' }
    ],
    booking: [
      { path: 'invoiceNo', label: '委托号' },
      { path: 'shipper.name', label: '托运人' },
      { path: 'consignee.name', label: '收货人' },
      { path: 'pol', label: '起运港' },
      { path: 'pod', label: '目的港' },
      { path: 'freightTerms', label: '运费条款' }
    ],
    packing: [
      { path: 'totals.boxCount', label: '总箱数' },
      { path: 'totals.gw', label: '总毛重' },
      { path: 'items.nameCn', label: '中文品名' }
    ],
    declare: [
      { path: 'items.nameCn', label: '中文品名' },
      { path: 'items.hsCode', label: '海关编码' },
      { path: 'items.qty', label: '申报数量' }
    ]
  };

  return {
    PH_RE: PH_RE,
    buildDocData: buildDocData,
    amountInWords: amountInWords,
    scanTemplate: scanTemplate,
    buildLabelMap: buildLabelMap,
    mapFieldLabel: mapFieldLabel,
    fillByFieldLabels: fillByFieldLabels,
    fillTemplate: fillTemplate,
    addLogo: addLogo,
    makeBuiltinInvoiceTemplate: makeBuiltinInvoiceTemplate,
    makeBuiltinBookingTemplate: makeBuiltinBookingTemplate,
    REQUIRED_FIELDS: REQUIRED_FIELDS
  };
});
