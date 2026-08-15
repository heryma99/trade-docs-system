/* L3 解析层：表头自动检测 + 别名词典列映射（源忠实：不造数据，剔空行） */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else { root.TD = root.TD || {}; root.TD.parser = factory(); }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------- 别名词典 ----------
  var JST_ALIASES = {
    orderNo:  ['内部单号', '内部订单号', '订单号', '聚水潭单号', '线上单号', 'ERP单号', '单号'],
    onlineNo: ['线上订单号', '网店订单号', '平台单号'],
    sku:      ['商家编码', '商品编码', 'SKU', 'sku', '规格编码', '商家SKU'],
    name:     ['商品名称', '品名', '货品名称', '名称'],
    qty:      ['数量', '下单数量', '销售数量'],
    price:    ['单价', '销售单价', '标准单价'],
    amount:   ['金额', '总金额', '销售金额', '应收金额'],
    buyer:    ['买家昵称', '客户名称', '买家', '客户', '分销商'],
    receiver: ['收货人', '收件人', '收货人姓名'],
    address:  ['收货地址', '地址', '详细地址'],
    country:  ['国家', '国家地区', '收货国家'],
    phone:    ['电话', '手机', '联系电话', '收货人电话'],
    shop:     ['店铺', '店铺名称', 'shop']
  };

  var PACKING_ALIASES = {
    boxNo:   ['箱号', '箱唛号', 'CTN NO', 'CARTON NO', 'Carton No', '箱子编号', 'NO.', '序号箱号'],
    orderNo: ['聚水潭单号', '内部单号', '订单号', '内部订单号', 'ERP单号', '关联单号', '单号'],
    sku:     ['商家编码', 'SKU', 'sku', '商品编码', '货号', '规格编码', '款号'],
    name:    ['商品名称', '品名', '英文品名', '货品名称', 'DESCRIPTION', 'Description'],
    qty:     ['数量', '装箱数量', 'QTY', 'Qty', 'PCS', '件数'],
    nw:      ['净重', '净重KG', 'N.W', 'N.W.', 'NW', 'N.W(KG)', '净重(KG)', '净重（KG）', '净重(kg)', '净重（kg）'],
    gw:      ['重量', '毛重', '毛重KG', 'G.W', 'G.W.', 'GW', 'G.W(KG)', '毛重(KG)', '毛重（KG）', '毛重(kg)', '毛重（kg）'],
    length:  ['长', '长CM', 'L', '长(CM)', '长（CM）', '长cm'],
    width:   ['宽', '宽CM', 'W', '宽(CM)', '宽（CM）', '宽cm'],
    height:  ['高', '高CM', 'H', '高(CM)', '高（CM）', '高cm'],
    volume:  ['体积', '体积CBM', 'CBM', '体积(CBM)', '立方', '方数', '体积重', '体积重(kg)', '体积重（kg）', '体积重KG'],
    boxSpec: ['纸箱规格', '箱规', '外箱规格', '纸箱尺寸', '外箱尺寸'],
    boxWeight: ['包装箱重量（千克）', '包装箱重量', '纸箱重量', '箱皮重', '箱重', 'boxWeight', 'Box Weight', 'Carton Weight', 'Tare Weight']
  };

  // ---------- 工具 ----------
  function cellText(v) {
    if (v === null || v === undefined) return '';
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map(function (t) { return t.text; }).join('');
      if (v.result !== undefined) return cellText(v.result);
      if (v.text !== undefined) return String(v.text);
      if (v instanceof Date) return v.toISOString().slice(0, 10);
      if (v.hyperlink) return String(v.text || v.hyperlink);
      return String(v);
    }
    return String(v);
  }
  function norm(s) { return cellText(s).replace(/\s+/g, '').replace(/[（(].*?[)）]/g, function (m) { return m; }); }
  function toNum(v) {
    var s = cellText(v).replace(/[,，\s]/g, '');
    if (s === '') return 0;
    var n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  /** ExcelJS worksheet → 二维数组（1-based转0-based，值已文本/原样保留） */
  function sheetToRows(ws) {
    var rows = [];
    ws.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      var arr = [];
      row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        arr[colNumber - 1] = cell.value;
      });
      rows[rowNumber - 1] = arr;
    });
    for (var i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
    return rows;
  }

  /** 表头自动检测：扫前maxScan行，命中别名最多的行为表头行 */
  function scanHeader(rows, aliases, minHits, maxScan) {
    minHits = minHits || 2; maxScan = maxScan || 15;
    var best = { headerRow: -1, colMap: {}, hits: 0 };
    var keys = Object.keys(aliases);
    for (var r = 0; r < Math.min(rows.length, maxScan); r++) {
      var colMap = {}, hits = 0;
      for (var c = 0; c < (rows[r] || []).length; c++) {
        var t = norm(rows[r][c]);
        if (!t) continue;
        for (var k = 0; k < keys.length; k++) {
          var key = keys[k];
          if (colMap[key] !== undefined) continue;
          var al = aliases[key];
          for (var a = 0; a < al.length; a++) {
            if (t === al[a].replace(/\s+/g, '') || t.toUpperCase() === al[a].toUpperCase()) {
              colMap[key] = c; hits++; break;
            }
          }
        }
      }
      if (hits > best.hits) best = { headerRow: r, colMap: colMap, hits: hits };
    }
    if (best.hits < minHits) return null;
    return best;
  }

  /** 解析聚水潭订单导出：按 orderNo 聚合，明细行合并 */
  function parseJstOrders(rows) {
    var h = scanHeader(rows, JST_ALIASES, 3);
    if (!h) throw new Error('未识别到聚水潭订单表头（至少需含订单号/商家编码/数量等3列）');
    var cm = h.colMap;
    if (cm.orderNo === undefined) throw new Error('缺少订单号列');
    if (cm.sku === undefined) throw new Error('缺少商家编码/SKU列');
    if (cm.qty === undefined) throw new Error('缺少数量列');
    var map = {}, order;
    var lastNo = ''; // 合并单元格续行沿用上一单号
    for (var r = h.headerRow + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var no = cellText(row[cm.orderNo]).trim();
      var sku = cellText(row[cm.sku]).trim();
      var qty = toNum(row[cm.qty]);
      if (!no && !sku) continue;           // 空行剔除
      if (!no) no = lastNo;                // 续行
      if (!no || !sku || qty <= 0) continue; // 源忠实：无单号/无SKU/0数量剔除
      lastNo = no;
      order = map[no];
      if (!order) {
        order = map[no] = {
          orderNo: no, source: 'jst_import', items: [],
          buyer: cellText(row[cm.buyer] !== undefined ? row[cm.buyer] : '').trim(),
          receiver: cm.receiver !== undefined ? cellText(row[cm.receiver]).trim() : '',
          address: cm.address !== undefined ? cellText(row[cm.address]).trim() : '',
          country: cm.country !== undefined ? cellText(row[cm.country]).trim() : '',
          phone: cm.phone !== undefined ? cellText(row[cm.phone]).trim() : '',
          onlineNo: cm.onlineNo !== undefined ? cellText(row[cm.onlineNo]).trim() : ''
        };
      }
      var item = { sku: sku, qty: qty };
      if (cm.name !== undefined) item.name = cellText(row[cm.name]).trim();
      if (cm.price !== undefined) item.price = toNum(row[cm.price]);
      if (cm.amount !== undefined) item.amount = toNum(row[cm.amount]);
      // 同SKU合并
      var exist = null;
      for (var i = 0; i < order.items.length; i++) if (order.items[i].sku === sku && (order.items[i].price || 0) === (item.price || 0)) { exist = order.items[i]; break; }
      if (exist) { exist.qty += qty; exist.amount = (exist.amount || 0) + (item.amount || 0); }
      else order.items.push(item);
    }
    var list = Object.keys(map).map(function (k) { return map[k]; });
    if (!list.length) throw new Error('文件中没有有效订单行');
    return list;
  }

  /** 从「纸箱规格」字符串提取三个数字（如 58*38*37 / 60*30*40.5） */
  function parseBoxSpec(s) {
    var nums = cellText(s).match(/[\d.]+/g);
    if (!nums || nums.length < 3) return null;
    var vals = [];
    for (var i = 0; i < nums.length && vals.length < 3; i++) {
      var n = parseFloat(nums[i]);
      if (!isNaN(n) && n > 0) vals.push(n);
    }
    if (vals.length < 3) return null;
    return { length: vals[0], width: vals[1], height: vals[2] };
  }

  /** 解析装箱清单：逐箱行，源忠实（不造箱号，空行剔除，箱号续行沿用） */
  function parsePacking(rows, opts) {
    opts = opts || {};
    // v2026-08-09：净重自动算数据源 — 优先级：源文件 nw → 行/箱 boxWeight → box_specs[sku].weight → 固定值
    var boxSpecLib = opts.boxSpecLib || (typeof window !== 'undefined' && window.BOX_SPECS) || {};
    var fixedBoxWeight = +opts.fixedBoxWeight || 0;
    var h = scanHeader(rows, PACKING_ALIASES, 2);
    if (!h) throw new Error('未识别到装箱清单表头（至少需含箱号/SKU/数量等2列）');
    var cm = h.colMap;
    if (cm.sku === undefined) throw new Error('缺少SKU列');
    if (cm.qty === undefined) throw new Error('缺少数量列');
    var boxes = [], lastBox = '', lastNo = '';
    var seenBoxKey = {}; // 净重 fallback 用的「箱首行」标记（同箱 boxWeight 只减一次）
    for (var r = h.headerRow + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var sku = cellText(row[cm.sku]).trim();
      var qty = toNum(row[cm.qty]);
      var boxNo = cm.boxNo !== undefined ? cellText(row[cm.boxNo]).trim() : '';
      var orderNo = cm.orderNo !== undefined ? cellText(row[cm.orderNo]).trim() : '';
      // 合计行/空行剔除
      if (!sku) continue;
      if (/^(合计|总计|TOTAL|Total)/.test(sku)) continue;
      if (qty <= 0) continue;
      if (!boxNo) boxNo = lastBox; else lastBox = boxNo;
      if (!orderNo) orderNo = lastNo; else lastNo = orderNo;
      var b = { boxNo: boxNo, orderNo: orderNo, sku: sku, qty: qty };
      if (cm.name !== undefined) b.name = cellText(row[cm.name]).trim();
      if (cm.nw !== undefined) b.nw = toNum(row[cm.nw]);
      if (cm.gw !== undefined) b.gw = toNum(row[cm.gw]);
      if (cm.boxWeight !== undefined) b.boxWeight = toNum(row[cm.boxWeight]);
      // v2026-08-09：净重自动算 — 同箱首行用 (gw − boxWeight) 算 nw，后续行不重复减（兼容现有 totalNw 按行累加）
      if (!b.nw && b.gw) {
        var boxKey = String(orderNo == null ? '' : orderNo) + '/' + String(boxNo == null ? '' : boxNo).trim();
        if (!seenBoxKey[boxKey]) {
          var bw = b.boxWeight || (boxSpecLib[sku] || {}).weight || fixedBoxWeight;
          if (bw) {
            b.boxWeight = +bw;
            b.nw = Math.max(0, b.gw - b.boxWeight);
            seenBoxKey[boxKey] = 1;
          }
        }
      }
      if (cm.length !== undefined) b.length = toNum(row[cm.length]);
      if (cm.width !== undefined) b.width = toNum(row[cm.width]);
      if (cm.height !== undefined) b.height = toNum(row[cm.height]);
      if (cm.volume !== undefined) b.volume = toNum(row[cm.volume]);
      if (cm.boxSpec !== undefined) b.boxSpec = cellText(row[cm.boxSpec]).trim();
      // 尺寸列缺失时，尝试从「纸箱规格」解析
      if ((!b.length || !b.width || !b.height) && b.boxSpec) {
        var spec = parseBoxSpec(b.boxSpec);
        if (spec) {
          if (!b.length) b.length = spec.length;
          if (!b.width) b.width = spec.width;
          if (!b.height) b.height = spec.height;
        }
      }
      // 体积列缺失时，用尺寸计算 CBM
      if ((b.volume === undefined || b.volume === 0) && b.length && b.width && b.height) {
        b.volume = Math.round(b.length * b.width * b.height / 1000000 * 10000) / 10000;
      }
      boxes.push(b);
    }
    if (!boxes.length) throw new Error('装箱清单中没有有效数据行');
    // v1.4.52：混装去重——毛重/体积/体积重是「箱」级属性，按唯一箱号只计一次，避免同箱多SKU重复累加
    var orderNos = {}, boxNos = {}, totalQty = 0, totalNw = 0, totalGw = 0, totalVol = 0, totalVolW = 0;
    var seenBox = {};
    boxes.forEach(function (b) {
      if (b.orderNo) orderNos[b.orderNo] = 1;
      var boxKey = String(b.orderNo == null ? '' : b.orderNo) + '/' + String(b.boxNo == null ? '' : b.boxNo).trim(); // v1.4.54：复合键，跨订单同箱号(1978182/1 与 1978183/1)不再误合并
      if (b.boxNo) boxNos[boxKey] = 1;
      totalQty += b.qty; totalNw += b.nw || 0;
      if (boxKey !== '/' && seenBox[boxKey]) return;   // 同一(订单/箱号)已在前面计过箱级重量/体积，跳过
      if (boxKey !== '/') seenBox[boxKey] = 1;
      totalGw += b.gw || 0;
      var L = Number(b.length) || 0, W = Number(b.width) || 0, H = Number(b.height) || 0;
      if (L && W && H) { totalVol += L * W * H / 1000000; totalVolW += L * W * H / 6000; }
      else { totalVol += Number(b.volume) || 0; totalVolW += Number(b.volumeWeight) || 0; }
    });
    return {
      boxes: boxes,
      orderNos: Object.keys(orderNos),
      totals: {
        boxCount: Object.keys(boxNos).length || boxes.length,
        qty: totalQty,
        nw: Math.round(totalNw * 1000) / 1000,
        gw: Math.round(totalGw * 1000) / 1000,
        volume: Math.round(totalVol * 10000) / 10000,
        volumeWeight: Math.round(totalVolW * 1000) / 1000
      }
    };
  }

  /** 简单内容hash（防重复导入） */
  function hashRows(rows) {
    var s = '';
    for (var r = 0; r < rows.length; r++) s += (rows[r] || []).map(cellText).join('|') + '\n';
    var h = 5381;
    for (var i = 0; i < s.length; i++) { h = ((h << 5) + h + s.charCodeAt(i)) | 0; }
    return 'h' + (h >>> 0).toString(36) + '_' + s.length;
  }

  return {
    JST_ALIASES: JST_ALIASES,
    PACKING_ALIASES: PACKING_ALIASES,
    cellText: cellText,
    toNum: toNum,
    sheetToRows: sheetToRows,
    scanHeader: scanHeader,
    parseJstOrders: parseJstOrders,
    parsePacking: parsePacking,
    hashRows: hashRows
  };
});
