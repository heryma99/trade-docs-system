/* 全场景测试共享模块（UMD）：同时供 Node 端 run_tests.js 与浏览器端 selftest.html 调用
 * 调用方负责注入依赖：{ ExcelJS, parser, validator, engine, adapters, log }
 * runAll(deps) 返回 Promise<{ passed, failed, failures }>
 */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else { root.TD = root.TD || {}; root.TD.tests = factory(); }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function runAll(deps) {
    var ExcelJS = deps.ExcelJS, parser = deps.parser, validator = deps.validator,
      engine = deps.engine, adapters = deps.adapters, log = deps.log || function () {};

    var passed = 0, failed = 0;
    var failures = [];
    function T(name, fn) {
      return Promise.resolve().then(fn).then(function () {
        passed++; log('  ✅ ' + name);
      }).catch(function (e) {
        failed++; failures.push({ name: name, msg: e.message });
        log('  ❌ ' + name + ' → ' + e.message);
      });
    }
    function assert(cond, msg) { if (!cond) throw new Error(msg || 'assert failed'); }
    function eq(a, b, msg) { if (a !== b) throw new Error((msg || '') + ' 期望 ' + JSON.stringify(b) + ' 实际 ' + JSON.stringify(a)); }

    // ---------- 构造xlsx工具 ----------
    function makeXlsx(headerRow, dataRows, preRows) {
      var wb = new ExcelJS.Workbook();
      var ws = wb.addWorksheet('Sheet1');
      (preRows || []).forEach(function (r) { ws.addRow(r); });
      ws.addRow(headerRow);
      dataRows.forEach(function (r) { ws.addRow(r); });
      return wb.xlsx.writeBuffer().then(function (buf) {
        var wb2 = new ExcelJS.Workbook();
        return wb2.xlsx.load(buf).then(function () {
          return parser.sheetToRows(wb2.worksheets[0]);
        });
      });
    }

    // ---------- 模拟主数据 ----------
    var declareMap = {
      'SKU-A': { sku: 'SKU-A', nameCn: '塑料杯', nameEn: 'PLASTIC CUP', hsCode: '3924100000', declarePrice: 1.5, material: 'PLASTIC', usage: 'HOME', brand: 'NO BRAND', model: 'SKU-A', nw: 0.1, gw: 0.12, unit: 'PCS', origin: 'CN' },
      'SKU-B': { sku: 'SKU-B', nameEn: 'COTTON T-SHIRT', nameCn: '棉T恤', hsCode: '6109100021', declarePrice: 3.2, material: 'COTTON', usage: 'WEAR', brand: 'NO BRAND', model: 'SKU-B', nw: 0.2, gw: 0.22, unit: 'PCS', origin: 'CN' },
      'SKU#特殊/字符': { sku: 'SKU#特殊/字符', nameCn: '特殊', nameEn: 'SPECIAL', hsCode: '9999999999', declarePrice: 9.9, material: 'MIX', usage: 'X', brand: 'NB', model: 'M1', nw: 0.05, gw: 0.06, unit: 'PCS', origin: 'CN' }
    };
    var shipper = { name: 'SHENZHEN DEMO TRADING CO., LTD.', address: 'SHENZHEN, CHINA', tel: '0755-123456' };
    var consignee = { name: 'US BUYER LLC', address: 'LA, USA', tel: '+1-555' };

    function mkOrder(no, items, extra) { return Object.assign({ orderNo: no, source: 'jst_import', items: items }, extra || {}); }
    function mkPacking(boxes) {
      var orderNos = {}, boxNos = {};
      var qty = 0, nw = 0, gw = 0, vol = 0;
      boxes.forEach(function (b) {
        if (b.orderNo) orderNos[b.orderNo] = 1; if (b.boxNo) boxNos[b.boxNo] = 1;
        qty += b.qty; nw += b.nw || 0; gw += b.gw || 0; vol += b.volume || 0;
      });
      return { boxes: boxes, orderNos: Object.keys(orderNos), totals: { boxCount: Object.keys(boxNos).length || boxes.length, qty: qty, nw: Math.round(nw * 1000) / 1000, gw: Math.round(gw * 1000) / 1000, volume: Math.round(vol * 10000) / 10000 } };
    }

    return (function main() {
      var chain = Promise.resolve();
      function step(fn) { chain = chain.then(fn); return chain; }

      log('\n══════════ 组1: 解析层（L3） ══════════');
      step(function () { return T('1.1 聚水潭订单解析：标准表头+多行合并同单', function () {
        return makeXlsx(['内部单号', '商家编码', '商品名称', '数量', '单价', '客户名称', '国家'],
          [['SO001', 'SKU-A', '塑料杯', 100, 1.5, '客户甲', 'US'], ['SO001', 'SKU-B', '棉T恤', 50, 3.2, '客户甲', 'US'], ['SO002', 'SKU-A', '塑料杯', 30, 1.5, '客户乙', 'DE']])
          .then(function (rows) { var orders = parser.parseJstOrders(rows); eq(orders.length, 2, '订单数'); eq(orders[0].items.length, 2, 'SO001明细数'); eq(orders[0].buyer, '客户甲'); eq(orders[1].items[0].qty, 30); });
      }); });
      step(function () { return T('1.2 聚水潭解析：表头前有标题行（自动检测表头行）', function () {
        return makeXlsx(['订单号', 'SKU', '数量'], [['SO010', 'SKU-A', 10]], [['某公司订单导出报表'], [''], ['导出时间: 2026-07-30']])
          .then(function (rows) { var orders = parser.parseJstOrders(rows); eq(orders.length, 1); eq(orders[0].orderNo, 'SO010'); });
      }); });
      step(function () { return T('1.3 聚水潭解析：单号合并单元格续行沿用', function () {
        return makeXlsx(['内部单号', '商家编码', '数量'], [['SO020', 'SKU-A', 5], ['', 'SKU-B', 8]])
          .then(function (rows) { var orders = parser.parseJstOrders(rows); eq(orders.length, 1); eq(orders[0].items.length, 2); });
      }); });
      step(function () { return T('1.4 聚水潭解析：0数量/空SKU行剔除（源忠实）', function () {
        return makeXlsx(['内部单号', '商家编码', '数量'], [['SO030', 'SKU-A', 10], ['SO030', '', 99], ['SO031', 'SKU-B', 0], ['', '', '']])
          .then(function (rows) { var orders = parser.parseJstOrders(rows); eq(orders.length, 1, '仅SO030有效'); eq(orders[0].items.length, 1); });
      }); });
      step(function () { return T('1.5 聚水潭解析：乱表头文件应报错而非误解析', function () {
        return makeXlsx(['甲', '乙', '丙'], [['a', 'b', 'c']]).then(function (rows) {
          try { parser.parseJstOrders(rows); throw new Error('本应报错'); }
          catch (e) { assert(/表头/.test(e.message), '报错信息应提示表头: ' + e.message); }
        });
      }); });
      step(function () { return T('1.6 装箱清单解析：标准格式+合计行剔除+箱号续行', function () {
        return makeXlsx(['箱号', '聚水潭单号', 'SKU', '数量', '净重', '毛重', '长', '宽', '高'],
          [['CTN001', 'SO001', 'SKU-A', 60, 6, 7.2, 50, 40, 30], ['', '', 'SKU-B', 30, 6, 6.6, 0, 0, 0], ['CTN002', 'SO001', 'SKU-A', 40, 4, 4.8, 50, 40, 30], ['合计', '', 'TOTAL', 0, 16, 18.6, 0, 0, 0]])
          .then(function (rows) { var pk = parser.parsePacking(rows); eq(pk.boxes.length, 3, '有效行数'); eq(pk.boxes[1].boxNo, 'CTN001', '箱号续行'); eq(pk.totals.boxCount, 2, '箱数'); eq(pk.totals.qty, 130); });
      }); });
      step(function () { return T('1.7 装箱清单：空文件报错', function () {
        return makeXlsx(['箱号', 'SKU', '数量'], []).then(function (rows) {
          try { parser.parsePacking(rows); throw new Error('本应报错'); }
          catch (e) { assert(/没有有效数据/.test(e.message), e.message); }
        });
      }); });
      step(function () { return T('1.8 重复导入检测：同内容hash一致，不同内容hash不同', function () {
        return Promise.all([makeXlsx(['箱号', 'SKU', '数量'], [['C1', 'SKU-A', 10]]), makeXlsx(['箱号', 'SKU', '数量'], [['C1', 'SKU-A', 10]]), makeXlsx(['箱号', 'SKU', '数量'], [['C1', 'SKU-A', 11]])])
          .then(function (rs) { eq(parser.hashRows(rs[0]), parser.hashRows(rs[1]), '同内容'); assert(parser.hashRows(rs[0]) !== parser.hashRows(rs[2]), '不同内容hash应不同'); });
      }); });
      step(function () { return T('1.9 英文表头装箱清单（CTN NO/QTY/N.W/G.W）', function () {
        return makeXlsx(['CTN NO', 'SKU', 'QTY', 'N.W', 'G.W'], [['1', 'SKU-A', 20, 2, 2.4]])
          .then(function (rows) { var pk = parser.parsePacking(rows); eq(pk.boxes[0].qty, 20); eq(pk.boxes[0].gw, 2.4); });
      }); });

      log('\n══════════ 组2: 校验引擎（L6） ══════════');
      step(function () { return T('2.1 单号匹配：完全一致通过', function () {
        var r = validator.matchOrderNos(['SO001', 'SO002'], ['SO002', 'SO001']); assert(r.ok);
      }); });
      step(function () { return T('2.2 单号匹配：选了订单但箱单缺失 → 阻断', function () {
        var r = validator.matchOrderNos(['SO001', 'SO002'], ['SO001']); assert(!r.ok); eq(r.missingInPacking[0], 'SO002');
      }); });
      step(function () { return T('2.3 单号匹配：箱单多出未勾选单号 → 阻断', function () {
        var r = validator.matchOrderNos(['SO001'], ['SO001', 'SO099']); assert(!r.ok); eq(r.extraInPacking[0], 'SO099');
      }); });
      step(function () { return T('2.4 SKU数量勾稽：一致通过', function () {
        var orders = [mkOrder('SO001', [{ sku: 'SKU-A', qty: 100 }, { sku: 'SKU-B', qty: 30 }])];
        var pk = mkPacking([{ boxNo: 'C1', orderNo: 'SO001', sku: 'SKU-A', qty: 60 }, { boxNo: 'C2', orderNo: 'SO001', sku: 'SKU-A', qty: 40 }, { boxNo: 'C2', orderNo: 'SO001', sku: 'SKU-B', qty: 30 }]);
        assert(validator.checkSkuQty(orders, pk).ok);
      }); });
      step(function () { return T('2.5 SKU数量勾稽：数量不一致 → 差异表', function () {
        var orders = [mkOrder('SO001', [{ sku: 'SKU-A', qty: 100 }])];
        var pk = mkPacking([{ boxNo: 'C1', orderNo: 'SO001', sku: 'SKU-A', qty: 90 }]);
        var r = validator.checkSkuQty(orders, pk); assert(!r.ok); eq(r.diffs[0].diff, -10);
      }); });
      step(function () { return T('2.6 SKU勾稽：箱单出现订单没有的SKU → 阻断', function () {
        var orders = [mkOrder('SO001', [{ sku: 'SKU-A', qty: 10 }])];
        var pk = mkPacking([{ boxNo: 'C1', orderNo: 'SO001', sku: 'SKU-A', qty: 10 }, { boxNo: 'C1', orderNo: 'SO001', sku: 'SKU-X', qty: 5 }]);
        var r = validator.checkSkuQty(orders, pk); assert(!r.ok); assert(r.diffs.some(function (d) { return d.sku === 'SKU-X' && d.orderQty === 0; }));
      }); });
      step(function () { return T('2.7 多订单合并勾稽：跨单聚合一致通过', function () {
        var orders = [mkOrder('SO001', [{ sku: 'SKU-A', qty: 60 }]), mkOrder('SO002', [{ sku: 'SKU-A', qty: 40 }])];
        var pk = mkPacking([{ boxNo: 'C1', orderNo: 'SO001', sku: 'SKU-A', qty: 60 }, { boxNo: 'C2', orderNo: 'SO002', sku: 'SKU-A', qty: 40 }]);
        assert(validator.matchOrderNos(orders.map(function (o) { return o.orderNo; }), pk.orderNos).ok);
        assert(validator.checkSkuQty(orders, pk).ok);
      }); });
      step(function () { return T('2.8 申报要素反查：齐全通过，缺失阻断', function () {
        var ok = validator.resolveDeclare([{ sku: 'SKU-A', qty: 1 }], declareMap); assert(ok.ok); eq(ok.items[0].hsCode, '3924100000');
        var bad = validator.resolveDeclare([{ sku: 'SKU-UNKNOWN', qty: 1 }], declareMap); assert(!bad.ok); assert(bad.missing[0].lacks.length > 0);
      }); });
      step(function () { return T('2.9 必填字段校验：缺失识别', function () {
        var r = validator.requiredFieldCheck({ invoiceNo: 'INV1', shipper: { name: '' } }, [{ path: 'invoiceNo', label: '发票号' }, { path: 'shipper.name', label: '发货人' }, { path: 'pod', label: '目的港' }]);
        assert(!r.ok); eq(r.missing.length, 2);
      }); });
      step(function () { return T('2.10 状态机：合法流转与越权拦截', function () {
        assert(validator.canTransition('draft', 'validated')); assert(validator.canTransition('validated', 'confirmed'));
        assert(validator.canTransition('confirmed', 'exported')); assert(!validator.canTransition('draft', 'exported'), 'draft不可直接导出');
        assert(!validator.canTransition('draft', 'confirmed'), 'draft不可直接确认'); assert(!validator.canTransition('exported', 'draft'), 'exported不可回退');
      }); });
      step(function () { return T('2.11 综合校验validateDocument：全通过场景', function () {
        var orders = [mkOrder('SO001', [{ sku: 'SKU-A', qty: 100, price: 1.5 }])];
        var pk = mkPacking([{ boxNo: 'C1', orderNo: 'SO001', sku: 'SKU-A', qty: 100, nw: 10, gw: 12 }]);
        var data = engine.buildDocData({ kind: 'invoice', orders: orders, packing: pk, meta: { invoiceNo: 'INV001', incoterms: 'FOB', pol: 'SHENZHEN', pod: 'LA', currency: 'USD' }, shipper: shipper, consignee: consignee, declareMap: declareMap });
        var rep = validator.validateDocument({ kind: 'invoice', orders: orders, packing: pk, data: data, requiredFields: engine.REQUIRED_FIELDS.invoice, declareMap: declareMap });
        assert(rep.ok, JSON.stringify(rep.blocks));
      }); });

      log('\n══════════ 组3: 生成引擎（L5）与数据组装 ══════════');
      step(function () { return T('3.1 buildDocData：金额/数量/大写金额正确', function () {
        var orders = [mkOrder('SO001', [{ sku: 'SKU-A', qty: 100, price: 1.5 }, { sku: 'SKU-B', qty: 50, price: 3.2 }])];
        var data = engine.buildDocData({ kind: 'invoice', orders: orders, meta: { currency: 'USD' }, shipper: shipper, consignee: consignee, declareMap: declareMap });
        eq(data.totals.qty, 150); eq(data.totals.amount, 310, '100*1.5+50*3.2=310'); assert(/THREE HUNDRED AND TEN/.test(data.amountInWords), data.amountInWords);
      }); });
      step(function () { return T('3.2 buildDocData：订单无单价时用申报价兜底', function () {
        var orders = [mkOrder('SO001', [{ sku: 'SKU-A', qty: 10 }])];
        var data = engine.buildDocData({ kind: 'invoice', orders: orders, meta: {}, shipper: shipper, consignee: consignee, declareMap: declareMap });
        eq(data.items[0].price, 1.5, '申报价兜底'); eq(data.items[0]._priceSource, 'master');
      }); });
      step(function () { return T('3.3 buildDocData：装箱清单重量按SKU分摊，总重取箱单权威值', function () {
        var orders = [mkOrder('SO001', [{ sku: 'SKU-A', qty: 100, price: 1 }])];
        var pk = mkPacking([{ boxNo: 'C1', orderNo: 'SO001', sku: 'SKU-A', qty: 100, nw: 10, gw: 12, volume: 0.06 }]);
        var data = engine.buildDocData({ kind: 'invoice', orders: orders, packing: pk, meta: {}, shipper: shipper, consignee: consignee, declareMap: declareMap });
        eq(data.items[0].nw, 10); eq(data.totals.gw, 12); eq(data.totals.boxCount, 1); eq(data.totals.volume, 0.06);
      }); });
      step(function () { return T('3.4 跨订单同SKU同价合并明细', function () {
        var orders = [mkOrder('SO001', [{ sku: 'SKU-A', qty: 60, price: 1.5 }]), mkOrder('SO002', [{ sku: 'SKU-A', qty: 40, price: 1.5 }])];
        var data = engine.buildDocData({ kind: 'invoice', orders: orders, meta: {}, shipper: shipper, consignee: consignee, declareMap: declareMap });
        eq(data.items.length, 1); eq(data.items[0].qty, 100); eq(data.orderNos, 'SO001, SO002');
      }); });
      step(function () { return T('3.5 小数精度：0.1+0.2类精度不漂移', function () {
        var orders = [mkOrder('SO001', [{ sku: 'SKU-A', qty: 3, price: 0.1 }, { sku: 'SKU-B', qty: 1, price: 0.2 }])];
        var data = engine.buildDocData({ kind: 'invoice', orders: orders, meta: {}, shipper: shipper, consignee: consignee, declareMap: declareMap });
        eq(data.totals.amount, 0.5);
      }); });
      step(function () { return T('3.6 内置发票模板生成+占位符扫描', function () {
        var wb = engine.makeBuiltinInvoiceTemplate(ExcelJS); var scan = engine.scanTemplate(wb);
        assert(scan.fields.indexOf('invoiceNo') >= 0, '含invoiceNo'); assert(scan.itemFields.indexOf('items.qty') >= 0, '含items.qty'); assert(scan.itemsRow > 0);
      }); });
      step(function () { return T('3.7 模板填充：表头+明细多行展开+合计正确写入', function () {
        var wb = engine.makeBuiltinInvoiceTemplate(ExcelJS);
        var orders = [mkOrder('SO001', [{ sku: 'SKU-A', qty: 100, price: 1.5 }, { sku: 'SKU-B', qty: 50, price: 3.2 }, { sku: 'SKU#特殊/字符', qty: 7, price: 9.9 }])];
        var data = engine.buildDocData({ kind: 'invoice', orders: orders, meta: { invoiceNo: 'INV-TEST-01', incoterms: 'FOB', pol: 'SHENZHEN', pod: 'LOS ANGELES', currency: 'USD' }, shipper: shipper, consignee: consignee, declareMap: declareMap });
        engine.fillTemplate(wb, data);
        return wb.xlsx.writeBuffer().then(function (buf) {
          var wb2 = new ExcelJS.Workbook(); return wb2.xlsx.load(buf).then(function () {
            var rows = parser.sheetToRows(wb2.worksheets[0]);
            var flat = rows.map(function (r) { return r.map(function (c) { return parser.cellText(c); }).join('|'); }).join('\n');
            assert(flat.indexOf('INV-TEST-01') >= 0, '发票号已填'); assert(flat.indexOf('PLASTIC CUP') >= 0, '明细1品名');
            assert(flat.indexOf('COTTON T-SHIRT') >= 0, '明细2品名（行展开）'); assert(flat.indexOf('SPECIAL') >= 0, '明细3特殊字符SKU');
            assert(flat.indexOf('{{') < 0, '不应残留任何占位符');
            var foundQty = false; wb2.worksheets[0].eachRow(function (r) { r.eachCell(function (c) { if (c.value === 100) foundQty = true; }); });
            assert(foundQty, '数量100应为数字类型');
          });
        });
      }); });
      step(function () { return T('3.8 边界：120行明细模板行复制不丢行', function () {
        var wb = engine.makeBuiltinInvoiceTemplate(ExcelJS); var items = [];
        for (var i = 0; i < 120; i++) items.push({ sku: 'BULK-' + String(i).padStart(3, '0'), qty: i + 1, price: 1 });
        var dm = {}; items.forEach(function (it) { dm[it.sku] = { sku: it.sku, nameEn: 'BULK ITEM ' + it.sku, hsCode: '1111111111', declarePrice: 1, material: 'M', model: it.sku, nw: 0.01, gw: 0.01, unit: 'PCS' }; });
        var orders = [mkOrder('SO-BULK', items)];
        var data = engine.buildDocData({ kind: 'invoice', orders: orders, meta: { invoiceNo: 'INV-BULK' }, shipper: shipper, consignee: consignee, declareMap: dm });
        engine.fillTemplate(wb, data);
        return wb.xlsx.writeBuffer().then(function (buf) {
          var wb2 = new ExcelJS.Workbook(); return wb2.xlsx.load(buf).then(function () {
            var flat = parser.sheetToRows(wb2.worksheets[0]).map(function (r) { return r.map(function (c) { return parser.cellText(c); }).join('|'); }).join('\n');
            assert(flat.indexOf('BULK-000') >= 0 && flat.indexOf('BULK-119') >= 0, '首尾明细都在'); assert(flat.indexOf('TOTAL') >= 0, '合计行仍在（未被覆盖）');
          });
        });
      }); });
      step(function () { return T('3.9 超长品名不报错', function () {
        var wb = engine.makeBuiltinInvoiceTemplate(ExcelJS); var longName = 'VERY LONG PRODUCT NAME '.repeat(20);
        var dm = { 'L1': { sku: 'L1', nameEn: longName, hsCode: '1', declarePrice: 1, material: 'M', model: 'L1', unit: 'PCS' } };
        var data = engine.buildDocData({ kind: 'invoice', orders: [mkOrder('SO-L', [{ sku: 'L1', qty: 1, price: 1 }])], meta: { invoiceNo: 'INV-L' }, shipper: shipper, consignee: consignee, declareMap: dm });
        engine.fillTemplate(wb, data); return wb.xlsx.writeBuffer();
      }); });
      step(function () { return T('3.10 内置订舱单模板生成+填充+回读', function () {
        var wb = engine.makeBuiltinBookingTemplate(ExcelJS);
        var orders = [mkOrder('SO001', [{ sku: 'SKU-A', qty: 100, price: 1.5 }])];
        var pk = mkPacking([{ boxNo: 'C1', orderNo: 'SO001', sku: 'SKU-A', qty: 100, nw: 10, gw: 12, volume: 0.06 }]);
        var data = engine.buildDocData({ kind: 'booking', orders: orders, packing: pk, meta: { invoiceNo: 'BK-001', pol: 'SHENZHEN', pod: 'HAMBURG', etd: '2026-08-15', vessel: 'MSC OSCAR V.123', containerType: '40HQ', containerQty: '1', freightTerms: 'FREIGHT PREPAID', incoterms: 'FOB', agent: 'DEMO FORWARDER' }, shipper: shipper, consignee: consignee, declareMap: declareMap });
        engine.fillTemplate(wb, data);
        return wb.xlsx.writeBuffer().then(function (buf) {
          var wb2 = new ExcelJS.Workbook(); return wb2.xlsx.load(buf).then(function () {
            var flat = parser.sheetToRows(wb2.worksheets[0]).map(function (r) { return r.map(function (c) { return parser.cellText(c); }).join('|'); }).join('\n');
            assert(flat.indexOf('BK-001') >= 0, '委托号'); assert(flat.indexOf('MSC OSCAR V.123') >= 0, '船名航次'); assert(flat.indexOf('40HQ x 1') >= 0, '柜型柜量'); assert(flat.indexOf('12 KG') >= 0, '总毛重'); assert(flat.indexOf('{{') < 0, '无残留占位符');
          });
        });
      }); });
      step(function () { return T('3.11 无占位符模板填充：不崩溃且报告0替换', function () {
        var wb = new ExcelJS.Workbook(); var ws = wb.addWorksheet('S'); ws.getCell('A1').value = '纯静态模板';
        var data = engine.buildDocData({ kind: 'invoice', orders: [mkOrder('S1', [{ sku: 'SKU-A', qty: 1, price: 1 }])], meta: {}, shipper: shipper, consignee: consignee, declareMap: declareMap });
        var r = engine.fillTemplate(wb, data); eq(r.replaced.length, 0);
      }); });

      log('\n══════════ 组4: 数据源适配器（L2）与降级 ══════════');
      step(function () { return T('4.1 远程适配器：正常拉取', function () {
        var fakeFetch = function () { return Promise.resolve({ ok: true, json: function () { return Promise.resolve([{ sku: 'R-1', hsCode: '123', updatedAt: 100 }]); } }); };
        var remote = new adapters.RemoteHttpAdapter({ baseURL: 'http://x' }, fakeFetch);
        return remote.pull('declare_reqs').then(function (rows) { eq(rows.length, 1); eq(rows[0].sku, 'R-1'); });
      }); });
      step(function () { return T('4.2 远程适配器：HTTP 500 → 报错（供上层降级）', function () {
        var fakeFetch = function () { return Promise.resolve({ ok: false, status: 500 }); };
        var remote = new adapters.RemoteHttpAdapter({ baseURL: 'http://x' }, fakeFetch);
        return remote.pull('declare_reqs').then(function () { throw new Error('本应报错'); }).catch(function (e) { assert(/500/.test(e.message)); });
      }); });
      step(function () { return T('4.3 远程不可达 → syncFromRemote降级本地不抛异常', function () {
        var fakeFetch = function () { return Promise.reject(new Error('network down')); };
        var remote = new adapters.RemoteHttpAdapter({ baseURL: 'http://x' }, fakeFetch);
        var fakeDb = { keyOf: function () { return 'sku'; }, bulkPut: function () { return Promise.resolve(0); } };
        return adapters.syncFromRemote(fakeDb, remote, 'declare_reqs').then(function (r) { assert(!r.ok); eq(r.source, 'local'); assert(/降级/.test(r.error)); });
      }); });
      step(function () { return T('4.4 远程返回非数组 → 降级', function () {
        var fakeFetch = function () { return Promise.resolve({ ok: true, json: function () { return Promise.resolve({ bad: 'shape' }); } }); };
        var remote = new adapters.RemoteHttpAdapter({ baseURL: 'http://x' }, fakeFetch);
        var fakeDb = { keyOf: function () { return 'sku'; }, bulkPut: function () { return Promise.resolve(0); } };
        return adapters.syncFromRemote(fakeDb, remote, 'declare_reqs').then(function (r) { assert(!r.ok); });
      }); });
      step(function () { return T('4.5 未配置URL → 明确报错', function () {
        var remote = new adapters.RemoteHttpAdapter({}, function () { return Promise.resolve({ ok: true, json: function () { return Promise.resolve([]); } }); });
        return remote.pull('parties').then(function () { throw new Error('本应报错'); }).catch(function (e) { assert(/未配置/.test(e.message)); });
      }); });
      step(function () { return T('4.6 聚水潭API占位适配器：未启用时明确提示', function () {
        var jst = new adapters.JstApiAdapter({});
        return jst.available().then(function (av) {
          eq(av, false);
          return jst.list('orders').then(function () { throw new Error('本应报错'); }).catch(function (e) { assert(/聚水潭/.test(e.message)); });
        });
      }); });

      log('\n══════════ 组5: 端到端模拟用户全链路 ══════════');
      step(function () { return T('5.1 链路A：聚水潭订单+装箱清单 → 校验 → 发票导出（合并2单一票）', function () {
        return makeXlsx(['内部单号', '商家编码', '数量', '单价'], [['SO100', 'SKU-A', 60, 1.5], ['SO100', 'SKU-B', 20, 3.2], ['SO101', 'SKU-A', 40, 1.5]]).then(function (ordRows) {
          var orders = parser.parseJstOrders(ordRows); eq(orders.length, 2);
          return makeXlsx(['箱号', '聚水潭单号', 'SKU', '数量', '净重', '毛重'], [['C1', 'SO100', 'SKU-A', 60, 6, 7], ['C1', 'SO100', 'SKU-B', 20, 4, 4.4], ['C2', 'SO101', 'SKU-A', 40, 4, 4.8]]).then(function (pkRows) {
            var pk = parser.parsePacking(pkRows);
            var data = engine.buildDocData({ kind: 'invoice', orders: orders, packing: pk, meta: { invoiceNo: 'INV-E2E-A', incoterms: 'CIF', pol: 'SHENZHEN', pod: 'NEW YORK', currency: 'USD' }, shipper: shipper, consignee: consignee, declareMap: declareMap });
            var rep = validator.validateDocument({ kind: 'invoice', orders: orders, packing: pk, data: data, requiredFields: engine.REQUIRED_FIELDS.invoice, declareMap: declareMap });
            assert(rep.ok, '校验应通过: ' + JSON.stringify(rep.blocks));
            var st = 'draft'; assert(validator.canTransition(st, 'validated')); st = 'validated'; assert(validator.canTransition(st, 'confirmed')); st = 'confirmed'; assert(validator.canTransition(st, 'exported'));
            var wb = engine.makeBuiltinInvoiceTemplate(ExcelJS); engine.fillTemplate(wb, data);
            return wb.xlsx.writeBuffer().then(function (buf) {
              assert(buf.byteLength > 3000, '导出文件应为有效xlsx');
              var wb2 = new ExcelJS.Workbook(); return wb2.xlsx.load(buf).then(function () {
                var flat = parser.sheetToRows(wb2.worksheets[0]).map(function (r) { return r.map(function (c) { return parser.cellText(c); }).join('|'); }).join('\n');
                assert(flat.indexOf('SO100, SO101') >= 0, '合并单号'); assert(flat.indexOf('INV-E2E-A') >= 0);
                var total = 0; wb2.worksheets[0].eachRow(function (r) { r.eachCell(function (c) { if (c.value === 214) total = c.value; }); });
                eq(total, 214, '总金额应为214且为数字单元格');
              });
            });
          });
        });
      }); });
      step(function () { return T('5.1b 总金额精确核算（独立复核）', function () {
        var orders = [mkOrder('SO100', [{ sku: 'SKU-A', qty: 60, price: 1.5 }, { sku: 'SKU-B', qty: 20, price: 3.2 }]), mkOrder('SO101', [{ sku: 'SKU-A', qty: 40, price: 1.5 }])];
        var data = engine.buildDocData({ kind: 'invoice', orders: orders, meta: {}, shipper: shipper, consignee: consignee, declareMap: declareMap });
        eq(data.totals.amount, 214);
      }); });
      step(function () { return T('5.2 链路B：装箱清单直生订单 → 发票 → 订舱单（无聚水潭订单）', function () {
        return makeXlsx(['箱号', '订单号', 'SKU', '数量', '净重', '毛重', '体积'], [['C1', 'PK-GEN-01', 'SKU-A', 50, 5, 6, 0.03], ['C2', 'PK-GEN-01', 'SKU-B', 25, 5, 5.5, 0.03]]).then(function (pkRows) {
          var pk = parser.parsePacking(pkRows);
          var groups = {}; pk.boxes.forEach(function (x) { var no = x.orderNo || 'PK-AUTO'; groups[no] = groups[no] || {}; groups[no][x.sku] = groups[no][x.sku] || { sku: x.sku, qty: 0 }; groups[no][x.sku].qty += x.qty; });
          var orders = Object.keys(groups).map(function (no) { return { orderNo: no, source: 'from_packing', items: Object.values(groups[no]) }; });
          eq(orders.length, 1); eq(orders[0].source, 'from_packing');
          assert(validator.matchOrderNos(orders.map(function (o) { return o.orderNo; }), pk.orderNos).ok); assert(validator.checkSkuQty(orders, pk).ok);
          var invData = engine.buildDocData({ kind: 'invoice', orders: orders, packing: pk, meta: { invoiceNo: 'INV-B', incoterms: 'FOB', pol: 'YANTIAN', pod: 'FELIXSTOWE', currency: 'USD' }, shipper: shipper, consignee: consignee, declareMap: declareMap });
          var invRep = validator.validateDocument({ kind: 'invoice', orders: orders, packing: pk, data: invData, requiredFields: engine.REQUIRED_FIELDS.invoice, declareMap: declareMap }); assert(invRep.ok, JSON.stringify(invRep.blocks));
          var bkData = engine.buildDocData({ kind: 'booking', orders: orders, packing: pk, meta: { invoiceNo: 'BK-B', pol: 'YANTIAN', pod: 'FELIXSTOWNE', freightTerms: 'FREIGHT COLLECT' }, shipper: shipper, consignee: consignee, declareMap: declareMap });
          var bkRep = validator.validateDocument({ kind: 'booking', orders: orders, packing: pk, data: bkData, requiredFields: engine.REQUIRED_FIELDS.booking }); assert(bkRep.ok, JSON.stringify(bkRep.blocks));
          var wb = engine.makeBuiltinBookingTemplate(ExcelJS); engine.fillTemplate(wb, bkData); return wb.xlsx.writeBuffer().then(function (buf) { assert(buf.byteLength > 3000); });
        });
      }); });
      step(function () { return T('5.3 链路C（异常流）：单号不匹配被阻断，修正后通过', function () {
        var orders = [mkOrder('SO200', [{ sku: 'SKU-A', qty: 10 }])];
        var wrongPk = mkPacking([{ boxNo: 'C1', orderNo: 'SO999', sku: 'SKU-A', qty: 10 }]);
        var data = engine.buildDocData({ kind: 'invoice', orders: orders, packing: wrongPk, meta: { invoiceNo: 'X', incoterms: 'FOB', pol: 'A', pod: 'B', currency: 'USD' }, shipper: shipper, consignee: consignee, declareMap: declareMap });
        var rep = validator.validateDocument({ kind: 'invoice', orders: orders, packing: wrongPk, data: data, requiredFields: engine.REQUIRED_FIELDS.invoice, declareMap: declareMap });
        assert(!rep.ok, '错箱单必须阻断'); assert(rep.blocks.some(function (b) { return b.type === 'orderNo'; }));
        var rightPk = mkPacking([{ boxNo: 'C1', orderNo: 'SO200', sku: 'SKU-A', qty: 10, nw: 1, gw: 1.2 }]);
        var rep2 = validator.validateDocument({ kind: 'invoice', orders: orders, packing: rightPk, data: data, requiredFields: engine.REQUIRED_FIELDS.invoice, declareMap: declareMap }); assert(rep2.ok);
      }); });
      step(function () { return T('5.4 链路D（异常流）：申报要素缺失阻断 → 补录后通过', function () {
        var orders = [mkOrder('SO300', [{ sku: 'NEW-SKU', qty: 5, price: 2 }])];
        var pk = mkPacking([{ boxNo: 'C1', orderNo: 'SO300', sku: 'NEW-SKU', qty: 5 }]);
        var dm = Object.assign({}, declareMap);
        var rep1 = validator.validateDocument({ kind: 'invoice', orders: orders, packing: pk, data: {}, requiredFields: [], declareMap: dm }); assert(!rep1.ok);
        dm['NEW-SKU'] = { sku: 'NEW-SKU', nameEn: 'NEW', hsCode: '123', declarePrice: 2, material: 'M', model: 'NEW-SKU' };
        var rep2 = validator.validateDocument({ kind: 'invoice', orders: orders, packing: pk, data: {}, requiredFields: [], declareMap: dm }); assert(rep2.ok);
      }); });

      return chain.then(function () {
        log('\n══════════════════════════════════');
        log('结果: ' + passed + ' 通过 / ' + failed + ' 失败 / 共 ' + (passed + failed));
        if (failed) { log('\n失败明细:'); failures.forEach(function (f) { log('  ✗ ' + f.name + ': ' + f.msg); }); }
        else { log('✅ 全部测试通过'); }
        return { passed: passed, failed: failed, failures: failures };
      });
    })();
  }

  return { runAll: runAll };
});
