/* 真实业务模板测试（仅 Node 端）：校验 9 个由 build_real_templates.js 生成的模板
 * 能被 ExcelJS 加载、scanTemplate 扫描到占位符、fillTemplate 试填充不报错。
 * 不放入 shared_tests.js（后者被浏览器 selftest.html 复用，不能用 fs）。 */
'use strict';
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const engine = require(path.join(__dirname, '..', 'js', 'engine.js'));

function b64ToBytes(b64) {
  const buf = Buffer.from(b64, 'base64');
  const ab = new Uint8Array(buf.length);
  for (let i = 0; i < buf.length; i++) ab[i] = buf[i];
  return ab;
}

function runAll() {
  return new Promise(function (resolve) {
    const results = [];
    function assert(cond, name) { results.push({ ok: !!cond, name: name }); }
    let fail = 0;
    const listSrc = fs.readFileSync(path.join(__dirname, '..', 'js', 'real_templates.js'), 'utf8');
    const m = listSrc.match(/window\.TD\.realTemplates\s*=\s*(\[[\s\S]*\]);/);
    assert(!!m, 'real_templates.js 可解析');
    if (!m) { return resolve({ failed: 1, count: results.length, results: results }); }
    const list = JSON.parse(m[1]);
    assert(list.length === 15, '共 15 个真实模板 (实际 ' + list.length + ')');

    // 申报信息表镜像（飞书双表合并）
    global.window = global.window || {};
    require(path.join(__dirname, '..', 'js', 'declare_data.js'));
    const dd = (global.window.TD && global.window.TD.declareData) || [];
    assert(dd.length >= 10000, '申报信息表镜像条数 >= 10000 (实际 ' + dd.length + ')');
    const withHs = dd.filter(function (r) { return r.hsCode; }).length;
    assert(withHs > dd.length * 0.9, '申报信息表海关编码覆盖率 > 90% (' + withHs + '/' + dd.length + ')');

    const sampleDoc = function () {
      return engine.buildDocData({
        orders: [{ orderNo: 'INV-TEST', items: [
          { sku: 'SKU-1', name: 'Test', price: 9.9, qty: 3 },
          { sku: 'SKU-2', name: 'Another', price: 5.5, qty: 2 }
        ] }],
        meta: { invoiceNo: 'INV-TEST', invoiceDate: '2026-07-30', pol: 'SHENZHEN', pod: 'LA', customsType: 'CIF', transport: 'BY SEA', incoterms: 'CIF' },
        shipper: { name: 'SHIPPER', address: 'A', tel: '1' },
        consignee: { name: 'CONSIGNEE', address: 'B', tel: '2' },
        declareMap: {}
      });
    };

    const KINDS = { invoice: 4, booking: 8, packing: 1, declare: 2 };
    const kindCount = {};
    list.forEach(function (t) { kindCount[t.kind] = (kindCount[t.kind] || 0) + 1; });
    Object.keys(KINDS).forEach(function (k) {
      assert(kindCount[k] === KINDS[k], '类型统计 ' + k + ' = ' + KINDS[k] + ' (实际 ' + (kindCount[k] || 0) + ')');
    });

    let done = 0;
    list.forEach(function (t) {
      const wb = new ExcelJS.Workbook();
      wb.xlsx.load(b64ToBytes(t.fileBufB64)).then(function () {
        const scan = engine.scanTemplate(wb);
        const total = scan.fields.length + scan.itemFields.length;
        assert(total > 0, t.id + ' 扫描到占位符 (' + scan.fields.length + '+' + scan.itemFields.length + ')');
        assert(scan.itemFields.length > 0, t.id + ' 含明细占位符');
        try {
          engine.fillTemplate(wb, sampleDoc());
          assert(true, t.id + ' 试填充无异常');
        } catch (e) {
          assert(false, t.id + ' 试填充异常: ' + e.message);
        }
        done++;
        if (done === list.length) {
          fail = results.filter(function (r) { return !r.ok; }).length;
          resolve({ failed: fail, count: results.length, results: results });
        }
      }).catch(function (e) {
        assert(false, t.id + ' 加载失败: ' + e.message);
        done++;
        if (done === list.length) {
          fail = results.filter(function (r) { return !r.ok; }).length;
          resolve({ failed: fail, count: results.length, results: results });
        }
      });
    });
  });
}

module.exports = { runAll: runAll };
