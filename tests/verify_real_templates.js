/* 验证生成的真实模板：加载 -> 扫描占位符 -> 试填充，确认无报错且占位符正确 */
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const engine = require('../js/engine.js');

// 读取 real_templates.js 里的数组
const src = fs.readFileSync(path.join(__dirname, '..', 'js', 'real_templates.js'), 'utf8');
const m = src.match(/window\.TD\.realTemplates\s*=\s*(\[[\s\S]*\]);/);
if (!m) { console.error('无法解析 real_templates.js'); process.exit(1); }
const list = JSON.parse(m[1]);

function b64ToBytes(b64) {
  const buf = Buffer.from(b64, 'base64');
  const ab = new Uint8Array(buf.length);
  for (let i = 0; i < buf.length; i++) ab[i] = buf[i];
  return ab;
}

// 构造一个最小 docData 用于试填充
function sampleDoc(kind) {
  return engine.buildDocData({
    orders: [{ orderNo: 'INV-TEST', items: [
      { sku: 'SKU-1', name: 'Test Product', price: 9.9, qty: 3 },
      { sku: 'SKU-2', name: 'Another', price: 5.5, qty: 2 }
    ] }],
    meta: { invoiceNo: 'INV-TEST', invoiceDate: '2026-07-30', pol: 'SHENZHEN', pod: 'LOS ANGELES', customsType: 'CIF', transport: 'BY SEA', incoterms: 'CIF' },
    shipper: { name: 'SHIPPER CO', address: 'ADDR', tel: '123' },
    consignee: { name: 'CONSIGNEE LLC', address: 'ADDR2', tel: '456' },
    declareMap: {}
  });
}

(async () => {
  let fail = 0;
  for (const t of list) {
    try {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(b64ToBytes(t.fileBufB64));
      const scan = engine.scanTemplate(wb);
      // 试填充
      const data = sampleDoc(t.kind);
      const filled = engine.fillTemplate(wb, data);
      const okItems = scan.itemFields.length > 0;
      const status = (scan.fields.length + scan.itemFields.length) > 0 ? 'OK' : 'WARN(无占位符)';
      console.log(`\n[${status}] ${t.id} (${t.kind}) sheet=${wb.worksheets[0].name}`);
      console.log(`  表头字段(${scan.fields.length}): ${scan.fields.slice(0, 8).join(', ')}${scan.fields.length > 8 ? ' ...' : ''}`);
      console.log(`  明细字段(${scan.itemFields.length}): ${scan.itemFields.slice(0, 10).join(', ')}${scan.itemFields.length > 10 ? ' ...' : ''}`);
      console.log(`  填充: 替换${filled.replaced.length} 未解析${filled.unresolved.length}${filled.unresolved.length ? ' -> ' + filled.unresolved.slice(0, 6).join(',') : ''}`);
      if (!okItems) fail++;
    } catch (e) {
      fail++;
      console.log(`\n[FAIL] ${t.id}: ${e.message}`);
    }
  }
  console.log('\n==== 验证完成，失败数=' + fail + ' ====');
  process.exit(fail ? 1 : 0);
})();
