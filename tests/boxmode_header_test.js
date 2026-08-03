/* 验证 engine.buildDocData(boxMode) + fillTemplate 的：
 *  1) 明细按箱×SKU展开，填充 boxNo/length/width/height/gw
 *  2) 模板表头标签识别：发件人(发货人)段从 shipper 主数据填入空单元格
 *  3) 值列样本残留（Ser hi/US/19808/美国…）被清理
 */
const X = require('exceljs');
const path = require('path');
const engine = require('../js/engine.js');

function cellStr(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (!v) return '';
  if (v.richText) return v.richText.map(t => t.text).join('');
  if (typeof v === 'string') return v;
  if (v.text !== undefined) return v.text;
  if (v.formula !== undefined) return v.result !== undefined ? String(v.result) : '';
  return String(v);
}

(async () => {
  const tplPath = path.resolve(__dirname, '../templates/tpl_real_yafeng_air.xlsx');
  const wb = new X.Workbook();
  await wb.xlsx.readFile(tplPath);

  const declareMap = {
    '104-1': { nameEn: 'Shoulder Bag', nameCn: '单肩包', hsCode: '4202220000', material: 'PU', brand: 'JW PEI', model: '104-1', declarePrice: 16.42, usage: 'Put things' },
    '2T68-1': { nameEn: 'Handbag', nameCn: '手提包', hsCode: '4202220000', material: 'PU', brand: 'JW PEI', model: '2T68-1', declarePrice: 19.73, usage: 'Put things' }
  };
  const orders = [{ orderNo: 'INV20260730-002', items: [
    { sku: '104-1', qty: 5, price: 16.42 },
    { sku: '2T68-1', qty: 8, price: 19.73 }
  ] }];
  const packing = {
    boxes: [
      { boxNo: 'BOX1', sku: '104-1', qty: 5, gw: 73.65, nw: 60, length: 58, width: 38, height: 37 },
      { boxNo: 'BOX2', sku: '2T68-1', qty: 8, gw: 73.65, nw: 60, length: 58, width: 38, height: 37 }
    ],
    totals: { boxCount: 2, nw: 120, gw: 147.3, volume: 0.16 }
  };
  const shipper = { name: 'SHENZHEN XXX CO LTD', company: 'SHENZHEN XXX CO LTD', address: 'No.1 Industrial Rd\nNanshan District\nShenzhen', city: 'Shenzhen', state: 'Guangdong', zip: '518000', tel: '+86-755-88888888', contact: 'Mr Li', email: 'sales@xxx.com', country: 'CN', taxNo: '91440300XXXXXXXX' };
  const consignee = { name: 'JW PEI AP LIMITED', address: '100 DEMO STREET, LOS ANGELES, CA, USA', tel: '+1-000-000-0000', contact: '', city: 'LOS ANGELES', zip: '90001', email: 'us@jwpei.com', country: 'US', taxNo: '' };

  const data = engine.buildDocData({ kind: 'invoice', orders, packing, meta: { invoiceNo: 'INV20260730-002' }, boxMode: true, shipper, consignee, notify: null, declareMap });
  console.log('=== buildDocData (boxMode) ===');
  console.log('items.length =', data.items.length);
  console.log('first item:', JSON.stringify(data.items[0]));
  console.log('totals:', JSON.stringify(data.totals));

  const fill = engine.fillTemplate(wb, data);
  console.log('\n=== fillTemplate unresolved (应为空或仅可选字段) ===', fill.unresolved);

  const ws = wb.worksheets[0];
  console.log('\n=== 表头区关键单元格 ===');
  const check = [
    [1, 1], [2, 1], [4, 1], [5, 1], [6, 1], [9, 1], [11, 1], [13, 1], [14, 1],
    [5, 9], [4, 9], [6, 9], [7, 9], [8, 9], [9, 9], [10, 9], [11, 9], [12, 9], [13, 9], [14, 9]
  ];
  check.forEach(([r, c]) => console.log(`R${r}C${c} =`, JSON.stringify(cellStr(ws.getRow(r).getCell(c)))));
  console.log('\n=== 样本残留检查（应为空）===');
  [[5, 2], [5, 3], [5, 4], [2, 2], [10, 2], [11, 2], [12, 2], [16, 2]].forEach(([r, c]) => console.log(`R${r}C${c} =`, JSON.stringify(cellStr(ws.getRow(r).getCell(c)))));
  console.log('\n=== 明细首行（箱号/尺寸/毛重）===');
  for (let r = 19; r <= 20; r++) {
    const row = ws.getRow(r);
    console.log(`R${r}: boxNo=${cellStr(row.getCell(1))} gw=${cellStr(row.getCell(2))} len=${cellStr(row.getCell(3))} w=${cellStr(row.getCell(4))} h=${cellStr(row.getCell(5))} qty=${cellStr(row.getCell(9))} sku=${cellStr(row.getCell(24))}`);
  }
  // 断言
  let ok = true;
  function assert(c, m) { if (!c) { ok = false; console.log('❌ FAIL:', m); } else console.log('✅', m); }
  assert(data.items.length === 2, '按箱展开 = 2 行');
  assert(cellStr(ws.getRow(19).getCell(1)) === 'BOX1', 'R19 箱号=BOX1');
  assert(cellStr(ws.getRow(19).getCell(3)) === '58', 'R19 长=58');
  assert(cellStr(ws.getRow(4).getCell(10)) === 'SHENZHEN XXX CO LTD', '发件人公司(R4C9)→R4C10 填入');
  assert(cellStr(ws.getRow(12).getCell(10)) === '+86-755-88888888', '发件人电话(R12C9)→R12C10 填入');
  assert(cellStr(ws.getRow(5).getCell(10)) === 'No.1 Industrial Rd', '发件人地址一(R5C9)→R5C10 填入');
  assert(cellStr(ws.getRow(8).getCell(10)) === 'Shenzhen', '发件人城市(R8C9)→R8C10');
  assert(cellStr(ws.getRow(11).getCell(10)) === 'CN', '发件人国家代码(R11C9)→R11C10');
  assert(cellStr(ws.getRow(5).getCell(2)) !== 'Ser hi', '样本 Ser hi 已清');
  assert(cellStr(ws.getRow(12).getCell(2)) !== '美国', '样本 美国 已清');
  console.log('\n' + (ok ? '🎉 ALL PASS' : '⚠️ SOME FAIL'));
  process.exit(ok ? 0 : 1);
})().catch(e => { console.error('ERR', e); process.exit(2); });
