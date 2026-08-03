/* 测试：模板无占位符但带真实表头时，仍能按表头填充明细 */
const path = require('path');
const ExcelJS = require('exceljs');
const engine = require(path.join(__dirname, '..', 'js', 'engine.js'));

function cellText(v) {
  if (v && v.richText) return v.richText.map(t => t.text).join('');
  return v === undefined || v === null ? '' : String(v);
}

async function buildHeaderOnlyTemplate() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('INVOICE');
  ws.getCell('A1').value = 'COMMERCIAL INVOICE';
  ws.getCell('A3').value = 'INVOICE NO.:';
  ws.getCell('B3').value = '{{invoiceNo}}';
  // 表头行（无占位符）
  const headers = ['NO.', 'SKU', 'DESCRIPTION', 'QTY', 'UNIT PRICE', 'AMOUNT'];
  headers.forEach((h, i) => { ws.getRow(5).getCell(i + 1).value = h; });
  // 空数据行（模拟物流商模板只有表头+一行示例空行）
  headers.forEach((h, i) => { ws.getRow(6).getCell(i + 1).value = ''; });
  return wb;
}

async function buildMixedTemplate() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('INVOICE');
  const headers = ['序号', '款号', '中文品名', '数量', 'HS编码', '材质'];
  headers.forEach((h, i) => { ws.getRow(1).getCell(i + 1).value = h; });
  const placeholders = ['{{items.no}}', '{{items.sku}}', '{{items.nameCn}}', '{{items.qty}}', '{{items.hsCode}}', '{{items.material}}'];
  placeholders.forEach((p, i) => { ws.getRow(2).getCell(i + 1).value = p; });
  return wb;
}

async function main() {
  // 1) 纯表头模板
  let wb = await buildHeaderOnlyTemplate();
  const data = {
    invoiceNo: 'INV001',
    items: [
      { no: 1, sku: 'SKU-A', nameEn: 'LED Light', qty: 100, price: 5.5, amount: 550 },
      { no: 2, sku: 'SKU-B', nameEn: 'Cable', qty: 200, price: 1.2, amount: 240 }
    ],
    totals: { qty: 300, amount: 790 }
  };
  let res = engine.fillTemplate(wb, data);
  let ws = wb.worksheets[0];
  if (cellText(ws.getCell('B3').value) !== 'INV001') throw new Error('表头占位符未填充');
  if (Number(ws.getCell('A6').value) !== 1) throw new Error('NO.未按表头填充，实际=' + ws.getCell('A6').value);
  if (cellText(ws.getCell('B6').value) !== 'SKU-A') throw new Error('SKU未按表头填充');
  if (Number(ws.getCell('D6').value) !== 100) throw new Error('QTY未按表头填充');
  if (Number(ws.getCell('E6').value) !== 5.5) throw new Error('UNIT PRICE未按表头填充，实际=' + ws.getCell('E6').value);
  if (Number(ws.getCell('F6').value) !== 550) throw new Error('AMOUNT未按表头填充');
  if (Number(ws.getCell('A7').value) !== 2) throw new Error('第二行NO.未填充');
  console.log('✅ 纯表头模板自动识别并填充明细：' + data.items.length + ' 行');

  // 2) 表头+占位符混合模板
  wb = await buildMixedTemplate();
  res = engine.fillTemplate(wb, data);
  ws = wb.worksheets[0];
  if (cellText(ws.getCell('B2').value) !== 'SKU-A') throw new Error('款号列占位符未填充');
  if (Number(ws.getCell('D2').value) !== 100) throw new Error('数量列未填充');
  console.log('✅ 表头+占位符混合模板填充正常');

  // 3) scanTemplate 应返回 itemHeaderMap
  wb = await buildHeaderOnlyTemplate();
  const scan = engine.scanTemplate(wb);
  if (scan.itemsRow !== 6) throw new Error('未识别到纯表头模板的明细行，实际=' + scan.itemsRow);
  if (!scan.itemHeaderMap['2']) throw new Error('未建立 SKU 列映射');
  if (scan.itemHeaderMap['2'] !== 'sku') throw new Error('SKU 列映射错误');
  console.log('✅ scanTemplate 返回表头映射：' + JSON.stringify(scan.itemHeaderMap));
}

if (require.main === module) {
  main().catch(e => { console.error('❌ ' + e.message); process.exit(1); });
}
module.exports = { main };
