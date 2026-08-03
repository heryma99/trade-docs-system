/* 验证 .xls → .xlsx 前端转换链路（用 vendor 里的 SheetJS + ExcelJS） */
const fs = require('fs');
const path = require('path');

// 加载 vendor 中的浏览器版 SheetJS（UMD，Node 也可用）
const XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
const ExcelJS = require('exceljs');

async function xlsToXlsx(arrayBuf) {
  const data = new Uint8Array(arrayBuf);
  const wb = XLSX.read(data, { type: 'array', cellFormula: false, cellNF: true });
  const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  if (out instanceof ArrayBuffer) return out;
  return out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength);
}

async function main() {
  const buf = fs.readFileSync(path.join(__dirname, 'sample.xls'));
  const xlsxBuf = await xlsToXlsx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(xlsxBuf);
  const ws = wb.worksheets[0];
  const a1 = ws.getCell('A1').value;
  const b2 = ws.getCell('B2').value;
  if (a1 !== '订单号') throw new Error('A1 期望「订单号」，实际 ' + a1);
  if (b2 !== 'ABC123') throw new Error('B2 期望 ABC123，实际 ' + b2);
  console.log('✅ .xls 自动转 .xlsx 并读取成功：A1=' + a1 + ', B2=' + b2);
}

module.exports = { main };

if (require.main === module) {
  main().catch(e => { console.error('❌ ' + e.message); process.exit(1); });
}
