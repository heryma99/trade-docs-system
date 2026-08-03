/* 验证 GBK 编码 .csv → .xlsx 前端转换链路（用 vendor 里的 SheetJS + ExcelJS） */
const fs = require('fs');
const path = require('path');
const XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
const ExcelJS = require('exceljs');

async function anyToXlsx(arrayBuf, fileName) {
  const fn = fileName || '';
  const isCsv = /\.csv$/i.test(fn);
  const bytes = new Uint8Array(arrayBuf);
  let wb;
  if (isCsv) {
    // 先探测编码：合法 UTF-8 直接用；否则（典型中文 GBK/GB2312）按 gbk 解码，再交给 SheetJS
    let text;
    try {
      text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    } catch (e) {
      text = new TextDecoder('gbk').decode(bytes);
    }
    wb = XLSX.read(text, { type: 'string', raw: true });
  } else {
    wb = XLSX.read(bytes, { type: 'array', cellFormula: false, cellNF: true });
  }
  const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  if (out instanceof ArrayBuffer) return out;
  return out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength);
}

async function main() {
  const files = ['sample.csv', 'sample_utf8.csv'];
  for (const f of files) {
    const buf = fs.readFileSync(path.join(__dirname, f));
    const ab = await anyToXlsx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength), f);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(ab);
    const a1 = wb.worksheets[0].getCell('A1').value;
    const b3 = wb.worksheets[0].getCell('B3').value;
    if (a1 !== '订单号') throw new Error('[' + f + '] A1 期望「订单号」实际 ' + a1);
    if (b3 !== '中文品名XYZ') throw new Error('[' + f + '] B3 期望「中文品名XYZ」实际 ' + b3);
    console.log('✅ ' + f + ' 自动转 .xlsx 并正确识别中文：A1=' + a1 + ', B3=' + b3);
  }
}

module.exports = { main };
if (require.main === module) {
  main().catch(e => { console.error('❌ ' + e.message); process.exit(1); });
}
