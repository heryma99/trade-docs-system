/* 用 v1.4.2 的真实转换链路验证真实业务 .xls 文件（不再弹旧错误） */
const fs = require('fs');
const path = require('path');
const XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
const ExcelJS = require('exceljs');

async function xlsToXlsx(arrayBuf) {
  const data = new Uint8Array(arrayBuf);
  const wb = XLSX.read(data, { type: 'array', cellFormula: false, cellNF: true });
  const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  if (out instanceof ArrayBuffer) return out;
  return out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength);
}

const files = [
  'D:\\模板\\订舱单\\DETRANS  BOOKING air -SA.xls',
  'D:\\模板\\订舱单\\BOOKING  FORM.xls',
  'D:\\模板\\订舱单\\CHR Shenzhen Booking Form.xls',
];

(async function () {
  let allOk = true;
  for (const f of files) {
    try {
      const buf = fs.readFileSync(f);
      const ab = await xlsToXlsx(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(ab);
      const ws = wb.worksheets[0];
      const rows = ws.rowCount;
      console.log('✅ ' + path.basename(f) + ' → 转换成功，首表「' + ws.name + '」共 ' + rows + ' 行');
    } catch (e) {
      allOk = false;
      console.error('❌ ' + path.basename(f) + ' → ' + e.message);
    }
  }
  if (!allOk) process.exit(1);
  console.log('\n全部真实 .xls 文件均自动转换并成功读取，不会再弹旧版报错。');
})();
