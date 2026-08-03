const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const parser = require(path.join(__dirname, '..', 'js', 'parser.js'));

async function main() {
  const file = path.join(__dirname, 'sample_packing.xlsx');
  const buf = fs.readFileSync(file);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength));
  const rows = parser.sheetToRows(wb.worksheets[0]);
  const pk = parser.parsePacking(rows);

  if (pk.totals.boxCount !== 2) throw new Error('箱数期望2，实际' + pk.totals.boxCount);
  if (pk.totals.qty !== 20) throw new Error('总数量期望20，实际' + pk.totals.qty);
  if (pk.totals.gw !== 40.94) throw new Error('总毛重期望40.94，实际' + pk.totals.gw);
  if (pk.totals.nw !== 38.64) throw new Error('总净重期望38.64，实际' + pk.totals.nw);
  if (pk.totals.volume !== 51.48) throw new Error('总体积期望51.48，实际' + pk.totals.volume);

  // 验证尺寸从纸箱规格解析并计算体积
  const b0 = pk.boxes[0];
  if (b0.length !== 58 || b0.width !== 38 || b0.height !== 37) {
    throw new Error('尺寸解析错误：' + JSON.stringify({ l: b0.length, w: b0.width, h: b0.height }));
  }
  if (!b0.boxSpec || b0.boxSpec !== '58*38*37') throw new Error('boxSpec 未保留');

  // 验证款号作为 SKU
  const skus = pk.boxes.map(b => b.sku);
  if (!skus.includes('1C131-2')) throw new Error('未识别款号作为SKU');

  console.log('✅ 装箱清单按实际表头解析通过：箱数=' + pk.totals.boxCount + ', 总数量=' + pk.totals.qty + ', 总体积=' + pk.totals.volume);
}

module.exports = { main };
if (require.main === module) main().catch(e => { console.error('❌ ' + e.message); process.exit(1); });
