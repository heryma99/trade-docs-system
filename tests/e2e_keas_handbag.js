/* 验证修复：长 nameEn（如 Handbag 7 字符）不再被列宽截断 + wrap 换行 */
const path = require('path');
const ExcelJS = require('exceljs');
const engine = require(path.join(__dirname, '..', 'js', 'engine.js'));

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, '..', 'templates', 'tpl_real_keas_booking.xlsx'));
  const ws = wb.worksheets[0];

  // 用户截图里的实际货物
  const data = {
    invoiceNo: 'BK20260803-006',
    invoiceDate: '2026-08-03',
    shipper: { name: '示例贸易公司 DEMO TRADING CO.,LTD.', address: 'ROOM 101, BUILDING A' },
    consignee: { name: 'JW PEI AP LIMITED', address: 'UNIT 1102 11/F, 29 AUSTIN ROAD, TSIM SHATSUI, KL, HongKong, CHINA' },
    notify: { name: 'DEMO IMPORT LLC', address: 'SAME AS CONSIGNEE' },
    vessel: '',
    pol: 'SHENZHEN, CHINA',
    pod: 'us',
    shippingMarks: '',
    items: [
      { boxCount: 1, nameEn: 'Shoulder Bag', gw: 73.65, volume: 0.3, hsCode: '4202220000', boxNo: '1' },
      { boxCount: 1, nameEn: 'Handbag', gw: 73.65, volume: 0.3, hsCode: '4202220000', boxNo: '2' },
      { boxCount: 1, nameEn: 'Handbag', gw: 73.65, volume: 0.3, hsCode: '4202220000', boxNo: '3' },
      { boxCount: 1, nameEn: 'Hangbag', gw: 73.65, volume: 0.3, hsCode: '4202220000', boxNo: '4' },
      { boxCount: 1, nameEn: 'Hangbag', gw: 73.65, volume: 0.3, hsCode: '4202220000', boxNo: '5' },
      { boxCount: 1, nameEn: 'Hangbag', gw: 73.65, volume: 0.3, hsCode: '4202220000', boxNo: '6' },
      { boxCount: 1, nameEn: 'Hangbag', gw: 73.65, volume: 0.3, hsCode: '4202220000', boxNo: '7' },
      { boxCount: 1, nameEn: 'Hangbag', gw: 73.65, volume: 0.3, hsCode: '4202220000', boxNo: '8' },
      { boxCount: 1, nameEn: 'Hangbag', gw: 73.65, volume: 0.3, hsCode: '4202220000', boxNo: '9' },
      { boxCount: 1, nameEn: 'Hangbag', gw: 73.65, volume: 0.3, hsCode: '4202220000', boxNo: '10' }
    ]
  };

  const filled = engine.fillTemplate(wb, data);
  console.log('unresolved:', filled.unresolved.length, '| replaced:', filled.replaced.length);

  console.log('\n=== column widths after fill ===');
  for (let c = 1; c <= 14; c++) {
    const col = ws.getColumn(c);
    console.log('  col', c, col.letter, 'width=', col.width);
  }

  console.log('\n=== item row D cell wrapText ===');
  for (let r = 36; r <= 45; r++) {
    const cell = ws.getRow(r).getCell(4);
    const v = cell.value;
    const vStr = (v && v.richText) ? v.richText.map(t => t.text).join('') : (typeof v === 'string' ? v : '');
    const wrap = cell.alignment ? cell.alignment.wrapText : '?';
    console.log('  D' + r + ' = "' + vStr + '" wrap=' + wrap);
  }

  await wb.xlsx.writeFile('templates/_e2e_keas_handbag.xlsx');
  console.log('\nwritten: templates/_e2e_keas_handbag.xlsx');
})();
