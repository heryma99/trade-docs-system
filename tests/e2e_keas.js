/* 端到端验证：KEAS 订舱单模板填充后是否 1:1 还原（不破坏区块式占位符） */
const path = require('path');
const ExcelJS = require('exceljs');
const engine = require(path.join(__dirname, '..', 'js', 'engine.js'));

function _cellStr(cell) {
  var v = cell && cell.value;
  if (!v) return '';
  if (v.richText) return v.richText.map(t => t.text).join('');
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return String(v);
  return '';
}

async function main() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, '..', 'templates', 'tpl_real_keas_booking.xlsx'));
  const ws = wb.worksheets[0];

  const data = {
    invoiceNo: 'KEAS-2026-0001',
    invoiceDate: '2026-07-31',
    shipper: { name: 'SHENZHEN HEROMA TRADING CO., LTD.', address: 'ROOM 101, BUILDING A, NO.1 NANHAI RD, SHENZHEN, GUANGDONG, CHINA' },
    consignee: { name: 'ABC LOGISTICS LLC', address: '123 MAIN STREET, LOS ANGELES, CA 90001, USA' },
    notify: { name: 'SAME AS CONSIGNEE', address: '' },
    vessel: 'EVER GIVEN V.123',
    pol: 'YANTIAN, SHENZHEN',
    pod: 'LOS ANGELES',
    shippingMarks: 'N/M',
    items: [
      { boxCount: 5, nameEn: 'LED LIGHT', gw: 120.5, volume: 0.6, hsCode: '940540', boxNo: '1-5' },
      { boxCount: 3, nameEn: 'PLASTIC TOY', gw: 80.2, volume: 0.4, hsCode: '950300', boxNo: '6-8' },
      { boxCount: 2, nameEn: 'CERAMIC MUG', gw: 50.0, volume: 0.3, hsCode: '691200', boxNo: '9-10' }
    ]
  };

  let filled;
  try {
    filled = engine.fillTemplate(wb, data);
  } catch (e) {
    console.log('FILL ERROR:', e && e.stack || e);
    process.exit(1);
  }

  console.log('=== unresolved placeholders ===');
  console.log(JSON.stringify(filled.unresolved));
  console.log('=== replaced count ===', filled.replaced.length);

  // 输出文件
  const out = path.join('D:/WB文件/2026-07-30-09-36-10/贸易单证系统/templates/_e2e_keas_filled.xlsx');
  await wb.xlsx.writeFile(out);
  console.log('=== written ===', out);

  // 打印关键单元格
  const keys = ['A2','A3','A9','A10','A16','A17','I2','K2','I4','K4','A22','A23','A25','A26','A28','A29','A31','A32','I32','H13','H14','H15','H16','H17','H22','H25','H28'];
  console.log('=== header cells ===');
  keys.forEach(k => {
    const v = _cellStr(ws.getCell(k));
    console.log(k + ' = ' + JSON.stringify(v.replace(/\n/g, ' / ')));
  });

  console.log('=== detail rows ===');
  for (let r = 36; r <= 38; r++) {
    const row = ws.getRow(r);
    const cells = [];
    for (let c = 1; c <= 13; c++) {
      const v = _cellStr(row.getCell(c));
      if (v) cells.push(_colLet(c) + v.replace(/\n/g, ' / '));
    }
    console.log('R' + r + ': ' + cells.join(' | '));
  }

  console.log('=== merges (' + (ws.model.merges || []).length + ') ===');
  (ws.model.merges || []).forEach(m => console.log('  ' + m));

  function _colLet(n){let s='';while(n>0){const r=(n-1)%26;s=String.fromCharCode(65+r)+s;n=Math.floor((n-1)/26);}return s;}
}
main();
