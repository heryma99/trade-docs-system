const path = require('path');
const ExcelJS = require('exceljs');
const fs = require('fs');
const engine = require(path.join(__dirname, '..', 'js', 'engine.js'));
const preview = require(path.join(__dirname, '..', 'js', 'preview.js'));

(async () => {
  // 从 real_templates.js 拿 KEAS logo
  const fs2 = require('fs');
  const vm = require('vm');
  const src2 = fs2.readFileSync(path.join(__dirname, '..', 'js', 'real_templates.js'), 'utf8');
  const sb = { window: {} };
  vm.createContext(sb);
  vm.runInContext(src2, sb);
  let logo = null;
  const list = sb.window.TD.realTemplates;
  const keas = list.find(t => /keas/i.test(t.id) || /keas/i.test(t.name));
  if (keas && keas.logo) {
    logo = keas.logo;
    console.log('KEAS logo ext:', keas.logo.ext, 'dataB64 len:', keas.logo.dataB64.length);
  }

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, '..', 'templates', 'tpl_real_keas_booking.xlsx'));

  const data = {
    invoiceNo: 'BK20260803-006',
    invoiceDate: '2026-08-03',
    shipper: { name: '示例贸易公司 DEMO TRADING CO.,LTD.', address: 'ROOM 101, BUILDING A' },
    consignee: { name: 'JW PEI AP LIMITED', address: 'UNIT 1102 11/F, 29 AUSTIN ROAD, TSIM SHATSUI, KL, HongKong, CHINA' },
    notify: { name: 'DEMO IMPORT LLC', address: 'SAME AS CONSIGNEE' },
    vessel: '', pol: 'SHENZHEN, CHINA', pod: 'us', shippingMarks: '',
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

  engine.fillTemplate(wb, data, { logo: logo });
  const html = preview.wbToHtml(wb);
  const fullHtml = '<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Arial,sans-serif;padding:20px}</style></head><body>' + html + '</body></html>';
  fs.writeFileSync('templates/_preview_keas_v1424c.html', fullHtml);
  console.log('written size:', fullHtml.length);
})();
