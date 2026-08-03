/* 渲染填充后的 KEAS 工作簿为 HTML 预览（用 ui.js 的 wbToHtml 委托 preview.wbToHtml） */
const path = require('path');
const ExcelJS = require('exceljs');
const fs = require('fs');

// 用 jsdom 模拟 window/document 给 preview.js
const { JSDOM } = require('jsdom');
const dom = new JSDOM('<!doctype html><html><body><div id="x"></div></body></html>');
global.window = dom.window;
global.document = dom.window.document;
global.navigator = dom.window.navigator;
global.HTMLElement = dom.window.HTMLElement;
global.Node = dom.window.Node;

const engine = require(path.join(__dirname, '..', 'js', 'engine.js'));
const preview = require(path.join(__dirname, '..', 'js', 'preview.js'));
console.log('preview keys:', Object.keys(preview));

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(__dirname, '..', 'templates', 'tpl_real_keas_booking.xlsx'));

  // 用户截图里的 10 条货物
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

  // 提取 logo (用 build_real_templates.js 的方式)
  const JSZip = require('jszip');
  const srcBuf = fs.readFileSync(path.join(__dirname, '..', 'templates', 'tpl_real_keas_booking.xlsx'));
  const zip = await JSZip.loadAsync(srcBuf);
  let logo = null;
  const mediaFiles = Object.keys(zip.files).filter(p => p.startsWith('xl/media/'));
  if (mediaFiles.length) {
    const media = mediaFiles[0];
    const ext = media.split('.').pop().toLowerCase();
    const data = await zip.files[media].async('base64');
    logo = { dataB64: data, ext: ext };
  }
  console.log('logo ext:', logo ? logo.ext : 'none', 'size:', logo ? logo.dataB64.length : 0);

  engine.fillTemplate(wb, data, { logo: logo });

  const html = preview.wbToHtml(wb);
  const fullHtml = '<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Arial,sans-serif;padding:20px}</style></head><body>' + html + '</body></html>';
  fs.writeFileSync('templates/_preview_keas_v1424c.html', fullHtml);
  console.log('written: templates/_preview_keas_v1424c.html size:', fullHtml.length);
})();
