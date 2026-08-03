/* 端到端模拟：模板 logo 是否真正出现在预览 HTML 中 */
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const engine = require(path.join(__dirname, '..', 'js', 'engine.js'));
const preview = require(path.join(__dirname, '..', 'js', 'preview.js'));

// 读 real_templates.js 找 KEAS 入口（手动解析，因为 real_templates 是 window.TD = {...}）
const rtPath = path.join(__dirname, '..', 'js', 'real_templates.js');
const rtContent = fs.readFileSync(rtPath, 'utf8');

// 用 eval 在 Node 里加载（real_templates 是浏览器格式）
let realTemplates;
try {
  const vm = require('vm');
  const sandbox = { window: {}, console: console, global: {} };
  sandbox.window.TD = {};
  sandbox.global = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(rtContent + '\n;global._td = window.TD;', sandbox);
  realTemplates = sandbox._td.realTemplates || sandbox._td.real_templates;
} catch (e) {
  console.log('eval real_templates err:', e.message);
  process.exit(1);
}

if (!realTemplates) {
  console.log('realTemplates not found');
  process.exit(1);
}

let keasKey, entry, keasIdx;
const keys = Object.keys(realTemplates);
console.log('total entries:', keys.length, 'first 5 keys:', keys.slice(0, 5));
for (let i = 0; i < keys.length; i++) {
  const e = realTemplates[keys[i]];
  const str = JSON.stringify(e).slice(0, 500);
  if (/K-EAS|KEAS/i.test(str) || (e && e.name && /keas/i.test(e.name))) {
    keasKey = keys[i]; entry = e; keasIdx = i; break;
  }
}
if (!entry) {
  // 退而求其次，找含 logo 的第一个
  for (let i = 0; i < keys.length; i++) {
    if (realTemplates[keys[i]].logo) { keasKey = keys[i]; entry = realTemplates[keys[i]]; keasIdx = i; break; }
  }
  console.log('KEAS not by name, fallback to first with logo');
}
console.log('KEAS entry idx:', keasKey, '| name:', entry && (entry.name || entry.id));
if (!entry) { process.exit(1); }
console.log('entry keys:', Object.keys(entry).slice(0, 10));
console.log('entry.logo?', entry.logo ? `YES (${entry.logo.ext}, ${entry.logo.dataB64.length} chars)` : 'NULL');

(async () => {
  const fileB64 = entry.fileBufB64 || entry.fileB64 || entry.file_buf_b64;
  if (!fileB64) { console.log('no fileBufB64 in entry, keys:', Object.keys(entry)); process.exit(1); }
  const wb = new ExcelJS.Workbook();
  const fileBuf = Buffer.from(fileB64, 'base64');
  await wb.xlsx.load(fileBuf);
  const ws = wb.worksheets[0];

  const data = {
    invoiceNo: 'BK20260803-006',
    invoiceDate: '2026-08-03',
    shipper: { name: '示例发件公司 DEMO TRADING CO., LTD.', address: 'JW PEI AP LIMITED UNIT 1102 11/F, 29 AUSTIN ROAD, TSIM SHA TSUI, KL, Hong Kong, CHINA' },
    consignee: { name: 'DEMO IMPORT LLC', address: 'SAME AS CONSIGNEE' },
    notify: { name: 'SAME AS CONSIGNEE', address: '' },
    pol: 'SHENZHEN, CHINA',
    pod: 'LOS ANGELES',
    shippingMarks: 'N/M',
    items: []
  };
  // 10 条货物，匹配用户截图
  for (let i = 1; i <= 10; i++) {
    data.items.push({
      boxCount: 1, nameEn: i === 1 ? 'Shoulder Bag' : 'Handbag', gw: 73.65,
      volume: 0.1, hsCode: '420222', boxNo: String(i), ctns: 1,
      containerNo: 'CARTON-1', sealNo: String(i),
      dims: '58x38x37'
    });
  }

  const fillRes = engine.fillTemplate(wb, data, { logo: entry.logo || null });
  console.log('unresolved:', JSON.stringify(fillRes.unresolved));
  console.log('replaced:', fillRes.replaced.length);

  // 检查 ws.getImages
  const imgs = ws.getImages();
  console.log('ws.getImages().length:', imgs.length);
  if (imgs.length) {
    const r = imgs[0].range;
    console.log('  im.imageId:', imgs[0].imageId, '| tl col/row:', r.tl.nativeCol + ',' + r.tl.nativeRow, '| br col/row:', r.br.nativeCol + ',' + r.br.nativeRow);
  }

  // 渲染预览
  const html = preview.wbToHtml(wb);
  console.log('preview html size:', html.length);
  console.log('contains <img:', html.includes('<img'));
  console.log('contains data:image:', html.includes('data:image'));
  console.log('contains data:image/jpeg:', html.includes('data:image/jpeg'));
  console.log('contains data:image/png:', html.includes('data:image/png'));
  // 截一段 img 标签
  const m = html.match(/<img[^>]{0,200}/);
  console.log('first <img>:', m ? m[0].slice(0, 180) : 'NONE');

  // 写一个 demo 预览
  const out = path.join(__dirname, '..', 'templates', '_preview_keas_v1422.html');
  fs.writeFileSync(out, '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>' + html + '</body></html>');
  console.log('preview written to:', out);
})().catch(e => console.log('ERR:', e));