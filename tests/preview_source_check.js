// 验证方案 B：模板库预览直显源文件原样（含 LOGO + 样张公司名/地址）
const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const vm = require('vm');

const ROOT = path.join(__dirname, '..');
const rtContent = fs.readFileSync(path.join(ROOT, 'js/real_templates.js'), 'utf8');
const sandbox = { window: {}, console };
sandbox.window.TD = {};
vm.createContext(sandbox);
vm.runInContext(rtContent, sandbox);
const realTemplates = sandbox.window.TD.realTemplates;

// 取 KEAS
const entry = realTemplates.find(t => /keas/i.test(t.id));
console.log('KEAS entry:', entry && entry.id);
console.log('has previewBufB64:', !!entry.previewBufB64, '| len:', entry.previewBufB64 ? entry.previewBufB64.length : 0);
console.log('has logo:', !!entry.logo);

(async () => {
  const ExcelJS2 = require('exceljs');
  const wb = new ExcelJS2.Workbook();
  const buf = Buffer.from(entry.previewBufB64, 'base64');
  await wb.xlsx.load(buf);
  console.log('source rows:', wb.worksheets[0].rowCount);

  // 方案 B 预览：addLogo 贴回
  if (entry.logo && entry.logo.dataB64) {
    // 复用 engine.addLogo 逻辑（此处内联等价实现）
    const logoBuf = Buffer.from(entry.logo.dataB64, 'base64');
    const imgId = wb.addImage({ buffer: logoBuf, extension: entry.logo.ext });
    wb.worksheets[0].addImage(imgId, { tl: entry.logo.from, br: entry.logo.to });
  }
  const imgs = wb.worksheets[0].getImages();
  console.log('ws.getImages().length after addLogo:', imgs.length);

  // 取 A5 单元格（KEAS 源文件为样张公司名）
  const a5 = wb.worksheets[0].getCell('A5').value;
  console.log('A5 (源文件样张):', JSON.stringify(a5));

  // 渲染 HTML（等价于 ui.js 模板库预览）
  const previewMod = require(path.join(ROOT, 'js/preview.js'));
  const html = previewMod.wbToHtml(wb);
  console.log('--- preview html ---');
  console.log('has <img> logo:', html.includes('<img'));
  console.log('has sample text JW PEI:', html.includes('JW PEI') || html.includes('PEI'));
  console.log('html length:', html.length);
  fs.writeFileSync(path.join(ROOT, 'templates', '_preview_source_keas.html'), html);
  console.log('wrote templates/_preview_source_keas.html');
})().catch(e => { console.error('ERR', e.message); process.exit(1); });
