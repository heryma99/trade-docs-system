// Compare KEAS source (user) vs repo tpl vs real_templates.js baked
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const ExcelJS = require('exceljs');

const ROOT = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统';
const USER_SRC = 'D:/模板/订舱单/New KEAS Booking_SI_VGM Form_Sea_南区Rev.xlsx';
const REPO_TPL = path.join(ROOT, 'templates/tpl_real_keas_booking.xlsx');

async function inspectFile(label, filePath) {
  console.log('\n========== ' + label + ' ==========');
  console.log('path:', filePath);
  if (!fs.existsSync(filePath)) { console.log('NOT EXISTS'); return; }
  const stat = fs.statSync(filePath);
  console.log('size:', stat.size, 'bytes');
  const z = await JSZip.loadAsync(fs.readFileSync(filePath));
  const media = Object.keys(z.files).filter(k => k.startsWith('xl/media/') && !k.endsWith('/'));
  const drawings = Object.keys(z.files).filter(k => k.match(/xl\/drawings\/drawing\d+\.xml$/));
  console.log('xl/media files:', media.length, media.slice(0, 6));
  console.log('xl/drawings xml files:', drawings.length, drawings.slice(0, 6));
  if (drawings.length) {
    const xml = await z.file(drawings[0]).async('string');
    const cols = xml.match(/col="\d+"/g) || [];
    const rows = xml.match(/row="\d+"/g) || [];
    const blip = xml.match(/r:embed="(rId\d+)"/);
    console.log('drawing1 anchors col:', cols.join(','), '| row:', rows.join(','));
    console.log('drawing1 blip embed:', blip ? blip[1] : 'none');
    // also check vml drawing
    const vml = Object.keys(z.files).filter(k => k.match(/vmlDrawing\d+\.vml$/));
    console.log('vml drawings:', vml);
  }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  const ws = wb.worksheets[0];
  console.log('sheet:', ws.name, '| rows:', ws.rowCount, '| merges:', (ws.model.merges || []).length, '| images(ExcelJS):', (ws.getImages() || []).length);
  ['A1','A2','A5','A13','A36','B36','I36'].forEach(a => {
    const c = ws.getCell(a);
    if (c.value !== null && c.value !== undefined) console.log('  ' + a + ':', JSON.stringify(c.value).slice(0, 100));
  });
}

async function main() {
  await inspectFile('USER SOURCE (用户上传的最新源文件)', USER_SRC);
  await inspectFile('REPO tpl_real_keas_booking.xlsx (仓库内 bake 输入)', REPO_TPL);

  const rtContent = fs.readFileSync(path.join(ROOT, 'js/real_templates.js'), 'utf8');
  const start = rtContent.indexOf('"tpl_real_keas_booking"');
  if (start < 0) { console.log('!! KEAS block not found in real_templates.js'); return; }
  const chunk = rtContent.slice(start, start + 8000);
  const m = chunk.match(/"fileBufB64":\s*"([A-Za-z0-9+/=]+)"/);
  if (!m) { console.log('\n!! KEAS fileBufB64 not found'); return; }
  const b64 = m[1];
  const buf = Buffer.from(b64, 'base64');
  console.log('\n========== real_templates.js BAKED ==========');
  console.log('fileBufB64 length:', b64.length, '| decoded size:', buf.length);
  const tmp = path.join(ROOT, 'templates/_keas_baked_tmp.xlsx');
  fs.writeFileSync(tmp, buf);
  await inspectFile('REAL_TEMPLATES bake (解码后的入库版)', tmp);
  try { fs.unlinkSync(tmp); } catch(e) {}
}

main().catch(e => { console.error('FATAL', e.message); process.exit(1); });
