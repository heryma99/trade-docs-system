/* 探查 6 个订舱单样张：sheet 名 + 主表前若干行内容，用于定位 detailRow。 */
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const ExcelJS = require('exceljs');

const SRC_DIRS = ['D:/模板/发票模板', 'D:/模板/订舱单'];
const OUT_DIR = path.join(__dirname, '..', 'templates');
const PY = process.env.PYTHON || (process.env.HOME + '/.workbuddy/binaries/python/envs/default/Scripts/python.exe');

const ONLY = process.env.ONLY;
const CAP = process.env.CAP ? parseInt(process.env.CAP, 10) : 45;
const FILES = ONLY ? [ONLY] : [
  'Aramex  Booking Form  .xlsx',
  'BOOKING  FORM.xls',
  'CHR Shenzhen Booking Form.xls',
  'DETRANS  BOOKING air -SA.xls',
  'GEODIS BOOKING+SI form(1).xlsx',
  'New KEAS Booking_SI_VGM Form_Sea_南区Rev.xlsx',
];

function resolveSrc(f) {
  for (const d of SRC_DIRS) {
    const cand = path.join(d, f);
    if (fs.existsSync(cand)) return cand;
  }
  throw new Error('找不到: ' + f);
}
function cellStr(c) {
  const v = c.value;
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(t => t.text).join('');
    if (v.formula) return '=' + v.formula;
    if (v.text) return v.text;
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return '';
  }
  return String(v);
}
(async () => {
  for (const f of FILES) {
    let p = resolveSrc(f);
    if (f.toLowerCase().endsWith('.xls')) {
      const tmp = path.join(OUT_DIR, '_insp_' + path.basename(f) + '.xlsx');
      const r = spawnSync(PY, ['tests/xls_to_xlsx.py', p, tmp], { encoding: 'utf8' });
      if (r.status !== 0) { console.log('\n### ' + f + '  ❌ xls转换失败\n' + (r.stderr || r.stdout)); continue; }
      p = tmp;
    }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(p);
    console.log('\n══════════════════════════════════════════');
    console.log('### ' + f);
    console.log('sheets: ' + wb.worksheets.map(w => w.name + '(rows=' + w.rowCount + ')').join('  |  '));
    // 选主表 = rowCount 最大的
    const main = wb.worksheets.slice().sort((a, b) => b.rowCount - a.rowCount)[0];
    console.log('--- 主表 [' + main.name + '] 前 45 行 ---');
    for (let r = 1; r <= Math.min(CAP, main.rowCount); r++) {
      const row = main.getRow(r);
      const parts = [];
      row.eachCell({ includeEmpty: false }, (c, col) => {
        const s = cellStr(c).replace(/\s+/g, ' ').slice(0, 24);
        if (s) parts.push(col + ':' + s);
      });
      if (parts.length) console.log(String(r).padStart(3) + ' | ' + parts.join('  '));
    }
    if (p.indexOf('_insp_') >= 0) { try { fs.unlinkSync(p); } catch (e) {} }
  }
})().catch(e => { console.error(e); process.exit(1); });
