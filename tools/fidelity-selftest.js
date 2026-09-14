/* 保真核验工具 · 自测（用真实订舱单模板 + 人工篡改来验证灵敏度） */
const ExcelJS = require('C:/Users/cn/.workbuddy/binaries/node/workspace/node_modules/exceljs');
const core = require('D:/WB文件/2026-08-21-16-22-08/trade-docs-system/tools/fidelity-core.js');
const TPL = 'D:/模板/订舱单/空模板/DETRANS  BOOKING air -GEODIS.xlsx';

let pass = 0, fail = 0;
function ok(n, c, e) { if (c) { pass++; console.log('  PASS  ' + n); } else { fail++; console.log('  FAIL  ' + n + (e ? ('  -> ' + e) : '')); } }

async function load(p) { const wb = new ExcelJS.Workbook(); await wb.xlsx.readFile(p); return wb; }
async function clone(wb) { const buf = await wb.xlsx.writeBuffer(); const w = new ExcelJS.Workbook(); await w.xlsx.load(buf); return w; }

(async () => {
  /* 先看模板基本特征，便于理解对齐假设 */
  let t0 = await load(TPL);
  console.log('模板：' + t0.worksheets[0].name +
    ' | sheets=' + t0.worksheets.map(w => w.name).join('/') +
    ' | rowCount=' + t0.worksheets[0].rowCount +
    ' | columnCount=' + t0.worksheets[0].columnCount);
  console.log('明细占位行(findItemRow) = ' + core.findItemRow(t0.worksheets[0]) +
    '（-1 表示该模板无 {{items.*}} 占位符，走表头标签映射填充）');
  console.log('合并区数量 = ' + core.mergesOf(t0.worksheets[0]).length);
  console.log('');

  /* T1 往返自比：检验「xlsx 写→读」本身是否引入假差异 */
  let tpl = await load(TPL), out = await clone(tpl);
  let r = core.compareWorkbooks(tpl, out);
  ok('T1 同一文件往返自比 → 无差异', r.stats.total === 0,
    'B=' + r.stats.B + ' A=' + r.stats.A + ' kinds=' + JSON.stringify(r.stats.byKind));

  /* T2 灵敏度：改边框 */
  tpl = await load(TPL); out = await clone(tpl);
  out.worksheets[0].getCell('A1').border = { top: { style: 'double' }, left: { style: 'thin' } };
  r = core.compareWorkbooks(tpl, out);
  ok('T2 边框被改 → 检出样式差异(B)', r.diffs.some(d => d.kind === 'style' && d.desc.indexOf('边框') >= 0),
    JSON.stringify(r.stats.byKind));

  /* T3 灵敏度：改底色 */
  tpl = await load(TPL); out = await clone(tpl);
  out.worksheets[0].getCell('B3').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00FF' } };
  r = core.compareWorkbooks(tpl, out);
  ok('T3 底色被改 → 检出', r.diffs.some(d => d.kind === 'style' && d.desc.indexOf('底色') >= 0), '');

  /* T4 灵敏度：改字体 */
  tpl = await load(TPL); out = await clone(tpl);
  out.worksheets[0].getCell('C4').font = { name: 'Arial', size: 20, bold: true };
  r = core.compareWorkbooks(tpl, out);
  ok('T4 字体被改 → 检出', r.diffs.some(d => d.kind === 'style' && d.desc.indexOf('字体') >= 0), '');

  /* T5 灵敏度：改对齐 */
  tpl = await load(TPL); out = await clone(tpl);
  out.worksheets[0].getCell('D5').alignment = { horizontal: 'center', vertical: 'top', wrapText: true };
  r = core.compareWorkbooks(tpl, out);
  ok('T5 对齐被改 → 检出', r.diffs.some(d => d.kind === 'style' && d.desc.indexOf('对齐') >= 0), '');

  /* T6 值差异（模板无占位符 → 应判为 B 类） */
  tpl = await load(TPL); out = await clone(tpl);
  out.worksheets[0].getCell('D6').value = 'ZZZ-CHANGED';
  r = core.compareWorkbooks(tpl, out);
  ok('T6 值被改 → 报 B 类值差异', r.diffs.some(d => d.kind === 'value' && d.level === 'B'), '');

  /* T7 占位符 → 值：应归 A 类（预期填充） */
  tpl = await load(TPL); out = await clone(tpl);
  tpl.worksheets[0].getCell('E7').value = '{{shipper.name}}';
  out.worksheets[0].getCell('E7').value = 'ACME CO';
  r = core.compareWorkbooks(tpl, out);
  ok('T7 占位符填充 → 归 A 类', r.diffs.some(d => d.level === 'A' && d.kind === 'value'),
    JSON.stringify(r.stats.byKind));

  /* T8 列宽 */
  tpl = await load(TPL); out = await clone(tpl);
  out.worksheets[0].getColumn(3).width = (tpl.worksheets[0].getColumn(3).width || 10) + 12;
  r = core.compareWorkbooks(tpl, out);
  ok('T8 列宽被放大 → 检出（验证「只增不减」问题）', r.diffs.some(d => d.kind === 'width'),
    JSON.stringify(r.widthDiffs.slice(0, 3)));

  /* T9 合并区丢失 */
  tpl = await load(TPL); out = await clone(tpl);
  const ms = core.mergesOf(tpl.worksheets[0]);
  if (ms.length) {
    out.worksheets[0].unMergeCells(ms[0]);
    r = core.compareWorkbooks(tpl, out);
    ok('T9 合并区丢失 → 检出', r.diffs.some(d => d.kind === 'merge' && d.level === 'B'), '被移除的合并=' + ms[0]);
  } else { console.log('  SKIP  T9（模板无合并区）'); }

  /* T10 引擎等价模拟：插行 + 按位移做样式 1:1 还原（这正是引擎 ④ 步做的事）
   *     → 期望工具报「0 样式差异」，只报行数变化 */
  async function engineSim(AT, D, skipRows) {
    const t = await load(TPL), o = await clone(t);
    const tWs = t.worksheets[0], oWs = o.worksheets[0];
    const SNAP = {};
    tWs.eachRow({ includeEmpty: true }, (row, rn) => {
      row.eachCell({ includeEmpty: true }, (c, cn) => {
        const s = c.style;
        if (s && (s.border || s.fill || s.font || s.alignment)) {
          SNAP[rn] = SNAP[rn] || {}; SNAP[rn][cn] = JSON.parse(JSON.stringify(s));
        }
      });
    });
    oWs.spliceRows(AT, 0, ...Array(D).fill([]));
    /* 模拟引擎：给「插入的新行」复制模板明细槽位行的样式（边框等） */
    const slotStyles = {};
    tWs.getRow(AT - 1).eachCell({ includeEmpty: true }, (c, cn) => {
      slotStyles[cn] = JSON.parse(JSON.stringify(c.style || {}));
    });
    for (let k = 0; k < D; k++) {
      const nrow = oWs.getRow(AT + k);
      Object.keys(slotStyles).forEach(cn => { nrow.getCell(+cn).style = JSON.parse(JSON.stringify(slotStyles[cn])); });
    }
    Object.keys(SNAP).forEach(rn => {
      const tr = +rn, fr = (tr <= AT - 1) ? tr : tr + D;
      if (skipRows && tr >= skipRows[0] && tr <= skipRows[1]) return;   // 故意漏还原
      Object.keys(SNAP[rn]).forEach(cn => {
        oWs.getRow(fr).getCell(+cn).style = JSON.parse(JSON.stringify(SNAP[rn][cn]));
      });
    });
    return core.compareWorkbooks(t, o);
  }

  r = await engineSim(8, 3, null);
  console.log('     [引擎等价模拟] ' + JSON.stringify(r.align[0]));
  ok('T10 引擎等价模拟 → 样式差异为 0', r.stats.styleDiff === 0,
    'styleDiff=' + r.stats.styleDiff + ' 样例=' + JSON.stringify(r.diffs.filter(d => d.kind === 'style').slice(0, 2).map(d => d.desc.slice(0, 80))));
  ok('T10 能检测到明细插行 delta=3', r.align[0].delta === 3, 'delta=' + r.align[0].delta);
  ok('T10 未插行区域 0 误报', !r.diffs.some(d => d.kind === 'style'),
    JSON.stringify(r.diffs.filter(d => d.kind === 'style').slice(0, 2)));

  /* T10b 灵敏度：故意漏还原 3 行 → 工具必须能抓到（否则「没问题」就是假结论） */
  r = await engineSim(8, 3, [20, 22]);
  ok('T10b 漏还原 3 行 → 工具能检出', r.stats.styleDiff > 0, 'styleDiff=' + r.stats.styleDiff);

  /* T11 sheet 数量差异 */
  tpl = await load(TPL); out = await clone(tpl);
  out.addWorksheet('ExtraSheet');
  r = core.compareWorkbooks(tpl, out);
  ok('T11 成品多出 sheet → 检出', r.diffs.some(d => d.kind === 'sheet'), '');

  console.log('');
  console.log('RESULT: pass=' + pass + ' fail=' + fail);
  if (fail > 0) process.exitCode = 1; else console.log('REGRESSION PASS: 工具行为符合预期');
})().catch(e => { console.error('ERROR', e); process.exitCode = 1; });
