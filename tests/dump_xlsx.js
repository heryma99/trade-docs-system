/* 解析 xlsx 模板结构：表名/尺寸/合并单元格/单元格内容 */
const ExcelJS = require('exceljs');
const path = process.argv[2];
const maxRows = parseInt(process.argv[3] || '60', 10);

function cellVal(c) {
  let v = c.value;
  if (v == null) return null;
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(t => t.text).join('');
    if (v.formula) return '=' + v.formula + (v.result != null ? ' -> ' + v.result : '');
    if (v.text) return v.text;
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return JSON.stringify(v);
  }
  return v;
}

(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  console.log('FILE: ' + path);
  wb.eachSheet((ws) => {
    console.log('\n=== SHEET "' + ws.name + '" rows=' + ws.rowCount + ' cols=' + ws.columnCount + ' ===');
    const merges = ws.model.merges || [];
    if (merges.length) console.log('MERGES: ' + merges.join(' '));
    // 列宽
    const widths = (ws.columns || []).slice(0, 20).map(c => c && c.width ? Math.round(c.width) : '-').join(',');
    console.log('COLWIDTHS: ' + widths);
    const lim = Math.min(ws.rowCount, maxRows);
    for (let r = 1; r <= lim; r++) {
      const row = ws.getRow(r);
      const cells = [];
      row.eachCell({ includeEmpty: false }, (c, colNum) => {
        const v = cellVal(c);
        if (v !== null && String(v).trim() !== '') cells.push(c.address + '=' + JSON.stringify(String(v).slice(0, 60)));
      });
      if (cells.length) console.log('R' + r + ': ' + cells.join(' | '));
    }
    if (ws.rowCount > maxRows) console.log('...(共 ' + ws.rowCount + ' 行，已截断)');
  });
})().catch(e => { console.error('ERROR: ' + e.message); process.exit(1); });
