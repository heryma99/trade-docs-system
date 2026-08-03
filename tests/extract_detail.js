/* 提取每个 xlsx 的明细表头行（含列标签）与 1 行样本，用于字段映射 */
const ExcelJS = require('exceljs');
const files = process.argv.slice(2);

const HINTS = ['品名','海关','HS','品牌','型号','材质','用途','数量','单价','金额','箱号','毛重','净重','长宽高','申报','QTY','SKU','Reference','FBA','Ref','ASIN','FNSKU','箱数','件数','体积','尺寸','规格','CTN','GW','NW','PURPOSE','MATERIAL','MODEL','BRAND','DESCRIPTION','EORI','VAT'];

function cellVal(c) {
  let v = c.value;
  if (v == null) return null;
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(t => t.text).join('');
    if (v.formula) return '=' + v.formula;
    if (v.text) return v.text;
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return JSON.stringify(v);
  }
  return v;
}
function score(s) {
  if (s == null) return 0;
  s = String(s);
  let n = 0;
  for (const h of HINTS) if (s.indexOf(h) >= 0) n++;
  return n;
}

(async () => {
  for (const f of files) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(f);
    console.log('\n########## FILE: ' + f.split('/').pop());
    wb.eachSheet((ws) => {
      // 找明细表头行：得分最高且 >=3
      let best = -1, bestScore = 0;
      for (let r = 1; r <= Math.min(ws.rowCount, 60); r++) {
        const row = ws.getRow(r);
        let sc = 0;
        row.eachCell({ includeEmpty: false }, (c) => { sc += score(cellVal(c)); });
        if (sc > bestScore) { bestScore = sc; best = r; }
      }
      if (best < 0 || bestScore < 3) { console.log('  SHEET "' + ws.name + '": 无明细表头'); return; }
      console.log('  SHEET "' + ws.name + '" 明细表头行=R' + best + ' (score=' + bestScore + ')');
      const hdr = ws.getRow(best);
      const cols = [];
      hdr.eachCell({ includeEmpty: true }, (c, col) => {
        const v = cellVal(c);
        if (v !== null && String(v).trim() !== '') cols.push(col + ':' + String(v).slice(0, 30));
      });
      console.log('    表头: ' + cols.join(' | '));
      // 样本 1 行
      const s1 = ws.getRow(best + 1);
      const sc = [];
      s1.eachCell({ includeEmpty: true }, (c, col) => {
        const v = cellVal(c);
        if (v !== null && String(v).trim() !== '') sc.push(col + ':' + String(v).slice(0, 24));
      });
      if (sc.length) console.log('    样本: ' + sc.join(' | '));
    });
  }
})().catch(e => { console.error('ERROR: ' + e.message); process.exit(1); });
