/* 全模板填充测试：扫描 templates/ 下所有 xlsx，模拟数据后检查 unresolved 占位符 */
const X = require('exceljs');
const path = require('path');
const fs = require('fs');
const engine = require('../js/engine.js');

function round(n, d) { return Math.round((n + Number.EPSILON) * Math.pow(10, d)) / Math.pow(10, d); }

const MOCK_DECLARE = {
  '104-1': { nameEn: 'Shoulder Bag', nameCn: '单肩包', hsCode: '4202220000', material: 'PU', brand: 'JW PEI', model: '104-1', declarePrice: 16.42, usage: 'Put things', nw: 0.2, gw: 0.25, unit: 'PCS', origin: 'CN' },
  '2T68-1': { nameEn: 'Handbag', nameCn: '手提包', hsCode: '4202220000', material: 'PU', brand: 'JW PEI', model: '2T68-1', declarePrice: 19.73, usage: 'Put things', nw: 0.3, gw: 0.35, unit: 'PCS', origin: 'CN' },
  '2T68-2': { nameEn: 'Handbag', nameCn: '手提包', hsCode: '4202220000', material: 'PU', brand: 'JW PEI', model: '2T68-2', declarePrice: 19.73, usage: 'Put things', nw: 0.3, gw: 0.35, unit: 'PCS', origin: 'CN' },
  '2T68-3': { nameEn: 'Handbag', nameCn: '手提包', hsCode: '4202220000', material: 'PU', brand: 'JW PEI', model: '2T68-3', declarePrice: 19.73, usage: 'Put things', nw: 0.3, gw: 0.35, unit: 'PCS', origin: 'CN' },
  'JH306B04-1': { nameEn: 'Handbag', nameCn: '手提包', hsCode: '4202220000', material: 'PU', brand: 'NO BRAND', model: 'JH306B04-1', declarePrice: 8, usage: 'Put things', nw: 0.3, gw: 0.35, unit: 'PCS', origin: 'CN' }
};

const MOCK_ORDERS = [
  { orderNo: 'ORD1', items: [
    { sku: '104-1', qty: 5, price: 16.42 },
    { sku: '2T68-1', qty: 8, price: 19.73 },
    { sku: '2T68-2', qty: 2, price: 19.73 },
    { sku: '2T68-3', qty: 5, price: 19.73 },
    { sku: 'JH306B04-1', qty: 39, price: 8 }
  ]}
];

const MOCK_PACKING = {
  boxes: [
    { boxNo: 'BOX1', sku: '104-1', qty: 5, gw: 73.65, nw: 60, length: 58, width: 38, height: 37 },
    { boxNo: 'BOX2', sku: '2T68-1', qty: 8, gw: 73.65, nw: 60, length: 58, width: 38, height: 37 },
    { boxNo: 'BOX3', sku: '2T68-2', qty: 2, gw: 73.65, nw: 60, length: 58, width: 38, height: 37 },
    { boxNo: 'BOX4', sku: '2T68-3', qty: 5, gw: 73.65, nw: 60, length: 58, width: 38, height: 37 },
    { boxNo: 'BOX5', sku: 'JH306B04-1', qty: 39, gw: 441.9, nw: 360, length: 58, width: 38, height: 37 }
  ],
  totals: { boxCount: 5, qty: 59, nw: 600, gw: 737.25, volume: 0.4 }
};

const SHIPPER = { name: 'SHIP', company: 'SHIP CO', address: 'No.1', city: 'Shenzhen', state: 'GD', zip: '518000', tel: '+86', contact: 'Mr Li', email: 'a@b.com', country: 'CN', taxNo: 'X' };
const CONSIGNEE = { name: 'JW', company: 'JW LLC', address: 'US', city: 'LA', state: 'CA', zip: '90001', tel: '+1', contact: '', email: 'u@j.com', country: 'US', taxNo: '' };

function countStyles(ws) {
  let border = 0, fill = 0, font = 0, align = 0;
  ws.eachRow({ includeEmpty: true }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      const st = cell.style;
      if (!st) return;
      if (st.border) border++;
      if (st.fill && st.fill.fgColor) fill++;
      if (st.font && (st.font.bold || st.font.size || st.font.name || st.font.color)) font++;
      if (st.alignment) align++;
    });
  });
  return { border: border, fill: fill, font: font, align: align };
}

function countMerges(ws) { return (ws.model.merges || []).length; }

function cellText(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (v && v.richText) return v.richText.map(t => t.text).join('');
  if (typeof v === 'string') return v;
  if (v && v.text !== undefined) return String(v.text);
  return v == null ? '' : String(v);
}

function getMergeAt(ws, row, col) {
  for (const m of (ws.model.merges || [])) {
    const mm = (typeof m === 'string') ? m.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/) : null;
    if (!mm) continue;
    const c1 = colCharToNum(mm[1]), r1 = parseInt(mm[2], 10), c2 = colCharToNum(mm[3]), r2 = parseInt(mm[4], 10);
    if (row >= r1 && row <= r2 && col >= c1 && col <= c2) return { top: r1, left: c1, bottom: r2, right: c2 };
  }
  return null;
}
function colCharToNum(s) { let n = 0; for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }

async function tempWriteRead(wb) {
  const tmp = path.join(require('os').tmpdir(), 'td_filltest_' + Date.now() + '_' + Math.floor(Math.random() * 1e6) + '.xlsx');
  await wb.xlsx.writeFile(tmp);
  const wb2 = await new X.Workbook().xlsx.readFile(tmp);
  try { fs.unlinkSync(tmp); } catch (e) {}
  return wb2;
}

async function run() {
  const dir = path.join(__dirname, '../templates');
  const files = fs.readdirSync(dir).filter(f => f.endsWith('.xlsx')).sort();
  let fail = 0, pass = 0;
  console.log('全模板填充测试开始，共 ' + files.length + ' 个模板\n');
  for (const f of files) {
    const fp = path.join(dir, f);
    let wb;
    try { wb = await new X.Workbook().xlsx.readFile(fp); } catch (e) { console.log('❌ ' + f + ' 读取失败: ' + e.message); fail++; continue; }
    const before = countStyles(wb.worksheets[0]);
    const beforeMerges = countMerges(wb.worksheets[0]);
    const scan = engine.scanTemplate(wb);
    const kind = /booking|订舱|book|aramex|chr|detrans|geodis|keas|zytd|yuntu/i.test(f) ? 'booking' : 'invoice';
    const itemFields = scan.itemFields || [];
    const boxMode = !!(itemFields.some(fld => /(boxNo|length|width|height)/.test(fld)));
    const data = engine.buildDocData({
      kind: kind,
      orders: MOCK_ORDERS,
      packing: MOCK_PACKING,
      meta: { invoiceNo: 'INV20260731-001', invoiceDate: '2026-07-31' },
      boxMode: boxMode,
      shipper: SHIPPER,
      consignee: CONSIGNEE,
      notify: null,
      declareMap: kind === 'invoice' ? MOCK_DECLARE : {}
    });
    let fillRes;
    try { fillRes = engine.fillTemplate(wb, data); } catch (e) { console.log('❌ ' + f + ' 填充抛错: ' + e.message); fail++; continue; }
    // 写回并重新读取，验证 ExcelJS 写回不丢样式（边框/底色/字体/对齐 1:1 保留）
    let wbOut;
    try { wbOut = await tempWriteRead(wb); } catch (e) { console.log('❌ ' + f + ' (boxMode=' + boxMode + ') 写回抛错: ' + e.message); fail++; continue; }
    const after = countStyles(wbOut.worksheets[0]);
    const afterMerges = countMerges(wbOut.worksheets[0]);
    const styleLost = (after.border < before.border) || (after.fill < before.fill) || (after.font < before.font) || (after.align < before.align);
    const mergeLost = afterMerges < beforeMerges;
    const unresolved = (fillRes.unresolved || []).filter((v, i, a) => a.indexOf(v) === i).sort();
    const expectedMayEmpty = ['items.asin', 'items.note', 'items.costPrice', 'items.fnsku'];
    const realIssues = unresolved.filter(u => expectedMayEmpty.indexOf(u) < 0);
    const styleMsg = styleLost
      ? ('❌ 样式丢失 [前' + JSON.stringify(before) + ' 后' + JSON.stringify(after) + ']')
      : ('样式 1:1 [边' + before.border + '/' + after.border + ' 底' + before.fill + '/' + after.fill + ' 字' + before.font + '/' + after.font + ' 齐' + before.align + '/' + after.align + ']');
    if (styleLost) { console.log('❌ ' + f + ' (' + kind + ', boxMode=' + boxMode + ') ' + styleMsg); fail++; continue; }
    if (mergeLost) { console.log('❌ ' + f + ' (' + kind + ', boxMode=' + boxMode + ') 合并单元格丢失 [前' + beforeMerges + ' 后' + afterMerges + ']'); fail++; continue; }
    if (realIssues.length) {
      console.log('⚠️  ' + f + ' (' + kind + ', boxMode=' + boxMode + ') ' + styleMsg + ' 有未解析占位符: ' + realIssues.join(' '));
      if (unresolved.length !== realIssues.length) console.log('    （可空字段: ' + unresolved.filter(u => expectedMayEmpty.indexOf(u) >= 0).join(' ') + '）');
    } else {
      console.log('✅ ' + f + ' (' + kind + ', boxMode=' + boxMode + ') ' + styleMsg + ' 合并' + beforeMerges + '/' + afterMerges + (unresolved.length ? ' 仅可空字段缺失' : ''));
      pass++;
    }
  }
    console.log('\n结果: 通过 ' + pass + ' / 失败 ' + fail + ' / 总计 ' + files.length);
    if (fail) process.exit(1);
}

async function main() { await run(); }

if (require.main === module) run().catch(e => { console.error(e); process.exit(1); });
module.exports = { run: run, main: main };
