/* 从飞书两个「申报信息」合并生成 js/declare_data.js（window.TD.declareData）。
 * 规则（用户指定）：
 *   - 主文档 Rv6ysNJjGhYpnptFL60cYUtFnMe（SKU主数据，9638行）字段更全，作为首选来源；
 *   - 备用文档 KImPstLu9h6Kj4tfY5qcKQuknee（申报专用，45037行）补齐主文档缺失的
 *     申报中文名/申报金额/海关编码/重量等申报关键项（同 SKU 匹配，仅填空字段）。
 * 用法：node tests/build_declare_data.js
 * 依赖：lark-cli（--as user 读本人资源）；落盘再读，避免管道。
 */
'use strict';
const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');

const PRIMARY = {
  token: 'Rv6ysNJjGhYpnptFL60cYUtFnMe', sheetId: '0uWwtB',
  url: 'https://ncncod7t3txg.feishu.cn/sheets/Rv6ysNJjGhYpnptFL60cYUtFnMe?sheet=0uWwtB',
  ranges: ['A1:T9638']
};
const FALLBACK = {
  token: 'KImPstLu9h6Kj4tfY5qcKQuknee', sheetId: '37f666',
  url: 'https://ncncod7t3txg.feishu.cn/sheets/KImPstLu9h6Kj4tfY5qcKQuknee?sheet=37f666',
  ranges: ['A1:T20000', 'A20001:T40000', 'A40001:T45037']
};
const OUT_JS = path.join(__dirname, '..', 'js', 'declare_data.js');

function blank(v) { return v === undefined || v === null || String(v).trim() === ''; }
function num(v) { if (blank(v)) return null; const n = parseFloat(String(v).replace(/,/g, '')); return isNaN(n) ? null : n; }
function trim(v) { return blank(v) ? '' : String(v).trim(); }

// CNY/USD 折算汇率（备用文档「申报金额」假设为元/CNY）。若备用文档实为 USD，将 RATE 设为 1 且 fallback currency 改 'USD'。
const RATE = 7.2;
function toUsd(v, cur) { if (blank(v)) return null; const n = (cur === 'USD') ? v : v / RATE; return Math.round(n * 1000) / 1000; }

function fetchSheet(doc, range, outFile) {
  const cmd = 'lark-cli sheets +csv-get --url "' + doc.url + '" --sheet-id ' + doc.sheetId +
    ' --as user --range ' + range + ' --format json > "' + outFile + '" 2>"' + outFile + '.err"';
  const r = spawnSync(cmd, { shell: true, encoding: 'utf8', cwd: path.join(__dirname, '..') });
  if (r.status !== 0) throw new Error('lark-cli 失败 ' + range + ': ' + (r.stderr || '') + (fs.existsSync(outFile + '.err') ? fs.readFileSync(outFile + '.err', 'utf8').slice(0, 300) : ''));
  if (!fs.existsSync(outFile)) throw new Error('未生成输出文件 ' + outFile);
  const txt = fs.readFileSync(outFile, 'utf8');
  let json; try { json = JSON.parse(txt); } catch (e) { throw new Error('JSON 解析失败 ' + outFile + ': ' + e.message + ' 前200字符=' + txt.slice(0, 200)); }
  if (!json.ok) throw new Error('飞书返回 ok=false: ' + (json.msg || ''));
  return json.data && json.data.annotated_csv ? json.data.annotated_csv : '';
}

// 解析 annotated_csv：每行 [row=N] a,b,c,...
function parseCsv(text) {
  const lines = String(text).split('\n');
  const rows = [];
  for (const line of lines) {
    const m = line.match(/^\[row=\d+\]\s?(.*)$/);
    if (!m) continue;
    rows.push(csvLine(m[1]));
  }
  return rows;
}
function csvLine(s) {
  const out = []; let cur = ''; let q = false;
  for (let i = 0; i < s.length; i++) {
    const c = s[i];
    if (q) {
      if (c === '"' && s[i + 1] === '"') { cur += '"'; i++; }
      else if (c === '"') q = false;
      else cur += c;
    } else {
      if (c === '"') q = true;
      else if (c === ',') { out.push(cur); cur = ''; }
      else cur += c;
    }
  }
  out.push(cur);
  return out;
}

function findHeader(rows, predicate) {
  for (let i = 0; i < rows.length; i++) if (predicate(rows[i])) return i;
  return -1;
}

console.log('抓取主文档...');
let primaryCsv = '';
for (const rg of PRIMARY.ranges) {
  const f = path.join(__dirname, '_pri_' + rg.replace(/:/g, '_') + '.json');
  primaryCsv += '\n' + fetchSheet(PRIMARY, rg, f);
}
const priRows = parseCsv(primaryCsv);
const priH = findHeader(priRows, r => trim(r[1]) === '商品编码');
if (priH < 0) throw new Error('未找到主文档表头(商品编码)');
const priData = priRows.slice(priH + 1).filter(r => !blank(r[1]));

console.log('抓取备用文档(分块)...');
let fbCsv = '';
for (const rg of FALLBACK.ranges) {
  const f = path.join(__dirname, '_fb_' + rg.replace(/:/g, '_') + '.json');
  fbCsv += '\n' + fetchSheet(FALLBACK, rg, f);
}
const fbRows = parseCsv(fbCsv);
const fbH = findHeader(fbRows, r => trim(r[0]) === 'sku_id');
if (fbH < 0) throw new Error('未找到备用文档表头(sku_id)');
const fbData = fbRows.slice(fbH + 1).filter(r => !blank(r[0]));

console.log('主文档数据行: ' + priData.length + '，备用文档数据行: ' + fbData.length);

// 主文档 → 记录（列索引见脚本顶部说明）
const primaryMap = {}; const primaryList = [];
for (const p of priData) {
  const sku = trim(p[1]);
  const rec = {
    sku: sku,
    nameCn: trim(p[4]), nameEn: trim(p[5]),
    hsCode: trim(p[10]),
    declarePrice: num(p[6]),           // 申报价（已为USD，来自销售单价/USD）
    declarePriceRaw: num(p[6]),        // 原币值（USD）
    material: trim(p[13]), brand: trim(p[12]), usage: trim(p[15]),
    unit: trim(p[11]), nw: num(p[9]), gw: null,
    currency: 'USD', model: sku, origin: 'CN', ver: 1
  };
  primaryMap[sku] = rec; primaryList.push(rec);
}

// 备用文档 → 补充映射（仅保留含申报有效信息的行，过滤虚拟/test 噪声）
const fbMap = {};
for (const f of fbData) {
  const sku = trim(f[0]);
  const fb = {
    nameCn: trim(f[8]),                // 申报中文名
    declarePriceRaw: num(f[9]),        // 申报金额（原币值，假设 CNY/元）
    declarePrice: toUsd(num(f[9]), 'CNY'), // 统一折算 USD
    currency: 'CNY',                   // 备用文档申报金额币种（假设元）
    hsCode: trim(f[10]),               // 海关编码
    nw: num(f[3]), gw: num(f[3])       // 重量
  };
  if (blank(fb.nameCn) && blank(fb.declarePrice) && blank(fb.hsCode) && blank(fb.nw)) continue;
  fbMap[sku] = fb;
}

// 字段级合并：主文档优先，空字段用备用补齐
// 注意：重量(nw/gw)只在「主文档本身也缺净重」时才取备用文档，避免主文档 kg 与备用文档
// 重量列单位不一致导致同 SKU 重量错配（如主文档 0.48kg 被备用 660 覆盖）。
let supplemented = 0, weightSkipped = 0;
for (const rec of primaryList) {
  const fb = fbMap[rec.sku];
  if (!fb) continue;
  if (blank(rec.hsCode)) { rec.hsCode = fb.hsCode; supplemented++; }
  if ((rec.declarePrice === null || rec.declarePrice === 0) && !blank(fb.declarePrice)) {
    rec.declarePrice = fb.declarePrice;        // 已是 USD
    rec.declarePriceRaw = fb.declarePriceRaw;  // 原币值
    rec.currency = fb.currency;                // 原币种
    supplemented++;
  }
  if (blank(rec.nameCn)) { rec.nameCn = fb.nameCn; supplemented++; }
  if (blank(rec.nw)) { rec.nw = fb.nw; rec.gw = fb.gw || fb.nw; supplemented++; }
  else { weightSkipped++; }
}

// 备用文档独有 SKU（主文档没有），且含申报有效信息 → 也纳入（扩大覆盖）
let addedFallbackOnly = 0;
for (const sku of Object.keys(fbMap)) {
  if (primaryMap[sku]) continue;
  const fb = fbMap[sku];
  primaryList.push({
    sku: sku, nameCn: fb.nameCn, nameEn: '',
    hsCode: fb.hsCode, declarePrice: fb.declarePrice, declarePriceRaw: fb.declarePriceRaw,
    material: '', brand: '', usage: '', unit: '',
    nw: fb.nw, gw: fb.gw, currency: fb.currency, model: sku, origin: 'CN', ver: 1
  });
  addedFallbackOnly++;
}

// 清理：declarePrice 0/null 统一为 0（前端 number 输入）；gw 缺失用 nw
for (const rec of primaryList) {
  rec.declarePrice = (rec.declarePrice === null) ? 0 : rec.declarePrice;
  if (blank(rec.gw) && !blank(rec.nw)) rec.gw = rec.nw;
  rec.nw = rec.nw === null ? 0 : rec.nw;
  rec.gw = rec.gw === null ? 0 : rec.gw;
}

const out = '/* 自动生成：由 tests/build_declare_data.js 从飞书两个申报信息合并（主文档优先，缺字段用备用补齐；申报价统一折算USD，币种见 currency 字段）。\n' +
  '   主文档(Rv6y)优先，缺字段用备用文档(KImP)同 SKU 补齐。请勿手改，重跑脚本即可。 */\n' +
  'window.TD = window.TD || {};\n' +
  'window.TD.declareData = ' + JSON.stringify(primaryList, null, 0) + ';\n';
fs.writeFileSync(OUT_JS, out, 'utf8');

console.log('合并完成：总记录 ' + primaryList.length +
  '（主文档 ' + priData.length + ' + 备用独有 ' + addedFallbackOnly + '），字段补齐次数 ' + supplemented +
  '，重量因单位不一致跳过覆盖 ' + weightSkipped);
console.log('已写出 ' + OUT_JS + ' 大小 ' + (fs.statSync(OUT_JS).size / 1024).toFixed(1) + ' KB');

// 清理临时文件
['_pri_', '_fb_'].forEach(pref => {
  for (const rg of PRIMARY.ranges.concat(FALLBACK.ranges)) {
    const base = path.join(__dirname, pref + rg.replace(/:/g, '_') + '.json');
    try { fs.unlinkSync(base); } catch (e) {}
    try { fs.unlinkSync(base + '.err'); } catch (e) {}
  }
});
