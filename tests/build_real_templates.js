/* 把 8 个真实样张转成带占位符的模板 xlsx，并生成 js/real_templates.js（base64 供种子入库）。
 * 用法: node tests/build_real_templates.js
 * 依赖: exceljs (Node) + python(openpyxl) 用于 .xls 转换
 */
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const ExcelJS = require('exceljs');
const JSZip = require('jszip');

const SRC_DIRS = ['D:/模板/发票模板', 'D:/模板/订舱单'];
const OUT_DIR = path.join(__dirname, '..', 'templates');
const OUT_JS = path.join(__dirname, '..', 'js', 'real_templates.js');
const PY = process.env.PYTHON || (process.env.HOME + '/.workbuddy/binaries/python/envs/default/Scripts/python.exe');

// ---------- 字段映射字典 ----------
// 明细列标签子串 -> items 字段路径（null 表示剔除/留空，如图片）
const ITEM_MAP = [
  // —— 订舱单专用明细列（高优先级，置于通用规则之前，避免被泛化规则误伤）——
  ['总件数', 'boxCount'], ['货物数', 'boxCount'], ['NO. OF PIECES', 'boxCount'],
  ['NO. OF PACKAGE', 'boxCount'], ['NO. AND TYPES OF PKG', 'boxCount'], ['PACKAGES', 'boxCount'],
  ['货物描述', 'nameEn'], ['DESCRIPTION OF GOODS', 'nameEn'], ['NATURE OF GOODS', 'nameEn'], ['GOODS DESCRIPTION', 'nameEn'],
  ['GROSS WEIGHT', 'gw'], ['ACTUAL GROSS WEIGHT', 'gw'], ['MEASUREMENT', 'volume'], ['MEAS', 'volume'],
  ['NET WEIGHT', 'nw'], ['VOLUME WEIGHT', 'volumeWeight'], ['CHARGEABLE WEIGHT', 'volumeWeight'],
  ['DIMENSION', 'dims'], ['DIMENSIONS', 'dims'], ['H.S.CODE', 'hsCode'],
  ['PACKAGE TYPE', 'boxSpec'], ['CTNR NO', 'boxNo'], ['箱型', 'boxSpec'],
  ['SEAL NO.', null], ['S/O NO.', null], ['PO NO.', null], ['SHIPPING MARKS', null], ['MARKS & NUMBERS', null], ['唛头', null],
  ['货箱编号', 'boxNo'], ['FBA箱号', 'fbaNo'], ['箱号', 'boxNo'], ['No.of Pkgs', 'boxNo'],
  ['子单号', 'subOrderNo'], ['款号', 'model'], ['型号', 'model'], ['Model', 'model'], ['MODEL', 'model'],
  ['英文品名', 'nameEn'], ['中文品名', 'nameCn'], ['品名', 'nameEn'],
  ['海关编码', 'hsCode'], ['商品编码', 'hsCode'], ['HS CODE', 'hsCode'],
  ['品牌', 'brand'], ['牌子', 'brand'], ['BRAND', 'brand'],
  ['材质', 'material'], ['MATERIAL', 'material'],
  ['用途', 'usage'], ['PURPOSE', 'usage'],
  ['申报数量', 'qty'], ['数量单位', 'unit'], ['数量', 'qty'], ['QTY', 'qty'],
  ['申报单价', 'price'], ['申报总价', 'amount'], ['总申报价值', 'amount'], ['Vauel', 'amount'],
  ['VALUE', 'amount'], ['金额', 'amount'], ['单价', 'price'],
  ['采购价格', 'costPrice'], ['销售价格', 'costPrice'],
  ['箱数', 'boxCount'], ['件数', 'ctns'],
  ['货箱重量', 'gw'], ['单箱重量', 'gw'], ['单件毛重', 'gw'], ['总毛重', 'gw'], ['产品毛重', 'gw'], ['重量', 'gw'], ['毛重', 'gw'], ['G.W', 'gw'], ['GW', 'gw'],
  ['货箱净重', 'boxNw'], ['单件净重', 'nw'], ['单个产品重量', 'nw'], ['总净重', 'nw'], ['产品净重', 'nw'], ['净重', 'nw'], ['单重', 'nw'], ['N.W', 'nw'], ['NW', 'nw'],
  ['长度', 'length'], ['宽度', 'width'], ['高度', 'height'],
  ['长(cm', 'length'], ['宽(cm', 'width'], ['高(cm', 'height'],
  ['长CM', 'length'], ['宽CM', 'width'], ['高CM', 'height'],
  ['长宽高', 'dims'],
  ['产品图片', null], ['图片', null],
  ['纸箱规格', 'boxSpec'], ['体积重', 'volumeWeight'],
  ['体积', 'volume'], ['VOLUME', 'volume'],
  ['规格/尺寸', 'dims'], ['产品尺寸', 'dims'], ['尺寸', 'dims'], ['规格', 'dims'],
  ['Reference ID', 'refId'], ['ReferenceID', 'refId'], ['Ref ID', 'refId'],
  ['ASIN', 'asin'], ['FNSKU', 'fnsku'], ['SKU', 'sku'],
  ['原产国', 'origin'], ['原产地', 'origin'], ['目的国', 'destCountry'],
  ['币种', 'currency'], ['带电', 'electrified'], ['带磁', 'magnetic'], ['产品性质', 'nature'],
  ['货物新旧', 'condition'], ['成交方式', 'tradeTerm'],
  ['销售链接', 'note'], ['销售地址', 'note'], ['产品销售链接', 'note'], ['备注', 'note'], ['配货', 'note'],
  ['电池类型', 'batteryType'], ['品牌类型', 'brandType'], ['出口享惠', 'exportPrefer'],
  ['税则号', 'taxNo'], ['单箱单个产品数量', 'ctnQty'], ['总箱单个产品数量', 'totalCtnQty'],
];
// 表头区 "Label：Value" -> 字段路径
const HEADER_MAP = [
  // —— 订舱单表头专用（船名/航次/柜型/运费条款/贸易术语/港口等）——
  ['船名', 'vessel'], ['航次', 'vessel'], ['VESSEL', 'vessel'], ['VOYAGE', 'vessel'], ['OCEAN VESSEL', 'vessel'], ['船名航次', 'vessel'],
  ['ETD', 'etd'], ['ETA', 'etd'], ['船期', 'etd'],
  ['柜型', 'containerType'], ['柜量', 'containerQty'], ['CONTAINER', 'containerType'],
  ['FREIGHT', 'freightTerms'], ['运费条款', 'freightTerms'],
  ['贸易术语', 'incoterms'], ['贸易条款', 'incoterms'], ['INCOTERMS', 'incoterms'], ['TRADE TERM', 'incoterms'],
  ['订舱号', 'invoiceNo'], ['BOOKING', 'invoiceNo'],
  ['品名概述', 'goodsSummary'], ['COMMODITY', 'goodsSummary'],
  ['PORT OF LOADING', 'pol'], ['PLACE OF RECEIPT', 'pol'],
  ['PORT OF DISCHARGE', 'pod'], ['PLACE OF DELIVERY', 'pod'],
  ['订舱代理', 'agent'], ['AGENT', 'agent'],
  // 具体优先（避免被泛化规则误伤）
  ['进口商地址', 'consignee.address'], ['收件人地址', 'consignee.address'],
  ['收件人姓名', 'consignee.contact'], ['收件人公司', 'consignee.name'],
  ['收件人省州', 'consignee.state'], ['收件人城市', 'consignee.city'],
  ['收件人邮编', 'consignee.zip'], ['收件人电话', 'consignee.tel'], ['收件人邮箱', 'consignee.email'],
  ['进口商名称', 'consignee.name'], ['进口商', 'consignee.name'],
  ['客户订单号', 'orderNos'], ['客户单号', 'orderNos'], ['订单号', 'orderNos'],
  ['发票号', 'invoiceNo'], ['INVOICE', 'invoiceNo'],
  ['日期', 'invoiceDate'], ['DATE', 'invoiceDate'],
  ['合同', 'contractNo'], ['CONTRACT', 'contractNo'],
  ['收件人', 'consignee.name'], ['收货人', 'consignee.name'], ['CONSIGNEE', 'consignee.name'],
  ['发货人', 'shipper.name'], ['SHIPPER', 'shipper.name'],
  ['起运', 'pol'], ['FROM', 'pol'], ['目的', 'pod'], ['TO', 'pod'], ['POD', 'pod'],
  ['预计总件数', 'totals.boxCount'], ['总件数', 'totals.boxCount'], ['总箱数', 'totals.boxCount'],
  ['预计重量', 'totals.gw'], ['总毛重', 'totals.gw'], ['毛重', 'totals.gw'],
  ['总净重', 'totals.nw'], ['净重', 'totals.nw'],
  ['体积', 'totals.volume'], ['VOLUME', 'totals.volume'],
  ['报关方式', 'customsType'], ['成交方式', 'customsType'],
  ['运输方式', 'transport'], ['服务', 'transport'],
  ['带电', 'dangerous'], ['带磁', 'dangerous'], ['保险', 'remark'],
  ['Reference', 'meta.refId'], ['Ref', 'meta.refId'],
  ['税号', 'consignee.taxNo'], ['VAT', 'consignee.taxNo'], ['EORI', 'consignee.eori'],
  ['通知', 'notify.name'],
];
const TOTAL_FIELDS = ['qty', 'gw', 'nw', 'boxCount', 'volume', 'amount', 'ctns'];

function norm(s) { return String(s == null ? '' : s).toLowerCase(); }
function mapItem(label) {
  const l = norm(label);
  for (const [sub, path] of ITEM_MAP) if (l.indexOf(norm(sub)) >= 0) return path;
  return null;
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
function headerReplace(v) {
  const seps = ['：', ':'];
  for (const sep of seps) {
    const idx = v.indexOf(sep);
    if (idx > 0 && idx < v.length - 1) {
      const label = v.slice(0, idx), val = v.slice(idx + 1);
      if (/^\s*[\dA-Za-z]/.test(val) || val.length) {
        const L = norm(label);
        for (const [sub, field] of HEADER_MAP) if (L.indexOf(norm(sub)) >= 0)
          return label + sep + '{{' + field + '}}';
      }
    }
  }
  // 整格即标签（无分隔符）：直接整格替换为占位符
  const L = norm(v);
  for (const [sub, field] of HEADER_MAP) if (L.indexOf(norm(sub)) >= 0) return '{{' + field + '}}';
  return null;
}
function _colNum(s) { let n = 0; for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
function _colLet(n) { let s = ''; while (n > 0) { let r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }

// 把合并区域主单元格（左上角）设值；不要清空从属单元格，否则 ExcelJS 写回时主值也会丢失
function setMergeValue(ws, range, value) {
  const m = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
  if (!m) return;
  const c1 = _colNum(m[1]), r1 = parseInt(m[2], 10);
  ws.getCell(r1, c1).value = value;
}

/** 从源 xlsx 提取第一个内嵌图片（通常是 logo）及其单元格锚点，返回 {data:Buffer, ext, from:{col,row}, to:{col,row}} 或 null */
async function extractFirstImage(srcPath) {
  try {
    const z = await JSZip.loadAsync(fs.readFileSync(srcPath));
    // 找第一个 drawing XML
    const drawingName = Object.keys(z.files).find(n => /^xl\/drawings\/drawing\d+\.xml$/.test(n));
    if (!drawingName) return null;
    const drawingXml = await z.file(drawingName).async('string');
    // 取第一个带 a:blip 的 twoCellAnchor
    const anchorMatch = drawingXml.match(/<xdr:twoCellAnchor[^>]*>(?:(?!<xdr:twoCellAnchor).)*<a:blip[^>]*r:embed="([^"]+)"[\s\S]*?<\/xdr:twoCellAnchor>/);
    if (!anchorMatch) return null;
    const anchor = anchorMatch[0];
    const embedId = anchorMatch[1];
    const fromCol = parseInt((anchor.match(/<xdr:from>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>/) || [])[1], 10);
    const fromRow = parseInt((anchor.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/) || [])[1], 10);
    const toCol = parseInt((anchor.match(/<xdr:to>[\s\S]*?<xdr:col>(\d+)<\/xdr:col>/) || [])[1], 10);
    const toRow = parseInt((anchor.match(/<xdr:to>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/) || [])[1], 10);
    // 通过 drawing rels 找到 media 文件名
    const relsName = drawingName.replace('/drawings/', '/drawings/_rels/') + '.rels';
    const relsXml = await z.file(relsName).async('string');
    const relMatch = relsXml.match(new RegExp('<Relationship[^>]*Id="' + embedId + '"[^>]*Target="([^"]+)"'));
    if (!relMatch) return null;
    let mediaPath = relMatch[1].replace(/^\.\.\//, '');
    if (!mediaPath.startsWith('xl/')) mediaPath = 'xl/' + mediaPath;
    const mediaFile = z.file(mediaPath);
    if (!mediaFile) return null;
    const data = await mediaFile.async('nodebuffer');
    const ext = (mediaPath.match(/\.([^.]+)$/) || [,'png'])[1].toLowerCase();
    return { data, ext, from: { col: fromCol, row: fromRow }, to: { col: toCol, row: toRow } };
  } catch (e) {
    return null;
  }
}

// KEAS 模板结构特殊（收发人为上下区块而非 Label:Value），单独处理以保证 1:1
function transformKeas(ws, t) {
  // 1. 收发人/通知人区块：保留标签行，数据区置为对应占位符
  setMergeValue(ws, 'A2:G2', 'SHIPPER (发货人)');
  setMergeValue(ws, 'A3:G8', '{{shipper.name}}\n{{shipper.address}}');
  setMergeValue(ws, 'A9:G9', 'CONSIGNEE (收货人)');
  setMergeValue(ws, 'A10:G15', '{{consignee.name}}\n{{consignee.address}}');
  setMergeValue(ws, 'A16:G16', 'NOTIFY PARTY (通知人)');
  setMergeValue(ws, 'A17:G21', '{{notify.name}}\n{{notify.address}}');

  // 2. 右上单据编号区
  setMergeValue(ws, 'I2:J3', 'SHIPPING ORDER NO.');
  setMergeValue(ws, 'K2:M3', '{{invoiceNo}}');
  setMergeValue(ws, 'I4:J5', 'CARGO READY DATE');
  setMergeValue(ws, 'K4:M5', '{{invoiceDate}}');

  // 3. 船名航次 / 港口
  setMergeValue(ws, 'A23:G24', '{{vessel}}');
  setMergeValue(ws, 'A26:D27', '{{pol}}');
  setMergeValue(ws, 'E26:G27', '{{pol}}');
  setMergeValue(ws, 'A29:D30', '{{pod}}');
  setMergeValue(ws, 'E29:G30', '{{pod}}');

  // 4. 红色横幅 / 分栏标题
  setMergeValue(ws, 'A31:M31', 'PARTICULARS FURNISHED BY SHIPPER');
  setMergeValue(ws, 'A32:H33', 'BOOKING INFORMATION (委托信息)');
  setMergeValue(ws, 'I32:M33', 'SHIPPING INSTRUCTION INFORMATION');

  // 5. 明细区：保留 R34 主表头、R35 子表头；R36 起只留一行占位模式，多余样例行清空
  // 主占位行（引擎将按此模式复制到每条明细）
  setMergeValue(ws, 'A36:A45', '{{shippingMarks}}');
  setMergeValue(ws, 'B36:C36', '{{items.boxCount}}');
  setMergeValue(ws, 'D36:F36', '{{items.nameEn}}');
  ws.getCell('G36').value = '{{items.gw}}';
  ws.getCell('H36').value = '{{items.volume}}';
  ws.getCell('I36').value = '{{items.hsCode}}';
  ws.getCell('J36').value = null; // S/O NO. 留空
  setMergeValue(ws, 'K36:K36', '{{items.boxNo}}');
  ws.getCell('L36').value = null; // SEAL NO. 留空
  ws.getCell('M36').value = null; // SIZE 留空

  // 6. 清空 R37-R45 的样本科目数据（合并区域只清主单元格，避免 ExcelJS 写回丢主值 bug）
  const mergeMap = {};
  (ws.model.merges || []).forEach((m) => {
    const mm = m.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!mm) return;
    const c1 = _colNum(mm[1]), r1 = parseInt(mm[2], 10), c2 = _colNum(mm[3]), r2 = parseInt(mm[4], 10);
    for (let rr = r1; rr <= r2; rr++) {
      for (let cc = c1; cc <= c2; cc++) {
        mergeMap[rr + ',' + cc] = { top: r1, left: c1, master: rr === r1 && cc === c1 };
      }
    }
  });
  for (let r = 37; r <= 45; r++) {
    const row = ws.getRow(r);
    row.eachCell({ includeEmpty: true }, (c, col) => {
      const info = mergeMap[r + ',' + col];
      if (info && !info.master) return; // 只清主单元格，从属格由合并显示
      c.value = null;
    });
  }

  // 7. 清空明细区下方、VGM 区之前的空样例行（保留 VGM 区 keepFrom 起的内容）
  const nRows = ws.rowCount;
  const clearEnd = t.keepFrom ? (t.keepFrom - 1) : nRows;
  for (let r = 46; r <= clearEnd; r++) {
    const row = ws.getRow(r);
    let hasItemPh = false;
    row.eachCell({ includeEmpty: false }, (c) => { if (/\{\{\s*items\./.test(cellStr(c))) hasItemPh = true; });
    if (hasItemPh) continue;
    row.eachCell({ includeEmpty: true }, (c) => { c.value = null; });
    row.height = null;
  }

  // 8. 仅清理「被清空的样例行范围」内的悬空合并，保留表头/货描/VGM 结构合并
  const merges = (ws.model.merges || []).slice();
  for (const m of merges) {
    const mm = m.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!mm) continue;
    const r1 = parseInt(mm[2], 10), r2 = parseInt(mm[4], 10);
    if (r1 >= 46 && r2 <= clearEnd) {
      try { ws.unMergeCells(m); } catch (e) {}
    }
  }
  return ws;
}

// 整本预处理：剔除内嵌图片、清空共享公式簿记、把所有公式转为值（避免写回时报 shared formula 错误）
function prepWorkbook(wb) {
  try { if (wb._images) wb._images = []; } catch (e) {}
  try { if (wb._sharedFormulas) wb._sharedFormulas = {}; } catch (e) {}
  wb.eachSheet((ws) => {
    try { ws.images = []; } catch (e) {}
    if (ws._images) ws._images = [];
    if (ws._sharedFormulas) ws._sharedFormulas = {};
    const rc = ws.rowCount;
    for (let r = 1; r <= rc; r++) {
      const row = ws.getRow(r);
      row.eachCell({ includeEmpty: false }, (c) => {
        const v = c.value;
        if (v && typeof v === 'object' && (v.formula !== undefined || v.sharedFormula !== undefined)) {
          c.value = (v.result !== undefined && v.result !== null) ? v.result : '';
        }
      });
    }
  });
}

// ---------- 模板配置 ----------
const TEMPLATES = [
  { id: 'tpl_real_yafeng_air', name: '亚丰·美国空派装箱申报模板', kind: 'invoice', carrier: '亚丰空派',
    src: '1929234美国空派11箱.xlsx', sheet: '亚丰--模板', detailRow: 18, placeholderRows: 3 },
  { id: 'tpl_real_meixian_sea', name: '美线海运发票模板(亚丰2.0)', kind: 'invoice', carrier: '亚丰海运',
    src: '1963959海运发票最新模板.xlsx', sheet: '美线模板2.0', detailRow: 27, placeholderRows: 3 },
  { id: 'tpl_real_yafeng_sea', name: '亚丰海运装箱申报模板', kind: 'invoice', carrier: '亚丰海运',
    src: '模板.xlsx', sheet: '亚丰海运模板', detailRow: 1, placeholderRows: 3 },
  { id: 'tpl_real_yafeng_air2', name: '亚丰空运装箱申报模板', kind: 'invoice', carrier: '亚丰空运',
    src: '模板.xlsx', sheet: '亚丰空运运模板', detailRow: 1, placeholderRows: 3 },
  { id: 'tpl_real_zytd_booking', name: '中运通达FBA订单模板', kind: 'booking', carrier: '中运通达',
    src: '订单模板(中运通达).xlsx', sheet: '多个订单可以建立多个sheet', detailRow: 20, placeholderRows: 3 },
  { id: 'tpl_real_yuntu_booking', name: '云途B2B批量下单模板', kind: 'booking', carrier: '云途',
    src: 'B2B批量下单模板_云途.xlsx', sheet: '批量下单', detailRow: 1, placeholderRows: 3 },
  { id: 'tpl_real_yuntu_declare', name: '云途单票申报信息模板(B2B)', kind: 'declare', carrier: '云途',
    src: '云途单票申报信息模板(B2B).xlsx', sheet: '装箱明细（必填）', detailRow: 1, placeholderRows: 3 },
  { id: 'tpl_real_packing_fmt', name: '装箱单格式模板', kind: 'packing', carrier: '通用',
    src: '装箱单 格式.xls', sheet: 'Sheet1', detailRow: 6, placeholderRows: 3 },
  { id: 'tpl_real_maidan_declare', name: '买单要素模板', kind: 'declare', carrier: '通用',
    src: '买单要素模板.xls', sheet: 'Sheet1', detailRow: 4, placeholderRows: 3 },
  // —— 第二批：6 个真实订舱单模板（D:/模板/订舱单）——
  { id: 'tpl_real_aramex_booking', name: 'Aramex 国际航空托运书', kind: 'booking', carrier: 'Aramex',
    src: 'Aramex  Booking Form.xlsx', sheet: 'Sheet1', detailRow: 20, placeholderRows: 2, keepFrom: 23 },
  { id: 'tpl_real_booking_form', name: '通用订舱委托书(BOOKING FORM)', kind: 'booking', carrier: '通用',
    src: 'BOOKING  FORM.xls', sheet: 'Sheet1', detailRow: 24, placeholderRows: 3 },
  { id: 'tpl_real_chr_booking', name: 'CHR 深圳海运订舱单', kind: 'booking', carrier: 'CHR',
    src: 'CHR Shenzhen Booking Form.xls', sheet: 'BOOKING FORM-SEA ', detailRow: 32, placeholderRows: 3, keepFrom: 36 },
  { id: 'tpl_real_detrans_booking', name: 'DETRANS 空运委托书', kind: 'booking', carrier: 'DETRANS',
    src: 'DETRANS  BOOKING air -SA.xls', sheet: 'Booking Form', detailRow: 26, placeholderRows: 3, keepFrom: 36 },
  { id: 'tpl_real_geodis_booking', name: 'GEODIS 海运托运单(SI)', kind: 'booking', carrier: 'GEODIS',
    src: 'GEODIS BOOKING+SI form(1).xlsx', sheet: 'SHIPPING ORDER V8.0', detailRow: 38, placeholderRows: 3, keepFrom: 46 },
  { id: 'tpl_real_keas_booking', name: 'KEAS 订舱+SI+VGM表单(海运)', kind: 'booking', carrier: 'KEAS',
    src: 'New KEAS Booking_SI_VGM Form_Sea_南区Rev.xlsx', sheet: 'Booking_SI_VGM Form', detailRow: 35, placeholderRows: 3, keepFrom: 46 },
];

function resolveSrc(t) {
  let p = null;
  for (const d of SRC_DIRS) {
    const cand = path.join(d, t.src);
    if (fs.existsSync(cand)) { p = cand; break; }
  }
  if (!p) throw new Error('找不到源文件: ' + t.src);
  if (t.src.toLowerCase().endsWith('.xls')) {
    const tmp = path.join(OUT_DIR, '_conv_' + t.id + '.xlsx');
    const r = spawnSync(PY, ['tests/xls_to_xlsx.py', p, tmp], { encoding: 'utf8' });
    if (r.status !== 0) throw new Error('xls转换失败 ' + t.src + '\n' + (r.stderr || r.stdout));
    return tmp;
  }
  return p; // .xlsx 直接用原始文件（ExcelJS 可正常读），图片在写出后由 normalize 剥离
}

async function transform(wb, t) {
  prepWorkbook(wb);
  // 容错定位 sheet：精确 -> 去首尾空格 -> 忽略大小写去空格，避免真实样张 sheet 名带尾随空格导致失败
  let ws = wb.getWorksheet(t.sheet);
  if (!ws) {
    const want = String(t.sheet).trim().toLowerCase();
    wb.eachSheet((o) => { if (!ws && o.name && o.name.trim().toLowerCase() === want) ws = o; });
  }
  if (!ws) throw new Error('找不到 sheet: ' + t.sheet + ' in ' + t.src);
  // 仅保留目标 sheet（源文件常含多个 sheet，避免把无关整表也写进模板）
  const target = ws;
  const toRemove = [];
  wb.eachSheet((o) => { if (o !== target) toRemove.push(o.id); });
  toRemove.forEach((id) => { try { wb.removeWorksheet(id); } catch (e) {} });

  // KEAS 模板为「标签在上、数据区块在下」的特殊布局，单独走自定义转换
  if (t.id === 'tpl_real_keas_booking') {
    return transformKeas(ws, t);
  }

  const detailRow = t.detailRow;

  // 1) 表头区（detailRow 以上）：替换 "Label：Value" + 剔除图片公式
  for (let r = 1; r < detailRow; r++) {
    const row = ws.getRow(r);
    row.eachCell({ includeEmpty: false }, (c) => {
      let v = cellStr(c); if (!v) return;
      if (/DISPIMG|_xlfn/i.test(v)) { c.value = ''; return; }
      const rep = headerReplace(v);
      if (rep !== null) c.value = rep;
    });
  }

  // 2) 明细表头行：记录每列 -> items 字段；剔除图片列
  const hdrRow = ws.getRow(detailRow);
  const colMap = {};
  hdrRow.eachCell({ includeEmpty: true }, (c, col) => {
    const label = cellStr(c);
    if (label && /DISPIMG|_xlfn/i.test(label)) { c.value = ''; colMap[col] = null; return; }
    colMap[col] = label ? mapItem(label) : null;
  });

  // 3) 数据块 -> N 个占位行（保留表头样式）
  const dataStart = detailRow + 1;
  const nRows = ws.rowCount;
  const N = t.placeholderRows || 3;
  for (let i = 0; i < N; i++) {
    const row = ws.getRow(dataStart + i);
    row.height = hdrRow.height;
    hdrRow.eachCell({ includeEmpty: true }, (hc, col) => {
      const c = row.getCell(col);
      try { c.style = JSON.parse(JSON.stringify(hc.style || {})); } catch (e) { c.style = hc.style; }
      const path = colMap[col];
      c.value = (path ? '{{items.' + path + '}}' : '');
    });
  }

  // 4) 合计行：定位并替换为 {{totals.xxx}}
  let totalRow = -1;
  for (let r = detailRow; r <= nRows; r++) {
    const row = ws.getRow(r);
    let hit = false;
    row.eachCell({ includeEmpty: false }, (c) => { if (/合计|总计|TOTAL/i.test(cellStr(c))) hit = true; });
    if (hit) { totalRow = r; break; }
  }
  if (totalRow > 0) {
    const trow = ws.getRow(totalRow);
    trow.eachCell({ includeEmpty: true }, (c, col) => {
      const v = cellStr(c);
      if (/合计|总计|TOTAL/i.test(v)) return; // 标签文字保留
      const label = colMap[col];
      if (label && TOTAL_FIELDS.indexOf(label) >= 0) {
        const f = (label === 'ctns') ? 'boxCount' : label;
        c.value = '{{totals.' + f + '}}';
      } else if (/^=|^[\d]/.test(v)) {
        c.value = '';
      }
    });
  }

  // 5) 清空占位行之后、合计行(或 keepFrom 之前)的剩余数据行（值+样式），得到紧凑模板。
  //    keepFrom（可选）：明确“明细数据区之后、从该行起必须保留”的页脚起始行（空运提单的
  //    免责声明/签字/费用表等表单内容在货品行之下，不能一并清空）。
  //    注意：keepFrom 是“保留起点”，不是截断点；从 keepFrom 到源文件末尾的页脚都要保留。
  const clearEnd = (totalRow > 0) ? (totalRow - 1) : (t.keepFrom ? (t.keepFrom - 1) : nRows);
  for (let r = dataStart + N; r <= clearEnd; r++) {
    const row = ws.getRow(r);
    row.eachCell({ includeEmpty: true }, (c) => { c.value = null; try { c.style = {}; } catch (e) {} });
    row.height = null;
  }
  // 截断：保留 表头 + N 占位行 + 约 20 行空白录入行；若有合计行则至少保留到合计行；
  // 若明确 keepFrom，则保留 keepFrom 之后的所有页脚内容，不截断。
  const EXTRA = 20;
  let keep;
  if (totalRow > 0) {
    keep = totalRow;
  } else if (t.keepFrom) {
    keep = ws._rows ? ws._rows.length : nRows;
  } else {
    keep = Math.min(clearEnd, dataStart + N + EXTRA);
  }
  if (ws._rows && ws._rows.length > keep) ws._rows.length = keep;

  // 6) 清理合并：仅清理「真正被清空的填充行」范围内的合并，避免悬空合并；
  //    表头行(detailRow)与占位样例行（含货描表结构合并，如 Aramex 的 E20:F20/I20:K20、
  //    B21:C21 等）一律保留，否则会破坏模板的表格格式，导致预览/导出显示不全。
  const merges = (ws.model.merges || []).slice();
  const clearFrom = dataStart + N;
  for (const m of merges) {
    const top = parseInt(m.match(/^([A-Z]+)(\d+)/)[2], 10);
    if (top >= clearFrom && top <= clearEnd) {
      try { ws.unMergeCells(m); } catch (e) {}
    }
  }
  return ws;
}

(async () => {
  if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });
  const out = [];
  for (const t of TEMPLATES) {
    let srcPath = resolveSrc(t);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(srcPath);
    const ws = await transform(wb, t);
    const outPath = path.join(OUT_DIR, t.id + '.xlsx');
    const tmpPath = path.join(OUT_DIR, '_raw_' + t.id + '.xlsx');
    await wb.xlsx.writeFile(tmpPath);
    // 剥离内嵌图片/图形/批注，避免体积膨胀与读回报错
    const r = spawnSync(PY, ['tests/normalize_xlsx.py', tmpPath, outPath], { encoding: 'utf8' });
    if (r.status !== 0) throw new Error('normalize失败 ' + t.id + '\n' + (r.stderr || r.stdout));
    try { fs.unlinkSync(tmpPath); } catch (e) {}
    const b64 = fs.readFileSync(outPath).toString('base64');
    const entry = { id: t.id, name: t.name, kind: t.kind, carrier: t.carrier, builtin: false, fileBufB64: b64 };
    // 提取源文件 logo 图片（第一个内嵌图）供预览/导出复现
    const img = await extractFirstImage(srcPath);
    if (img) {
      entry.logo = {
        dataB64: img.data.toString('base64'),
        ext: img.ext,
        from: img.from,
        to: img.to
      };
      console.log('    logo ' + img.ext + ' ' + img.data.length + ' bytes @ ' + JSON.stringify(img.from) + '-' + JSON.stringify(img.to));
    }
    out.push(entry);
    console.log('  ✓ ' + t.id + '  (' + t.kind + ')  rows=' + ws.rowCount + '  ' + (b64.length / 1024).toFixed(1) + 'KB');
    // 清理临时转换文件
    if (srcPath.indexOf('_conv_') >= 0) { try { fs.unlinkSync(srcPath); } catch (e) {} }
  }
  const header = "/* 自动生成：由 tests/build_real_templates.js 从 14 个真实业务样张（8发票/申报/装箱 + 6订舱单）转换。请勿手改，重跑脚本即可。 */\n";
  const body = "window.TD = window.TD || {};\nwindow.TD.realTemplates = " + JSON.stringify(out, null, 2) + ";\n";
  fs.writeFileSync(OUT_JS, header + body);
  console.log('\n生成 ' + out.length + ' 个模板 -> ' + OUT_JS);
})().catch(e => { console.error('ERROR: ' + e.message + '\n' + e.stack); process.exit(1); });
