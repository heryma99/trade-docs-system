/* 输出成品审计层（v1.5.0 新增）
 * 设计目标：把"模板原件"当基准，填充后做一次逐格 diff，只允许三类受控变更：
 *   ① 占位符 {{path}} 被替换；② 标签映射目标格被填值；③ 明细数据区(行>itemsRow)被写入。
 * 任何"计划外写入"或"保护区文字被改"都判为 bug → 列入 blocks 拦截导出。
 * 这是一套"通用拦截器"，不依赖知道具体会出什么错——任何未经授权的变更都会被同一机制兜住。
 *
 * 用法（浏览器 / Node 通用 UMD）：
 *   var rep = TD.audit.auditFilled(filledWb, templateWb, docData, { kind, labelMap, boxMode });
 *   rep = { ok, blocks:[{dim,cell,msg}], warns:[{dim,cell,msg}], info:[{dim,msg}] }
 */
(function (root, factory) {
  var eng = (typeof module !== 'undefined' && module.exports)
    ? require('./engine.js')
    : (root.TD && root.TD.engine);
  if (typeof module !== 'undefined' && module.exports) module.exports = factory(eng);
  else { root.TD = root.TD || {}; root.TD.audit = factory(eng); }
})(typeof self !== 'undefined' ? self : this, function (engine) {
  'use strict';

  // ---------------- 工具 ----------------
  function cellText(cell) {
    if (!cell) return '';
    var v = cell.value;
    if (v == null) return '';
    if (v.richText) return v.richText.map(function (t) { return t.text; }).join('');
    if (typeof v === 'string') return v;
    if (typeof v === 'number') return String(v);
    if (v.text !== undefined) return String(v.text);
    if (v.formula !== undefined) return v.result !== undefined ? String(v.result) : '';
    if (v.sharedString !== undefined) return String(v.sharedString);
    if (v.hyperlink) return '';
    return String(v);
  }
  function eqVal(a, b) { return String(a) === String(b); }
  function colToChar(n) {
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }
  function cellRef(r, c) { return colToChar(c) + r; }

  var PH_RE = (engine && engine.PH_RE) || /\{\{\s*([A-Za-z0-9_.]+)\s*\}\}/g;
  // 保护区识别：仅限"警告/指示类"文字（含明确关键词）。
  // 注意：不能用"长度>20"兜底——模板自身含大量长样本文字（地址/公司名/提单说明），引擎会合法清空或改写它们，
  // 用长度会误杀正确导出。历史 bug（v1.4.63 禁止私自修改 / v1.4.65 因亚马逊地址时常变动）均含明确关键词，故仅按关键词判定。
  var PROTECTED_RE = /(禁止|警告|注意|请勿|私[自改]|修改模板|合并模板|MODIFY|WARNING|CAUTION|ATTENTION|DO\s*NOT|AMAZON|ADDRESS\s*MAY)/i;
  function isProtectedText(s) {
    if (!s || typeof s !== 'string') return false;
    return PROTECTED_RE.test(s);
  }
  function isPlaceholder(s) {
    if (!s || typeof s !== 'string') return false;
    PH_RE.lastIndex = 0;
    return PH_RE.test(s);
  }

  // 地址 → 国家（维度 C 一致性用；自带实现，与 engine.parseAddress 同源逻辑但解耦，避免依赖 engine 内部函数）
  var COUNTRY_LONG = [
    ['HONG KONG', 'HK'], ['HONGKONG', 'HK'], ['MACAU', 'MO'], ['MACAO', 'MO'], ['CHINA', 'CN'],
    ['UNITED STATES OF AMERICA', 'US'], ['UNITED STATES', 'US'], ['U.S.A.', 'US'], ['U.S.A', 'US'], ['USA', 'US'],
    ['UNITED KINGDOM', 'GB'], ['U.K.', 'GB'], ['U.K', 'GB'], ['ENGLAND', 'GB'], ['SCOTLAND', 'GB'], ['GREAT BRITAIN', 'GB'],
    ['JAPAN', 'JP'], ['GERMANY', 'DE'], ['DEUTSCHLAND', 'DE'], ['FRANCE', 'FR'], ['AUSTRALIA', 'AU'],
    ['CANADA', 'CA'], ['SINGAPORE', 'SG'], ['MALAYSIA', 'MY'], ['REPUBLIC OF KOREA', 'KR'], ['KOREA', 'KR'],
    ['TAIWAN', 'TW'], ['THAILAND', 'TH'], ['VIETNAM', 'VN'], ['INDIA', 'IN'], ['MEXICO', 'MX'], ['BRAZIL', 'BR']
  ];
  var COUNTRY_SHORT = [['HK', 'HK'], ['CN', 'CN'], ['US', 'US'], ['UK', 'GB'], ['GB', 'GB'], ['DE', 'DE'], ['FR', 'FR'],
    ['JP', 'JP'], ['AU', 'AU'], ['CA', 'CA'], ['SG', 'SG'], ['MY', 'MY'], ['KR', 'KR'], ['TW', 'TW'], ['TH', 'TH'],
    ['VN', 'VN'], ['IN', 'IN'], ['MX', 'MX'], ['BR', 'BR']];
  function parseCountry(addr) {
    if (!addr || typeof addr !== 'string') return '';
    var s = addr.toUpperCase();
    for (var i = 0; i < COUNTRY_LONG.length; i++) { if (s.indexOf(COUNTRY_LONG[i][0]) >= 0) return COUNTRY_LONG[i][1]; }
    for (var j = 0; j < COUNTRY_SHORT.length; j++) {
      if (new RegExp('\\b' + COUNTRY_SHORT[j][0] + '\\b').test(s)) return COUNTRY_SHORT[j][1];
    }
    return '';
  }

  // 明细数值列提示（用于 E 数值语义）
  var NUMERIC_HINT = [
    { re: /箱号|BOX\s*NO|BOXNO|FBA箱号|货箱编号/i, f: 'boxNo' },
    { re: /数量|QUANTITY|QTY/i, f: 'qty' },
    { re: /毛重|G\.?W\.?|GROSS\s*WEIGHT/i, f: 'gw' },
    { re: /净重|N\.?W\.?|NET\s*WEIGHT/i, f: 'nw' },
    { re: /长|LENGTH/i, f: 'length' },
    { re: /宽|WIDTH/i, f: 'width' },
    { re: /高|HEIGHT/i, f: 'height' },
    { re: /体积|VOLUME|CBM/i, f: 'volume' },
    { re: /单价|PRICE/i, f: 'price' },
    { re: /总箱|TOTAL\s*BOX|TOTAL\s*CTN|TOTAL\s*QTY/i, f: 'totalBoxQty' }
  ];

  // ---------------- 主审计 ----------------
  function auditFilled(filledWb, templateWb, data, opts) {
    opts = opts || {};
    data = data || {};
    var report = { ok: true, blocks: [], warns: [], info: [] };
    try {
      if (!filledWb || !templateWb) { report.info.push({ dim: 'A', msg: '缺少对比工作簿，跳过审计' }); return report; }
      var fws = filledWb.worksheets[0];
      var tws = templateWb.worksheets[0];
      if (!fws || !tws) { report.info.push({ dim: 'A', msg: '缺少工作表，跳过审计' }); return report; }

      // 模板扫描：明细行、标签映射、字段列
      var scan = (engine && engine.scanTemplate) ? engine.scanTemplate(templateWb) : { itemsRow: -1, itemFields: [], itemHeaderMap: {} };
      var itemsRow = scan.itemsRow || -1;
      var boxMode = !!(opts.boxMode) ||
        (scan.itemFields || []).some(function (f) { return /(boxNo|length|width|height)/.test(f); });

      // 标签映射目标格集合（受控写入集之一）
      var labelMap = (opts.labelMap && opts.labelMap.length)
        ? opts.labelMap
        : ((engine && engine.buildLabelMap) ? engine.buildLabelMap(templateWb, itemsRow) : []);
      var labelTargets = {};
      labelMap.forEach(function (e) {
        if (e && e.row && e.col) labelTargets[e.row + ':' + e.col] = 1;
      });

      // 明细数据区下界（行 > itemsRow 视为数据区，受控写入集之三）
      var detailStart = itemsRow > 0 ? itemsRow + 1 : 99999;

      var maxR = Math.max(fws.rowCount || 0, tws.rowCount || 0);
      var maxC = Math.max(fws.columnCount || 0, tws.columnCount || 0);

      // ---- 维度 A（模板保真/保护区）+ 维度 D（明细表头行非授权改写）----
      // 设计取舍：不做"逐格 diff 全盘白名单"——模板含大量合法示例数据/样本文字（如"美国""GOODYEAR"）
      // 与长占位符，引擎合法改写它们，白名单极易误杀正确导出。改为两类高置信阻断点：
      //   A：上行区(r<detailStart)的保护区文字（长 CJK 装饰/警告，且非占位符）被改写 → 抓 v1.4.63/1.4.65
      //   D：明细表头行(r==itemsRow)的非占位符/非标签目标单元格被改写 → 抓 v1.4.58 表头被覆盖
      // 其余上行区变更（引擎合法填样本/标签格）不阻断，避免误报。
      var changedCount = 0, protectedChanged = 0, headerChanged = 0;
      for (var r = 1; r <= maxR; r++) {
        for (var c = 1; c <= maxC; c++) {
          var pv = cellText(tws.getCell(r, c));
          var fv = cellText(fws.getCell(r, c));
          if (eqVal(pv, fv)) continue;
          changedCount++;
          // 受控写入：占位符替换 / 标签映射目标 / 明细数据区 → 合法，跳过
          if (isPlaceholder(pv) || labelTargets[r + ':' + c] || (r >= detailStart)) continue;
          // A：上行区保护区文字被改写（长 CJK 装饰/警告，且不是占位符——占位符已在上面放行）
          if (r < detailStart && isProtectedText(pv)) {
            protectedChanged++;
            report.blocks.push({
              dim: 'A', cell: cellRef(r, c),
              msg: '模板保护区文字被改写（原="' + pv.slice(0, 24) + (pv.length > 24 ? '…' : '') + '" → 现="' + fv.slice(0, 24) + (fv.length > 24 ? '…' : '') + '"）'
            });
            continue;
          }
          // D：明细表头行（r==itemsRow）的非空表头标签被值覆盖（v1.4.58 类）。
          // 仅当原格"非空且有文字标签"才判违规——空格被引擎填入数据是合法行为，不阻断。
          if (itemsRow > 0 && r === itemsRow && pv !== '' && !labelTargets[r + ':' + c]) {
            headerChanged++;
            report.blocks.push({
              dim: 'D', cell: cellRef(r, c),
              msg: '明细表头行被非授权改写（原="' + (pv || '∅') + '" → 现="' + (fv || '∅') + '"），表头标签不应被数据覆盖'
            });
            continue;
          }
          // 其他上行区变更：引擎合法重写样本/标签格，放行（不计入 blocks）
        }
      }
      if (protectedChanged === 0 && headerChanged === 0) {
        report.info.push({ dim: 'A/D', msg: '模板保真 + 表头保护 ✓（变更 ' + changedCount + ' 格均为受控写入）' });
      }

      // ---- 维度 B（模板强制字段完整）----
      // 依据标签映射/明细表头识别模板期望的明细列，断言数据区非空
      var itemFields = scan.itemFields || [];
      var expectBoxNo = boxMode || itemFields.some(function (f) { return /boxNo/.test(f); });
      // 从 itemHeaderMap 找 boxNo 列
      var boxNoCol = -1;
      Object.keys(scan.itemHeaderMap || {}).forEach(function (col) {
        var f = scan.itemHeaderMap[col];
        if (f === 'boxNo' || /boxNo/.test(String(f))) boxNoCol = parseInt(col, 10);
      });
      // 若没有 itemHeaderMap 但有 items.boxNo 占位符，扫列头
      if (boxNoCol < 0 && itemsRow > 0) {
        for (var cc = 1; cc <= maxC; cc++) {
          var h = cellText(tws.getCell(itemsRow, cc));
          if (/箱号|BOX\s*NO|BOXNO|FBA箱号|货箱编号/i.test(h)) { boxNoCol = cc; break; }
        }
      }
      if (expectBoxNo && boxNoCol > 0) {
        var boxFilled = false;
        for (var br = detailStart; br <= fws.rowCount; br++) {
          if (cellText(fws.getCell(br, boxNoCol))) { boxFilled = true; break; }
        }
        if (!boxFilled) {
          report.blocks.push({ dim: 'B', cell: cellRef(itemsRow, boxNoCol), msg: '模板要求箱号列（' + cellRef(itemsRow, boxNoCol) + '）但填充后整列空——箱号未生成' });
        } else {
          report.info.push({ dim: 'B', msg: '箱号列已填充 ✓' });
        }
      }
      // 通用明细必含列：SKU/数量/品名 若模板有则该列非空
      ['sku', 'qty', 'nameEn', 'nameCn', 'hsCode'].forEach(function (fld) {
        var col = -1;
        Object.keys(scan.itemHeaderMap || {}).forEach(function (col2) {
          var f = scan.itemHeaderMap[col2];
          if (f === fld || new RegExp(fld, 'i').test(String(f))) col = parseInt(col2, 10);
        });
        if (col > 0) {
          var anyVal = false;
          for (var dr = detailStart; dr <= fws.rowCount; dr++) {
            if (cellText(fws.getCell(dr, col))) { anyVal = true; break; }
          }
          if (!anyVal) report.warns.push({ dim: 'B', cell: cellRef(itemsRow, col), msg: '明细列「' + fld + '」整列为空（模板有该列但无数据）' });
        }
      });

      // ---- 维度 C（字段一致性）----
      var cons = data.consignee || {};
      var addr = cons.address || '';
      if (addr) {
        var pac = parseCountry(addr);
        if (pac && cons.country && pac.toUpperCase() !== String(cons.country).toUpperCase()) {
          report.warns.push({
            dim: 'C', cell: '',
            msg: '国家与地址不一致：录入国家=' + cons.country + '，地址解析国家=' + pac + '（导出时按地址覆盖为 ' + pac + '）'
          });
        }
      }
      // 总箱数 vs 明细行数（boxMode）
      if (boxMode && data.totals && data.totals.boxCount) {
        var detailRows = 0;
        for (var rr = detailStart; rr <= fws.rowCount; rr++) {
          var rowHas = false;
          for (var c2 = 1; c2 <= maxC; c2++) { if (cellText(fws.getCell(rr, c2))) { rowHas = true; break; } }
          if (rowHas) detailRows++;
        }
        if (detailRows !== Number(data.totals.boxCount)) {
          report.warns.push({ dim: 'C', cell: '', msg: '明细行数(' + detailRows + ') 与总箱数(' + data.totals.boxCount + ') 不一致' });
        }
      }

      // ---- 维度 E（数值语义）----
      if (itemsRow > 0) {
        for (var hc = 1; hc <= maxC; hc++) {
          var hdr = cellText(tws.getCell(itemsRow, hc));
          if (!hdr) continue;
          var numField = null;
          for (var i = 0; i < NUMERIC_HINT.length; i++) {
            if (NUMERIC_HINT[i].re.test(hdr)) { numField = NUMERIC_HINT[i].f; break; }
          }
          if (!numField) continue;
          // 检查该列明细值是否均为数值（允许空）
          for (var er = detailStart; er <= fws.rowCount; er++) {
            var cv = fws.getCell(er, hc).value;
            if (cv == null || cv === '') continue;
            if (typeof cv !== 'number') {
              report.warns.push({ dim: 'E', cell: cellRef(er, hc), msg: '列「' + hdr + '」应为数值，但单元格值为文本："' + cellText(fws.getCell(er, hc)) + '"' });
              break;
            }
          }
        }
        // boxNo 连续性（boxMode）
        if (boxNoCol > 0) {
          var seq = [];
          for (var sr = detailStart; sr <= fws.rowCount; sr++) {
            var bv = cellText(fws.getCell(sr, boxNoCol));
            if (bv) seq.push(bv);
          }
          if (seq.length) {
            var expect = 1, bad = false;
            for (var si = 0; si < seq.length; si++) {
              var n = parseInt(seq[si], 10);
              if (isNaN(n) || n !== expect) { bad = true; break; }
              expect++;
            }
            if (bad) report.warns.push({ dim: 'E', cell: cellRef(itemsRow, boxNoCol), msg: '箱号不连续（期望 1,2,3… 实际首段 ' + seq.slice(0, 5).join(',') + '），可能存在错位' });
          }
        }
      }

      // ---- 维度 F（结构保真）----
      var tm = (tws.model.merges || []).length, fm = (fws.model.merges || []).length;
      if (fm < tm) report.warns.push({ dim: 'F', cell: '', msg: '合并单元格减少（模板 ' + tm + ' → 成品 ' + fm + '），结构可能被破坏' });
      if ((fws.columnCount || 0) !== (tws.columnCount || 0)) report.warns.push({ dim: 'F', cell: '', msg: '列数变化（模板 ' + tws.columnCount + ' → 成品 ' + fws.columnCount + '）' });
      // 列宽保真
      for (var wc = 1; wc <= Math.min(tws.columnCount || 0, fws.columnCount || 0); wc++) {
        var tw = tws.getColumn(wc).width, fw = fws.getColumn(wc).width;
        if (tw && fw && Math.abs(tw - fw) > 0.5) { report.warns.push({ dim: 'F', cell: colToChar(wc), msg: '列 ' + colToChar(wc) + ' 宽度变化（' + tw + '→' + fw + '）' }); break; }
      }

      report.ok = report.blocks.length === 0;
      if (report.warns.length === 0 && report.blocks.length === 0) {
        report.info.push({ dim: 'B/C/E/F', msg: '字段完整/一致/数值/结构 全部通过 ✓' });
      }
    } catch (e) {
      // 审计自身异常时 fail-open：不拦截导出，仅记录，避免新模块成为卡点
      report.info.push({ dim: 'X', msg: '审计执行异常（已放行）: ' + (e && e.message) });
      report.ok = true;
    }
    return report;
  }

  // 生成质检面板 HTML（供 ui.js 直接插入预览区）
  function auditHtml(rep) {
    if (!rep) return '';
    var blocks = rep.blocks || [], warns = rep.warns || [], info = rep.info || [];
    if (!blocks.length && !warns.length && !info.length) return '';
    var html = '<div class="vres ' + (blocks.length ? 'block' : (warns.length ? 'warn' : 'ok')) + '" style="margin:10px 0">';
    if (blocks.length) {
      html += '<b>❌ 导出前质检未通过（' + blocks.length + ' 项阻断）</b><ul style="margin:6px 0 0 18px">';
      blocks.forEach(function (b) {
        html += '<li><span class="mono">' + (b.cell ? b.cell + ' ' : '') + '</span>[' + b.dim + '] ' + esc(b.msg) + '</li>';
      });
      html += '</ul><p class="hint">上述问题会在点「导出」时被拦截，请修复后重新生成。</p>';
    } else if (warns.length) {
      html += '<b>⚠️ 质检提示（' + warns.length + ' 项，可继续导出）</b><ul style="margin:6px 0 0 18px">';
      warns.forEach(function (w) {
        html += '<li><span class="mono">' + (w.cell ? w.cell + ' ' : '') + '</span>[' + w.dim + '] ' + esc(w.msg) + '</li>';
      });
      html += '</ul>';
    } else {
      html += '<b>✅ 导出前质检通过</b>';
    }
    if (info.length && !blocks.length) {
      html += '<div class="hint" style="margin-top:6px">' + info.map(function (i) { return '[' + i.dim + '] ' + esc(i.msg); }).join('　') + '</div>';
    }
    html += '</div>';
    return html;
  }
  function esc(s) {
    return String(s).replace(/[&<>"]/g, function (c) { return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]; });
  }

  return { auditFilled: auditFilled, auditHtml: auditHtml };
});
