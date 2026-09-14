/* 商单易 · 模板保真核验核心（v1.0）
 * 作用：把「上传的模板」与「系统导出的成品」逐单元格对比，并自动区分两类差异：
 *   A 类（预期填充）：模板单元格是 {{占位符}} / 明细占位符 → 成品被换成实际值，属正常，不计入问题
 *   B 类（保真差异）：样式(字体/底色/边框/对齐/数字格式)、合并区、列宽、行高 不一致 → 需要关注
 * 纯逻辑、无 DOM 依赖，浏览器与 Node 均可复用（便于自测）。
 */
(function (global) {
  'use strict';

  var RE_PH = /\{\{/;                 // 任意占位符
  var RE_ITEM_PH = /\{\{\s*items\./;  // 明细占位符

  /* ---------- 工具 ---------- */
  function col2num(s) {
    var n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n;
  }
  function num2col(n) {
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    return s;
  }
  function addr(r, c) { return num2col(c) + r; }

  function colorKey(c) {
    if (!c) return '';
    if (c.argb) return String(c.argb).toUpperCase();
    if (c.theme !== undefined && c.theme !== null) return 'theme' + c.theme + (c.tint ? ('/' + c.tint) : '');
    if (c.indexed !== undefined && c.indexed !== null) return 'idx' + c.indexed;
    return '';
  }
  function sideKey(s) { return s ? ((s.style || 'none') + ':' + colorKey(s.color)) : 'none'; }

  /* 把 ExcelJS 的 style 归一化成可比较字符串（只取会影响观感的项，避免假差异） */
  function normStyle(st) {
    st = st || {};
    var f = st.font || {}, fill = st.fill || {}, b = st.border || {}, al = st.alignment || {};
    return [
      'F[' + [f.name || '', f.size || '', f.bold ? 1 : 0, f.italic ? 1 : 0, f.underline ? 1 : 0, colorKey(f.color)].join('|') + ']',
      'L[' + [fill.type || '', fill.pattern || '', colorKey(fill.fgColor)].join('|') + ']',
      'B[' + [sideKey(b.top), sideKey(b.right), sideKey(b.bottom), sideKey(b.left)].join('|') + ']',
      'A[' + [al.horizontal || '', al.vertical || '', al.wrapText ? 1 : 0, al.textRotation || 0].join('|') + ']',
      'N[' + (st.numFmt || '') + ']'
    ].join(';');
  }

  /* 单元格文本（兼容 富文本 / 公式 / 超链接 / 日期） */
  function cellText(cell) {
    if (!cell) return '';
    var v = cell.value;
    if (v === null || v === undefined) return '';
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map(function (t) { return t.text || ''; }).join('');
      if (v.formula !== undefined) return '=' + String(v.formula);   // 只比公式本体，不比结果（结果随数据变）
      if (v.sharedFormula !== undefined) return '=' + String(v.sharedFormula);
      if (v.text !== undefined) return String(v.text);
      if (v instanceof Date) return v.toISOString().slice(0, 10);
      if (v.error) return String(v.error);
      return JSON.stringify(v);
    }
    if (typeof v === 'number') return String(v);
    return String(v);
  }

  /* 合并区集合（归一化 "A1:B2"） */
  function mergesOf(ws) {
    var out = [];
    try {
      var m = (ws.model && ws.model.merges) || [];
      for (var i = 0; i < m.length; i++) out.push(String(m[i]).toUpperCase());
    } catch (e) {}
    return out.sort();
  }

  /* 行签名：该行所有非空单元格文本（占位符视作空，便于「模板槽位行」与「成品数据行」可比） */
  function rowSignature(row) {
    var parts = [];
    try {
      row.eachCell({ includeEmpty: false }, function (c) {
        var t = cellText(c).trim();
        if (!t) return;
        if (RE_PH.test(t)) return;        // 占位符不算签名内容
        parts.push(t);
      });
    } catch (e) {}
    return parts.join('\u0001');
  }

  /* 找模板里第一处明细占位行 */
  function findItemRow(ws) {
    var found = -1;
    try {
      ws.eachRow({ includeEmpty: false }, function (row, rn) {
        if (found !== -1) return;
        row.eachCell({ includeEmpty: false }, function (c) {
          if (found === -1 && RE_ITEM_PH.test(cellText(c))) found = rn;
        });
      });
    } catch (e) {}
    return found;
  }

  /* 行「样式签名」：该行各单元格归一化样式（不含值）。空白模板与成品的值必然不同，
   * 但未插行区域的样式应一一对应，故用样式签名做对齐。 */
  var EMPTY_STYLE = null;
  function rowStyleSig(row) {
    if (EMPTY_STYLE === null) EMPTY_STYLE = normStyle({});
    var parts = [];
    try {
      row.eachCell({ includeEmpty: true }, function (c, cn) {
        var s = normStyle(c.style);
        parts.push(cn + ':' + (s === EMPTY_STYLE ? '' : s));
      });
    } catch (e) {}
    return parts.join('|');
  }

  /* 自动判定行对齐方式 */
  function detectAlign(tplWs, outWs) {
    var tMax = tplWs.rowCount || 0, oMax = outWs.rowCount || 0;
    var itemRow = findItemRow(tplWs);

    /* 模式 1：模板有 {{items.*}} → 明细槽位行即插行分界，delta = 行数差 */
    if (itemRow > 0) {
      return {
        mode: 'placeholder', alignBase: itemRow, delta: oMax - tMax, itemRow: itemRow,
        tplRowCount: tMax, outRowCount: oMax, pre: -1, score: -1
      };
    }

    /* 模式 2：无占位符 → 样式签名前缀匹配 + 在后半段搜索最佳 delta */
    var tSig = [], oSig = [], i, j;
    for (i = 1; i <= tMax; i++) tSig[i] = rowStyleSig(tplWs.getRow(i));
    for (j = 1; j <= oMax; j++) oSig[j] = rowStyleSig(outWs.getRow(j));

    var pre = 0, lim = Math.min(tMax, oMax);
    while (pre + 1 <= lim && tSig[pre + 1] === oSig[pre + 1]) pre++;

    var best = { delta: oMax - tMax, score: -1 };
    var lo = -Math.min(10, tMax), hi = Math.min(40, Math.max(10, oMax - tMax + 10));
    for (var d = lo; d <= hi; d++) {
      var sc = 0;
      for (var r = pre + 1; r <= tMax; r++) {
        var rr = r + d;
        if (rr >= pre + 1 && rr <= oMax && tSig[r] === oSig[rr]) sc++;
      }
      if (sc > best.score) best = { delta: d, score: sc };
    }
    return {
      mode: 'stylesig', alignBase: pre, delta: best.delta, itemRow: -1,
      tplRowCount: tMax, outRowCount: oMax, pre: pre, score: best.score
    };
  }

  /* ---------- 主流程：对比两个 workbook ---------- */
  function compareWorkbooks(tplWb, outWb, opts) {
    opts = opts || {};
    var diffs = [];
    var align = [];
    var tplNames = tplWb.worksheets.map(function (w) { return w.name; });
    var outNames = outWb.worksheets.map(function (w) { return w.name; });

    var sheetsOnlyTpl = tplNames.filter(function (n) { return outNames.indexOf(n) < 0; });
    var sheetsOnlyOut = outNames.filter(function (n) { return tplNames.indexOf(n) < 0; });
    sheetsOnlyOut.forEach(function (n) {
      diffs.push({ sheet: n, kind: 'sheet', level: 'B', desc: '成品多出 sheet「' + n + '」（模板没有）' });
    });

    /* 只对比「模板第一个 sheet」—— 引擎只读 worksheets[0]，这正是实际参与生成的那张 */
    var tplWs = tplWb.worksheets[0];
    if (!tplWs) return { ok: false, error: '模板没有任何 sheet' };
    var outWs = null;
    for (var i = 0; i < outWb.worksheets.length; i++) {
      if (outWb.worksheets[i].name === tplWs.name) { outWs = outWb.worksheets[i]; break; }
    }
    if (!outWs) outWs = outWb.worksheets[0];
    var sheetName = tplWs.name;

    if (outWs.name !== tplWs.name) {
      diffs.push({ sheet: sheetName, kind: 'sheet', level: 'B', desc: '首个 sheet 名称不一致：模板「' + tplWs.name + '」 vs 成品「' + outWs.name + '」' });
    }

    /* 1) 计算行对齐（两种模式，见 detectAlign）：
     *    · placeholder：模板含 {{items.*}} → 明细槽位行即插行分界
     *    · stylesig   ：空白模板（走表头标签映射填充）→ 样式签名前缀匹配 + delta 搜索 */
    var AL = detectAlign(tplWs, outWs);
    var tplItemRow = AL.itemRow;
    var tplRowCount = AL.tplRowCount;
    var outRowCount = AL.outRowCount;
    var delta = AL.delta;
    var alignBase = AL.alignBase;
    function mapRow(r) { return (delta !== 0 && r > alignBase) ? (r + delta) : r; }
    align.push({
      sheet: sheetName, mode: AL.mode, tplItemRow: tplItemRow, tplRowCount: tplRowCount,
      outRowCount: outRowCount, delta: delta, alignBase: alignBase,
      prefixMatched: AL.pre, matchScore: AL.score
    });

    /* 2) 列宽对比（引擎注释表明「只增不减」，故此处会暴露放宽的列） */
    var maxCol = Math.max(tplWs.columnCount || 0, outWs.columnCount || 0, 1);
    var widthDiffs = [];
    for (var c = 1; c <= maxCol; c++) {
      var tw = tplWs.getColumn(c).width, ow = outWs.getColumn(c).width;
      var tn = (tw === undefined || tw === null) ? null : Math.round(tw * 100) / 100;
      var on = (ow === undefined || ow === null) ? null : Math.round(ow * 100) / 100;
      if (tn !== on) {
        widthDiffs.push({ col: num2col(c), tpl: tn, out: on });
        diffs.push({
          sheet: sheetName, kind: 'width', level: 'B', addr: num2col(c) + '列',
          desc: '列宽不同：模板 ' + tn + ' → 成品 ' + on
        });
      }
    }

    /* 3) 行高对比（按对齐后的行号） */
    var maxRow = Math.max(tplRowCount, outRowCount);
    for (var r2 = 1; r2 <= maxRow; r2++) {
      var tr = tplWs.getRow(r2), orow = outWs.getRow(mapRow(r2));
      var th = (tr && tr.height) ? Math.round(tr.height * 100) / 100 : null;
      var oh = (orow && orow.height) ? Math.round(orow.height * 100) / 100 : null;
      if (th !== oh) {
        diffs.push({
          sheet: sheetName, kind: 'height', level: 'B', addr: r2 + '行',
          desc: '行高不同：模板 ' + th + ' → 成品 ' + oh
        });
      }
    }

    /* 4) 合并区对比 */
    var tplM = mergesOf(tplWs), outM = mergesOf(outWs);
    var onlyTpl = tplM.filter(function (x) { return outM.indexOf(x) < 0; });
    var onlyOut = outM.filter(function (x) { return tplM.indexOf(x) < 0; });
    onlyTpl.forEach(function (m) {
      diffs.push({ sheet: sheetName, kind: 'merge', level: 'B', addr: m, desc: '合并区仅在模板存在（成品丢失）：' + m });
    });
    /* 成品多出的合并区：若其行号落在「明细扩展区」之内，视为插行产生的合理位移 */
    onlyOut.forEach(function (m) {
      var rm = parseInt(String(m).replace(/^[A-Z]+/, ''), 10);
      var lvl = (delta !== 0 && rm > alignBase) ? 'A' : 'B';
      diffs.push({
        sheet: sheetName, kind: 'merge', level: lvl, addr: m,
        desc: (lvl === 'A' ? '合并区随明细插行位移（预期）：' : '合并区仅在成品存在（模板没有）：') + m
      });
    });

    /* 5) 逐单元格：样式 + 值（按对齐后的行号） */
    var cellChecked = 0, cellStyleDiff = 0, cellValueDiff = 0, valueExpected = 0;
    for (var rr = 1; rr <= tplRowCount; rr++) {
      var tRow = tplWs.getRow(rr);
      var fr = mapRow(rr);
      var oRow = outWs.getRow(fr);
      var maxC = Math.max((tRow.cellCount || 0), (oRow.cellCount || 0), 0);
      for (var cc = 1; cc <= maxC; cc++) {
        var tc = tRow.getCell(cc), oc = oRow.getCell(cc);
        cellChecked++;
        var ts = normStyle(tc.style), os = normStyle(oc.style);
        var tv = cellText(tc), ov = cellText(oc);
        var tIsPh = RE_PH.test(tv);
        var tHasVal = tv !== '', oHasVal = ov !== '';

        /* 样式差异 */
        if (ts !== os) {
          cellStyleDiff++;
          diffs.push({
            sheet: sheetName, kind: 'style', level: 'B',
            addr: addr(rr, cc), addrOut: addr(fr, cc),
            desc: '样式不同 @ ' + addr(rr, cc) + (fr !== rr ? ('→' + addr(fr, cc)) : '') + '：' + styleDiffBrief(ts, os)
          });
          continue;   // 样式已报，值差异不再重复报同一格
        }

        /* 值差异：模板含占位符 → A 类（预期填充） */
        if (tv !== ov) {
          if (tIsPh) {
            valueExpected++;
            diffs.push({
              sheet: sheetName, kind: 'value', level: 'A',
              addr: addr(rr, cc), addrOut: addr(fr, cc),
              desc: '占位符已填充（预期）：' + brief(tv) + ' → ' + brief(ov)
            });
          } else if (tHasVal !== oHasVal || (tHasVal && oHasVal)) {
            cellValueDiff++;
            diffs.push({
              sheet: sheetName, kind: 'value', level: 'B',
              addr: addr(rr, cc), addrOut: addr(fr, cc),
              desc: '值不同 @ ' + addr(rr, cc) + (fr !== rr ? ('→' + addr(fr, cc)) : '') + '：模板 ' + brief(tv) + ' → 成品 ' + brief(ov)
            });
          }
        }
      }
    }

    /* 5.5) 明细扩展「新增行」的样式检查：引擎插行后必须把模板明细行的样式(含边框)整份复制过去。
     *      这里以模板的明细槽位行为期望，逐列比对新行的样式——不一致即为真问题（明细区丢边框）。 */
    if (delta > 0) {
      var slotRow = (tplItemRow > 0) ? tplItemRow : alignBase;
      var slotStyles = [], slotCols = [];
      tplWs.getRow(slotRow).eachCell({ includeEmpty: true }, function (c, cn) {
        slotCols.push(cn); slotStyles[cn] = normStyle(c.style);
      });
      for (var nr = alignBase + 1; nr <= alignBase + delta; nr++) {
        var nRowObj = outWs.getRow(nr);
        slotCols.forEach(function (cn) {
          var ns = normStyle(nRowObj.getCell(cn).style);
          if (ns !== slotStyles[cn]) {
            diffs.push({
              sheet: sheetName, kind: 'style', level: 'B',
              addr: addr(nr, cn), addrOut: addr(nr, cn),
              desc: '明细新增行未继承模板槽位行样式 @ ' + addr(nr, cn) + '（参照模板第' + slotRow + '行）：' + styleDiffBrief(slotStyles[cn], ns)
            });
          }
        });
      }
    }

    /* 6) 计数 */
    var stats = { total: diffs.length, A: 0, B: 0, byKind: {} };
    diffs.forEach(function (d) {
      if (d.level === 'A') stats.A++; else stats.B++;
      stats.byKind[d.kind] = (stats.byKind[d.kind] || 0) + 1;
    });
    stats.cellChecked = cellChecked;
    stats.styleDiff = cellStyleDiff;
    stats.valueDiffReal = cellValueDiff;
    stats.valueExpected = valueExpected;
    stats.widthDiff = widthDiffs.length;
    stats.sheetsOnlyTpl = sheetsOnlyTpl;
    stats.sheetsOnlyOut = sheetsOnlyOut;

    return {
      ok: true, sheet: sheetName, diffs: diffs, stats: stats,
      align: align, widthDiffs: widthDiffs,
      tplSheets: tplNames, outSheets: outNames
    };
  }

  function styleDiffBrief(a, b) {
    var la = a.split(';'), lb = b.split(';'), out = [];
    var nameMap = { F: '字体', L: '底色', B: '边框', A: '对齐', N: '数字格式' };
    for (var i = 0; i < Math.max(la.length, lb.length); i++) {
      if (la[i] !== lb[i]) {
        var k = (la[i] || lb[i] || '').charAt(0);
        out.push((nameMap[k] || k) + ' [' + (la[i] || '-') + ' → ' + (lb[i] || '-') + ']');
      }
    }
    return out.join('，');
  }
  function brief(s, n) {
    s = String(s == null ? '' : s).replace(/\s+/g, ' ');
    n = n || 40;
    return s.length > n ? (s.slice(0, n) + '…') : (s === '' ? '(空)' : s);
  }

  var api = {
    compareWorkbooks: compareWorkbooks,
    normStyle: normStyle,
    cellText: cellText,
    mergesOf: mergesOf,
    rowSignature: rowSignature,
    rowStyleSig: rowStyleSig,
    findItemRow: findItemRow,
    detectAlign: detectAlign
  };
  global.FidelityCore = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
