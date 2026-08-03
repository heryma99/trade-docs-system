/* L? 预览层：workbook首表 → 1:1 HTML 预览
 * 忠实还原模板格式：列宽 / 行高 / 合并单元格(rowspan,colspan) / 底色 / 字体(粗体·字号·颜色) / 对齐 / 边框。
 * 纯函数，无 DOM 依赖，可在浏览器与 Node 共用（便于单测）。UMD。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./parser.js'));
  } else {
    root.TD = root.TD || {};
    root.TD.preview = factory(root.TD.parser);
  }
})(typeof self !== 'undefined' ? self : this, function (parser) {
  'use strict';

  function esc(s) {
    return String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  // 标准 Office 主题色板（theme 索引 → 基准 hex），覆盖绝大多数默认主题模板
  // Office 2016+ 主题语义：0=lt1 窗口背景(白) 1=dk1 窗口文字(黑) 2=lt2 3=dk2 4-8=accent1-5 9=hlink 10=folHlink 11=lt2-variant
  var XLSX_THEME = ['FFFFFF', '000000', '1F4E79', 'E7E6E6', '4472C4', 'ED7D31', 'FFC000', '70AD47', '7030A0', 'C00000', 'BFBFBF', '808080'];

  function _hx(n) { n = Math.max(0, Math.min(255, Math.round(n))); return ('0' + n.toString(16)).slice(-2); }
  function _mix(hex, target, t) {
    var r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
    var tr = (target >> 16) & 255, tg = (target >> 8) & 255, tb = target & 255;
    return _hx(r + (tr - r) * t) + _hx(g + (tg - g) * t) + _hx(b + (tb - b) * t);
  }
  function _lum(hex) {
    var r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16);
    return 0.299 * r + 0.587 * g + 0.114 * b;
  }
  function _resolveColor(c) {
    if (!c) return null;
    if (c.rgb) return ('' + c.rgb).slice(-6);
    if (c.argb) return ('' + c.argb).slice(-6);
    if (c.theme !== undefined && c.theme !== null) {
      var base = XLSX_THEME[c.theme] || '000000';
      if (c.tint) return c.tint < 0 ? _mix(base, 0x000000, -c.tint) : _mix(base, 0xFFFFFF, c.tint);
      return base;
    }
    return null;
  }
  function _borderSide(b) {
    if (!b || !b.style) return null;
    var w = (b.style === 'medium' || b.style === 'thick') ? '1.5px' : '1px';
    var st = (b.style === 'dashed') ? 'dashed' : (b.style === 'dotted' ? 'dotted' : 'solid');
    var col = _resolveColor(b.color) || '999999';
    return w + ' ' + st + ' #' + col;
  }
  function _colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
  function _parseMergeRef(m) {
    if (typeof m === 'string') {
      var mm = m.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      if (!mm) return null;
      return { top: parseInt(mm[2], 10) - 1, left: _colNum(mm[1]) - 1, bottom: parseInt(mm[4], 10) - 1, right: _colNum(mm[3]) - 1 };
    }
    if (m && m.top !== undefined) return { top: m.top, left: m.left, bottom: m.bottom, right: m.right };
    return null;
  }

  /** workbook首表 → 1:1 HTML 预览表（忠实还原模板格式） */
  function wbToHtml(wb) {
    var ws = wb.worksheets[0];
    if (!ws) return '<div class="empty">模板无工作表</div>';
    var cellText = parser && parser.cellText ? parser.cellText : function (v) { return v == null ? '' : String(v); };

    // 1) 合并范围
    var merges = [];
    try { (ws.model.merges || []).forEach(function (m) { var r = _parseMergeRef(m); if (r) merges.push(r); }); } catch (e) {}
    var covered = {}, spanOf = {};
    merges.forEach(function (m) {
      spanOf[(m.top + 1) + ',' + (m.left + 1)] = { rs: m.bottom - m.top + 1, cs: m.right - m.left + 1 };
      for (var r = m.top; r <= m.bottom; r++) for (var c = m.left; c <= m.right; c++) {
        if (r === m.top && c === m.left) continue;
        covered[(r + 1) + ',' + (c + 1)] = true;
      }
    });

    // 2) 最大列 & 列宽
    var maxCol = ws.columnCount || 0;
    merges.forEach(function (m) { if (m.right + 1 > maxCol) maxCol = m.right + 1; });
    ws.eachRow({ includeEmpty: true }, function (row, rn) {
      row.eachCell({ includeEmpty: true }, function (cell, cn) { if (cn > maxCol) maxCol = cn; });
    });
    if (!maxCol) maxCol = 12; else if (maxCol > 40) maxCol = 40;
    var colWs = [];
    for (var c = 1; c <= maxCol; c++) { var col = ws.getColumn(c); colWs[c] = (col && col.width) ? Math.round(col.width * 7) : 0; }

    // 3) 渲染
    var html = '<table class="preview-sheet"><colgroup>';
    for (var c2 = 1; c2 <= maxCol; c2++) html += '<col' + (colWs[c2] ? (' style="width:' + colWs[c2] + 'px"') : '') + '>';
    html += '</colgroup>';
    ws.eachRow({ includeEmpty: true }, function (row, rn) {
      if (rn > 200) return;
      var trStyle = row.height ? ' style="height:' + Math.round(row.height * 1.333) + 'px"' : '';
      html += '<tr' + trStyle + '>';
      for (var c = 1; c <= maxCol; c++) {
        if (covered[rn + ',' + c]) continue;
        var cell = row.getCell(c);
        var sp = spanOf[rn + ',' + c];
        var s = cellText(cell.value);
        var style = '', st = cell.style || {}, hasFontColor = false;
        if (st.fill && st.fill.fgColor) { var bg = _resolveColor(st.fill.fgColor); if (bg) style += 'background:#' + bg + ';'; }
        if (st.font) { var f = st.font;
          if (f.bold) style += 'font-weight:700;';
          if (f.italic) style += 'font-style:italic;';
          if (f.size) style += 'font-size:' + f.size + 'px;';
          if (f.name) style += 'font-family:"' + f.name + '";';
          if (f.color) { var fc = _resolveColor(f.color); if (fc) { style += 'color:#' + fc + ';'; hasFontColor = true; } }
          if (f.underline) style += 'text-decoration:underline;';
        }
        // 深底色且无显式字体色 → 自动转白字，避免黑底黑字不可读
        if (!hasFontColor && bg && _lum(bg) < 140) style += 'color:#fff;';
        if (st.alignment) { var a = st.alignment;
          if (a.horizontal === 'center') style += 'text-align:center;';
          else if (a.horizontal === 'right') style += 'text-align:right;';
          else if (a.horizontal === 'left') style += 'text-align:left;';
          if (a.vertical === 'middle') style += 'vertical-align:middle;';
          else if (a.vertical === 'top') style += 'vertical-align:top;';
          else if (a.vertical === 'bottom') style += 'vertical-align:bottom;';
          if (a.wrapText) style += 'white-space:normal;';
        }
        if (st.border) { var b = st.border, parts = [], bt = _borderSide(b.top), bl = _borderSide(b.left), br = _borderSide(b.right), bb = _borderSide(b.bottom);
          if (bt) parts.push('border-top:' + bt); if (bl) parts.push('border-left:' + bl); if (br) parts.push('border-right:' + br); if (bb) parts.push('border-bottom:' + bb);
          if (parts.length) style += parts.join(';') + ';';
        }
        var attr = '';
        if (sp) attr += ' rowspan="' + sp.rs + '" colspan="' + sp.cs + '"';
        if (style) attr += ' style="' + style + '"';
        html += '<td' + attr + '>' + esc(s) + '</td>';
      }
      html += '</tr>';
    });
    // 4) 内嵌图片（如 logo）：用绝对定位覆盖在表格对应单元格之上
    try {
      var imgs = ws.getImages ? ws.getImages() : [];
      for (var ig = 0; ig < imgs.length; ig++) {
        var im = imgs[ig];
        var media = wb.media && wb.media[im.imageId];
        if (!media || !media.buffer) continue;
        var b64 = '';
        if (typeof Buffer !== 'undefined') {
          b64 = media.buffer.toString('base64');
        } else if (typeof btoa !== 'undefined') {
          // 浏览器：ArrayBuffer/Uint8Array -> base64
          var u8 = media.buffer;
          var bin = '';
          var len = u8.length;
          var chunk = 0x8000;
          for (var bi = 0; bi < len; bi += chunk) {
            bin += String.fromCharCode.apply(null, u8.subarray(bi, Math.min(len, bi + chunk)));
          }
          b64 = btoa(bin);
        }
        var ext = (media.extension || 'png').toLowerCase();
        var mime = ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : 'image/' + ext;
        // 计算像素位置：累加 col/row 之前的列宽/行高
        // 自动行高/列宽(ExcelJS 返回 undefined)需回退默认值，否则图片 top/height 累加为 0 → 不可见。
        //   默认行高 ≈ 15pt × 1.333 ≈ 20px；默认列宽 ≈ 8.43 × 7 ≈ 59px
        function colX(idx) { var x = 0; for (var i = 1; i <= idx; i++) x += colWs[i] || 59; return x; }
        function rowY(idx) { var y = 0; for (var i = 1; i <= idx; i++) { var rh = ws.getRow(i).height; y += rh ? Math.round(rh * 1.333) : 20; } return y; }
        var r0 = im.range.tl; var r1 = im.range.br;
        var x0 = colX(r0.nativeCol); var y0 = rowY(r0.nativeRow);
        var x1 = colX(r1.nativeCol + 1); var y1 = rowY(r1.nativeRow + 1);
        var w = x1 - x0; var h = y1 - y0;
        var extra = '';
        if (r0.nativeColOff) extra += 'left:' + (x0 + r0.nativeColOff / 9525) + 'px;';
        else extra += 'left:' + x0 + 'px;';
        if (r0.nativeRowOff) extra += 'top:' + (y0 + r0.nativeRowOff / 9525) + 'px;';
        else extra += 'top:' + y0 + 'px;';
        html += '<img src="data:' + mime + ';base64,' + b64 + '" style="position:absolute;' + extra + 'width:' + w + 'px;height:' + h + 'px;pointer-events:none;z-index:5;"/>';
      }
    } catch (e) {}
    return html + '</table>';
  }

  return { wbToHtml: wbToHtml };
});
