/* L7 导出：workbook → xlsx 下载（浏览器）/ buffer（Node测试） */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else { root.TD = root.TD || {}; root.TD.exporter = factory(); }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function toBuffer(wb) { return wb.xlsx.writeBuffer(); }

  /** v1.6.21（仅订舱单）：导出前 LOGO 保真闸门。
   *  背景：订舱单模板多含货代 LOGO（GEODIS 1 张 PNG / KLN 4 张位图）。
   *  ExcelJS 写出只认 r:embed，图片字节虽在，但若引擎/样式处理过程丢了图片关系，
   *  会静默产出一张没有 LOGO 的订舱单 —— 而 LOGO 是承运商单据的必备要素。
   *  这里在【真正写出前】比对图片数，不通过就直接抛错（不产出文件），由调用方提示用户。 */
  function assertLogoPreserved(wb) {
    var before = 0;
    try { before = (wb && wb._mediaCount) || 0; } catch (e) {}
    if (!before) return;                       // 模板本无图片 → 无需校验
    var after = 0;
    try {
      (wb.worksheets || []).forEach(function (w) {
        if (w && w.getImages) after += w.getImages().length;
      });
      if (!after && wb.model && wb.model.media) after = wb.model.media.length;
    } catch (e) {}
    if (after < before) {
      var err = new Error('LOGO 丢失：模板原有 ' + before + ' 张图片，填充后仅剩 ' + after +
        ' 张。已阻止导出，避免生成缺少货代 LOGO 的订舱单。请反馈该模板以修复。');
      err.code = 'LOGO_LOST';
      err.logoBefore = before;
      err.logoAfter = after;
      throw err;
    }
  }

  function download(wb, filename) {
    assertLogoPreserved(wb);   // v1.6.21：LOGO 保真闸门（抛错则阻止导出）
    return toBuffer(wb).then(function (buf) {
      var blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url; a.download = filename;
      document.body.appendChild(a); a.click();
      setTimeout(function () { document.body.removeChild(a); URL.revokeObjectURL(url); }, 500);
      return blob;
    });
  }

  function safeName(s) { return String(s || '').replace(/[\\/:*?"<>|]/g, '_'); }

  return { toBuffer: toBuffer, download: download, safeName: safeName, assertLogoPreserved: assertLogoPreserved };
});
