/* L7 导出：workbook → xlsx 下载（浏览器）/ buffer（Node测试） */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else { root.TD = root.TD || {}; root.TD.exporter = factory(); }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function toBuffer(wb) { return wb.xlsx.writeBuffer(); }

  function download(wb, filename) {
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

  return { toBuffer: toBuffer, download: download, safeName: safeName };
});
