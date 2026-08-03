/* 验证 preview.wbToHtml 的 1:1 还原能力：合并单元格 / 列宽 / 行高 / 底色 / 边框 / 字体
 * 用真实订舱单模板跑，断言 HTML 含 rowspan/colspan/style，且不抛错。
 */
const X = require('exceljs');
const path = require('path');
const preview = require('../js/preview.js');

function load(f) { return new Promise((res, rej) => { const wb = new X.Workbook(); wb.xlsx.readFile(path.resolve(__dirname, '../templates', f)).then(() => res(wb)).catch(rej); }); }

let ok = true;
function assert(c, m) { if (!c) { ok = false; console.log('  ❌ FAIL:', m); } else console.log('  ✅', m); }

(async () => {
  const files = ['tpl_real_aramex_booking.xlsx', 'tpl_real_keas_booking.xlsx', 'tpl_real_booking_form.xlsx'];
  for (const f of files) {
    console.log('=== ' + f + ' ===');
    const wb = await load(f);
    const ws = wb.worksheets[0];
    let html = '';
    let threw = null;
    try { html = preview.wbToHtml(wb); } catch (e) { threw = e; }
    assert(!threw, '渲染不抛错' + (threw ? ' (' + threw.message + ')' : ''));
    assert(typeof html === 'string' && html.indexOf('<table') === 0, '输出为 <table>');

    // 合并单元格：仅当模板本身有合并时才要求还原
    const merges = (ws.model.merges || []).length;
    assert(merges === 0 || (html.indexOf('rowspan') >= 0 && html.indexOf('colspan') >= 0),
      '合并单元格已还原（模板合并数=' + merges + '）');

    // 列宽：仅当模板有列宽时才要求还原
    let anyWidth = false;
    for (let c = 1; c <= (ws.columnCount || 0); c++) { const col = ws.getColumn(c); if (col && col.width) { anyWidth = true; break; } }
    assert(!anyWidth || /<col style="width:\d+px">/.test(html), '列宽已还原（模板有列宽=' + anyWidth + '）');

    // 行高：仅当模板有行高时才要求还原
    let anyH = false;
    ws.eachRow({ includeEmpty: true }, function (row) { if (row.height) anyH = true; });
    assert(!anyH || /<tr style="height:\d+px">/.test(html), '行高已还原（模板有行高=' + anyH + '）');

    // 样式：应含 background / border / 字体 中至少一项
    assert(/background:#|border-(top|left|right|bottom):|font-(weight|size)/.test(html),
      '含单元格样式(底色/边框/字体)');
  }
  console.log('\n' + (ok ? '🎉 preview 1:1 渲染测试 ALL PASS' : '⚠️ preview 测试存在 FAIL'));
  process.exit(ok ? 0 : 1);
})().catch(e => { console.error('ERR', e); process.exit(2); });
