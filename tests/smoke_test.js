/* 浏览器侧冒烟测试：用 jsdom + fake-indexeddb 真实加载 index.html 的 DOM 与全部脚本，
 * 验证 ui.js 的 init()（打开 IndexedDB → 种子数据 → render）不抛异常。
 * 这是 Node 端引擎测试覆盖不到的浏览器专用代码（db.js/ui.js）的运行时校验。 */
'use strict';
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');
const { indexedDB, IDBKeyRange } = require('fake-indexeddb');

const ROOT = path.join(__dirname, '..');
const errors = [];

const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const htmlNoScripts = html.replace(/<script[\s\S]*?<\/script>/g, '');
const dom = new JSDOM(htmlNoScripts, { runScripts: 'outside-only', pretendToBeVisual: true, url: 'http://localhost/' });
const window = dom.window;
window.indexedDB = indexedDB;
window.IDBKeyRange = IDBKeyRange;
const origErr = window.console.error.bind(window.console);
window.console.error = function () { errors.push(Array.from(arguments).map(String).join(' ')); origErr.apply(null, arguments); };
window.addEventListener('error', function (e) { errors.push('window.error: ' + (e.error && e.error.stack || e.message)); });

function ev(rel) { window.eval(fs.readFileSync(path.join(ROOT, rel), 'utf8')); }

try {
  ev('vendor/exceljs.min.js');
  ['js/db.js', 'js/adapters.js', 'js/parser.js', 'js/validator.js', 'js/engine.js', 'js/exporter.js', 'js/declare_data.js', 'js/seed.js', 'js/real_templates.js', 'js/ui.js'].forEach(ev);
} catch (e) {
  console.log('❌ 脚本执行期异常: ' + e.stack);
  process.exit(1);
}

// 等待 init() 异步完成（db.open + seed + render）
setTimeout(function () {
  let ok = true;
  const TD = window.TD;
  if (errors.length) { ok = false; console.log('❌ 捕获到运行时错误:\n  ' + errors.join('\n  ')); }
  if (!TD || !TD.db || !TD.engine || !TD.parser || !TD.validator || !TD.adapters || !TD.seed || !TD.exporter) {
    ok = false; console.log('❌ TD 命名空间未正确挂载: ' + Object.keys(TD || {}));
  }
  const mainHtml = (window.document.getElementById('main') || {}).innerHTML || '';
  if (/初始化失败/.test(mainHtml)) { ok = false; console.log('❌ init 报告初始化失败: ' + mainHtml.slice(0, 200)); }
  // 首屏默认页（订单管理）应已渲染出标题
  if (!/订单管理/.test(mainHtml)) { ok = false; console.log('❌ 首屏未渲染默认订单页'); }
  // 页签导航应存在
  const tabCount = window.document.querySelectorAll('#tabs .tab').length;
  if (tabCount === 0) { ok = false; console.log('❌ 未检测到页签导航'); }

  if (!ok) { process.exit(1); }

  // 进一步验证：种子数据确实写入了 IndexedDB（模板/收发货人/申报要素）
  Promise.all([TD.db.all('templates'), TD.db.all('parties'), TD.db.all('declare_reqs')])
    .then(function (res) {
      const [tpls, parties, declares] = res;
      const checks = [
        ['内置+真实模板入库(>=17)', tpls.length >= 17],
        ['含真实模板(亚丰海运)', tpls.some(function (t) { return t.id === 'tpl_real_yafeng_sea'; })],
        ['含 packing/declare 类型', tpls.some(function (t) { return t.kind === 'packing'; }) && tpls.some(function (t) { return t.kind === 'declare'; })],
        ['示例收发货人', parties.length >= 1],
        ['飞书申报信息表镜像入库(>=10000)', declares.length >= 10000]
      ];
      const failed = checks.filter(function (c) { return !c[1]; });
      if (failed.length) {
        console.log('❌ 种子数据缺失: ' + failed.map(function (c) { return c[0]; }).join(', '));
        process.exit(1);
      }
      console.log('✅ 浏览器侧冒烟测试通过：index.html 加载、IndexedDB 打开、种子写入、首屏 render 均无异常');
      console.log('   TD 子模块: ' + Object.keys(TD).join(', '));
      console.log('   种子数据: 模板 ' + tpls.length + ' / 收发货人 ' + parties.length + ' / 申报要素 ' + declares.length);
      console.log('   页签数: ' + tabCount);
      process.exit(0);
    })
    .catch(function (e) { console.log('❌ 读取种子数据失败: ' + e.stack); process.exit(1); });
}, 5000);
