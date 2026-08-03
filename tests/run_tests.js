/* Node 端测试运行器：注入依赖，先跑共享测试模块 tests/shared_tests.js，
 * 再跑真实模板测试 tests/real_templates_test.js（Node 专用，复用 exceljs）。
 * 浏览器端 selftest.html 仅复用 shared_tests.js，保证两套用例不漂移。 */
'use strict';
const path = require('path');
const ExcelJS = require('exceljs');
const parser = require(path.join(__dirname, '..', 'js', 'parser.js'));
const validator = require(path.join(__dirname, '..', 'js', 'validator.js'));
const engine = require(path.join(__dirname, '..', 'js', 'engine.js'));
const adapters = require(path.join(__dirname, '..', 'js', 'adapters.js'));
const shared = require(path.join(__dirname, 'shared_tests.js'));
const realTpl = require(path.join(__dirname, 'real_templates_test.js'));

(async function () {
  const res = await shared.runAll({ ExcelJS, parser, validator, engine, adapters, log: console.log });
  const rt = await realTpl.runAll();
  const xlsConv = require(path.join(__dirname, 'xls_convert_test.js'));
  await xlsConv.main();
  const csvConv = require(path.join(__dirname, 'csv_convert_test.js'));
  await csvConv.main();
  const pkParse = require(path.join(__dirname, 'packing_parse_test.js'));
  await pkParse.main();
  const hdrFill = require(path.join(__dirname, 'template_header_fill_test.js'));
  await hdrFill.main();
  const allTpl = require(path.join(__dirname, 'all_templates_fill_test.js'));
  await allTpl.main();
  const failed = (res.failed || 0) + (rt.failed || 0);
  console.log('\n==== 汇总: 共享用例 ' + res.count + ' (失败 ' + (res.failed || 0) + ') | 真实模板 ' + rt.count + ' (失败 ' + (rt.failed || 0) + ') | .xls/.csv转换 通过 | 装箱清单解析 通过 | 模板表头智能填充 通过 | 全模板填充 通过 ====');
  if (failed) process.exit(1);
})();
