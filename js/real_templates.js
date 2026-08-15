/* v1.4.59 起清空内嵌模板：单一真源改为 GitHub userdata.json（stores.templates），启动 pullShared 全量同步。
   如需恢复内嵌，重跑 tests/build_real_templates.js 重新生成本文件。 */
window.TD = window.TD || {};
window.TD.realTemplates = [];
