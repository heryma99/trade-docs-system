/* L1 UI层：页签 + CRUD + 发票/订舱单双向导 */
(function () {
  'use strict';
  var db = TD.db, parser = TD.parser, validator = TD.validator, engine = TD.engine,
      exporter = TD.exporter, adapters = TD.adapters, seed = TD.seed;

  var $main = document.getElementById('main');
  var _initSkipPull = false; // v1.4.49：watchdog 卡死时用户可手动置 true 跳过拉取
  var CARRIERS = ['通用', '亚丰', '安速', '亦邦', '合联', '艾杜克'];
  var INCOTERMS = ['FOB', 'CIF', 'CFR', 'EXW', 'DDP', 'DDU', 'DAP', 'FCA'];
  var state = { tab: 'orders', wiz: null, bwiz: null };
  var FX_RATE = 7.2; // CNY per USD：备用文档申报金额(假设元)折算USD汇率，需与 tests/build_declare_data.js 的 RATE 保持一致
  function fxToUsd(v, cur) { if (cur === 'USD' || v === undefined || v === null) return v; return Math.round((v / FX_RATE) * 1000) / 1000; }

  // ---------- 基础helper ----------
  function esc(s) { return String(s === undefined || s === null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function kindBadge(k) {
    if (k === 'invoice') return '<span class="badge blue">发票</span>';
    if (k === 'booking') return '<span class="badge purple">订舱单</span>';
    if (k === 'packing') return '<span class="badge green">装箱单</span>';
    if (k === 'declare') return '<span class="badge yellow">申报</span>';
    return '<span class="badge gray">' + esc(k) + '</span>';
  }
  function toast(msg, type) {
    var box = document.getElementById('toast');
    var d = document.createElement('div');
    d.className = 'toast-item ' + (type || '');
    d.textContent = msg;
    box.appendChild(d);
    setTimeout(function () { d.remove(); }, 3500);
  }
  function showModal(html) {
    document.getElementById('modal').innerHTML = html;
    document.getElementById('modal-mask').classList.remove('hidden');
  }
  function closeModal() { document.getElementById('modal-mask').classList.add('hidden'); }
  document.getElementById('modal-mask').addEventListener('click', function (e) { if (e.target.id === 'modal-mask') closeModal(); });
  window.TDUI = { closeModal: closeModal }; // 供内联onclick

  // v1.5.0 e2e 测试 seam（纯增量，不改变任何业务行为；仅暴露向导内部状态供自动化测试点击真实门禁）
  window.__TD_E2E__ = {
    wiz: function () { return state.wiz; },
    setWiz: function (o) { if (!state.wiz) state.wiz = { step: 4 }; Object.assign(state.wiz, o); },
    render: function () { renderWizStep(); },
    // 直接进入发票向导 step4 并渲染（注入态跳过真实填充），绕开逐步导航的异步竞争，供 e2e 确定性点击门禁
    mountInvoiceStep4: async function (inject) {
      state.tab = 'invoice';
      var main = document.getElementById('main') || document.body;
      main.innerHTML = '<h2>生成发票</h2><div id="wiz-body"></div>';
      if (!state.wiz) state.wiz = { step: 4 };
      Object.assign(state.wiz, inject);
      await renderWizStep();
    }
  };

  function confirmBox(msg, danger) {
    return new Promise(function (resolve) {
      showModal('<h3>' + (danger ? '⚠️ ' : '') + '确认操作</h3><p style="margin:14px 0">' + msg + '</p>' +
        '<div style="text-align:right;display:flex;gap:8px;justify-content:flex-end">' +
        '<button class="btn ghost" id="cf-no">取消</button><button class="btn ' + (danger ? 'danger' : '') + '" id="cf-yes">确认</button></div>');
      document.getElementById('cf-no').onclick = function () { closeModal(); resolve(false); };
      document.getElementById('cf-yes').onclick = function () { closeModal(); resolve(true); };
    });
  }
  function fmtTime(t) { if (!t) return ''; var d = new Date(t); return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0') + ' ' + String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0'); }
  function today() { return new Date().toISOString().slice(0, 10); }
  /** 用 SheetJS 把旧版 .xls / .csv 转成标准 .xlsx ArrayBuffer，再交给 ExcelJS。
   *  CSV 先用 TextDecoder 探测编码（非法 UTF-8 则按中文 GBK 解码），再交给 SheetJS 解析 */
  async function anyToXlsx(arrayBuf, fileName) {
    if (typeof XLSX === 'undefined') {
      throw new Error('缺少文件转换库，请刷新页面后重试');
    }
    var fn = fileName || '';
    var isCsv = /\.csv$/i.test(fn);
    var bytes = new Uint8Array(arrayBuf);
    var wb;
    if (isCsv) {
      var text;
      try { text = new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
      catch (e) { text = new TextDecoder('gbk').decode(bytes); }
      wb = XLSX.read(text, { type: 'string', raw: true });
    } else {
      wb = XLSX.read(bytes, { type: 'array', cellFormula: false, cellNF: true });
    }
    var out = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    if (out instanceof ArrayBuffer) return out;
    return out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength);
  }
  /**
   * 加载工作簿（自动探测格式，兼容 .xlsx / .xls 老格式 / .csv，不依赖 fileName）
   * magic bytes: .xlsx=PK\x03\x04 / .xls=D0\xCF\x11\xE0 (CFB) / 其它按 .csv 试
   */
  async function loadWb(buf, fileName) {
    var fn = fileName || '';
    // 先把任意形态的 fileBuf 规范化成 ArrayBuffer（base64-string / {type:Buffer,data:[]} / TypedArray 都接），避免老 sync 残留的 base64-string 让 loadWb 误判
    var ab = _asArrayBuffer(buf);
    var u8 = new Uint8Array(ab);
    var isZip = u8.length >= 4 && u8[0] === 0x50 && u8[1] === 0x4B && (u8[2] === 0x03 || u8[2] === 0x05);
    var isCfb = u8.length >= 8 && u8[0] === 0xD0 && u8[1] === 0xCF && u8[2] === 0x11 && u8[3] === 0xE0;
    var isCsvHint = /\.csv$/i.test(fn);
    var isOldXlsHint = /\.xls$/i.test(fn) && !/\.xlsx$/i.test(fn);
    // 自动探测: 老 .xls 格式（CFB 头且不是 ZIP） 或 显式 .csv 提示 → 必转 .xlsx
    var needConvert = isCsvHint || (!isZip && isCfb) || (!isZip && !isCfb && isOldXlsHint);
    var formatGuess = isZip ? 'xlsx' : (isCfb ? 'xls-old' : (isCsvHint ? 'csv' : (isOldXlsHint ? 'xls-hint' : 'unknown')));
    if (needConvert) {
      try {
        var fakeFn = fn || (isCfb ? 'upload.xls' : 'upload.csv');
        ab = await anyToXlsx(ab, fakeFn);
      } catch (cv) {
        throw new Error('文件无法解析（探测为 ' + formatGuess + '，转换失败: ' + (cv.message || cv) + '）。请用 Excel/WPS 另存为 .xlsx 后再试');
      }
    }
    try {
      var wb = new ExcelJS.Workbook();
      await wb.xlsx.load(ab);
      wb.__formatGuess = formatGuess; // 业务侧可选读
      return wb;
    } catch (err) {
      var msg = err.message || '';
      if (/zip|central directory|end of central/i.test(msg)) {
        throw new Error('文件不是有效的 .xlsx（探测格式: ' + formatGuess + '，可能是 .xls 旧格式、.csv 或文件已损坏），请用 Excel/WPS 另存为 .xlsx 后再试');
      }
      throw err;
    }
  }
  function val(id) { var e = document.getElementById(id); return e ? e.value.trim() : ''; }

  /** workbook首表 → 1:1 HTML预览表（还原列宽/行高/合并/字体/底色/边框），实现见 js/preview.js */
  function wbToHtml(wb) { return TD.preview.wbToHtml(wb); }

  // ---------- 页签路由 ----------
  document.getElementById('tabs').addEventListener('click', function (e) {
    var btn = e.target.closest('.tab');
    if (!btn) return;
    document.querySelectorAll('.tab').forEach(function (t) { t.classList.remove('active'); });
    btn.classList.add('active');
    state.tab = btn.dataset.tab;
    render();
  });

  function render() {
    var fn = PAGES[state.tab];
    $main.innerHTML = '<div class="loading">加载中…</div>';
    Promise.resolve(fn()).then(function (html) {
      if (typeof html === 'string') $main.innerHTML = html;
      var binder = BINDERS[state.tab];
      if (binder) binder();
    }).catch(function (e) {
      $main.innerHTML = '<div class="card vres block">页面渲染失败: ' + esc(e.message) + '</div>';
      console.error(e);
    });
  }

  var PAGES = {}, BINDERS = {};

  // ================= 订单页 =================
  PAGES.orders = async function () {
    var orders = await db.all('orders');
    orders.sort(function (a, b) { return b.createdAt - a.createdAt; });
    var rows = orders.map(function (o) {
      var qty = (o.items || []).reduce(function (s, i) { return s + i.qty; }, 0);
      return '<tr><td><input type="checkbox" class="ord-ck" value="' + o.id + '"></td>' +
        '<td class="mono">' + esc(o.orderNo) + '</td>' +
        '<td>' + (o.source === 'from_packing' ? '<span class="badge purple">装箱清单直生</span>' : '<span class="badge blue">聚水潭导入</span>') + '</td>' +
        '<td>' + esc(o.buyer || '') + '</td><td>' + esc(o.country || '') + '</td>' +
        '<td class="num">' + (o.items || []).length + '</td><td class="num">' + qty + '</td>' +
        '<td>' + fmtTime(o.createdAt) + '</td>' +
        '<td><button class="btn sm ghost ord-view" data-id="' + o.id + '">详情</button> ' +
        '<button class="btn sm danger ord-del" data-id="' + o.id + '">删除</button></td></tr>';
    }).join('');
    return '<h2>订单管理</h2><div class="card"><div class="toolbar">' +
      '<label class="btn" style="display:inline-block">📥 导入聚水潭订单(.xlsx / .xls / .csv)<input type="file" id="jst-file" accept=".xlsx,.xls,.csv" style="display:none"></label>' +
      '<button class="btn danger" id="ord-del-sel">删除所选</button>' +
      '<span class="hint">聚水潭API自动拉取可在「设置」中配置凭证后开通；也可在「装箱清单」页直接由箱单生成订单</span></div>' +
      '<table class="grid"><tr><th><input type="checkbox" id="ord-ck-all"></th><th>订单号</th><th>来源</th><th>买家/客户</th><th>国家</th><th>SKU种数</th><th>总数量</th><th>创建时间</th><th>操作</th></tr>' +
      (rows || '<tr><td colspan="9" class="empty">暂无订单，请导入聚水潭订单或从装箱清单生成</td></tr>') + '</table></div>';
  };
  BINDERS.orders = function () {
    document.getElementById('jst-file').onchange = async function () {
      var f = this.files[0]; if (!f) return;
      try {
        var wb = await loadWb(await f.arrayBuffer(), f.name);
        var rows = parser.sheetToRows(wb.worksheets[0]);
        var list = parser.parseJstOrders(rows);
        var exist = await db.all('orders');
        var existNos = {}; exist.forEach(function (o) { existNos[o.orderNo] = 1; });
        var dup = list.filter(function (o) { return existNos[o.orderNo]; });
        var fresh = list.filter(function (o) { return !existNos[o.orderNo]; });
        if (dup.length && !(await confirmBox('有 ' + dup.length + ' 个订单号已存在（将跳过）：' + esc(dup.map(function (o) { return o.orderNo; }).slice(0, 5).join(', ')) + (dup.length > 5 ? '…' : '') + '。继续导入其余 ' + fresh.length + ' 单？'))) return;
        if (!fresh.length) { toast('没有新订单可导入', 'err'); return; }
        await db.bulkPut('orders', fresh);
        toast('成功导入 ' + fresh.length + ' 个订单', 'ok'); render();
      } catch (e) { toast('导入失败: ' + e.message, 'err'); }
    };
    document.getElementById('ord-ck-all').onchange = function () {
      var on = this.checked;
      document.querySelectorAll('.ord-ck').forEach(function (c) { c.checked = on; });
    };
    document.getElementById('ord-del-sel').onclick = async function () {
      var ids = Array.from(document.querySelectorAll('.ord-ck:checked')).map(function (c) { return c.value; });
      if (!ids.length) { toast('请先勾选订单', 'err'); return; }
      if (!(await confirmBox('确认删除所选 ' + ids.length + ' 个订单？', true))) return;
      for (var i = 0; i < ids.length; i++) await db.del('orders', ids[i]);
      toast('已删除', 'ok'); render();
    };
    document.querySelectorAll('.ord-del').forEach(function (b) {
      b.onclick = async function () {
        if (!(await confirmBox('确认删除该订单？', true))) return;
        await db.del('orders', b.dataset.id); toast('已删除', 'ok'); render();
      };
    });
    document.querySelectorAll('.ord-view').forEach(function (b) {
      b.onclick = async function () {
        var o = await db.get('orders', b.dataset.id);
        var items = (o.items || []).map(function (i, n) {
          return '<tr><td>' + (n + 1) + '</td><td class="mono">' + esc(i.sku) + '</td><td>' + esc(i.name || '') + '</td><td class="num">' + i.qty + '</td><td class="num">' + (i.price || '') + '</td></tr>';
        }).join('');
        showModal('<h3>订单详情 · ' + esc(o.orderNo) + '</h3>' +
          '<p class="hint">来源: ' + (o.source === 'from_packing' ? '装箱清单直生' : '聚水潭导入') + ' | 买家: ' + esc(o.buyer || '-') + ' | 收货人: ' + esc(o.receiver || '-') + ' | 国家: ' + esc(o.country || '-') + '</p>' +
          (o.address ? '<p class="hint">地址: ' + esc(o.address) + '</p>' : '') +
          '<table class="grid"><tr><th>#</th><th>SKU</th><th>品名</th><th>数量</th><th>单价</th></tr>' + items + '</table>' +
          '<div style="text-align:right;margin-top:12px"><button class="btn" onclick="TDUI.closeModal()">关闭</button></div>');
      };
    });
  };

  // ================= 装箱清单页 =================
  PAGES.packings = async function () {
    var pks = await db.all('packings');
    pks.sort(function (a, b) { return b.createdAt - a.createdAt; });
    var rows = pks.map(function (p) {
      // 汇总 SKU：按出现顺序去重，统计种类数与总件数
      var skuMap = {};
      (p.boxes || []).forEach(function (b) { skuMap[b.sku] = (skuMap[b.sku] || 0) + (b.qty || 0); });
      var skuList = Object.keys(skuMap);
      var skuTxt = skuList.length ? ('共' + skuList.length + '种：' + skuList.slice(0, 3).map(function (s) { return s + '×' + skuMap[s]; }).join(', ') + (skuList.length > 3 ? ' +' + (skuList.length - 3) : '')) : '-';
      return '<tr><td>' + esc(p.fileName) + '</td><td class="mono">' + esc((p.orderNos || []).join(', ') || '-') + '</td>' +
        '<td class="num">' + p.totals.boxCount + '</td><td class="num">' + p.totals.qty + '</td>' +
        '<td class="mono" title="' + esc(skuTxt) + '">' + esc(skuTxt.length > 35 ? skuTxt.slice(0, 35) + '…' : skuTxt) + '</td>' +
        '<td class="num">' + p.totals.gw + '</td><td class="num">' + p.totals.volume + '</td>' +
        '<td>' + fmtTime(p.createdAt) + '</td>' +
        '<td><button class="btn sm ghost pk-view" data-id="' + p.id + '">详情</button> ' +
        '<button class="btn sm ok pk-gen" data-id="' + p.id + '">生成订单</button> ' +
        '<button class="btn sm danger pk-del" data-id="' + p.id + '">删除</button></td></tr>';
    }).join('');
    return '<h2>装箱清单</h2><div class="card"><div class="toolbar">' +
      '<label class="btn" style="display:inline-block">📥 上传装箱清单(.xlsx / .xls / .csv)<input type="file" id="pk-file" accept=".xlsx,.xls,.csv" style="display:none"></label>' +
      '<span class="hint" style="margin-left:12px">箱重固定值(kg)：<input id="pk-fixed-bw" type="number" step="0.01" min="0" style="width:80px" placeholder="选填"></span>' +
      '<span class="hint">解析后可直接「生成订单」——针对已在其他系统核对过的装箱清单，无需再导聚水潭订单</span></div>' +
      '<table class="grid"><tr><th>文件名</th><th>关联单号</th><th>箱数</th><th>总数量</th><th>SKU 汇总</th><th>总毛重(KG)</th><th>体积(CBM)</th><th>导入时间</th><th>操作</th></tr>' +
      (rows || '<tr><td colspan="9" class="empty">暂无装箱清单</td></tr>') + '</table></div>';
  };
  BINDERS.packings = function () {
    // v2026-08-09：箱重固定值兜底（持久化到 localStorage；解析时传给 parsePacking.fixedBoxWeight）
    var fixedBw = +(localStorage.getItem('td.fixedBoxWeight') || 0);
    document.getElementById('pk-fixed-bw').value = fixedBw || '';
    document.getElementById('pk-fixed-bw').onchange = function () {
      var v = +this.value || 0; localStorage.setItem('td.fixedBoxWeight', v);
      toast(v ? '已设置箱重固定值 ' + v + ' kg（源文件无「包装箱重量」且 box_specs 查不到时启用）' : '已清空箱重固定值', 'ok');
    };
    document.getElementById('pk-file').onchange = async function () {
      var f = this.files[0]; if (!f) return;
      try {
        var wb = await loadWb(await f.arrayBuffer(), f.name);
        var rows = parser.sheetToRows(wb.worksheets[0]);
        var hash = parser.hashRows(rows);
        var exist = await db.all('packings');
        if (exist.some(function (p) { return p.hash === hash; })) { toast('该装箱清单已导入过（内容完全相同），已拦截重复导入', 'err'); return; }
        var fixed = +(document.getElementById('pk-fixed-bw').value || localStorage.getItem('td.fixedBoxWeight') || 0);
        var pk = parser.parsePacking(rows, { fixedBoxWeight: fixed });
        pk.fileName = f.name; pk.hash = hash;
        await db.put('packings', pk);
        toast('装箱清单解析成功：' + pk.totals.boxCount + '箱 / ' + pk.totals.qty + '件', 'ok'); render();
      } catch (e) { toast('解析失败: ' + e.message, 'err'); }
    };
    document.querySelectorAll('.pk-del').forEach(function (b) {
      b.onclick = async function () {
        if (!(await confirmBox('确认删除该装箱清单？', true))) return;
        await db.del('packings', b.dataset.id); toast('已删除', 'ok'); render();
      };
    });
    document.querySelectorAll('.pk-view').forEach(function (b) {
      b.onclick = async function () {
        var p = await db.get('packings', b.dataset.id);
        var rows = (p.boxes || []).slice(0, 200).map(function (x, n) {
          return '<tr><td>' + (n + 1) + '</td><td class="mono">' + esc(x.boxNo) + '</td><td class="mono">' + esc(x.orderNo || '') + '</td><td class="mono">' + esc(x.sku) + '</td><td>' + esc(x.name || '') + '</td><td class="num">' + x.qty + '</td><td class="num">' + (x.nw || '') + '</td><td class="num">' + (x.gw || '') + '</td></tr>';
        }).join('');
        showModal('<h3>装箱明细 · ' + esc(p.fileName) + '</h3>' +
          '<p class="hint">共 ' + p.boxes.length + ' 行' + (p.boxes.length > 200 ? '（仅显示前200行）' : '') + ' | 箱数 ' + p.totals.boxCount + ' | 数量 ' + p.totals.qty + ' | 净重 ' + p.totals.nw + ' | 毛重 ' + p.totals.gw + ' | 体积 ' + p.totals.volume + '</p>' +
          '<table class="grid"><tr><th>#</th><th>箱号</th><th>单号</th><th>SKU</th><th>品名</th><th>数量</th><th>净重</th><th>毛重</th></tr>' + rows + '</table>' +
          '<div style="text-align:right;margin-top:12px"><button class="btn" onclick="TDUI.closeModal()">关闭</button></div>');
      };
    });
    document.querySelectorAll('.pk-gen').forEach(function (b) {
      b.onclick = async function () {
        var p = await db.get('packings', b.dataset.id);
        var groups = {};
        p.boxes.forEach(function (x) {
          var no = x.orderNo || ('PK' + today().replace(/-/g, '') + '-' + p.id.slice(-4));
          if (!groups[no]) groups[no] = {};
          var g = groups[no];
          if (!g[x.sku]) g[x.sku] = { sku: x.sku, name: x.name || '', qty: 0 };
          g[x.sku].qty += x.qty;
        });
        var exist = await db.all('orders');
        var existNos = {}; exist.forEach(function (o) { existNos[o.orderNo] = 1; });
        var nos = Object.keys(groups), created = 0, skipped = [];
        for (var i = 0; i < nos.length; i++) {
          if (existNos[nos[i]]) { skipped.push(nos[i]); continue; }
          await db.put('orders', {
            orderNo: nos[i], source: 'from_packing', packingId: p.id,
            items: Object.values(groups[nos[i]]),
            buyer: '', receiver: '', address: '', country: ''
          });
          created++;
        }
        toast('已生成 ' + created + ' 个订单' + (skipped.length ? '，跳过已存在: ' + skipped.join(', ') : ''), created ? 'ok' : 'err');
      };
    });
  };

  // ================= 收发货人页 =================
  PAGES.parties = async function () {
    var ps = await db.all('parties');
    ps.sort(function (a, b) { return a.type < b.type ? -1 : 1; });
    var typeName = { shipper: '发货人 SHIPPER', consignee: '收货人 CONSIGNEE', notify: '通知人 NOTIFY' };
    var rows = ps.map(function (p) {
      return '<tr><td><span class="badge ' + (p.type === 'shipper' ? 'blue' : p.type === 'consignee' ? 'green' : 'gray') + '">' + typeName[p.type] + '</span></td>' +
        '<td>' + esc(p.name) + '</td><td>' + esc(p.address || '') + '</td><td>' + esc(p.tel || '') + '</td><td>' + esc(p.country || '') + '</td>' +
        '<td><button class="btn sm ghost pt-edit" data-id="' + p.id + '">编辑</button> <button class="btn sm danger pt-del" data-id="' + p.id + '">删除</button></td></tr>';
    }).join('');
    return '<h2>收发货人主数据</h2><div class="flexrow"><div class="card" style="flex:2">' +
      '<table class="grid"><tr><th>类型</th><th>名称</th><th>地址</th><th>电话</th><th>国家</th><th>操作</th></tr>' +
      (rows || '<tr><td colspan="6" class="empty">暂无数据</td></tr>') + '</table></div>' +
      '<div class="card form-col" style="flex:1;max-width:380px"><h3 id="pt-form-title">新增</h3>' +
      '<input type="hidden" id="pt-id">' +
      '<label class="req">类型</label><select id="pt-type"><option value="shipper">发货人 SHIPPER</option><option value="consignee">收货人 CONSIGNEE</option><option value="notify">通知人 NOTIFY</option></select>' +
      '<label class="req" style="margin-top:8px">名称（英文）</label><input id="pt-name">' +
      '<label style="margin-top:8px">仓库代码（如 LAX9 / FTW1 / GYR3，私人地址留空）</label><input id="pt-whcode" placeholder="海外仓 FBA 代码">' +
      '<label style="margin-top:8px">公司名（英文）</label><input id="pt-company" placeholder="如 JW PEI AP LIMITED">' +
      '<label style="margin-top:8px">地址（英文，可多行）</label><textarea id="pt-address" rows="2" onblur="window.__tdAddrBlur&&window.__tdAddrBlur(this)"></textarea>' +
      '<label style="margin-top:8px">城市</label><input id="pt-city">' +
      '<label style="margin-top:8px">省份/州</label><input id="pt-state">' +
      '<label style="margin-top:8px">邮编</label><input id="pt-zip">' +
      '<label style="margin-top:8px">电话</label><input id="pt-tel">' +
      '<label style="margin-top:8px">联系人</label><input id="pt-contact">' +
      '<label style="margin-top:8px">邮箱</label><input id="pt-email">' +
      '<label style="margin-top:8px">国家代码</label><input id="pt-country" placeholder="如 US / CN">' +
      '<label style="margin-top:8px">税号 TAX ID</label><input id="pt-taxno">' +
      '<label style="margin-top:8px">VAT 号</label><input id="pt-vatno" placeholder="欧盟 VAT，留空则用税号">' +
      '<label style="margin-top:8px">EORI</label><input id="pt-eori" placeholder="留空则用税号">' +
      '<div style="margin-top:12px;display:flex;gap:8px"><button class="btn" id="pt-save">保存</button><button class="btn ghost" id="pt-reset">清空</button></div></div></div>';
  };
  // v1.4.62：地址框失焦时保守解析国家/邮编/城市并回填到对应输入框（仅填空字段，可见可改，不覆盖已填值）
  // v1.4.64：国家匹配修正——长名命中即停 + 短码 \b 边界（修 "CAUSEWAY"→US 误判）；表单层仍仅填空（不惊扰用户已填）
  window.__tdAddrBlur = function (ta) {
    var addr = (ta && ta.value) || '';
    if (!addr.trim()) return;
    function g(id) { return document.getElementById(id); }
    var s = addr.toUpperCase();
    var longMap = [
      ['HONG KONG', 'HK'], ['HONGKONG', 'HK'], ['MACAU', 'MO'], ['MACAO', 'MO'], ['CHINA', 'CN'],
      ['UNITED STATES OF AMERICA', 'US'], ['UNITED STATES', 'US'], ['U.S.A.', 'US'], ['U.S.A', 'US'], ['USA', 'US'],
      ['UNITED KINGDOM', 'GB'], ['U.K.', 'GB'], ['U.K', 'GB'], ['ENGLAND', 'GB'], ['SCOTLAND', 'GB'], ['GREAT BRITAIN', 'GB'],
      ['JAPAN', 'JP'], ['GERMANY', 'DE'], ['DEUTSCHLAND', 'DE'], ['FRANCE', 'FR'], ['AUSTRALIA', 'AU'],
      ['CANADA', 'CA'], ['SINGAPORE', 'SG'], ['MALAYSIA', 'MY'], ['REPUBLIC OF KOREA', 'KR'], ['KOREA', 'KR'],
      ['TAIWAN', 'TW'], ['THAILAND', 'TH'], ['VIETNAM', 'VN'], ['INDIA', 'IN'], ['MEXICO', 'MX'], ['BRAZIL', 'BR']
    ];
    var country = '';
    for (var i = 0; i < longMap.length; i++) {
      if (s.indexOf(longMap[i][0]) >= 0) { country = longMap[i][1]; break; }
    }
    if (!country) {
      var shortMap = [['HK', 'HK'], ['CN', 'CN'], ['US', 'US'], ['UK', 'GB'], ['GB', 'GB'], ['DE', 'DE'], ['FR', 'FR'], ['JP', 'JP'], ['AU', 'AU'], ['CA', 'CA'], ['SG', 'SG'], ['MY', 'MY'], ['KR', 'KR'], ['TW', 'TW'], ['TH', 'TH'], ['VN', 'VN'], ['IN', 'IN'], ['MX', 'MX'], ['BR', 'BR']];
      for (var j = 0; j < shortMap.length; j++) {
        if (new RegExp('\\b' + shortMap[j][0] + '\\b').test(s)) { country = shortMap[j][1]; break; }
      }
    }
    var zm = addr.match(/\b(\d{5,6}|\d{5}-\d{4})\b/); var zip = zm ? zm[1] : '';
    var cityMap = ['HONG KONG', 'HONGKONG', 'SHENZHEN', 'GUANGZHOU', 'SHANGHAI', 'YIWU', 'NINGBO', 'TOKYO', 'OSAKA',
      'LOS ANGELES', 'NEW YORK', 'NEW JERSEY', 'NEWARK', 'CHICAGO', 'DALLAS', 'HOUSTON', 'MIAMI', 'SEATTLE', 'ATLANTA',
      'LONDON', 'MANCHESTER', 'SYDNEY', 'MELBOURNE', 'SINGAPORE', 'KUALA LUMPUR', 'BUSAN', 'SEOUL', 'TAIPEI', 'BANGKOK', 'HANOI'];
    var city = ''; cityMap.forEach(function (c) { if (s.indexOf(c) >= 0) city = c.replace('HONGKONG', 'HONG KONG'); });
    if (g('pt-country') && !g('pt-country').value && country) g('pt-country').value = country;
    if (g('pt-zip') && !g('pt-zip').value && zip) g('pt-zip').value = zip;
    if (g('pt-city') && !g('pt-city').value && city) g('pt-city').value = city;
  };
  BINDERS.parties = function () {
    function resetForm() { ['pt-id', 'pt-name', 'pt-whcode', 'pt-company', 'pt-address', 'pt-city', 'pt-state', 'pt-zip', 'pt-tel', 'pt-contact', 'pt-email', 'pt-country', 'pt-taxno', 'pt-vatno', 'pt-eori'].forEach(function (i) { var el = document.getElementById(i); if (el) el.value = ''; }); document.getElementById('pt-form-title').textContent = '新增'; }
    document.getElementById('pt-reset').onclick = resetForm;
    document.getElementById('pt-save').onclick = async function () {
      var name = val('pt-name');
      if (!name) { toast('名称必填', 'err'); return; }
      var obj = { type: val('pt-type'), name: name, company: val('pt-company'), warehouseCode: val('pt-whcode'), address: val('pt-address'), city: val('pt-city'), state: val('pt-state'), zip: val('pt-zip'), tel: val('pt-tel'), contact: val('pt-contact'), email: val('pt-email'), country: val('pt-country'), taxNo: val('pt-taxno'), vatNo: val('pt-vatno'), eori: val('pt-eori') };
      var id = val('pt-id');
      if (id) { var old = await db.get('parties', id); obj = Object.assign(old, obj); }
      // 用户接管 seed 占位数据：保存时自动剥离 isSeed 标记，避免下次 push 把 DEMO 推上团队库
      if (obj && obj.isSeed) delete obj.isSeed;
      await db.put('parties', obj);
      toast('已保存', 'ok'); render();
      autoSyncToTeam(); // v1.5.22 主数据变更自动同步
    };
    document.querySelectorAll('.pt-edit').forEach(function (b) {
      b.onclick = async function () {
        var p = await db.get('parties', b.dataset.id);
        document.getElementById('pt-id').value = p.id;
        document.getElementById('pt-type').value = p.type;
        document.getElementById('pt-name').value = p.name || '';
        document.getElementById('pt-whcode').value = p.warehouseCode || '';
        document.getElementById('pt-company').value = p.company || '';
        document.getElementById('pt-address').value = p.address || '';
        document.getElementById('pt-city').value = p.city || '';
        document.getElementById('pt-state').value = p.state || '';
        document.getElementById('pt-zip').value = p.zip || '';
        document.getElementById('pt-tel').value = p.tel || '';
        document.getElementById('pt-contact').value = p.contact || '';
        document.getElementById('pt-email').value = p.email || '';
        document.getElementById('pt-country').value = p.country || '';
        document.getElementById('pt-taxno').value = p.taxNo || '';
        var _vat = document.getElementById('pt-vatno'); if (_vat) _vat.value = p.vatNo || '';
        var _eori = document.getElementById('pt-eori'); if (_eori) _eori.value = p.eori || '';
        document.getElementById('pt-form-title').textContent = '编辑: ' + p.name;
      };
    });
    document.querySelectorAll('.pt-del').forEach(function (b) {
      b.onclick = async function () {
        if (!(await confirmBox('确认删除该收发货人？', true))) return;
        await db.del('parties', b.dataset.id); toast('已删除', 'ok'); render();
        autoSyncToTeam(); // v1.5.22 主数据变更自动同步
      };
    });
  };

  // ================= 申报信息页（分面浏览） =================
  var declareAll = [];
  var declareFilter = { q: '', missing: '', currency: '', brand: '', hsPrefix: '' };
  var declareExpand = { brand: false, hs: false };
  var declarePage = 1;          // 当前页（从1起）
  var declarePageSize = 100;    // 每页条数：50/100/200

  function declareRowHtml(d) {
    return '<tr data-sku="' + esc(d.sku) + '">' +
      '<td class="mono">' + esc(d.sku) + '</td>' +
      '<td class="mono">' + esc(d.styleCode || '') + '</td>' +
      '<td>' + esc(d.goodsName || '') + '</td>' +
      '<td>' + esc(d.shortName || '') + '</td>' +
      '<td>' + esc(d.nameCn || '') + '</td>' +
      '<td>' + esc(d.nameEn || '') + '</td>' +
      '<td class="mono">' + esc(d.hsCode || '') + '</td>' +
      '<td class="num">' + (d.declarePrice || '') + '</td>' +
      '<td class="mono">' + esc(d.currency || 'USD') + (d.declarePriceRaw && d.currency && d.currency !== 'USD' ? ' <span class="hint" title="原币值 ' + d.declarePriceRaw + ' ' + d.currency + '">≈' + d.declarePriceRaw + '</span>' : '') + '</td>' +
      '<td class="num">' + (d.purchasePriceCny != null ? d.purchasePriceCny : '') + '</td>' +
      '<td class="num">' + (d.purchasePriceCnyTax != null ? d.purchasePriceCnyTax : '') + '</td>' +
      '<td>' + esc(d.material || '') + '</td>' +
      '<td>' + esc(d.soleMaterial || '') + '</td>' +
      '<td>' + esc(d.brand || '') + '</td>' +
      '<td>' + esc(d.category || '') + '</td>' +
      '<td class="num">' + (d.nw || '') + '</td>' +
      '<td><button class="btn sm ghost dc-edit" data-sku="' + esc(d.sku) + '">编辑</button> <button class="btn sm danger dc-del" data-sku="' + esc(d.sku) + '">删除</button></td></tr>';
  }
  function declareMatches(d, f) {
    if (f.q) {
      var t = [d.sku, d.styleCode, d.goodsName, d.shortName, d.nameCn, d.nameEn, d.hsCode, d.material, d.soleMaterial, d.brand, d.category].join(' ').toLowerCase();
      if (t.indexOf(f.q) < 0) return false;
    }
    if (f.missing) {
      if (f.missing === 'hs' && d.hsCode) return false;
      if (f.missing === 'price' && d.declarePrice) return false;
      if (f.missing === 'material' && d.material) return false;
      if (f.missing === 'namecn' && d.nameCn) return false;
    }
    if (f.currency && (d.currency || 'USD') !== f.currency) return false;
    if (f.brand && (d.brand || '') !== f.brand) return false;
    if (f.hsPrefix && (d.hsCode || '').indexOf(f.hsPrefix) !== 0) return false;
    return true;
  }
  function declareMatchesExcept(d, f, excludeKey) {
    var f2 = {};
    for (var k in f) { if (k !== excludeKey) f2[k] = f[k]; }
    return declareMatches(d, f2);
  }
  function facetCounts(keyFn, f, excludeKey) {
    var map = {};
    declareAll.forEach(function (d) {
      if (excludeKey && !declareMatchesExcept(d, f, excludeKey)) return;
      var k = keyFn(d);
      if (k == null || k === '') return;
      map[k] = (map[k] || 0) + 1;
    });
    return map;
  }
  function hsChapter(d) {
    var h = d.hsCode || '';
    if (h.length >= 4) return h.slice(0, 4);
    return h ? h : '未填写';
  }
  function chipHtml(facet, val, label, count, activeVal) {
    var active = (activeVal === val) ? ' active' : '';
    return '<button class="chip' + active + '" data-facet="' + facet + '" data-val="' + esc(val) + '">' + esc(label) + (count != null ? ' <i>' + count + '</i>' : '') + '</button>';
  }

  PAGES.declares = async function () {
    declareAll = await db.all('declare_reqs');
    declareAll.sort(function (a, b) { return a.sku < b.sku ? -1 : 1; });
    return '<h2>申报信息（申报要素主数据 · 分面浏览）</h2>' +
      '<p class="hint">数据来源：本地《商品申报信息》表（D:/模板/商品申报信息.xlsx，约 13019 行），由 build_declare_from_xlsx.py 重建为 js/declare_data.js。点「🔄 并入申报主数据」以该表整表刷新本地——仅填空字段，<b>不覆盖</b>手填项。</p>' +
      '<div class="card"><div class="toolbar">' +
      '<button class="btn" id="dc-add">＋新增SKU</button>' +
      '<label class="btn ghost" style="display:inline-block">📥 从 xlsx/xls/csv 导入<input type="file" id="dc-file" accept=".xlsx,.xls,.csv" style="display:none"></label>' +
      '<button class="btn warn" id="dc-pull">🔄 并入申报主数据</button>' +
      '<input type="text" id="dc-search" class="input" placeholder="搜索 SKU / 款式编码 / 商品名称 / 品名 / HS编码 / 材质 / 品牌 / 分类" style="width:320px;margin-left:auto">' +
      '</div>' +
      '<div id="dc-facets" class="facets"></div>' +
      '<div id="dc-count" class="hint"></div>' +
      '<table class="grid"><thead><tr><th>SKU</th><th>款式编码</th><th>商品名称</th><th>商品简称</th><th>中文品名</th><th>英文品名</th><th>HS编码</th><th>申报价(USD)</th><th>币种</th><th>采购单价CNY</th><th>采购单价含税CNY</th><th>材质</th><th>鞋底材质</th><th>品牌</th><th>分类</th><th>单件净重</th><th>操作</th></tr></thead><tbody id="dc-tbody"></tbody></table>' +
      '<div id="dc-pager" class="pager"></div></div>';
  };
  function declareForm(d) {
    d = d || {};
    showModal('<h3>' + (d.sku ? '编辑' : '新增') + '申报要素</h3><div class="form-grid" style="margin-top:12px">' +
      '<div><label class="req">SKU</label><input id="dc-sku" value="' + esc(d.sku || '') + '"' + (d.sku ? ' readonly' : '') + '></div>' +
      '<div><label>款式编码</label><input id="dc-stylecode" value="' + esc(d.styleCode || '') + '"></div>' +
      '<div><label>商品名称</label><input id="dc-goodsname" value="' + esc(d.goodsName || '') + '"></div>' +
      '<div><label>商品简称</label><input id="dc-shortname" value="' + esc(d.shortName || '') + '"></div>' +
      '<div><label>中文品名</label><input id="dc-namecn" value="' + esc(d.nameCn || '') + '"></div>' +
      '<div><label class="req">英文品名</label><input id="dc-nameen" value="' + esc(d.nameEn || '') + '"></div>' +
      '<div><label class="req">HS编码</label><input id="dc-hs" value="' + esc(d.hsCode || '') + '"></div>' +
      '<div><label class="req">申报价(USD)</label><input id="dc-price" type="number" step="0.01" value="' + (d.declarePrice || '') + '" title="以USD计；若下方币种非USD，保存时自动折算为USD并记原币值"></div>' +
      '<div><label>币种</label><input id="dc-cur" value="' + esc(d.currency || 'USD') + '"></div>' +
      '<div><label>采购单价CNY</label><input id="dc-pcny" type="number" step="0.01" value="' + (d.purchasePriceCny != null ? d.purchasePriceCny : '') + '"></div>' +
      '<div><label>采购单价含税CNY</label><input id="dc-pcnytax" type="number" step="0.01" value="' + (d.purchasePriceCnyTax != null ? d.purchasePriceCnyTax : '') + '"></div>' +
      '<div><label class="req">材质</label><input id="dc-mat" value="' + esc(d.material || '') + '"></div>' +
      '<div><label>鞋底材质</label><input id="dc-sole" value="' + esc(d.soleMaterial || '') + '"></div>' +
      '<div><label>用途</label><input id="dc-usage" value="' + esc(d.usage || '') + '"></div>' +
      '<div><label>品牌</label><input id="dc-brand" value="' + esc(d.brand || 'NO BRAND') + '"></div>' +
      '<div><label>分类</label><input id="dc-cat" value="' + esc(d.category || '') + '"></div>' +
      '<div><label>型号</label><input id="dc-model" value="' + esc(d.model || '') + '"></div>' +
      '<div><label>单件净重KG</label><input id="dc-nw" type="number" step="0.001" value="' + (d.nw || '') + '"></div>' +
      '<div><label>单件毛重KG</label><input id="dc-gw" type="number" step="0.001" value="' + (d.gw || '') + '"></div>' +
      '<div><label>单位</label><input id="dc-unit" value="' + esc(d.unit || 'PCS') + '"></div>' +
      '<div><label>原产地</label><input id="dc-origin" value="' + esc(d.origin || 'CN') + '"></div></div>' +
      '<div style="text-align:right;margin-top:14px;display:flex;gap:8px;justify-content:flex-end">' +
      '<button class="btn ghost" onclick="TDUI.closeModal()">取消</button><button class="btn" id="dc-save">保存</button></div>');
    document.getElementById('dc-save').onclick = async function () {
      var sku = val('dc-sku');
      if (!sku) { toast('SKU必填', 'err'); return; }
      var dcCur = val('dc-cur') || 'USD';
      var dcRaw = parseFloat(val('dc-price')) || 0;
      await db.put('declare_reqs', {
        sku: sku, styleCode: val('dc-stylecode'), goodsName: val('dc-goodsname'), shortName: val('dc-shortname'),
        nameCn: val('dc-namecn'), nameEn: val('dc-nameen'), hsCode: val('dc-hs'),
        declarePrice: fxToUsd(dcRaw, dcCur), declarePriceRaw: dcRaw, currency: dcCur,
        purchasePriceCny: parseFloat(val('dc-pcny')) || 0, purchasePriceCnyTax: parseFloat(val('dc-pcnytax')) || 0,
        material: val('dc-mat'), soleMaterial: val('dc-sole'), usage: val('dc-usage'), brand: val('dc-brand'),
        category: val('dc-cat'),
        model: val('dc-model') || sku, nw: parseFloat(val('dc-nw')) || 0, gw: parseFloat(val('dc-gw')) || 0,
        unit: val('dc-unit') || 'PCS', origin: val('dc-origin') || 'CN', ver: (d.ver || 0) + 1
      });
      closeModal(); toast('已保存', 'ok'); render();
    };
  }
  BINDERS.declares = function () {
    document.getElementById('dc-add').onclick = function () { declareForm(null); };
    var searchInput = document.getElementById('dc-search');
    searchInput.oninput = function () { declareFilter.q = this.value.trim().toLowerCase(); declarePage = 1; repaintDeclares(); };
    var facetsEl = document.getElementById('dc-facets');
    facetsEl.addEventListener('click', function (e) {
      var chip = e.target.closest('.chip'); if (!chip) return;
      var facet = chip.dataset.facet, v = chip.dataset.val;
      if (facet === 'brand-more') { declareExpand.brand = !declareExpand.brand; repaintDeclares(); return; }
      if (facet === 'hs-more') { declareExpand.hs = !declareExpand.hs; repaintDeclares(); return; }
      if (facet === 'clear') { declareFilter = { q: '', missing: '', currency: '', brand: '', hsPrefix: '' }; searchInput.value = ''; declarePage = 1; repaintDeclares(); return; }
      declareFilter[facet] = (declareFilter[facet] === v) ? '' : v;
      declarePage = 1; repaintDeclares();
    });
    var pagerEl = document.getElementById('dc-pager');
    if (pagerEl) {
      pagerEl.addEventListener('click', function (e) {
        var btn = e.target.closest('.chip.pg'); if (!btn || btn.disabled) return;
        var matched = declareAll.filter(function (d) { return declareMatches(d, declareFilter); });
        var tp = Math.max(1, Math.ceil(matched.length / declarePageSize));
        var act = btn.dataset.act;
        if (act === 'first') declarePage = 1;
        else if (act === 'prev') declarePage = Math.max(1, declarePage - 1);
        else if (act === 'next') declarePage = Math.min(tp, declarePage + 1);
        else if (act === 'last') declarePage = tp;
        else if (btn.dataset.page) declarePage = parseInt(btn.dataset.page, 10) || 1;
        repaintDeclares();
      });
      pagerEl.addEventListener('change', function (e) {
        if (e.target.id === 'dc-pagesize') { declarePageSize = parseInt(e.target.value, 10) || 100; declarePage = 1; repaintDeclares(); }
        else if (e.target.id === 'dc-jump') { var v = parseInt(e.target.value, 10); if (v >= 1) declarePage = v; repaintDeclares(); }
      });
    }
    function bindDeclareRowButtons() {
      document.querySelectorAll('#dc-tbody .dc-edit').forEach(function (b) {
        b.onclick = async function () { declareForm(await db.get('declare_reqs', b.dataset.sku)); };
      });
      document.querySelectorAll('#dc-tbody .dc-del').forEach(function (b) {
        b.onclick = async function () {
          if (!(await confirmBox('确认删除 SKU ' + esc(b.dataset.sku) + ' 的申报要素？', true))) return;
          await db.del('declare_reqs', b.dataset.sku); toast('已删除', 'ok'); render();
        };
      });
    }
    function repaintDeclares() {
      var matched = declareAll.filter(function (d) { return declareMatches(d, declareFilter); });
      var total = matched.length;
      var pageSize = declarePageSize;
      var totalPages = Math.max(1, Math.ceil(total / pageSize));
      if (declarePage > totalPages) declarePage = totalPages;
      if (declarePage < 1) declarePage = 1;
      var start = (declarePage - 1) * pageSize;
      var pageRows = matched.slice(start, start + pageSize);
      var tbody = document.getElementById('dc-tbody');
      if (!pageRows.length) {
        tbody.innerHTML = '<tr><td colspan="17" class="empty">无匹配记录</td></tr>';
      } else {
        tbody.innerHTML = pageRows.map(declareRowHtml).join('');
      }
      bindDeclareRowButtons();
      document.getElementById('dc-count').innerHTML = '共匹配 <b>' + total + '</b> 条，全部 ' + declareAll.length + ' 条' +
        (total > 0 ? '（第 ' + declarePage + ' / ' + totalPages + ' 页，本页 ' + pageRows.length + ' 条）' : '');
      renderPager(totalPages);
      renderFacets();
    }
    function renderPager(totalPages) {
      var pager = document.getElementById('dc-pager');
      if (!pager) return;
      var html = '<div class="pg-left">每页显示 ' +
        '<select id="dc-pagesize" class="pg-sel">' +
        [50, 100, 200].map(function (n) { return '<option value="' + n + '"' + (n === declarePageSize ? ' selected' : '') + '>' + n + '</option>'; }).join('') +
        '</select> 条</div>';
      html += '<div class="pg-right">';
      html += '<button class="chip pg" data-act="first"' + (declarePage <= 1 ? ' disabled' : '') + '>« 首页</button>';
      html += '<button class="chip pg" data-act="prev"' + (declarePage <= 1 ? ' disabled' : '') + '>‹ 上一页</button>';
      var pages = [], cur = declarePage, tp = totalPages;
      if (tp <= 7) {
        for (var i = 1; i <= tp; i++) pages.push(i);
      } else {
        pages.push(1);
        var s = Math.max(2, cur - 2), e = Math.min(tp - 1, cur + 2);
        if (s > 2) pages.push('...');
        for (var j = s; j <= e; j++) pages.push(j);
        if (e < tp - 1) pages.push('...');
        pages.push(tp);
      }
      pages.forEach(function (p) {
        if (p === '...') html += '<span class="pg-gap">…</span>';
        else html += '<button class="chip pg' + (p === cur ? ' active' : '') + '" data-page="' + p + '">' + p + '</button>';
      });
      html += '<button class="chip pg" data-act="next"' + (declarePage >= totalPages ? ' disabled' : '') + '>下一页 ›</button>';
      html += '<button class="chip pg" data-act="last"' + (declarePage >= totalPages ? ' disabled' : '') + '>末页 »</button>';
      html += ' 跳至 <input id="dc-jump" class="pg-jump" type="number" min="1" max="' + totalPages + '" value="' + declarePage + '"> 页';
      html += '</div>';
      pager.innerHTML = html;
    }
    function renderFacets() {
      var f = declareFilter, html = '';
      // 完整性
      var mAll = (facetCounts(function () { return 'all'; }, f, 'missing')['all']) || 0;
      function missCount(field) {
        var n = 0;
        declareAll.forEach(function (d) {
          if (!declareMatchesExcept(d, f, 'missing')) return;
          if (field === 'hs' && !d.hsCode) n++;
          else if (field === 'price' && !d.declarePrice) n++;
          else if (field === 'material' && !d.material) n++;
          else if (field === 'namecn' && !d.nameCn) n++;
        });
        return n;
      }
      html += '<div class="facet-row"><span class="facet-label">完整性</span>' +
        chipHtml('missing', '', '全部', mAll, f.missing) +
        chipHtml('missing', 'hs', '缺HS编码', missCount('hs'), f.missing) +
        chipHtml('missing', 'price', '缺申报价', missCount('price'), f.missing) +
        chipHtml('missing', 'material', '缺材质', missCount('material'), f.missing) +
        chipHtml('missing', 'namecn', '缺中文名', missCount('namecn'), f.missing) + '</div>';
      // 币种
      var cur = facetCounts(function (d) { return d.currency || 'USD'; }, f, 'currency');
      html += '<div class="facet-row"><span class="facet-label">币种</span>';
      Object.keys(cur).sort(function (a, b) { return cur[b] - cur[a]; }).forEach(function (k) {
        html += chipHtml('currency', k, k, cur[k], f.currency);
      });
      html += '</div>';
      // 品牌
      var br = facetCounts(function (d) { return d.brand || '未填写'; }, f, 'brand');
      var brKeys = Object.keys(br).sort(function (a, b) { return br[b] - br[a]; });
      var brShow = declareExpand.brand ? brKeys : brKeys.slice(0, 12);
      html += '<div class="facet-row"><span class="facet-label">品牌</span>';
      brShow.forEach(function (k) { html += chipHtml('brand', k, k === '未填写' ? '(未填写)' : k, br[k], f.brand); });
      if (brKeys.length > 12) html += '<button class="chip more" data-facet="brand-more">' + (declareExpand.brand ? '收起' : '更多 ' + (brKeys.length - 12) + ' ▾') + '</button>';
      html += '</div>';
      // HS章节
      var hs = facetCounts(hsChapter, f, 'hsPrefix');
      var hsKeys = Object.keys(hs).sort(function (a, b) { return hs[b] - hs[a]; });
      var hsShow = declareExpand.hs ? hsKeys : hsKeys.slice(0, 12);
      html += '<div class="facet-row"><span class="facet-label">HS章节</span>';
      hsShow.forEach(function (k) { html += chipHtml('hsPrefix', k, k === '未填写' ? '(未填写)' : k, hs[k], f.hsPrefix); });
      if (hsKeys.length > 12) html += '<button class="chip more" data-facet="hs-more">' + (declareExpand.hs ? '收起' : '更多 ' + (hsKeys.length - 12) + ' ▾') + '</button>';
      html += '</div>';
      // 清除筛选
      if (f.q || f.missing || f.currency || f.brand || f.hsPrefix) {
        html += '<div class="facet-row"><button class="chip clear" data-facet="clear">✕ 清除全部筛选</button></div>';
      }
      facetsEl.innerHTML = html;
    }
    bindDeclareRowButtons();
    repaintDeclares();
    document.getElementById('dc-file').onchange = async function () {
      var f = this.files[0]; if (!f) return;
      try {
        var wb = await loadWb(await f.arrayBuffer(), f.name);
        var rows = parser.sheetToRows(wb.worksheets[0]);
        var aliases = {
          sku: ['SKU', 'sku', '商家编码', '商品编码'], nameCn: ['中文品名', '品名', '中文名称'], nameEn: ['英文品名', '英文名称', 'DESCRIPTION'],
          hsCode: ['HS编码', 'HS CODE', 'HSCODE', '海关编码'], declarePrice: ['申报单价', '申报价', '单价'], material: ['材质'],
          usage: ['用途'], brand: ['品牌'], model: ['型号'], nw: ['净重', '单件净重'], gw: ['毛重', '单件毛重'], origin: ['原产地', '产地']
        };
        var h = parser.scanHeader(rows, aliases, 3);
        if (!h) throw new Error('未识别表头（至少需SKU/HS编码/申报单价等3列）');
        var list = [];
        for (var r = h.headerRow + 1; r < rows.length; r++) {
          var row = rows[r] || [];
          var sku = parser.cellText(row[h.colMap.sku]).trim();
          if (!sku) continue;
          var o = { sku: sku };
          Object.keys(h.colMap).forEach(function (k) {
            if (k === 'sku') return;
            var v = row[h.colMap[k]];
            o[k] = (k === 'declarePrice' || k === 'nw' || k === 'gw') ? parser.toNum(v) : parser.cellText(v).trim();
          });
          o.unit = o.unit || 'PCS'; o.currency = o.currency || 'USD'; o.model = o.model || sku;
          if (o.currency !== 'USD' && o.declarePrice) { o.declarePriceRaw = o.declarePrice; o.declarePrice = fxToUsd(o.declarePrice, o.currency); }
          else if (o.declarePrice) { o.declarePriceRaw = o.declarePrice; }
          list.push(o);
        }
        if (!list.length) throw new Error('无有效数据行');
        await db.bulkPut('declare_reqs', list);
        toast('导入/更新 ' + list.length + ' 条申报要素', 'ok'); render();
      } catch (e) { toast('导入失败: ' + e.message, 'err'); }
    };
    document.getElementById('dc-pull').onclick = async function () {
      toast('正在以《商品申报信息》表刷新本地申报要素…');
      var r = await syncDeclaresFromMirror();
      if (r.ok) { toast('✅ 同步完成：新增 ' + r.added + ' 条、补齐 ' + r.merged + ' 条（不覆盖手填项）', 'ok'); }
      else {
        var cfg = await db.get('config', 'remote');
        if (cfg && cfg.value && cfg.value.baseURL) {
          var remote = new adapters.RemoteHttpAdapter(cfg.value);
          var rr = await adapters.syncFromRemote(db, remote, 'declare_reqs');
          if (rr.ok) { toast('远程同步成功，合并 ' + rr.merged + ' 条', 'ok'); }
          else toast('同步失败: ' + rr.error, 'err');
        } else toast('同步失败: ' + r.error, 'err');
      }
      render();
    };
  };

  // 从飞书《申报信息》镜像（js/declare_data.js，页面加载时已注入 window.TD.declareData）
  // 做字段级合并：本地已有 SKU 仅填空字段、不覆盖手填值；本地缺失 SKU 新增。
  async function syncDeclaresFromMirror() {
    var data = (window.TD && window.TD.declareData) || [];
    if (!data.length) return { ok: false, added: 0, merged: 0, error: '本地未找到申报信息主数据（js/declare_data.js），请重新部署系统' };
    var existing = await db.all('declare_reqs');
    var map = {}; existing.forEach(function (e) { map[e.sku] = e; });
    var FIELDS = ['nameCn', 'nameEn', 'hsCode', 'declarePrice', 'declarePriceRaw', 'material', 'brand', 'usage', 'unit', 'nw', 'gw', 'currency', 'model', 'origin'];
    function isBlank(v) { return v === undefined || v === null || v === '' || (typeof v === 'number' && isNaN(v)); }
    var added = 0, merged = 0;
    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (!r || !r.sku) continue;
      var cur = map[r.sku];
      if (!cur) { await db.put('declare_reqs', Object.assign({}, r)); added++; }
      else {
        var changed = false;
        FIELDS.forEach(function (k) { if (isBlank(cur[k]) && !isBlank(r[k])) { cur[k] = r[k]; changed = true; } });
        if (changed) { await db.put('declare_reqs', cur); merged++; }
      }
    }
    return { ok: true, added: added, merged: merged };
  }

  // ================= 模板库页 =================
  PAGES.templates = async function () {
    var ts = await db.all('templates');
    ts.sort(function (a, b) { return (a.kind + a.name) < (b.kind + b.name) ? -1 : 1; });
    var rows = ts.map(function (t) {
      // v1.5.25 类型列改为可编辑下拉：直接改分类 + 自动保存 + 自动团队同步
      var kindOpts = [
        ['invoice', '发票'], ['booking', '订舱单'], ['packing', '装箱单'], ['declare', '申报']
      ].map(function (o) { return '<option value="' + o[0] + '"' + (t.kind === o[0] ? ' selected' : '') + '>' + o[1] + '</option>'; }).join('');
      return '<tr><td>' + esc(t.name) + (t.builtin ? ' <span class="badge gray">内置</span>' : '') + '</td>' +
        '<td><select class="tp-kind-sel" data-id="' + t.id + '" style="padding:2px 6px;font-size:12px">' + kindOpts + '</select></td>' +
        '<td>' + esc(t.carrier || '') + '</td>' +
        '<td>' + (t.status === 'active' ? '<span class="badge green">启用</span>' : '<span class="badge red">停用</span>') + '</td>' +
        '<td>' + fmtTime(t.createdAt) + '</td>' +
        '<td><button class="btn sm ghost tp-prev" data-id="' + t.id + '">预览</button> ' +
        '<button class="btn sm ghost tp-map" data-id="' + t.id + '">字段</button> ' +
        '<button class="btn sm ' + (t.status === 'active' ? 'warn' : 'ok') + ' tp-toggle" data-id="' + t.id + '">' + (t.status === 'active' ? '停用' : '启用') + '</button> ' +
        '<button class="btn sm danger tp-del" data-id="' + t.id + '">删除</button></td></tr>';
    }).join('');
    var carrierOpts = CARRIERS.map(function (c) { return '<option>' + c + '</option>'; }).join('');
    function tplUploadCard() {
      var r = window.__lastTplUpload; if (!r) return '';
      if (r.ok) {
        var sz = r.size >= 1048576 ? (r.size / 1048576).toFixed(1) + 'MB' : Math.round(r.size / 1024) + 'KB';
        return '<div class="card" style="border-left:4px solid #2ecc71;background:rgba(46,204,113,.08)"><h3>✓ 上传并解析成功</h3>' +
          '<p>' + esc(r.name) + ' · ' + esc(r.kind) + ' · ' + sz + (r.converted ? ' · 已从 .xls 转为 .xlsx' : '') + '</p>' +
          '<p>Sheet：' + r.sheets.map(esc).join('、') + '</p>' +
          '<p>扫描到 <b>' + r.fields + '</b> 个表头占位符 + <b>' + r.itemFields + '</b> 个明细占位符</p></div>';
      }
      return '<div class="card" style="border-left:4px solid #e74c3c;background:rgba(231,76,60,.08)"><h3>✗ 上传解析失败</h3><p>' + esc(r.error) + '</p></div>';
    }
    var uploadCard = tplUploadCard();
    return (uploadCard ? uploadCard + '<br>' : '') + '<h2>模板库</h2><div class="card"><h3>上传新模板</h3><div class="form-grid">' +
      '<div><label class="req">模板名称</label><input id="tp-name" placeholder="如 亚丰发票模板v2"></div>' +
      '<div><label class="req">类型</label><select id="tp-kind"><option value="invoice">发票</option><option value="booking">订舱单(BOOKING FORM)</option><option value="packing">装箱单(PACKING LIST)</option><option value="declare">申报/买单要素</option></select></div>' +
      '<div><label>物流商</label><input id="tp-carrier" list="carrier-list" placeholder="通用"><datalist id="carrier-list">' + carrierOpts + '</datalist></div>' +
      '<div><label class="req">模板文件(.xlsx / .xls / .csv)</label><input type="file" id="tp-file" accept=".xlsx,.xls,.csv"></div></div>' +
      '<p class="hint">模板中用 <span class="mono">{{字段名}}</span> 做占位符（如 {{invoiceNo}} {{shipper.name}}），明细行用 <span class="mono">{{items.sku}} {{items.qty}}</span>，系统会自动按明细条数复制该行。上传后自动扫描占位符。</p>' +
      '<button class="btn" id="tp-upload">上传并解析</button></div>' +
      '<div class="card"><table class="grid"><tr><th>名称</th><th>类型</th><th>物流商</th><th>状态</th><th>上传时间</th><th>操作</th></tr>' +
      (rows || '<tr><td colspan="6" class="empty">暂无模板</td></tr>') + '</table></div>';
  };
  BINDERS.templates = function () {
    document.getElementById('tp-upload').onclick = async function () {
      var name = val('tp-name'), kind = val('tp-kind'), carrier = val('tp-carrier') || '通用';
      var f = document.getElementById('tp-file').files[0];
      if (!name || !f) { toast('模板名称和文件必填', 'err'); return; }
      try {
        var buf = await f.arrayBuffer();
        var wb = await loadWb(buf);
        var scan = engine.scanTemplate(wb);
        if (!scan.fields.length && !scan.itemFields.length) {
          if (!(await confirmBox('⚠️ 该模板中未扫描到任何 {{占位符}}，生成时将无法自动填充数据。仍要上传吗？'))) return;
        }
        autoSyncToTeam(); // v1.5.22 模板变更自动同步
        await db.put('templates', {
          name: name, kind: kind, carrier: carrier, status: 'active',
          fileBuf: buf, fileName: f.name, fileSize: buf.byteLength || buf.length || 0,
          uploadedAt: Date.now(),
          mapping: { required: engine.REQUIRED_FIELDS[kind] || [], scanned: scan, labelMap: engine.buildLabelMap(wb, scan.itemsRow || -1) }
        });
        window.__lastTplUpload = { ok: true, name: name, kind: kind, size: (buf.byteLength || buf.length || 0), sheets: wb.worksheets.map(function (w) { return w.name; }), fields: scan.fields.length, itemFields: scan.itemFields.length, converted: !!(wb.__formatGuess && wb.__formatGuess !== 'xlsx') };
        toast('模板已上传：扫描到 ' + scan.fields.length + ' 个表头字段 + ' + scan.itemFields.length + ' 个明细字段', 'ok');
        render();
      } catch (e) { window.__lastTplUpload = { ok: false, error: e.message }; toast('模板解析失败: ' + e.message, 'err'); render(); }
    };
    document.querySelectorAll('.tp-prev').forEach(function (b) {
      b.onclick = async function () {
        var t = await db.get('templates', b.dataset.id);
        // 方案 B：模板库预览直接显示源文件原样（含 LOGO + 样张公司名/地址），不走 normalize
        var buf = t.previewBuf || t.fileBuf;
        var wb = await loadWb(buf);
        // 源文件被 ExcelJS 载入时可能已带有内嵌图，避免重复贴图
        var hasImg = wb.worksheets[0].getImages && wb.worksheets[0].getImages().length > 0;
        if (t.logo && t.logo.dataB64 && !hasImg) engine.addLogo(wb, wb.worksheets[0], t.logo);
        showModal('<h3>模板预览 · ' + esc(t.name) + '</h3><div style="overflow:auto">' + wbToHtml(wb) + '</div>' +
          '<div style="text-align:right;margin-top:12px"><button class="btn" onclick="TDUI.closeModal()">关闭</button></div>');
      };
    });
    document.querySelectorAll('.tp-map').forEach(function (b) {
      b.onclick = async function () {
        var t = await db.get('templates', b.dataset.id);
        var wb = await loadWb(t.fileBuf);
        var scan = engine.scanTemplate(wb);
        // 字段映射候选（数据模型字段）
        var CAND = [
          { p: 'invoiceNo', l: '发票号' }, { p: 'invoiceDate', l: '发票日期' }, { p: 'contractNo', l: '合同号' }, { p: 'orderNos', l: '订单号' }, { p: 'refId', l: '参考号' },
          { p: 'shipper.name', l: '发货人名称' }, { p: 'shipper.company', l: '发货人公司' }, { p: 'shipper.address', l: '发货人地址' }, { p: 'shipper.city', l: '发货人城市' }, { p: 'shipper.state', l: '发货人省/州' }, { p: 'shipper.zip', l: '发货人邮编' }, { p: 'shipper.country', l: '发货人国家' }, { p: 'shipper.tel', l: '发货人电话' }, { p: 'shipper.email', l: '发货人邮箱' }, { p: 'shipper.taxNo', l: '发货人税号' },
          { p: 'consignee.name', l: '收货人名称' }, { p: 'consignee.company', l: '收货人公司' }, { p: 'consignee.address', l: '收货人地址' }, { p: 'consignee.city', l: '收货人城市' }, { p: 'consignee.state', l: '收货人省/州' }, { p: 'consignee.zip', l: '收货人邮编' }, { p: 'consignee.country', l: '收货人国家' }, { p: 'consignee.tel', l: '收货人电话' }, { p: 'consignee.email', l: '收货人邮箱' }, { p: 'consignee.taxNo', l: '收货人税号' },
          { p: 'incoterms', l: '贸易条款' }, { p: 'pol', l: '起运港' }, { p: 'pod', l: '目的港' }, { p: 'transport', l: '运输方式' }, { p: 'etd', l: '船期' }, { p: 'vessel', l: '船名航次' }, { p: 'containerType', l: '柜型' }, { p: 'containerQty', l: '柜量' }, { p: 'freightTerms', l: '运费条款' }, { p: 'paymentTerms', l: '付款条款' }, { p: 'customsType', l: '报关方式' }, { p: 'agent', l: '订舱代理' }, { p: 'shippingMarks', l: '唛头' }, { p: 'remark', l: '备注' }, { p: 'goodsSummary', l: '品名概述' },
          { p: 'totals.boxCount', l: '总箱数' }, { p: 'totals.gw', l: '总毛重' }, { p: 'totals.nw', l: '总净重' }, { p: 'totals.volume', l: '总体积' }, { p: 'totals.amount', l: '总金额' }, { p: 'totals.qty', l: '总数量' },
          { p: 'items.sku', l: 'SKU' }, { p: 'items.nameCn', l: '中文品名' }, { p: 'items.nameEn', l: '英文品名' }, { p: 'items.hsCode', l: '海关编码' }, { p: 'items.qty', l: '数量' }, { p: 'items.unit', l: '单位' }, { p: 'items.price', l: '单价' }, { p: 'items.amount', l: '金额' }, { p: 'items.nw', l: '净重' }, { p: 'items.gw', l: '毛重' }, { p: 'items.material', l: '材质' }, { p: 'items.brand', l: '品牌' }, { p: 'items.model', l: '型号' }, { p: 'items.boxNo', l: '箱号' }, { p: 'items.boxCount', l: '箱数' }, { p: 'items.dims', l: '尺寸' }, { p: 'items.origin', l: '原产国' }, { p: 'items.usage', l: '用途' }
        ];
        var lm = (t.mapping && t.mapping.labelMap && t.mapping.labelMap.length)
          ? t.mapping.labelMap.slice()
          : engine.buildLabelMap(wb, scan.itemsRow || -1);
        function rowHtml(e, i) {
          var opts = '<option value="">（不映射）</option>' + CAND.map(function (c) {
            return '<option value="' + c.p + '"' + (e.path === c.p ? ' selected' : '') + '>' + c.l + ' · ' + c.p + '</option>';
          }).join('');
          return '<tr><td class="mono" style="max-width:240px;overflow:hidden;text-overflow:ellipsis">' + esc(e.label) + '</td>' +
            '<td><select data-idx="' + i + '" style="max-width:260px">' + opts + '</select></td>' +
            '<td style="color:' + (e.resolved ? '#2ecc71' : '#e67e22') + '">' + (e.resolved ? '✓自动' : '⚠未识别') + '</td></tr>';
        }
        var html = '<h3>字段映射 · ' + esc(t.name) + '</h3>' +
          '<p class="hint">系统已自动识别模板里的中文标签并映射到数据字段。下拉可改；选「不映射」则该标签原样保留。改完自动存库。</p>' +
          '<div class="card"><table id="tp-maptable" class="grid"><tr><th>模板标签（如 收件人国家代码）</th><th>映射到字段</th><th>状态</th></tr>' +
          lm.map(rowHtml).join('') + '</table></div>' +
          '<div style="margin-top:10px"><button class="btn" id="tp-rescan">🔄 重新扫描标签</button> ' +
          '<button class="btn" onclick="TDUI.closeModal()">关闭</button></div>';
        showModal(html);
        var wrapRows = function () {
          var tbl = document.getElementById('tp-maptable');
          tbl.querySelectorAll('select').forEach(function (sel) {
            sel.onchange = async function () {
              var i = +sel.dataset.idx; lm[i].path = sel.value; lm[i].resolved = !!sel.value;
              t.mapping = t.mapping || {}; t.mapping.labelMap = lm;
              await db.put('templates', t); toast('已保存映射', 'ok'); autoSyncToTeam(); // v1.5.22
            };
          });
        };
        wrapRows();
        document.getElementById('tp-rescan').onclick = async function () {
          lm = engine.buildLabelMap(wb, scan.itemsRow || -1);
          t.mapping = t.mapping || {}; t.mapping.labelMap = lm;
          await db.put('templates', t);
          showModal('<h3>字段映射 · ' + esc(t.name) + '</h3>' +
            '<p class="hint">已重新扫描。下拉可改，改完自动存库。</p>' +
            '<div class="card"><table id="tp-maptable" class="grid"><tr><th>模板标签</th><th>映射到字段</th><th>状态</th></tr>' +
            lm.map(rowHtml).join('') + '</table></div>' +
            '<div style="margin-top:10px"><button class="btn" id="tp-rescan">🔄 重新扫描标签</button> ' +
            '<button class="btn" onclick="TDUI.closeModal()">关闭</button></div>');
          wrapRows();
          document.getElementById('tp-rescan').onclick = arguments.callee;
          toast('已重新扫描标签', 'ok');
        };
      };
    });
    document.querySelectorAll('.tp-toggle').forEach(function (b) {
      b.onclick = async function () {
        var t = await db.get('templates', b.dataset.id);
        t.status = t.status === 'active' ? 'disabled' : 'active';
        await db.put('templates', t);
        toast(t.status === 'active' ? '已启用' : '已停用（生成向导中将不可选）', 'ok'); render();
      };
    });
    // v1.5.25 模板类型下拉：改分类 → 自动保存 + 3秒后自动团队同步（v1.5.22 autoSyncToTeam）
    document.querySelectorAll('.tp-kind-sel').forEach(function (sel) {
      sel.onchange = async function () {
        var t = await db.get('templates', sel.dataset.id);
        if (!t || t.kind === sel.value) return;
        var old = t.kind; t.kind = sel.value; t.updatedAt = Date.now();
        await db.put('templates', t);
        toast('已改为「' + sel.options[sel.selectedIndex].text + '」（3秒后自动同步团队）', 'ok');
        autoSyncToTeam();
        render();
      };
    });
    document.querySelectorAll('.tp-del').forEach(function (b) {
      b.onclick = async function () {
        var docs = await db.all('documents');
        var refs = docs.filter(function (d) { return d.templateId === b.dataset.id; });
        var msg = refs.length ? '⚠️ 该模板已被 ' + refs.length + ' 条单证记录引用，删除后这些记录将无法重新导出。确认删除？' : '确认删除该模板？';
        if (!(await confirmBox(msg, true))) return;
        if (refs.length && !(await confirmBox('二次确认：真的要删除被引用的模板吗？', true))) return;
        await db.del('templates', b.dataset.id); toast('已删除', 'ok'); render(); autoSyncToTeam(); // v1.5.22
      };
    });
  };

  // ================= 生成发票向导 =================
  function wizSteps(cur, names) {
    return '<div class="steps">' + names.map(function (n, i) {
      return '<div class="step ' + (i === cur ? 'on' : i < cur ? 'done' : '') + '">' + (i + 1) + '. ' + n + '</div>';
    }).join('') + '</div>';
  }

  PAGES.invoice = async function () {
    if (!state.wiz) state.wiz = { step: 0, orderIds: [], packingId: '', templateId: '', carrier: '', channel: '', meta: {}, shipperId: '', consigneeId: '', notifyId: '', report: null, doc: null };
    var w = state.wiz;
    var names = ['勾选订单', '装箱清单&模板', '表头信息', '校验核对', '预览·确认·导出'];
    var html = '<h2>生成发票</h2>' + wizSteps(w.step, names) + '<div id="wiz-body"></div>';
    return html;
  };
  BINDERS.invoice = function () { renderWizStep(); };

  async function renderWizStep() {
    var w = state.wiz;
    var body = document.getElementById('wiz-body');
    if (!body) return;
    try {
      await renderWizStepInner();
    } catch (e) {
      console.error(e);
      body.innerHTML = '<div class="card vres block" style="padding:20px">' +
        '<h3 style="margin-top:0">⚠️ 步骤 ' + (w ? w.step : '?') + ' 渲染失败</h3>' +
        '<p style="color:#c0392b;margin:8px 0"><b>' + esc(e.message || String(e)) + '</b></p>' +
        '<pre style="background:#fafbfd;padding:8px;border-radius:6px;font-size:12px;max-height:200px;overflow:auto;white-space:pre-wrap">' + esc((e.stack || '').split('\n').slice(0, 4).join('\n')) + '</pre>' +
        '<p class="hint" style="margin-top:10px">💡 截图发给开发者可快速定位。常见诱因：①本地 IndexedDB 缺失模板数据；②v1.4.33 之前录入的旧数据 schema 与新版本不兼容；③模板 fileBuf 损坏。可尝试 <b>Ctrl+Shift+R 硬刷新</b> 重新走流程，或在「设置 → ⚠️ 危险操作」清空数据库后让种子数据重新注入。</p>' +
        (w ? '<div style="margin-top:10px;display:flex;gap:8px"><button class="btn ghost" id="err-wiz-home">← 回到发票向导 step 1</button></div>' : '') +
        '</div>';
      var back = document.getElementById('err-wiz-home');
      if (back && w) back.onclick = function () { w.step = 0; w.orderIds = []; w.packingId = ''; w.templateId = ''; w.meta = {}; w.report = null; w.doc = null; render(); };
    }
  }

  async function renderWizStepInner() {
    var w = state.wiz;
    var body = document.getElementById('wiz-body');

    // ---------- step0 选订单 ----------
    if (w.step === 0) {
      var orders = await db.all('orders');
      orders.sort(function (a, b) { return b.createdAt - a.createdAt; });
      var rows = orders.map(function (o) {
        var qty = (o.items || []).reduce(function (s, i) { return s + i.qty; }, 0);
        var ck = w.orderIds.indexOf(o.id) >= 0 ? ' checked' : '';
        return '<tr class="checkrow"><td><input type="checkbox" class="wz-ord" value="' + o.id + '"' + ck + '></td>' +
          '<td class="mono">' + esc(o.orderNo) + '</td><td>' + (o.source === 'from_packing' ? '<span class="badge purple">箱单直生</span>' : '<span class="badge blue">聚水潭</span>') + '</td>' +
          '<td>' + esc(o.buyer || '') + '</td><td class="num">' + (o.items || []).length + '</td><td class="num">' + qty + '</td></tr>';
      }).join('');
      body.innerHTML = '<div class="card"><p class="hint">可多选订单合并开一票（合并时将强校验装箱清单单号与SKU数量一致性）</p>' +
        '<table class="grid"><tr><th></th><th>订单号</th><th>来源</th><th>买家</th><th>SKU种数</th><th>数量</th></tr>' +
        (rows || '<tr><td colspan="6" class="empty">暂无订单，请先到「订单」或「装箱清单」页导入</td></tr>') + '</table>' +
        '<div style="margin-top:14px"><button class="btn" id="wz-next0">下一步 →</button></div></div>';
      document.getElementById('wz-next0').onclick = function () {
        var ids = Array.from(document.querySelectorAll('.wz-ord:checked')).map(function (c) { return c.value; });
        if (!ids.length) { toast('请至少勾选一个订单', 'err'); return; }
        w.orderIds = ids; w.step = 1; render();
      };
      return;
    }

    // ---------- step1 装箱清单&模板 ----------
    if (w.step === 1) {
      var orders = [], i;
      for (i = 0; i < w.orderIds.length; i++) orders.push(await db.get('orders', w.orderIds[i]));
      var selNos = orders.map(function (o) { return o.orderNo; });
      var pks = await db.all('packings');
      var _allInvTpls = (await db.all('templates')).filter(function (t) { return t.kind === 'invoice' && t.status === 'active'; });
      // v1.5.45：同步未拉完（fileBuf 空）的模板标记为残缺，排除在可用下拉外，避免用户选到 step4 才暴露
      var brokenTpls = _allInvTpls.filter(function (t) { return !t.fileBuf || !t.fileBuf.byteLength; });
      var tpls = _allInvTpls.filter(function (t) { return t.fileBuf && t.fileBuf.byteLength > 0; });
      // v1.4.45：当发票模板为空时，在 发票生成 页渲染一个明显的红色提示卡 + 「立即恢复」按钮，
      // 并在页面渲染时（不限 init 阶段）自动尝试一次静默远程拉取 + 拿到后立刻重渲。
      if (tpls.length === 0) {
        // v1.4.46：改用 _recoverInvoiceTemplates（add-only，绕过 threeWayMerge/db.clear）
        try {
          if (!window.__recovering_invoice__) {
            window.__recovering_invoice__ = 1;
            var prAuto = await _recoverInvoiceTemplates();
            if (prAuto && prAuto.ok && prAuto.got > 0) {
              var refreshed = (await db.all('templates')).filter(function (t) { return t.kind === 'invoice' && t.status === 'active'; });
              toast('已自动恢复 ' + refreshed.length + ' 个发票模板', 'ok');
              window.__recovering_invoice__ = 0;
              render(); return;
            }
            window.__recovering_invoice__ = 0;
          }
        } catch (e) { window.__recovering_invoice__ = 0; }
      }
      var banner = tpls.length === 0
        ? '<div id="wz-tpl-banner" class="card" style="border:2px solid #d97706;background:#fff8e1;padding:12px 14px;margin:0 0 12px 0"><div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap"><div style="flex:1;color:#92400e"><b>⚠ 本地无发票模板</b> —— 看起来团队共享库里的发票模板没拉下来。点击右侧按钮立即从团队库拉取。</div><button class="btn" id="wz-tpl-recover">立即恢复发票模板</button></div></div>'
        : '';
      // v1.5.45：有模板同步未完成（fileBuf 空）时，显示红色提示卡 + 重新同步按钮（不再让用户选到跑不到底的模板）
      var syncWarn = brokenTpls.length
        ? '<div id="wz-tpl-broken" class="card" style="border:2px solid #dc2626;background:#fef2f2;padding:12px 14px;margin:0 0 12px 0"><div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap"><div style="flex:1;color:#991b1b"><b>⚠ ' + brokenTpls.length + ' 个发票模板同步未完成</b>（模板文件拉取失败，暂不可用于导出）。点击右侧按钮重新从团队库拉取模板文件。</div><button class="btn warn" id="wz-tpl-resync2">↻ 重新同步模板</button></div></div>'
        : '';
      var pkRows = pks.map(function (p) {
        var inter = (p.orderNos || []).filter(function (n) { return selNos.indexOf(n) >= 0; });
        var hint = inter.length === selNos.length && inter.length === (p.orderNos || []).length ? '<span class="badge green">完全匹配</span>' :
          inter.length ? '<span class="badge yellow">部分匹配(' + inter.length + '/' + selNos.length + ')</span>' : '<span class="badge gray">无交集</span>';
        var ck = w.packingId === p.id ? ' checked' : '';
        return '<tr class="checkrow"><td><input type="radio" name="wz-pk" value="' + p.id + '"' + ck + '></td>' +
          '<td>' + esc(p.fileName) + '</td><td class="mono">' + esc((p.orderNos || []).join(', ')) + '</td><td class="num">' + p.totals.boxCount + '</td><td class="num">' + p.totals.qty + '</td><td>' + hint + '</td></tr>';
      }).join('');
      var tplOpts = tpls.map(function (t) { return '<option value="' + t.id + '"' + (w.templateId === t.id ? ' selected' : '') + '>' + esc(t.name) + '（' + esc(t.carrier) + '）</option>'; }).join('');
      body.innerHTML = banner + syncWarn + '<div class="card"><h3>① 选择装箱清单（用于单号/SKU/数量强校验及重量体积）</h3>' +
        '<p class="hint">📌 装箱清单决定「长(cm) / 宽(cm) / 高(cm) / 单箱重量」能否填出。海运类模板（如亚运发货最新）必须关联装箱清单，否则这些明细列为空。下方②选择模板后将在下一步预览核对。</p>' +
        '<table class="grid"><tr><th></th><th>文件名</th><th>关联单号</th><th>箱数</th><th>数量</th><th>与所选订单</th></tr>' +
        (pkRows || '<tr><td colspan="6" class="empty">暂无装箱清单，请先到「装箱清单」页上传</td></tr>') + '</table>' +
        '<h3 style="margin-top:16px">② 物流商 / 渠道 / 发票模板</h3><div class="form-grid">' +
        '<div><label>物流商</label><input id="wz-carrier" list="carrier-list2" value="' + esc(w.carrier) + '"><datalist id="carrier-list2">' + CARRIERS.map(function (c) { return '<option>' + c + '</option>'; }).join('') + '</datalist></div>' +
        '<div><label>渠道</label><input id="wz-channel" value="' + esc(w.channel) + '" placeholder="如 海运整柜/海快/美森"></div>' +
        '<div><label class="req">发票模板（仅显示启用中）</label><select id="wz-tpl"><option value="">请选择</option>' + tplOpts + '</select></div></div>' +
        '<div style="margin-top:14px;display:flex;gap:8px"><button class="btn ghost" id="wz-back1">← 上一步</button><button class="btn" id="wz-next1">下一步 →</button></div></div>';
      document.getElementById('wz-back1').onclick = function () { w.step = 0; render(); };
      // v1.4.45：发票模板空时的「立即恢复」按钮
      var rb = document.getElementById('wz-tpl-recover');
      if (rb) rb.onclick = async function () {
        rb.disabled = true; rb.textContent = '正在从团队库拉取…';
        try {
          // v1.4.46：改用 add-only 独立恢复函数，不动其他 store、不清表、不 merge
          var pr2 = await _recoverInvoiceTemplates();
          if (pr2 && pr2.ok) {
            var got = (await db.all('templates')).filter(function (t) { return t.kind === 'invoice' && t.status === 'active'; });
            if (got.length > 0) {
              toast('已从团队库拉取 ' + got.length + ' 个发票模板', 'ok');
              render();
            } else {
              toast('云端也没有可用发票模板（拉到 0）', 'err');
              rb.disabled = false; rb.textContent = '立即恢复发票模板';
            }
          } else {
            toast('拉取失败：' + (pr2 && pr2.error || '未知错误'), 'err');
            rb.disabled = false; rb.textContent = '立即恢复发票模板';
          }
        } catch (e) {
          toast('拉取异常：' + e.message, 'err');
          rb.disabled = false; rb.textContent = '立即恢复发票模板';
        }
      };
      document.getElementById('wz-next1').onclick = async function () {
        var pk = document.querySelector('input[name="wz-pk"]:checked');
        if (!pk) { toast('必须选择装箱清单（发票强校验依赖装箱清单）', 'err'); return; }
        var tpl = val('wz-tpl');
        if (!tpl) { toast('请选择发票模板', 'err'); return; }
        // v1.5.45：选中即校验 fileBuf，残缺模板直接拦截并提示重新同步（不再等到 step4 才暴露）
        var tplRec = await db.get('templates', tpl);
        if (!tplRec || !tplRec.fileBuf || !tplRec.fileBuf.byteLength) {
          toast('该模板文件未同步完整，请点上方「↻ 重新同步模板」', 'err');
          render(); return;
        }
        w.packingId = pk.value; w.templateId = tpl;
        w.carrier = val('wz-carrier'); w.channel = val('wz-channel');
        w.step = 2; render();
      };
      // v1.5.45：重新同步按钮 —— 重拉团队库（pullShared 对本地无 fileBuf 的模板会重新 fetch，可修复残缺）
      var rb2 = document.getElementById('wz-tpl-resync2');
      if (rb2) rb2.onclick = async function () {
        rb2.disabled = true; rb2.textContent = '正在重新同步…';
        try {
          var pr3 = await pullShared();
          if (pr3 && pr3.ok) {
            var got = (await db.all('templates')).filter(function (t) { return t.kind === 'invoice' && t.status === 'active' && t.fileBuf && t.fileBuf.byteLength > 0; });
            toast('已重新同步，可用发票模板 ' + got.length + ' 个', 'ok');
            render();
          } else {
            toast('同步失败：' + ((pr3 && pr3.error) || '未知'), 'err');
            rb2.disabled = false; rb2.textContent = '↻ 重新同步模板';
          }
        } catch (e) {
          toast('同步异常：' + e.message, 'err');
          rb2.disabled = false; rb2.textContent = '↻ 重新同步模板';
        }
      };
      return;
    }

    // ---------- step2 表头信息 ----------
    if (w.step === 2) {
      var orders = [];
      for (var i = 0; i < w.orderIds.length; i++) orders.push(await db.get('orders', w.orderIds[i]));
      var parties = await db.all('parties');
      var shippers = parties.filter(function (p) { return p.type === 'shipper'; });
      var consignees = parties.filter(function (p) { return p.type === 'consignee'; });
      var notifies = parties.filter(function (p) { return p.type === 'notify'; });
      var docs = await db.all('documents');
      var m = w.meta;
      if (!m.invoiceNo) m.invoiceNo = 'INV' + today().replace(/-/g, '') + '-' + String(docs.filter(function (d) { return d.kind === 'invoice'; }).length + 1).padStart(3, '0');
      if (!m.invoiceDate) m.invoiceDate = today();
      // 订单自带收货人
      var ordRecv = orders.find(function (o) { return o.receiver || o.address; });
      function opts(list, sel) { return '<option value="">请选择</option>' + list.map(function (p) { return '<option value="' + p.id + '"' + (sel === p.id ? ' selected' : '') + '>' + esc(p.name) + '</option>'; }).join(''); }
      body.innerHTML = '<div class="card"><h3>收发货人</h3><div class="form-grid">' +
        '<div><label class="req">SHIPPER 发货人</label><select id="wz-shipper">' + opts(shippers, w.shipperId) + '</select>' + (shippers.length === 0 ? '<p class="hint">📌 无发货人主数据。<button class="btn sm ghost" id="wz-use-ord-shipper">抓取订单的卖家/承运商为发货人</button></p>' : '') + '</div>' +
        '<div id="wz-shipper-info" class="party-info"></div>' +
        '<div><label class="req">CONSIGNEE 收货人</label><select id="wz-consignee">' + opts(consignees, w.consigneeId) + '</select>' +
        (ordRecv ? '<p class="hint">📌 订单自带收货人「' + esc(ordRecv.receiver || ordRecv.buyer) + '」<button class="btn sm ghost" id="wz-use-ord-recv">直接抓取使用</button></p>' : '<p class="hint">订单无收发货人信息，请从主数据选择 <button class="btn sm ghost" id="wz-goto-parties">去「收发货人」页维护</button></p>') + '</div>' +
        '<div id="wz-consignee-info" class="party-info"></div>' +
        '<div><label>NOTIFY 通知人</label><select id="wz-notify">' + opts(notifies, w.notifyId) + '</select></div></div>' +
        '<h3 style="margin-top:14px">发票要素</h3><div class="form-grid">' +
        '<div><label class="req">发票号</label><input id="wz-invno" value="' + esc(m.invoiceNo) + '"></div>' +
        '<div><label class="req">发票日期</label><input type="date" id="wz-invdate" value="' + esc(m.invoiceDate) + '"></div>' +
        '<div><label>合同号</label><input id="wz-contract" value="' + esc(m.contractNo || '') + '"></div>' +
        '<div><label class="req">贸易条款</label><select id="wz-inco">' + INCOTERMS.map(function (t) { return '<option' + (m.incoterms === t ? ' selected' : '') + '>' + t + '</option>'; }).join('') + '</select></div>' +
        '<div><label>付款方式</label><input id="wz-pay" value="' + esc(m.paymentTerms || 'T/T') + '"></div>' +
        '<div><label>运输方式</label><select id="wz-trans"><option' + (m.transport === 'BY SEA' ? ' selected' : '') + '>BY SEA</option><option' + (m.transport === 'BY AIR' ? ' selected' : '') + '>BY AIR</option><option' + (m.transport === 'BY EXPRESS' ? ' selected' : '') + '>BY EXPRESS</option><option' + (m.transport === 'BY RAIL' ? ' selected' : '') + '>BY RAIL</option></select></div>' +
        '<div><label class="req">起运港 POL</label><input id="wz-pol" value="' + esc(m.pol || 'SHENZHEN, CHINA') + '"></div>' +
        '<div><label class="req">目的港 POD</label><input id="wz-pod" value="' + esc(m.pod || '') + '"></div>' +
        '<div><label class="req">币种</label><select id="wz-cur"><option' + ((m.currency || 'USD') === 'USD' ? ' selected' : '') + '>USD</option><option' + (m.currency === 'RMB' ? ' selected' : '') + '>RMB</option><option' + (m.currency === 'EUR' ? ' selected' : '') + '>EUR</option></select></div>' +
        '<div><label>唛头</label><input id="wz-marks" value="' + esc(m.shippingMarks || 'N/M') + '"></div>' +
        '<div><label>备注</label><input id="wz-remark" value="' + esc(m.remark || '') + '"></div></div>' +
        '<div style="margin-top:14px;display:flex;gap:8px"><button class="btn ghost" id="wz-back2">← 上一步</button><button class="btn" id="wz-next2">下一步：校验 →</button></div></div>';
      var useBtn = document.getElementById('wz-use-ord-recv');
      if (useBtn) useBtn.onclick = async function () {
        autoSyncToTeam(); // v1.5.22 自动同步
        var p = await db.put('parties', { type: 'consignee', name: ordRecv.receiver || ordRecv.buyer, address: ordRecv.address || '', tel: ordRecv.phone || '', country: ordRecv.country || '', remark: '从订单 ' + ordRecv.orderNo + ' 抓取' });
        w.consigneeId = p.id; toast('已抓取订单收货人并存入主数据', 'ok'); renderWizStep();
      };
      var shipperBtn = document.getElementById('wz-use-ord-shipper');
      if (shipperBtn) shipperBtn.onclick = async function () {
        var so = orders.find(function (o) { return o.seller || o.sellerName || o.company; });
        if (!so) { toast('该订单没有卖家/发货人信息，请到「收发货人」页手动新增', 'err'); return; }
        autoSyncToTeam(); // v1.5.22 自动同步
        var p = await db.put('parties', { type: 'shipper', name: so.seller || so.sellerName || so.company, company: so.company || so.seller || '', address: so.sellerAddress || '', remark: '从订单 ' + so.orderNo + ' 抓取' });
        w.shipperId = p.id; toast('已抓取订单卖家为发货人并存入主数据', 'ok'); renderWizStep();
      };
      var gotoParties = document.getElementById('wz-goto-parties');
      if (gotoParties) gotoParties.onclick = function () {
        state.wiz = null; state.tab = 'parties';
        document.querySelectorAll('.tab').forEach(function (t) { t.classList.toggle('active', t.dataset.tab === 'parties'); });
        render();
      };
      // v1.4.50：选中收发货人后就地显示其摘要，避免选错看不出来（用户曾因模板自带样本残留误以为填错）
      function renderPartyInfo() {
        var allMap = {};
        parties.forEach(function (p) { allMap[p.id] = p; });
        ['wz-shipper-info|wz-shipper', 'wz-consignee-info|wz-consignee'].forEach(function (pair) {
          var parts = pair.split('|');
          var box = document.getElementById(parts[0]);
          var sel = document.getElementById(parts[1]);
          if (!box || !sel) return;
          var p = allMap[sel.value];
          if (!p) { box.innerHTML = '<span class="pi-empty">（未选择）</span>'; return; }
          var lines = [p.name, p.company, p.address, [p.country, p.state, p.city].filter(Boolean).join(' / '), p.tel, p.email, p.zip ? '邮编 ' + p.zip : '', p.warehouseCode ? '仓库代码 ' + p.warehouseCode : ''].filter(Boolean);
          box.innerHTML = '<div class="pi-card"><b>已选：' + esc(p.name) + '</b><br>' + lines.slice(1).map(function (l) { return esc(l); }).join('<br>') + '</div>';
        });
      }
      renderPartyInfo();
      var shSel = document.getElementById('wz-shipper'), coSel = document.getElementById('wz-consignee');
      if (shSel) shSel.onchange = renderPartyInfo;
      if (coSel) coSel.onchange = renderPartyInfo;
      document.getElementById('wz-back2').onclick = function () { w.step = 1; render(); };
      document.getElementById('wz-next2').onclick = function () {
        w.shipperId = val('wz-shipper'); w.consigneeId = val('wz-consignee'); w.notifyId = val('wz-notify');
        w.meta = {
          invoiceNo: val('wz-invno'), invoiceDate: val('wz-invdate'), contractNo: val('wz-contract'),
          incoterms: val('wz-inco'), paymentTerms: val('wz-pay'), transport: val('wz-trans'),
          pol: val('wz-pol'), pod: val('wz-pod'), currency: val('wz-cur'),
          shippingMarks: val('wz-marks'), remark: val('wz-remark')
        };
        w.step = 3; render();
      };
      return;
    }

    // ---------- step3 校验 ----------
    if (w.step === 3) {
      var ctx = await buildWizContext(w, 'invoice');
      var tpl = await db.get('templates', w.templateId);
      var itemFields = (tpl && tpl.mapping && tpl.mapping.scanned && tpl.mapping.scanned.itemFields) || [];
      var report = validator.validateDocument({
        kind: 'invoice', orders: ctx.orders, packing: ctx.packing,
        data: ctx.data, requiredFields: (tpl.mapping && tpl.mapping.required) || engine.REQUIRED_FIELDS.invoice,
        declareMap: ctx.declareMap, templateItemFields: itemFields
      });
      w.report = report;
      var html = '<div class="card"><h3>校验报告</h3>';
      if (report.ok) html += '<div class="vres pass">✅ 全部校验通过：单号匹配、SKU数量勾稽一致、申报要素齐全、必填字段完整</div>';
      report.blocks.forEach(function (b) {
        html += '<div class="vres block">❌ ' + esc(b.msg);
        if (b.diffs) {
          html += '<table class="grid" style="margin-top:8px"><tr><th>SKU</th><th>订单数量</th><th>箱单数量</th><th>差值</th></tr>' +
            b.diffs.map(function (d) { return '<tr><td class="mono">' + esc(d.sku) + '</td><td class="num">' + d.orderQty + '</td><td class="num">' + d.packingQty + '</td><td class="num" style="color:#b91c1c;font-weight:700">' + (d.diff > 0 ? '+' : '') + d.diff + '</td></tr>'; }).join('') + '</table>';
        }
        if (b.missing && b.type === 'declare') {
          var missingSkus = b.missing.filter(function (x) { return x.lacks.indexOf('申报要素缺失') >= 0; }).map(function (x) { return x.sku; });
          var incompleteSkus = b.missing.filter(function (x) { return x.lacks.indexOf('申报要素缺失') < 0; });
          var fieldHint = (b.needFields && b.needFields.length)
            ? '本模板用到的申报字段：' + b.needFields.map(function (f) { return { hsCode: 'HS编码', declarePrice: '申报价', material: '材质', nameEn: '英文品名', nameCn: '中文品名', brand: '品牌', usage: '用途', model: '型号', nw: '净重', gw: '毛重', origin: '原产地' }[f] || f; }).join('、')
            : '默认检查：HS编码、申报价、材质';
          html += '<p class="hint">' + fieldHint + '</p>' +
            '<table class="grid" style="margin-top:8px"><tr><th>SKU</th><th>缺失项</th><th>说明</th></tr>' +
            b.missing.map(function (x) {
              var note = x.lacks.indexOf('申报要素缺失') >= 0 ? '申报信息表中没有该 SKU' : '该 SKU 已存在但字段不全';
              return '<tr><td class="mono">' + esc(x.sku) + '</td><td>' + esc(x.lacks.join('、')) + '</td><td>' + note + '</td></tr>';
            }).join('') + '</table>' +
            '<p class="hint">共 ' + b.missing.length + ' 个 SKU 需要补录。' +
            (missingSkus.length ? '其中 ' + missingSkus.length + ' 个 SKU 在「申报信息」中不存在；' : '') +
            (incompleteSkus.length ? incompleteSkus.length + ' 个 SKU 已存在但缺字段。' : '') +
            '请到「申报信息」页补录后回来点「重新校验」。</p>' +
            '<button class="btn ghost" id="wz-to-declare">去「申报信息」页补录</button>';
        }
        html += '</div>';
      });
      html += '<div style="margin-top:14px;display:flex;gap:8px"><button class="btn ghost" id="wz-back3">← 上一步</button>' +
        '<button class="btn ghost" id="wz-recheck">🔄 重新校验</button>' +
        '<button class="btn ok" id="wz-next3"' + (report.ok ? '' : ' disabled') + '>校验通过，下一步：预览 →</button></div></div>';
      body.innerHTML = html;
      document.getElementById('wz-back3').onclick = function () { w.step = 2; render(); };
      document.getElementById('wz-recheck').onclick = function () { render(); };
      var toDeclareBtn = document.getElementById('wz-to-declare');
      if (toDeclareBtn) {
        toDeclareBtn.onclick = function () {
          state.wiz = null; state.tab = 'declares';
          document.querySelectorAll('.tab').forEach(function (t) { t.classList.toggle('active', t.dataset.tab === 'declares'); });
          render();
        };
      }
      document.getElementById('wz-next3').onclick = function () { if (w.report.ok) { w.step = 4; render(); } };
      return;
    }

    // ---------- step4 预览确认导出 ----------
    if (w.step === 4) {
      var tplName = w.templateId || '测试模板';
      if (w._wb && w._audit && w._data) {
        // e2e 注入态：复用注入的成品与审计，跳过真实填充（正常流程不会预置这三个字段，行为完全不变）
        var fillRes = { unresolved: [] };
        var wb = w._wb;
      } else {
      var ctx = await buildWizContext(w, 'invoice');
      var tpl = await db.get('templates', w.templateId);
      // 防御：模板不存在/损坏 → 自动回退到 step=1 重新选模板（不再黑屏）
      if (!tpl || !tpl.fileBuf) {
        toast('模板不可用（id=' + esc(w.templateId || '') + '），请重新选择', 'err');
        w.step = 1; w.templateId = ''; w.doc = null; render();
        return;
      }
      // 早期诊断 fileBuf 格式（避免直接抛 generic 错误）
      // 兼容：ArrayBuffer / Uint8Array / {data:[]} / {type:'Buffer', data:[]} / 其他损坏形式
      var _info = _inspectFileBuf(tpl.fileBuf);
      var _u8 = _info.bytes;
      var _fmt = (_u8.length >= 4 && _u8[0] === 0x50 && _u8[1] === 0x4B) ? 'xlsx'
                : (_u8.length >= 8 && _u8[0] === 0xD0 && _u8[1] === 0xCF && _u8[2] === 0x11 && _u8[3] === 0xE0) ? 'xls-old'
                : (_u8.length && _u8[0] >= 0x20 && _u8[0] <= 0x7E) ? 'csv-or-text' : 'unknown';
      if (_fmt === 'unknown' || _fmt === 'csv-or-text') {
        body.innerHTML = '<div class="card vres block" style="padding:20px">' +
          '<h3 style="margin-top:0">⚠️ 模板文件已损坏</h3>' +
          '<p>当前选中模板 <b>' + esc(tpl.name) + '</b> 的 fileBuf 已不是有效 xlsx（探测为 <code>' + _fmt + '</code>，形态 <code>' + esc(_info.shape) + '</code>，原始大小 ' + _info.size + ' 字节）。</p>' +
          '<p>可能原因：浏览器本地存储被异常清空、IndexedDB schema 不兼容、模板上传时网络中断等。</p>' +
          '<div style="margin-top:14px;display:flex;gap:8px;flex-wrap:wrap">' +
          '<button class="btn danger" id="wz-del-bad-tpl">🗑 删除该损坏模板并回 step 1</button>' +
          '<button class="btn ghost" id="wz-back4b">← 回到 step 1 重选模板</button>' +
          '<button class="btn warn" id="wz-back3b">← 回到 step 3 校验核对</button>' +
          '</div></div>';
        document.getElementById('wz-del-bad-tpl').onclick = async function () {
          if (!(await confirmBox('确认删除模板「' + esc(tpl.name) + '」？将回到 step 1 重选模板'))) return;
          await db.del('templates', tpl.id);
          toast('已删除坏模板「' + tpl.name + '」', 'ok');
          w.step = 1; w.templateId = ''; w.doc = null; render();
        };
        document.getElementById('wz-back4b').onclick = function () { w.step = 1; w.templateId = ''; w.doc = null; render(); };
        document.getElementById('wz-back3b').onclick = function () { w.step = 3; w.doc = null; render(); };
        return;
      }
      var wb = await loadWb(tpl.fileBuf, tpl.fileName);
      var fillRes = engine.fillTemplate(wb, ctx.data, { logo: tpl.logo || null, labelMap: ((tpl.mapping && tpl.mapping.labelMap) || []) });
      // v1.5.0 输出成品审计：以模板原件为基准做填充后逐格 diff，拦截未授权写入/保护区被改/箱号缺失等
      var tplWb = await loadWb(tpl.fileBuf, tpl.fileName);
      var audit = (window.TD && window.TD.audit) ? window.TD.audit.auditFilled(wb, tplWb, ctx.data, { kind: 'invoice', labelMap: ((tpl.mapping && tpl.mapping.labelMap) || []) }) : { ok: true, blocks: [], warns: [], info: [] };
      w._audit = audit; w._wb = wb; w._data = ctx.data;
      tplName = tpl.name;
      }
      // v1.5.1 产品图嵌入：必须在审计之后(审计比对fillTemplate后无图wb)，且 wbToHtml 之前(预览含图)
      await engine.embedProductImages(wb, w._data, {});
      // v1.5.7 导出单次嵌图上限 1000 的可见提示（防浏览器卡死，超量切片跳过）
      // v1.5.8 SKU 别名归一化后仍无图的 SKU 可见提示（便于排查"导出没图"）
      // v1.5.12 嵌图诊断条：直接显示引擎本次锁定的"产品图片列号 + 行号 + 命中来源（itemHeaderMap/精确扫描/宽泛扫描）"，便于排查"图飞到最右空白列"
      var embWarn = '';
      var ed = (typeof window !== 'undefined') ? window.__embDiag : null;
      if (ed) {
        var _srcLabel = (ed.sourceMap === 'itemHeaderMap') ? '复用 fillTemplate 字段映射'
                       : (ed.sourceMap === 'scan-exact') ? '精确表头扫描'
                       : (ed.sourceMap === 'scan-loose') ? '宽泛扫描'
                       : '未命中';
        embWarn += '<div class="vres info" style="font-size:12px">🖼 嵌图诊断：' +
          '锁定列号 C' + (ed.imgCol > 0 ? ed.imgCol : '?') +
          '、起始行 R' + (ed && wb._fillItemsRowNum ? wb._fillItemsRowNum : '?') +
          '、列号来源=' + _srcLabel +
          '、items=' + ed.itemsLen + '、待嵌图=' + ed.tasksLen + (ed.overflow ? '（溢出 '+ed.overflow+'）' : '') +
          '。若列号 ≠ 模板表头里的「产品图片」所在列，请把模板"明细表头行"的列名规范为「产品图片」或「商品图片」后再导出。</div>';
      }
      if (ed && ed.overflow > 0) {
        embWarn += '<div class="vres warn">⚠️ 本次导出 SKU 产品图共 ' + (ed.tasksLen + ed.overflow) + ' 张，已超出单次嵌图上限 1000，仅嵌入前 1000 张，其余 ' + ed.overflow + ' 张未嵌入（如需全部出图，请减少单次导出的 SKU 行数或分批导出）。</div>';
      }
      if (ed && ed.missed && ed.missed.length) {
        var mlist = ed.missed.slice(0, 30).map(esc).join('　') + (ed.missed.length > 30 ? (' …等 ' + ed.missed.length + ' 个') : '');
        embWarn += '<div class="vres warn">⚠️ 以下 ' + ed.missed.length + ' 个 SKU 在图库中无对应产品图（已跳过，不影响导出）：<span class="mono">' + mlist + '</span></div>';
      }
      // v1.5.50 取图失败清单：网络/镜像不可达导致 embed 拿不到真图，明确列出（不再静默空白）
      if (ed && ed.failed && ed.failed.length) {
        var flist = ed.failed.slice(0, 40).map(esc).join('　') + (ed.failed.length > 40 ? (' …等 ' + ed.failed.length + ' 个') : '');
        embWarn += '<div class="vres warn">⚠️ 以下 ' + ed.failed.length + ' 个 SKU 取图失败（镜像不可达/超时），导出后对应单元格为空白：<span class="mono">' + flist + '</span>。若整批都失败，请检查网络或换用 CloudStudio 国内镜像打开。</div>';
      }
      var confirmed = w.doc && w.doc.status === 'confirmed';
      body.innerHTML = '<div class="card"><h3>发票预览（模板: ' + esc(tplName) + '）</h3>' +
        (fillRes.unresolved.length ? '<div class="vres warn">⚠️ 以下占位符无数据（已置空）: <span class="mono">' + fillRes.unresolved.filter(function (v, i, a) { return a.indexOf(v) === i; }).map(esc).join('　') + '</span></div>' : '') +
        (window.TD && window.TD.audit ? window.TD.audit.auditHtml(w._audit) : '') +
        embWarn +
        '<div style="overflow:auto;border:1px solid #e3e8f0;border-radius:8px;padding:8px;background:#fafbfd">' + wbToHtml(wb) + '</div>' +
        '<div style="margin-top:14px;display:flex;gap:8px;align-items:center">' +
        '<button class="btn ghost" id="wz-back4">← 上一步</button>' +
        '<button class="btn warn" id="wz-confirm"' + (confirmed ? ' disabled' : '') + '>' + (confirmed ? '✓ 已确认' : '① 确认无误') + '</button>' +
        '<button class="btn ok" id="wz-export"' + (confirmed ? '' : ' disabled') + '>② 导出发票 xlsx</button>' +
        '<span class="hint">必须先「确认无误」才能导出（状态机: draft→validated→confirmed→exported）</span></div></div>';
      document.getElementById('wz-back4').onclick = function () { w.step = 3; w.doc = null; render(); };
      document.getElementById('wz-confirm').onclick = async function () {
        try {
        // v1.5.38 修复：二次进入 step4（上一步再下一步/换模板重生成）走 w._wb 缓存分支时 ctx 未定义 → ctx.data 抛 TypeError 致按钮静默失效。
        // 统一用 w._data（两种路径下都有值：首次=ctx.data 赋值，缓存=注入复用），彻底消除该缺陷。
        var doc = {
          kind: 'invoice', orderIds: w.orderIds.slice(), packingId: w.packingId, templateId: w.templateId,
          carrier: w.carrier, channel: w.channel, docNo: (w._data ? w._data.invoiceNo : ''),
          data: (w._data || ctx.data || {}), status: 'draft'
        };
        if (!validator.canTransition('draft', 'validated')) { toast('状态机异常', 'err'); return; }
        doc.status = 'validated';
        if (!validator.canTransition('validated', 'confirmed')) { toast('状态机异常', 'err'); return; }
        doc.status = 'confirmed';
        w.doc = await db.put('documents', doc);
        toast('已确认，可导出', 'ok'); renderWizStep();
        } catch (e) { console.error('[wz-confirm]', e); toast('确认失败: ' + (e && e.message || e), 'err'); }
      };
      document.getElementById('wz-export').onclick = async function () {
        if (!w.doc || !validator.canTransition(w.doc.status, 'exported')) { toast('请先确认无误', 'err'); return; }
        // v1.5.0 拦截门：审计有阻断项则禁止导出
        if (w._audit && w._audit.blocks && w._audit.blocks.length) {
          toast('导出前质检未通过，请查看预览区红框问题', 'err');
          showModal('<h3>❌ 导出前质检未通过（' + w._audit.blocks.length + ' 项）</h3>' + (window.TD.audit ? window.TD.audit.auditHtml(w._audit) : '') + '<div style="text-align:right;margin-top:12px"><button class="btn" onclick="TDUI.closeModal()">关闭</button></div>');
          return;
        }
        var fname = exporter.safeName((w.carrier ? w.carrier + '_' : '') + (w._data ? w._data.invoiceNo : '')) + '.xlsx';
        await exporter.download(w._wb, fname);
        w.doc.status = 'exported'; w.doc.exportedFile = fname;
        await db.put('documents', w.doc);
        toast('发票已导出: ' + fname, 'ok');
        state.wiz = null; state.tab = 'documents';
        document.querySelectorAll('.tab').forEach(function (t) { t.classList.toggle('active', t.dataset.tab === 'documents'); });
        render();
      };
      return;
    }
  }

  /** 组装向导上下文（发票/订舱共用） */
  async function buildWizContext(w, kind) {
    var orders = [];
    for (var i = 0; i < w.orderIds.length; i++) { var o = await db.get('orders', w.orderIds[i]); if (o) orders.push(o); }
    var packing = w.packingId ? await db.get('packings', w.packingId) : null;
    var declares = await db.all('declare_reqs');
    var declareMap = {}; declares.forEach(function (d) { declareMap[d.sku] = d; });
    var shipper = w.shipperId ? await db.get('parties', w.shipperId) : {};
    var consignee = w.consigneeId ? await db.get('parties', w.consigneeId) : {};
    var notify = w.notifyId ? await db.get('parties', w.notifyId) : null;
    var tpl = w.templateId ? await db.get('templates', w.templateId) : null;
    var itemFields = (tpl && tpl.mapping && tpl.mapping.scanned && tpl.mapping.scanned.itemFields) || [];
    // 兜底：旧数据或内置真实模板未缓存 scanned 时，现场扫描模板并写回数据库
    if (tpl && (!tpl.mapping || !tpl.mapping.scanned || !tpl.mapping.scanned.itemFields.length)) {
      try {
        var _wb = await loadWb(tpl.fileBuf);
        var _scan = engine.scanTemplate(_wb);
        itemFields = _scan.itemFields || [];
        tpl.mapping = tpl.mapping || {};
        tpl.mapping.scanned = _scan;
        await db.put('templates', tpl);
      } catch (e) {}
    }
    // v1.4.47：boxMode 触发条件同时看 ① 模板占位符 itemFields ② 现场扫描的 itemHeaderMap 是否映射长/宽/高/单箱重
    //          纯标签驱动的海运类模板（明细列是「长(cm)/宽(cm)/单箱重量」文字而非 {{items.lengthCm}}）也能触发 boxMode
    // v1.4.50 修：① 主动现场 scan 一次（不依赖 tpl 缓存，因为用户早期上传的模板 scan 当时是旧版本，没识别 lengthCm/widthCm/heightCm，tpl 缓存里没这些字段）② 兜底 if 条件写反问题（之前是 headerHasDims=true 才进重扫，但旧 tpl 缓存里就是空的 headerHasDims=false → 永远不重扫）
    var itemHeaderMap = (tpl && tpl.mapping && tpl.mapping.scanned && tpl.mapping.scanned.itemHeaderMap) || {};
    var headerHasBoxCols = Object.keys(itemHeaderMap).some(function (c) {
      var f = itemHeaderMap[c];
      return /^(boxNo|lengthCm|widthCm|heightCm|singleGw)$/.test(f);
    });
    // v1.4.50 主动现场 scan 一次（无论 tpl 缓存是否有 dims），避免早期上传模板因旧 scan 引擎没识别长宽高而永远不触发 boxMode
    if (tpl && !itemHeaderMap.__scanned && w._wb) {
      try {
        var _liveScan = engine.scanTemplate(w._wb);
        if (_liveScan && _liveScan.itemHeaderMap) {
          itemHeaderMap = _liveScan.itemHeaderMap;
          itemHeaderMap.__scanned = 1;
        }
      } catch (e) {}
      headerHasBoxCols = Object.keys(itemHeaderMap).some(function (c) {
        var f = itemHeaderMap[c];
        return /^(boxNo|lengthCm|widthCm|heightCm|singleGw)$/.test(f);
      });
    }
    // v1.4.62：只要关联了「含箱号的装箱清单」且模板明细含箱号类列（货箱编号/箱号/CARTON NO），就强制箱号模式，确保箱号一定显示
    var packingHasBoxes = !!(packing && (packing.boxes || []).length);
    var boxMode = !!(packing && (itemFields.some(function (f) { return /(boxNo|length|width|height)/.test(f); }) || (headerHasBoxCols && packingHasBoxes)));
    var data = engine.buildDocData({
      kind: kind, orders: orders, packing: packing, meta: w.meta, boxMode: boxMode,
      shipper: shipper || {}, consignee: consignee || {}, notify: notify, declareMap: declareMap
    });
    return { orders: orders, packing: packing, declareMap: declareMap, data: data };
  }

  // ================= 生成订舱单向导 =================
  PAGES.booking = async function () {
    if (!state.bwiz) state.bwiz = { step: 0, orderIds: [], packingId: '', templateId: '', carrier: '', meta: {}, shipperId: '', consigneeId: '', notifyId: '', report: null, doc: null };
    var names = ['勾选订单', '选模板', '订舱信息', '校验', '预览·确认·导出'];
    return '<h2>生成订舱单 BOOKING FORM</h2>' + wizSteps(state.bwiz.step, names) + '<div id="bwiz-body"></div>';
  };
  BINDERS.booking = function () { renderBookingStep(); };

  async function renderBookingStep() {
    var w = state.bwiz;
    var body = document.getElementById('bwiz-body');
    if (!body) return;
    try {
      await renderBookingStepInner();
    } catch (e) {
      console.error(e);
      body.innerHTML = '<div class="card vres block" style="padding:20px">' +
        '<h3 style="margin-top:0">⚠️ 订舱单步骤 ' + (w ? w.step : '?') + ' 渲染失败</h3>' +
        '<p style="color:#c0392b;margin:8px 0"><b>' + esc(e.message || String(e)) + '</b></p>' +
        '<pre style="background:#fafbfd;padding:8px;border-radius:6px;font-size:12px;max-height:200px;overflow:auto;white-space:pre-wrap">' + esc((e.stack || '').split('\n').slice(0, 4).join('\n')) + '</pre>' +
        '<p class="hint" style="margin-top:10px">💡 截图发给开发者可快速定位。可尝试 <b>Ctrl+Shift+R 硬刷新</b> 或在「设置 → ⚠️ 危险操作」清空数据库后重试。</p>' +
        (w ? '<div style="margin-top:10px"><button class="btn ghost" id="err-bwiz-home">← 回到订舱向导 step 1</button></div>' : '') +
        '</div>';
      var back = document.getElementById('err-bwiz-home');
      if (back && w) back.onclick = function () { w.step = 0; w.orderIds = []; w.packingId = ''; w.templateId = ''; w.meta = {}; w.report = null; w.doc = null; render(); };
    }
  }

  async function renderBookingStepInner() {
    var w = state.bwiz;
    var body = document.getElementById('bwiz-body');
    if (!body) return;

    if (w.step === 0) {
      var orders = await db.all('orders');
      orders.sort(function (a, b) { return b.createdAt - a.createdAt; });
      var rows = orders.map(function (o) {
        var qty = (o.items || []).reduce(function (s, i) { return s + i.qty; }, 0);
        var ck = w.orderIds.indexOf(o.id) >= 0 ? ' checked' : '';
        return '<tr class="checkrow"><td><input type="checkbox" class="bw-ord" value="' + o.id + '"' + ck + '></td>' +
          '<td class="mono">' + esc(o.orderNo) + '</td><td>' + (o.source === 'from_packing' ? '<span class="badge purple">箱单直生</span>' : '<span class="badge blue">聚水潭</span>') + '</td>' +
          '<td class="num">' + qty + '</td></tr>';
      }).join('');
      body.innerHTML = '<div class="card"><p class="hint">勾选本地保存的订单（含装箱清单直生订单）用于订舱</p>' +
        '<table class="grid"><tr><th></th><th>订单号</th><th>来源</th><th>数量</th></tr>' +
        (rows || '<tr><td colspan="4" class="empty">暂无订单</td></tr>') + '</table>' +
        '<div style="margin-top:14px"><button class="btn" id="bw-next0">下一步 →</button></div></div>';
      document.getElementById('bw-next0').onclick = function () {
        var ids = Array.from(document.querySelectorAll('.bw-ord:checked')).map(function (c) { return c.value; });
        if (!ids.length) { toast('请至少勾选一个订单', 'err'); return; }
        w.orderIds = ids; w.step = 1; render();
      };
      return;
    }

    if (w.step === 1) {
      var _allBkTpls = (await db.all('templates')).filter(function (t) { return t.kind === 'booking' && t.status === 'active'; });
      var tpls = _allBkTpls.filter(function (t) { return t.fileBuf && t.fileBuf.byteLength > 0; });
      var pks = await db.all('packings');
      var orders = []; for (var i = 0; i < w.orderIds.length; i++) orders.push(await db.get('orders', w.orderIds[i]));
      var selNos = orders.map(function (o) { return o.orderNo; });
      var tplOpts = tpls.map(function (t) { return '<option value="' + t.id + '"' + (w.templateId === t.id ? ' selected' : '') + '>' + esc(t.name) + '（' + esc(t.carrier) + '）</option>'; }).join('');
      var pkOpts = '<option value="">不关联（重量体积需手动确认）</option>' + pks.map(function (p) {
        var inter = (p.orderNos || []).filter(function (n) { return selNos.indexOf(n) >= 0; });
        return '<option value="' + p.id + '"' + (w.packingId === p.id ? ' selected' : '') + '>' + esc(p.fileName) + (inter.length ? '（匹配' + inter.length + '单）' : '') + '</option>';
      }).join('');
      body.innerHTML = '<div class="card"><div class="form-grid">' +
        '<div><label>物流商</label><input id="bw-carrier" list="carrier-list3" value="' + esc(w.carrier) + '"><datalist id="carrier-list3">' + CARRIERS.map(function (c) { return '<option>' + c + '</option>'; }).join('') + '</datalist></div>' +
        '<div><label class="req">订舱单模板（仅启用中）</label><select id="bw-tpl"><option value="">请选择</option>' + tplOpts + '</select></div>' +
        '<div><label>关联装箱清单（自动带出件毛体）</label><select id="bw-pk">' + pkOpts + '</select></div></div>' +
        '<div style="margin-top:14px;display:flex;gap:8px"><button class="btn ghost" id="bw-back1">← 上一步</button><button class="btn" id="bw-next1">下一步 →</button></div></div>';
      document.getElementById('bw-back1').onclick = function () { w.step = 0; render(); };
      document.getElementById('bw-next1').onclick = async function () {
        var tpl = val('bw-tpl');
        if (!tpl) { toast('请选择订舱单模板', 'err'); return; }
        // v1.5.45：选中即校验 fileBuf，残缺模板直接拦截
        var tplRec = await db.get('templates', tpl);
        if (!tplRec || !tplRec.fileBuf || !tplRec.fileBuf.byteLength) {
          toast('该订舱模板文件未同步完整，请到「设置-数据管理-重新同步」后重试', 'err');
          return;
        }
        w.templateId = tpl; w.carrier = val('bw-carrier'); w.packingId = val('bw-pk');
        w.step = 2; render();
      };
      return;
    }

    if (w.step === 2) {
      var parties = await db.all('parties');
      var docs = await db.all('documents');
      var m = w.meta;
      if (!m.invoiceNo) m.invoiceNo = 'BK' + today().replace(/-/g, '') + '-' + String(docs.filter(function (d) { return d.kind === 'booking'; }).length + 1).padStart(3, '0');
      if (!m.invoiceDate) m.invoiceDate = today();
      function opts(type, sel) {
        return '<option value="">请选择</option>' + parties.filter(function (p) { return p.type === type; }).map(function (p) { return '<option value="' + p.id + '"' + (sel === p.id ? ' selected' : '') + '>' + esc(p.name) + '</option>'; }).join('');
      }
      body.innerHTML = '<div class="card"><div class="form-grid">' +
        '<div><label class="req">委托号</label><input id="bw-no" value="' + esc(m.invoiceNo) + '"></div>' +
        '<div><label class="req">日期</label><input type="date" id="bw-date" value="' + esc(m.invoiceDate) + '"></div>' +
        '<div><label class="req">SHIPPER 托运人</label><select id="bw-shipper">' + opts('shipper', w.shipperId) + '</select></div>' +
        '<div><label class="req">CONSIGNEE 收货人</label><select id="bw-consignee">' + opts('consignee', w.consigneeId) + '</select></div>' +
        '<div><label>NOTIFY 通知人</label><select id="bw-notify">' + opts('notify', w.notifyId) + '</select></div>' +
        '<div><label class="req">起运港 POL</label><input id="bw-pol" value="' + esc(m.pol || 'SHENZHEN, CHINA') + '"></div>' +
        '<div><label class="req">目的港 POD</label><input id="bw-pod" value="' + esc(m.pod || '') + '"></div>' +
        '<div><label>ETD 船期</label><input type="date" id="bw-etd" value="' + esc(m.etd || '') + '"></div>' +
        '<div><label>船名航次</label><input id="bw-vessel" value="' + esc(m.vessel || '') + '"></div>' +
        '<div><label>柜型</label><select id="bw-ctype"><option value="">LCL散货</option>' + ['20GP', '40GP', '40HQ', '45HQ'].map(function (c) { return '<option' + (m.containerType === c ? ' selected' : '') + '>' + c + '</option>'; }).join('') + '</select></div>' +
        '<div><label>柜量</label><input id="bw-cqty" type="number" value="' + esc(m.containerQty || '1') + '"></div>' +
        '<div><label>贸易条款</label><select id="bw-inco">' + INCOTERMS.map(function (t) { return '<option' + (m.incoterms === t ? ' selected' : '') + '>' + t + '</option>'; }).join('') + '</select></div>' +
        '<div><label class="req">运费条款</label><select id="bw-freight"><option' + (m.freightTerms === 'FREIGHT PREPAID' ? ' selected' : '') + '>FREIGHT PREPAID</option><option' + (m.freightTerms === 'FREIGHT COLLECT' ? ' selected' : '') + '>FREIGHT COLLECT</option></select></div>' +
        '<div><label>订舱代理</label><input id="bw-agent" value="' + esc(m.agent || '') + '"></div>' +
        '<div><label>品名概述(英文)</label><input id="bw-goods" value="' + esc(m.goodsSummary || '') + '" placeholder="留空则自动取明细品名"></div>' +
        '<div><label>唛头</label><input id="bw-marks" value="' + esc(m.shippingMarks || 'N/M') + '"></div>' +
        '<div><label>危险品声明</label><input id="bw-dg" value="' + esc(m.dangerous || 'NON-DANGEROUS / GENERAL CARGO') + '"></div>' +
        '<div><label>报关方式</label><input id="bw-customs" value="' + esc(m.customsType || '买单报关') + '"></div>' +
        '<div><label>备注</label><input id="bw-remark" value="' + esc(m.remark || '') + '"></div></div>' +
        '<div style="margin-top:14px;display:flex;gap:8px"><button class="btn ghost" id="bw-back2">← 上一步</button><button class="btn" id="bw-next2">下一步：校验 →</button></div></div>';
      document.getElementById('bw-back2').onclick = function () { w.step = 1; render(); };
      document.getElementById('bw-next2').onclick = function () {
        w.shipperId = val('bw-shipper'); w.consigneeId = val('bw-consignee'); w.notifyId = val('bw-notify');
        w.meta = {
          invoiceNo: val('bw-no'), invoiceDate: val('bw-date'), pol: val('bw-pol'), pod: val('bw-pod'),
          etd: val('bw-etd'), vessel: val('bw-vessel'), containerType: val('bw-ctype'), containerQty: val('bw-cqty'),
          incoterms: val('bw-inco'), freightTerms: val('bw-freight'), agent: val('bw-agent'),
          goodsSummary: val('bw-goods'), shippingMarks: val('bw-marks'), dangerous: val('bw-dg'),
          customsType: val('bw-customs'), remark: val('bw-remark')
        };
        w.step = 3; render();
      };
      return;
    }

    if (w.step === 3) {
      var ctx = await buildWizContext(w, 'booking');
      var tpl = await db.get('templates', w.templateId);
      var itemFields = (tpl && tpl.mapping && tpl.mapping.scanned && tpl.mapping.scanned.itemFields) || [];
      var report = validator.validateDocument({
        kind: 'booking', orders: ctx.orders, packing: ctx.packing,
        data: ctx.data, requiredFields: (tpl.mapping && tpl.mapping.required) || engine.REQUIRED_FIELDS.booking,
        declareMap: null, skipPacking: !w.packingId, templateItemFields: itemFields
      });
      w.report = report;
      var html = '<div class="card"><h3>校验报告</h3>';
      if (report.ok) html += '<div class="vres pass">✅ 校验通过' + (w.packingId ? '（含装箱清单勾稽）' : '（未关联装箱清单，请人工核对件毛体）') + '</div>';
      report.blocks.forEach(function (b) {
        html += '<div class="vres block">❌ ' + esc(b.msg);
        if (b.diffs) html += '<table class="grid" style="margin-top:8px"><tr><th>SKU</th><th>订单数量</th><th>箱单数量</th><th>差值</th></tr>' + b.diffs.map(function (d) { return '<tr><td class="mono">' + esc(d.sku) + '</td><td class="num">' + d.orderQty + '</td><td class="num">' + d.packingQty + '</td><td class="num">' + d.diff + '</td></tr>'; }).join('') + '</table>';
        html += '</div>';
      });
      html += '<div style="margin-top:14px;display:flex;gap:8px"><button class="btn ghost" id="bw-back3">← 上一步</button>' +
        '<button class="btn ok" id="bw-next3"' + (report.ok ? '' : ' disabled') + '>下一步：预览 →</button></div></div>';
      body.innerHTML = html;
      document.getElementById('bw-back3').onclick = function () { w.step = 2; render(); };
      document.getElementById('bw-next3').onclick = function () { if (w.report.ok) { w.step = 4; render(); } };
      return;
    }

    if (w.step === 4) {
      var tplName = w.templateId || '测试模板';
      if (w._wb && w._audit && w._data) {
        // e2e 注入态：复用注入的成品与审计，跳过真实填充（正常流程不会预置，行为不变）
        var fillRes = { unresolved: [] };
        var wb = w._wb;
      } else {
      var ctx = await buildWizContext(w, 'booking');
      var tpl = await db.get('templates', w.templateId);
      // 防御：模板不可用 → 自动回退到 step=1 重新选择
      if (!tpl || !tpl.fileBuf) {
        toast('模板不可用（id=' + esc(w.templateId || '') + '），请重新选择', 'err');
        w.step = 1; w.templateId = ''; w.doc = null; render();
        return;
      }
      // 早期诊断 fileBuf 格式（同 invoice step=4）
      var _infob = _inspectFileBuf(tpl.fileBuf);
      var _u8b = _infob.bytes;
      var _fmtb = (_u8b.length >= 4 && _u8b[0] === 0x50 && _u8b[1] === 0x4B) ? 'xlsx'
                : (_u8b.length >= 8 && _u8b[0] === 0xD0 && _u8b[1] === 0xCF && _u8b[2] === 0x11 && _u8b[3] === 0xE0) ? 'xls-old'
                : (_u8b.length && _u8b[0] >= 0x20 && _u8b[0] <= 0x7E) ? 'csv-or-text' : 'unknown';
      if (_fmtb === 'unknown' || _fmtb === 'csv-or-text') {
        body.innerHTML = '<div class="card vres block" style="padding:20px">' +
          '<h3 style="margin-top:0">⚠️ 模板文件已损坏</h3>' +
          '<p>当前选中订舱模板 <b>' + esc(tpl.name) + '</b> 的 fileBuf 已被破坏（探测为 <code>' + _fmtb + '</code>，形态 <code>' + esc(_infob.shape) + '</code>，大小 ' + _infob.size + ' 字节）。</p>' +
          '<div style="margin-top:14px;display:flex;gap:8px;flex-wrap:wrap">' +
          '<button class="btn danger" id="bw-del-bad-tpl">🗑 删除该损坏模板并回 step 1</button>' +
          '<button class="btn ghost" id="bw-back4b">← 回到 step 1</button>' +
          '</div></div>';
        document.getElementById('bw-del-bad-tpl').onclick = async function () {
          if (!(await confirmBox('确认删除订舱模板「' + esc(tpl.name) + '」？'))) return;
          await db.del('templates', tpl.id);
          toast('已删除坏模板', 'ok');
          w.step = 1; w.templateId = ''; w.doc = null; render();
        };
        document.getElementById('bw-back4b').onclick = function () { w.step = 1; w.templateId = ''; w.doc = null; render(); };
        return;
      }
      var wb = await loadWb(tpl.fileBuf, tpl.fileName);
      var fillRes = engine.fillTemplate(wb, ctx.data, { logo: tpl.logo || null, labelMap: ((tpl.mapping && tpl.mapping.labelMap) || []) });
      // v1.5.0 输出成品审计
      var tplWb = await loadWb(tpl.fileBuf, tpl.fileName);
      var audit = (window.TD && window.TD.audit) ? window.TD.audit.auditFilled(wb, tplWb, ctx.data, { kind: 'booking', labelMap: ((tpl.mapping && tpl.mapping.labelMap) || []) }) : { ok: true, blocks: [], warns: [], info: [] };
      w._audit = audit; w._wb = wb; w._data = ctx.data;
      tplName = tpl.name;
      }
      var confirmed = w.doc && w.doc.status === 'confirmed';
      body.innerHTML = '<div class="card"><h3>订舱单预览（模板: ' + esc(tplName) + '）</h3>' +
        (fillRes.unresolved.length ? '<div class="vres warn">⚠️ 空占位符: <span class="mono">' + fillRes.unresolved.filter(function (v, i, a) { return a.indexOf(v) === i; }).map(esc).join('　') + '</span></div>' : '') +
        (window.TD && window.TD.audit ? window.TD.audit.auditHtml(w._audit) : '') +
        '<div style="overflow:auto;border:1px solid #e3e8f0;border-radius:8px;padding:8px;background:#fafbfd">' + wbToHtml(wb) + '</div>' +
        '<div style="margin-top:14px;display:flex;gap:8px;align-items:center">' +
        '<button class="btn ghost" id="bw-back4">← 上一步</button>' +
        '<button class="btn warn" id="bw-confirm"' + (confirmed ? ' disabled' : '') + '>' + (confirmed ? '✓ 已确认' : '① 确认无误') + '</button>' +
        '<button class="btn ok" id="bw-export"' + (confirmed ? '' : ' disabled') + '>② 导出订舱单 xlsx</button></div></div>';
      document.getElementById('bw-back4').onclick = function () { w.step = 3; w.doc = null; render(); };
      document.getElementById('bw-confirm').onclick = async function () {
        var doc = {
          kind: 'booking', orderIds: w.orderIds.slice(), packingId: w.packingId, templateId: w.templateId,
          carrier: w.carrier, docNo: (w._data ? w._data.invoiceNo : ''), data: (w._data || ctx.data), status: 'confirmed'
        };
        w.doc = await db.put('documents', doc);
        toast('已确认，可导出', 'ok'); renderBookingStep();
      };
      document.getElementById('bw-export').onclick = async function () {
        if (!w.doc || !validator.canTransition(w.doc.status, 'exported')) { toast('请先确认无误', 'err'); return; }
        // v1.5.0 拦截门
        if (w._audit && w._audit.blocks && w._audit.blocks.length) {
          toast('导出前质检未通过，请查看预览区红框问题', 'err');
          showModal('<h3>❌ 导出前质检未通过（' + w._audit.blocks.length + ' 项）</h3>' + (window.TD.audit ? window.TD.audit.auditHtml(w._audit) : '') + '<div style="text-align:right;margin-top:12px"><button class="btn" onclick="TDUI.closeModal()">关闭</button></div>');
          return;
        }
        var fname = exporter.safeName((w.carrier ? w.carrier + '_' : '') + 'BOOKING_' + (w._data ? w._data.invoiceNo : '')) + '.xlsx';
        await exporter.download(w._wb, fname);
        w.doc.status = 'exported'; w.doc.exportedFile = fname;
        await db.put('documents', w.doc);
        toast('订舱单已导出: ' + fname, 'ok');
        state.bwiz = null; state.tab = 'documents';
        document.querySelectorAll('.tab').forEach(function (t) { t.classList.toggle('active', t.dataset.tab === 'documents'); });
        render();
      };
      return;
    }
  }

  // ================= 单证记录页 =================
  PAGES.documents = async function () {
    var docs = await db.all('documents');
    docs.sort(function (a, b) { return b.createdAt - a.createdAt; });
    var stBadge = { draft: '<span class="badge gray">草稿</span>', validated: '<span class="badge yellow">已校验</span>', confirmed: '<span class="badge blue">已确认</span>', exported: '<span class="badge green">已导出</span>' };
    var rows = docs.map(function (d) {
      return '<tr><td>' + (d.kind === 'invoice' ? '🧾 发票' : '🚢 订舱单') + '</td><td class="mono">' + esc(d.docNo) + '</td>' +
        '<td class="mono">' + esc((d.data && d.data.orderNos) || '') + '</td><td>' + esc(d.carrier || '') + '</td>' +
        '<td>' + (stBadge[d.status] || d.status) + '</td><td>' + fmtTime(d.createdAt) + '</td>' +
        '<td><button class="btn sm ghost doc-view" data-id="' + d.id + '">预览</button> ' +
        '<button class="btn sm ok doc-re" data-id="' + d.id + '">重新导出</button> ' +
        '<button class="btn sm danger doc-del" data-id="' + d.id + '">删除</button></td></tr>';
    }).join('');
    return '<h2>单证记录</h2><div class="card"><table class="grid"><tr><th>类型</th><th>单证号</th><th>关联订单</th><th>物流商</th><th>状态</th><th>时间</th><th>操作</th></tr>' +
      (rows || '<tr><td colspan="7" class="empty">暂无生成记录</td></tr>') + '</table></div>';
  };
  BINDERS.documents = function () {
    document.querySelectorAll('.doc-del').forEach(function (b) {
      b.onclick = async function () {
        if (!(await confirmBox('确认删除该单证记录？', true))) return;
        await db.del('documents', b.dataset.id); toast('已删除', 'ok'); render();
      };
    });
    async function fillDocWb(d) {
      var tpl = await db.get('templates', d.templateId);
      if (!tpl) throw new Error('关联模板已删除，无法重新生成');
      var wb = await loadWb(tpl.fileBuf, tpl.fileName);
      engine.fillTemplate(wb, d.data, { logo: tpl.logo || null, labelMap: ((tpl.mapping && tpl.mapping.labelMap) || []) });
      return wb;
    }
    document.querySelectorAll('.doc-view').forEach(function (b) {
      b.onclick = async function () {
        try {
          var d = await db.get('documents', b.dataset.id);
          var wb = await fillDocWb(d);
          showModal('<h3>单证预览 · ' + esc(d.docNo) + '</h3><div style="overflow:auto">' + wbToHtml(wb) + '</div>' +
            '<div style="text-align:right;margin-top:12px"><button class="btn" onclick="TDUI.closeModal()">关闭</button></div>');
        } catch (e) { toast(e.message, 'err'); }
      };
    });
    document.querySelectorAll('.doc-re').forEach(function (b) {
      b.onclick = async function () {
        try {
          var d = await db.get('documents', b.dataset.id);
          if (d.status !== 'confirmed' && d.status !== 'exported') { toast('该单证尚未确认，不可导出（状态: ' + d.status + '）', 'err'); return; }
          var wb = await fillDocWb(d);
          var fname = d.exportedFile || (exporter.safeName(d.docNo) + '.xlsx');
          await exporter.download(wb, fname);
          if (d.status !== 'exported') { d.status = 'exported'; await db.put('documents', d); render(); }
          toast('已导出: ' + fname, 'ok');
        } catch (e) { toast(e.message, 'err'); }
      };
    });
  };

  // ================= 设置页 =================
  PAGES.settings = async function () {
    var remote = (await db.get('config', 'remote')) || { value: {} };
    var jst = (await db.get('config', 'jst')) || { value: {} };
    return '<h2>设置</h2>' +
      '<div class="card"><h3>☁️ 远程数据源（有服务器时启用在线反查/拉取）</h3><div class="form-grid">' +
      '<div><label>Base URL</label><input id="st-url" value="' + esc(remote.value.baseURL || '') + '" placeholder="https://your-server.com/api"></div>' +
      '<div><label>Token（可选）</label><input id="st-token" value="' + esc(remote.value.token || '') + '"></div></div>' +
      '<p class="hint">约定：GET {baseURL}/declare_reqs.json、/parties.json 返回数组；不可达时系统自动降级本地数据。</p>' +
      '<div style="display:flex;gap:8px"><button class="btn" id="st-save-remote">保存</button><button class="btn ghost" id="st-test">测试连接</button>' +
      '<button class="btn warn" id="st-pull-parties">拉取收发货人</button><button class="btn warn" id="st-pull-declares">拉取申报信息</button></div></div>' +
      '<div class="card"><h3>🔌 聚水潭开放平台（预留）</h3><div class="form-grid">' +
      '<div><label>app_key</label><input id="st-jst-key" value="' + esc(jst.value.appKey || '') + '"></div>' +
      '<div><label>app_secret</label><input id="st-jst-secret" type="password" value="' + esc(jst.value.appSecret || '') + '"></div></div>' +
      '<p class="hint">当前版本以手动导入xlsx为主；填入凭证后后续版本可开通API自动拉单。</p>' +
      '<button class="btn" id="st-save-jst">保存</button></div>' +
      '<div class="card"><h3>💾 数据备份与本地化</h3>' +
      '<p class="hint">所有数据存在浏览器 IndexedDB，换浏览器 / 清缓存 / 用他人电脑会丢失。可导出为「自包含 HTML」（数据嵌入文件，双击即用）或 JSON 备份。</p>' +
      '<div style="display:flex;gap:8px;flex-wrap:wrap">' +
      '<button class="btn ok" id="st-export-html">📦 导出自包含 HTML（推荐）</button>' +
      '<button class="btn" id="st-export-json">📤 导出 JSON 备份</button>' +
      '<button class="btn" id="st-import-json">📥 导入 JSON 备份</button>' +
      '</div></div>' +
      '<div class="card"><h3>🔄 团队共享主数据（GitHub）</h3>' +
      '<p class="hint"><b>三步同步团队数据：</b>① 打开 <a href="https://github.com/settings/personal-access-tokens" target="_blank" rel="noopener">github.com/settings/personal-access-tokens</a> 建<strong>细粒度 PAT</strong>（仅授权本仓库 <code>trade-docs-system</code>、权限 Contents = Read and write）；② 把整串 <code>github_pat_...</code> 粘贴到下面的「上传 Token」框（粘贴后显示圆点属正常，已自动存本地）；③ 点绿色「⬆️ 上传到团队库」。其他人打开系统自动拉取，无需再维护。</p>' +
      '<div class="form-grid">' +
      '<div><label>上传 Token（仅本地保存，不写进公开代码）</label><input id="sync-token" type="password" placeholder="整串粘贴 github_pat_ 开头的 token"></div>' +
      '<div><label><input type="checkbox" id="sync-auto"> 启动时自动从团队库拉取</label></div></div>' +
      '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:8px">' +
      '<button class="btn ok" id="sync-push">⬆️ 上传到团队库</button>' +
      '<button class="btn" id="sync-pull">⬇️ 从团队库拉取</button>' +
      '</div>' +
      '<p class="hint" id="sync-status"></p></div>' +
      '<div class="card"><h3>📂 模板管理（单一真源：团队库）</h3>' +
      '<p class="hint">模板不再内置/内嵌，全部来自 GitHub 团队库。下述按钮用于「始终展示最新」：重新同步会清空本地模板缓存、以团队库为准重建；批量导入文件夹可把本地文件夹整体上传覆盖团队库（同名文件更新、删掉的文件会从团队库移除）。</p>' +
      '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:8px">' +
      '<button class="btn ok" id="tpl-resync">🔄 清空本地缓存并重新同步</button>' +
      '<button class="btn" id="tpl-import-folder">📁 批量导入文件夹并上传团队库</button>' +
      '<input type="file" id="tpl-folder-input" webkitdirectory directory multiple style="display:none">' +
      '</div>' +
      '<p class="hint" id="tpl-mgmt-status"></p></div>' +
      '<div class="card"><h3>⚠️ 危险操作</h3>' +
      '<div style="display:flex;gap:8px;flex-wrap:wrap">' +
      '<button class="btn warn" id="st-scan-tpl">🛠 扫描并列出损坏的模板</button>' +
      '<button class="btn danger" id="st-wipe">清空全部本地数据（不可恢复）</button>' +
      '</div>' +
      '<div id="tpl-scan-result" class="hint" style="margin-top:8px"></div></div>' +
      '<footer class="note">贸易单证系统 · 纯前端本地存储(IndexedDB) · 七层解耦架构 · 模板占位符引擎</footer>';
  };
  BINDERS.settings = async function () {
    document.getElementById('st-save-remote').onclick = async function () {
      await db.put('config', { key: 'remote', value: { baseURL: val('st-url'), token: val('st-token') } });
      toast('远程数据源已保存', 'ok');
    };
    document.getElementById('st-test').onclick = async function () {
      var remote = new adapters.RemoteHttpAdapter({ baseURL: val('st-url'), token: val('st-token'), timeoutMs: 5000 });
      toast('测试中…');
      var ok = await remote.available();
      toast(ok ? '✅ 远程数据源可达' : '❌ 远程数据源不可达（将降级本地）', ok ? 'ok' : 'err');
    };
    function pullBtn(id, store, label) {
      document.getElementById(id).onclick = async function () {
        var remote = new adapters.RemoteHttpAdapter({ baseURL: val('st-url'), token: val('st-token') });
        var r = await adapters.syncFromRemote(db, remote, store);
        toast(r.ok ? '✅ ' + label + '同步成功，合并 ' + r.merged + ' 条' : '❌ ' + r.error, r.ok ? 'ok' : 'err');
      };
    }
    pullBtn('st-pull-parties', 'parties', '收发货人');
    document.getElementById('st-pull-declares').onclick = async function () {
      toast('正在以《商品申报信息》表刷新本地申报要素…');
      var r = await syncDeclaresFromMirror();
      if (r.ok) toast('✅ 申报信息同步完成：新增 ' + r.added + ' 条、补齐 ' + r.merged + ' 条', 'ok');
      else toast('同步失败: ' + r.error, 'err');
    };
    document.getElementById('st-save-jst').onclick = async function () {
      await db.put('config', { key: 'jst', value: { appKey: val('st-jst-key'), appSecret: val('st-jst-secret') } });
      toast('聚水潭凭证已保存（API拉取功能待开通）', 'ok');
    };
    document.getElementById('st-scan-tpl').onclick = async function () {
      var tpls = await db.all('templates');
      var bad = [], ok = 0;
      for (var i = 0; i < tpls.length; i++) {
        var t = tpls[i];
        var info = _inspectFileBuf(t.fileBuf);
        var isZip = info.bytes.length >= 4 && info.bytes[0] === 0x50 && info.bytes[1] === 0x4B;
        var isCfb = info.bytes.length >= 8 && info.bytes[0] === 0xD0 && info.bytes[1] === 0xCF && info.bytes[2] === 0x11 && info.bytes[3] === 0xE0;
        if (!info.size || (!isZip && !isCfb)) {
          bad.push({ id: t.id, name: t.name, kind: t.kind, size: info.size, shape: info.shape, format: !info.size ? 'empty' : (isCfb ? 'xls-old' : 'unknown') });
        } else { ok++; }
      }
      var el = document.getElementById('tpl-scan-result');
      if (!bad.length) {
        el.innerHTML = '<div class="vres ok">✅ 扫描完毕：' + tpls.length + ' 个模板全部健康（' + ok + ' 个有效 xlsx/xls）</div>';
        return;
      }
      el.innerHTML = '<div class="vres warn"><b>🔍 发现 ' + bad.length + ' 个损坏/格式异常模板（' + ok + ' 个健康）：</b><ul style="margin:6px 0 6px 20px">' +
        bad.map(function (b) { return '<li><code>' + esc(b.id.slice(0, 12)) + '…</code> 「' + esc(b.name) + '」 <span class="hint">(' + esc(b.kind) + ' · ' + esc(b.shape) + ' · ' + b.size + 'B)</span> ' +
          '<button class="btn sm danger tpl-del-bad" data-id="' + esc(b.id) + '" style="margin-left:8px">🗑 删除</button></li>'; }).join('') +
        '</ul>' +
        '<p class="hint" style="margin-top:6px">💡 这些模板的 fileBuf 已不是有效 xlsx（形态 <code>empty-object / null</code> 通常是 sync 拉空数据导致）。点删除可让向导 step 4 不再报错；坏模板里如果有真实发票样张，请在「模板管理」重新上传（建议同时另存为标准 .xlsx）。</p>' +
        '</div>';
      el.querySelectorAll('.tpl-del-bad').forEach(function (btn) {
        btn.onclick = async function () {
          if (!(await confirmBox('确认删除损坏模板 id=' + btn.dataset.id + '？'))) return;
          await db.del('templates', btn.dataset.id);
          toast('已删除', 'ok');
          document.getElementById('st-scan-tpl').click();
        };
      });
    };
    document.getElementById('st-wipe').onclick = async function () {
      if (!(await confirmBox('⚠️ 将清空订单/装箱清单/主数据/模板/单证记录全部本地数据！若已配置「团队共享+自动同步」，清空后会自动从团队库恢复主数据与模板；未配置则主数据将永久丢失。确认？', true))) return;
      if (!(await confirmBox('二次确认：真的要清空全部数据吗？', true))) return;
      for (var i = 0; i < db.STORES.length; i++) await db.clear(db.STORES[i]);
      toast('已清空，正在重新初始化…', 'ok');
      await seed.run(db, engine, ExcelJS);
      // v1.5.22 若已配置团队同步，清空后自动从团队库恢复主数据+模板，防止误清丢失
      try {
        var _sc = await loadSyncCfg();
        if (_sc && _sc.auto && _sc.token) {
          var _pr = await pullShared();
          if (_pr && _pr.ok) toast('已从团队库恢复 ' + _pr.merged + ' 条主数据/模板', 'ok');
        }
      } catch (e) { /* 离线则下次启动自动拉取 */ }
      render();
    };
    document.getElementById('st-export-html').onclick = exportSelfContainedHTML;
    document.getElementById('st-export-json').onclick = exportJSON;
    document.getElementById('st-import-json').onclick = importJSON;
    // ---------- 团队共享主数据 ----------
    function setSyncStatus(t) { var el = document.getElementById('sync-status'); if (el) el.textContent = t; }
    var scfg = await loadSyncCfg();
    document.getElementById('sync-token').value = scfg.token || '';
    document.getElementById('sync-auto').checked = scfg.auto;
    document.getElementById('sync-token').onchange = async function () { var c = await loadSyncCfg(); c.token = this.value; await saveSyncCfg(c); toast('Token 已保存到本地', 'ok'); };
    document.getElementById('sync-auto').onchange = async function () { var c = await loadSyncCfg(); c.auto = this.checked; await saveSyncCfg(c); toast(this.checked ? '已开启启动自动拉取' : '已关闭自动拉取', 'ok'); };
    document.getElementById('sync-push').onclick = async function () {
      var c = await loadSyncCfg();
      setSyncStatus('上传中（与团队库三方合并）…');
      var r = await pushShared(c.token);
      var okMsg = '✅ 已合并上传到团队库（' + (r.merged || 0) + ' 条' + (r.conflicts ? '，' + r.conflicts + ' 处冲突按最新时间合并' : '') + '），团队其他人刷新即可看到';
      setSyncStatus(r.ok ? okMsg : '❌ ' + r.error);
      toast(r.ok ? okMsg : '❌ ' + r.error, r.ok ? 'ok' : 'err');
    };
    document.getElementById('sync-pull').onclick = async function () {
      setSyncStatus('拉取中…');
      var r = await pullShared();
      setSyncStatus(r.ok ? '✅ 已拉取 ' + r.merged + ' 条到本地' : '❌ ' + r.error);
      toast(r.ok ? '✅ 拉取成功 ' + r.merged + ' 条' : '❌ ' + r.error, r.ok ? 'ok' : 'err');
      if (r.ok) render();
    };
    // ---------- 模板管理：重新同步 + 批量导入文件夹 ----------
    document.getElementById('tpl-resync').onclick = async function () {
      var el = document.getElementById('tpl-mgmt-status');
      el.textContent = '正在清空本地模板缓存并从团队库重新同步…';
      var r = await pullShared();
      if (r.ok) { el.textContent = '✅ 已以团队库为准重建本地模板（' + r.merged + ' 条）'; toast('✅ 重新同步成功，' + r.merged + ' 个模板', 'ok'); }
      else { el.textContent = '❌ ' + r.error; toast('❌ 同步失败: ' + r.error, 'err'); }
      render();
    };
    document.getElementById('tpl-import-folder').onclick = function () {
      document.getElementById('tpl-folder-input').click();
    };
    document.getElementById('tpl-folder-input').onchange = async function (ev) {
      var files = ev.target.files;
      if (!files || !files.length) return;
      var el = document.getElementById('tpl-mgmt-status');
      var c = await loadSyncCfg();
      if (!c.token) { el.textContent = '⚠️ 请先在上方「上传 Token」框填入细粒度 PAT 再批量导入'; toast('请先填写上传 Token', 'err'); return; }
      el.textContent = '正在读取 ' + files.length + ' 个文件并上传团队库…';
      try {
        var r = await _replaceCloudTemplatesByFolder(c.token, files);
        if (r.ok) { el.textContent = '✅ 已用文件夹 ' + r.count + ' 个模板覆盖团队库并重新同步'; toast('✅ 批量导入成功，' + r.count + ' 个模板已上传', 'ok'); }
        else { el.textContent = '❌ ' + r.error; toast('❌ 导入失败: ' + r.error, 'err'); }
      } catch (e) { el.textContent = '❌ ' + e.message; toast('❌ 导入失败: ' + e.message, 'err'); }
      render();
    };
  };

  // ---------- 数据本地化：嵌入 HTML / 导出导入 ----------
  function b64encodeUnicode(str) {
    var bytes = new TextEncoder().encode(str);
    var bin = '';
    bytes.forEach(function (b) { bin += String.fromCharCode(b); });
    return btoa(bin);
  }
  function b64decodeUnicode(str) {
    var bin = atob(str);
    var bytes = Uint8Array.from(bin, function (c) { return c.charCodeAt(0); });
    return new TextDecoder().decode(bytes);
  }

  // ---------- v1.4.39: fileBuf 多形态检测 + sync 序列化/反序列化 ----------
  // 诊断 fileBuf 的真实形态和大小（兼容 ArrayBuffer / TypedArray / {data:[]}/ {type:'Buffer',data:[]}/ 损坏）
  function _inspectFileBuf(x) {
    var shape = 'unknown', size = 0, bytes = new Uint8Array(0);
    if (x == null) { shape = 'null'; return { shape: shape, size: 0, bytes: bytes }; }
    if (x instanceof ArrayBuffer) { shape = 'ArrayBuffer'; size = x.byteLength; bytes = new Uint8Array(x); }
    else if (ArrayBuffer.isView(x)) { shape = 'TypedArray'; size = x.byteLength; bytes = new Uint8Array(x.buffer, x.byteOffset, x.byteLength); }
    else if (typeof x === 'string') { shape = 'base64-string'; size = x.length; try { var _bin = atob(x); bytes = Uint8Array.from(_bin, function (c) { return c.charCodeAt(0); }); } catch (e) { bytes = new Uint8Array(0); } }
    else if (typeof x === 'object') {
      var arr = (x.data && Array.isArray(x.data)) ? x.data : (Array.isArray(x) ? x : null);
      if (arr) { shape = (x.type === 'Buffer') ? '{type:Buffer,data:[]}' : '{data:[]}'; size = arr.length; bytes = Uint8Array.from(arr); }
      else { shape = 'empty-object'; size = 0; }
    }
    return { shape: shape, size: size, bytes: bytes };
  }
  // 把任意形态的 fileBuf 规范化为 ArrayBuffer（base64-string / {type:Buffer,data:[]} / TypedArray / ArrayBuffer 都接）
  // 用于 loadWb 等所有"用 fileBuf"的入口，幂等
  function _asArrayBuffer(x) {
    var info = _inspectFileBuf(x);
    if (info.shape === 'ArrayBuffer') return x;
    if (info.shape === 'TypedArray') return x.buffer;
    if (info.shape === 'base64-string') return info.bytes.buffer;
    if (info.shape === '{type:Buffer,data:[]}' || info.shape === '{data:[]}') return info.bytes.buffer;
    throw new Error('fileBuf 形态无效: ' + info.shape + ' (size=' + info.size + ')');
  }
  // 启动自愈：把 IndexedDB 里残留的 base64-string 模板就地还原成 ArrayBuffer（v1.4.39 之前或 fileBufShape 缺失情况下漏转的）
  async function _migrateBase64Templates() {
    try {
      var templates = await db.all('templates');
      var fixed = 0;
      for (var i = 0; i < templates.length; i++) {
        var t = templates[i];
        if (!t.fileBuf) continue;
        var info = _inspectFileBuf(t.fileBuf);
        if (info.shape !== 'base64-string') continue;
        var ab;
        try { ab = _asArrayBuffer(t.fileBuf); } catch (e) { continue; }
        var u8 = new Uint8Array(ab);
        var isValid = (u8.length >= 4 && u8[0] === 0x50 && u8[1] === 0x4B && (u8[2] === 0x03 || u8[2] === 0x05)) ||
                      (u8.length >= 8 && u8[0] === 0xD0 && u8[1] === 0xCF && u8[2] === 0x11 && u8[3] === 0xE0);
        if (!isValid) continue;
        t.fileBuf = ab;
        delete t.fileBufShape; delete t.fileBufSize;
        await db.put('templates', t);
        fixed++;
      }
      if (fixed) console.log('[v1.4.42 自愈] 已把 ' + fixed + ' 个 base64-string 模板还原为 ArrayBuffer');
      return fixed;
    } catch (e) {
      console.warn('_migrateBase64Templates failed', e);
      return 0;
    }
  }
  // 把对象里所有 fileBuf(ArrayBuffer/TypedArray) 转成 base64 字符串（push 序列化前调）
  function _preSerializeForSync(obj) {
    if (obj == null) return obj;
    if (Array.isArray(obj)) return obj.map(_preSerializeForSync);
    if (typeof obj === 'object' && !(obj instanceof ArrayBuffer) && !ArrayBuffer.isView(obj)) {
      var out = {};
      for (var k in obj) {
        if (k === 'fileBuf' && obj[k] != null) {
          var info = _inspectFileBuf(obj[k]);
          if (info.size > 0) {
            // ArrayBuffer → base64
            out.fileBuf = btoa(String.fromCharCode.apply(null, info.bytes));
            out.fileBufShape = info.shape;
            out.fileBufSize = info.size;
          } else {
            // 空 fileBuf 标 - 避免把空 {} 推上团队库污染别人
            out.fileBuf = null;
            out.fileBufBroken = true;
          }
        } else {
          out[k] = _preSerializeForSync(obj[k]);
        }
      }
      return out;
    }
    return obj;
  }
  // pull 后还原：把 fileBuf base64 / {type:'Buffer', data:[]} 还原成 ArrayBuffer
  function _postDeserializeFromSync(obj) {
    if (obj == null) return obj;
    if (Array.isArray(obj)) return obj.map(_postDeserializeFromSync);
    if (typeof obj === 'object' && !(obj instanceof ArrayBuffer) && !ArrayBuffer.isView(obj)) {
      var out = {};
      for (var k in obj) {
        out[k] = _postDeserializeFromSync(obj[k]);
      }
      // 还原 fileBuf
      if (typeof out.fileBuf === 'string' && out.fileBuf.length > 0 && out.fileBufShape) {
        try {
          var bin = atob(out.fileBuf);
          var u8 = Uint8Array.from(bin, function (c) { return c.charCodeAt(0); });
          out.fileBuf = u8.buffer;
        } catch (e) {}
      } else if (out.fileBuf && typeof out.fileBuf === 'object' && Array.isArray(out.fileBuf.data)) {
        // 旧版 {type:'Buffer', data:[]} 形态
        var u8b = Uint8Array.from(out.fileBuf.data);
        out.fileBuf = u8b.buffer;
      } else if (out.fileBuf && Array.isArray(out.fileBuf)) {
        var u8c = Uint8Array.from(out.fileBuf);
        out.fileBuf = u8c.buffer;
      } else if (out.fileBuf == null || (typeof out.fileBuf === 'object' && Object.keys(out.fileBuf).length === 0)) {
        out.fileBufBroken = true;
      }
      return out;
    }
    return obj;
  }
  // ---------- v1.5.24 模板外置 helpers ----------
  // Uint8Array/ArrayBuffer → base64（分块避免 apply 栈溢出）
  function _b64FromBytes(u8) {
    var CH = 0x8000, bin = '';
    for (var i = 0; i < u8.length; i += CH) {
      bin += String.fromCharCode.apply(null, u8.subarray(i, Math.min(i + CH, u8.length)));
    }
    return btoa(bin);
  }
  // 模板对象 → 团队库元数据（去 fileBuf，加 filePath 指向 templates/<id>.xlsx）
  function _tplToMeta(t) {
    var m = {};
    for (var k in t) {
      if (k !== 'fileBuf' && k !== 'fileBufShape' && k !== 'fileBufSize' && k !== 'previewBuf') m[k] = t[k];
    }
    m.filePath = SYNC_TPL_DIR + t.id + '.xlsx';
    var _info = _inspectFileBuf(t.fileBuf);
    if (_info.size > 0) m.fileSize = _info.size;
    return m;
  }
  // 上传单个模板文件到 GitHub templates/<id>.xlsx（已存在则带 sha 覆盖），返回是否成功
  async function _putTplFile(token, tpl) {
    var info = _inspectFileBuf(tpl.fileBuf);
    if (info.size <= 0) return false;
    // GitHub Contents API 单文件上限 100MB；留 5% 余量提前拦截，避免大文件上传超时/失败无提示
    if (info.size > 95 * 1024 * 1024) return false;
    var fname = tpl.id + '.xlsx';
    var encName = encodeURIComponent(fname);
    var b64 = _b64FromBytes(info.bytes);
    var sha = null;
    try {
      var g = await fetch(SYNC_TPL_API + encName + '?ref=' + SYNC_BRANCH, { headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json' } });
      if (g.ok) { var gj = await g.json(); sha = gj.sha; }
    } catch (e) {}
    var res = await fetch(SYNC_TPL_API + encName, {
      method: 'PUT',
      headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json', 'Content-Type': 'application/json' },
      body: JSON.stringify({ message: 'sync: template ' + tpl.name, content: b64, branch: SYNC_BRANCH, sha: sha })
    });
    return res.ok;
  }
  // 拉取模板文件（v1.5.37 同源优先 + raw 兜底，双镜像都能稳定拉到）：
  // 1) 先试当前站点同源 /templates/（CloudStudio / GitHub Pages 都自带模板目录 → 国内访问快且稳）
  // 2) 同源失败再回退 raw.githubusercontent.com（GitHub 直连，国内偶发慢但兜底可用）
  async function _fetchTplFile(fname) {
    var candidates = [];
    try {
      var base = location.href.split('#')[0].split('?')[0];
      base = base.slice(0, base.lastIndexOf('/') + 1);
      candidates.push(base + SYNC_TPL_DIR + encodeURIComponent(fname));
    } catch (e) {}
    candidates.push(SYNC_TPL_CDN + encodeURIComponent(fname));
    // v1.5.45：每个候选重试 3 次（国内访问 GitHub raw 不稳，单次失败率高），单次超时 10s，
    // 失败指数退避 0.5s/1s，整体大幅降低「模板文件拉取失败→fileBuf 空→step4 才暴露」的概率。
    for (var i = 0; i < candidates.length; i++) {
      for (var attempt = 0; attempt < 3; attempt++) {
        var ac = new AbortController();
        var to = setTimeout(function () { try { ac.abort(); } catch (e) {} }, 10000);
        try {
          var res = await fetch(candidates[i], { cache: 'no-store', signal: ac.signal });
          if (res.ok) { clearTimeout(to); return await res.arrayBuffer(); }
        } catch (e) {} finally { clearTimeout(to); }
        if (attempt < 2) { await new Promise(function (r) { setTimeout(r, 500 * (attempt + 1)); }); }
      }
    }
    return null;
  }
  // 启动时：若文件内置了用户数据快照，合并写回 IndexedDB（同名 key 覆盖）
  async function hydrateFromEmbedded() {
    var b64 = window.__USERDATA_B64__;
    if (!b64) return;
    try {
      var data = JSON.parse(b64decodeUnicode(b64));
      var total = 0;
      for (var i = 0; i < db.STORES.length; i++) {
        var s = db.STORES[i], list = data[s];
        if (list && list.length) { await db.bulkPut(s, list); total += list.length; }
      }
      if (total) toast('已从文件内置数据恢复 ' + total + ' 条记录', 'ok');
    } catch (e) {
      console.error('hydrateFromEmbedded failed', e);
    }
    window.__USERDATA_B64__ = null; // 防止重复写入
  }
  // ---------- 团队共享主数据（GitHub 仓库 userdata.json） ----------
  // 共享范围：用户维护、需团队复用的主数据（收发货人 + 自定义模板）。
  // 订单/装箱清单/单证记录属交易数据不强制共享；申报信息来自飞书也不塞（避免超 1MB 接口上限）。
  var SYNC_STORES = ['parties', 'templates'];
  var SYNC_REPO = 'heryma99/trade-docs-system';
  var SYNC_BRANCH = 'main';
  var SYNC_FILE = 'userdata.json';
  // 改用 raw.githubusercontent.com（GitHub raw 不缓存或缓存极短，几秒级生效）。
  // 历史曾用 jsDelivr CDN，但 s-maxage=12h + purge 沙箱不可达，会卡 4~12 小时，导致「拉取成功 0 条」。
  var SYNC_CDN = 'https://raw.githubusercontent.com/' + SYNC_REPO + '/' + SYNC_BRANCH + '/' + SYNC_FILE;
  var SYNC_API = 'https://api.github.com/repos/' + SYNC_REPO + '/contents/' + SYNC_FILE;
  // v1.5.24 模板外置：模板 xlsx 单独存 GitHub templates/ 目录，userdata.json 只存元数据+主数据。
  // 解决 userdata.json 内联模板 fileBuf 后 base64 超 1MB 接口上限、浏览器 push 必失败的 bug。
  var SYNC_TPL_DIR = 'templates/';
  var SYNC_TPL_CDN = 'https://raw.githubusercontent.com/' + SYNC_REPO + '/' + SYNC_BRANCH + '/' + SYNC_TPL_DIR;
  var SYNC_TPL_API = 'https://api.github.com/repos/' + SYNC_REPO + '/contents/' + SYNC_TPL_DIR;
  async function loadSyncCfg() {
    var c = await db.get('config', 'sync');
    // v1.5.42 自动同步默认开：只要「未显式设置 auto=false」就自动拉取。
    // 之前 bug：用户保存过 token(c.value 存在) 但 auto 未显式设置时 !!undefined=false → 不自动拉，需手动点。
    // 团队库公开仓库，拉取无需 token，默认开安全；仅用户显式取消勾选(auto===false)才停。
    var auto = true;
    if (c && c.value && c.value.auto === false) auto = false;
    return { token: (c && c.value && c.value.token) || '', auto: auto };
  }
  async function saveSyncCfg(cfg) {
    await db.put('config', { key: 'sync', value: { token: cfg.token || '', auto: !!cfg.auto } });
  }
  // 三方合并基准：上次成功 pull/push 时记录的服务端快照（用于 3-way merge，多人协同不互覆盖）
  async function loadSyncBase() {
    var c = await db.get('config', 'syncBase');
    return (c && c.value) || { stores: {} };
  }
  async function saveSyncBase(snapshot) {
    await db.put('config', { key: 'syncBase', value: snapshot });
  }
  function _toMap(list) {
    var m = {}; if (!list) return m;
    for (var i = 0; i < list.length; i++) { var r = list[i]; if (r && r.id) m[r.id] = r; }
    return m;
  }
  function _eq(a, b) { return JSON.stringify(a) === JSON.stringify(b); }
  function _newer(a, b) { return (a && a.updatedAt || 0) >= (b && b.updatedAt || 0) ? a : b; }
  // 3-way 合并：base=上次服务端快照, local=本地, remote=当前服务端。返回合并后数组（按 id 去重）。
  function threeWayMerge(baseList, localList, remoteList, onConflict) {
    var b = _toMap(baseList), l = _toMap(localList), r = _toMap(remoteList);
    var ids = {}, out = [];
    [b, l, r].forEach(function (m) { Object.keys(m).forEach(function (k) { ids[k] = 1; }); });
    Object.keys(ids).forEach(function (id) {
      var bv = b[id], lv = l[id], rv = r[id];
      if (lv && !rv && !bv) { out.push(lv); return; }                       // 本地新增
      if (!lv && rv && !bv) { out.push(rv); return; }                       // 远端新增
      if (!lv && !rv) return;                                              // 都不存在
      if (bv && lv && !rv) { if (!_eq(lv, bv)) out.push(lv); return; }      // 本地改/删
      if (bv && !lv && rv) { if (!_eq(rv, bv)) out.push(rv); return; }      // 本地删、远端改 → 远端赢
      if (lv && rv) {
        if (_eq(lv, rv)) { out.push(lv); return; }                         // 完全一致
        if (!bv) { if (onConflict) onConflict(); out.push(_newer(lv, rv)); return; } // 并发新增同 id
        if (_eq(lv, bv)) { out.push(rv); return; }                          // 本地未动、远端改
        if (_eq(rv, bv)) { out.push(lv); return; }                          // 远端未动、本地改
        if (onConflict) onConflict();                                       // 双方都改 → 按 updatedAt 取新
        out.push(_newer(lv, rv)); return;
      }
    });
    return out;
  }
  // ---------- 发票模板自愈（add-only，绕过 threeWayMerge，单独用） ----------
  // v1.4.46 抽出来：之前直接调 pullShared()，但 pullShared 会 db.clear + 三方合并，
  // 在「本地已有老 invoice + 远端有同 id 新 invoice」边界 case 下会把 invoice 模板全丢，
  // 而且会顺手把本地其他非 invoice 模板也一起清掉。
  // 这里只做"从云端取 invoice 模板 add-only 写入本地"，不动其他 store、不清表、不 merge。
  async function _recoverInvoiceTemplates() {
    try {
      var res = await fetch(SYNC_CDN, { cache: 'no-store' });
      if (!res.ok) return { ok: false, error: '云端 HTTP ' + res.status, got: 0 };
      var r = await res.json();
      if (!r || !r.stores || !Array.isArray(r.stores.templates)) {
        return { ok: false, error: '云端数据格式错误（无 stores.templates）', got: 0 };
      }
      var remoteList = r.stores.templates.filter(function (x) { return x && x.kind === 'invoice' && x.status === 'active' && !x.isSeed; });
      if (!remoteList.length) return { ok: true, got: 0, note: '云端没有可恢复的发票模板' };
      // 还原 fileBuf（v1.4.39 修复：push 上云的是 base64-string 或 {type:'Buffer',data:[]}）
      remoteList = remoteList.map(function (t) {
        var out = {};
        for (var k in t) out[k] = t[k];
        if (typeof out.fileBuf === 'string' && out.fileBuf.length > 0 && out.fileBufShape) {
          try {
            var bin = atob(out.fileBuf);
            var u8 = Uint8Array.from(bin, function (c) { return c.charCodeAt(0); });
            out.fileBuf = u8.buffer;
          } catch (e) {}
        } else if (out.fileBuf && typeof out.fileBuf === 'object' && Array.isArray(out.fileBuf.data)) {
          out.fileBuf = Uint8Array.from(out.fileBuf.data).buffer;
        } else if (Array.isArray(out.fileBuf)) {
          out.fileBuf = Uint8Array.from(out.fileBuf).buffer;
        }
        return out;
      });
      // add-only：已存在的同 id 不覆盖（避免覆盖本地更新的版本），只新增缺的
      var local = await db.all('templates');
      var localIds = {};
      local.forEach(function (x) { if (x && x.id) localIds[x.id] = 1; });
      var toPut = remoteList.filter(function (x) { return x && x.id && !localIds[x.id]; });
      if (toPut.length) await db.bulkPut('templates', toPut);
      return { ok: true, got: toPut.length, total: remoteList.length };
    } catch (e) {
      return { ok: false, error: e.message || String(e), got: 0 };
    }
  }

  // 拉取：公开 CDN，无需 token。与本地做三方合并（保留本地未上传的修改），并记录服务端快照为基准。
  async function pullShared() {
    var r;
    // v1.5.40 双通道拉取：同源(站点自带 userdata.json，GitHub Pages/CloudStudio 都有)优先，raw 兜底。
    // 之前只走 raw.githubusercontent.com(8s 超时)——国内访问 GitHub raw 经常超时静默失败，
    // 导致"其他人打开看不到模板/主数据"。同源部署时秒开；raw 仅当同源 404 时兜底。
    var fetchData = async function (url, ms) {
      var ac = new AbortController();
      var to = setTimeout(function () { try { ac.abort(); } catch (e) {} }, ms);
      try {
        var res = await fetch(url, { cache: 'no-store', signal: ac.signal });
        if (!res.ok) return { ok: false, err: 'HTTP ' + res.status };
        return { ok: true, data: await res.json() };
      } catch (e) {
        return { ok: false, err: (e && e.name === 'AbortError') ? 'timeout' : String(e.message || e) };
      } finally { clearTimeout(to); }
    };
    var sameOrigin = null;
    try {
      var base = location.href.split('#')[0].split('?')[0];
      base = base.slice(0, base.lastIndexOf('/') + 1);
      sameOrigin = base + SYNC_FILE;
    } catch (e) {}
    var f = sameOrigin ? await fetchData(sameOrigin, 8000) : { ok: false };
    if (!f.ok) {
      // 同源失败(如 GitHub Pages 尚未同步该文件) → raw 兜底，给更长时间
      var r2 = await fetchData(SYNC_CDN, 15000);
      if (r2.ok) { r = r2.data; }
      else { return { ok: false, error: '拉取失败（同源+raw 均不可达）: ' + (f.err || '') + ' / ' + (r2.err || '') }; }
    } else {
      r = f.data;
    }
    if (!r || !r.stores) return { ok: false, error: '共享数据格式错误' };
    var base = await loadSyncBase();
    var conflicts = 0, merged = 0;
    for (var i = 0; i < SYNC_STORES.length; i++) {
      var s = SYNC_STORES[i], list = r.stores[s];
      if (!list) continue;
      // 远端剔除 seed.js 内置的占位数据（isSeed:true），避免历史误混入团队库
      list = list.filter(function (x) { return !x || !x.isSeed; });
      if (s === 'templates') {
        // v1.4.59：模板以云端为唯一真源，全量替换本地（清掉历史内置/嵌入/老浏览器缓存残留的旧模板）。
        // v1.5.24：云端是元数据（filePath 指向 templates/<id>.xlsx），本地需组装 fileBuf：
        //   本地已有同 id 且 updatedAt 相同 → 复用本地 fileBuf（省网络）；否则 fetch 模板文件。
        var metas = _postDeserializeFromSync({ stores: { templates: list } }).stores.templates;
        var localTpls = await db.all('templates');
        var localById = {};
        localTpls.forEach(function (x) { if (x && x.id) localById[x.id] = x; });
        var out = [], needFetch = [];
        for (var ti = 0; ti < metas.length; ti++) {
          var mt = metas[ti];
          if (!mt || !mt.id) continue;
          var local = localById[mt.id];
          if (local && local.fileBuf && local.updatedAt === mt.updatedAt) { out.push(local); continue; }
          // 需要拉文件（并行 fetch）
          needFetch.push(mt);
        }
        var fetched = await Promise.all(needFetch.map(function (mt) { return _fetchTplFile(mt.id + '.xlsx').then(function (buf) { return { mt: mt, buf: buf }; }); }));
        fetched.forEach(function (f) {
          var t = {};
          for (var k in f.mt) t[k] = f.mt[k];
          if (f.buf && f.buf.byteLength > 0) { t.fileBuf = f.buf; }
          else t.fileBufBroken = true;
          out.push(t);
        });
        await db.clear('templates');
        await db.bulkPut('templates', out);
        merged += out.length;
        continue;
      }
      var local = await db.all(s);
      // 本地残留的未编辑占位也剔除，避免下次 push 把 DEMO 再次推上去
      local = local.filter(function (x) { return !x || !x.isSeed; });
      var m = threeWayMerge(base.stores ? base.stores[s] : null, local, list, function () { conflicts++; });
      await db.clear(s);
      await db.bulkPut(s, m);
      merged += m.length;
    }
    await saveSyncBase({ stores: r.stores }); // 服务端快照作为下次合并基准
    return { ok: true, merged: merged, conflicts: conflicts };
  }
  // v1.4.59：批量导入文件夹 → 以文件夹内容整体覆盖团队库 templates（同名更新、删掉的移除），再重新同步本地。
  async function _replaceCloudTemplatesByFolder(token, fileList) {
    // 1. 读当前云端 userdata.json（带 sha）
    var head = await fetch(SYNC_API + '?ref=' + SYNC_BRANCH, { headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json' } });
    if (!head.ok) return { ok: false, error: '读取云端失败 HTTP ' + head.status };
    var hj = await head.json();
    var sha = hj.sha;
    var remote = _postDeserializeFromSync(JSON.parse(b64decodeUnicode(hj.content)));
    var remoteTpls = (remote.stores && remote.stores.templates) || [];
    var byName = {};
    remoteTpls.forEach(function (t) { if (t && t.name) byName[t.name] = t; });
    // 2. 读本地选中的 xlsx
    var xlsx = Array.prototype.slice.call(fileList).filter(function (f) { return /\.xlsx?$/i.test(f.name) && f.name.indexOf('~$') !== 0; });
    if (!xlsx.length) return { ok: false, error: '未选中任何 xlsx 文件' };
    var newTpls = [];
    for (var i = 0; i < xlsx.length; i++) {
      var name = xlsx[i].name.replace(/\.xlsx?$/i, '');
      var carrier = name.indexOf('-') >= 0 ? name.split('-')[0] : '通用';
      var kind = /装箱/.test(name) ? 'packing' : (/申报|买单/.test(name) ? 'declare' : 'invoice');
      var existing = byName[name];
      var id = existing ? existing.id : ('tpl_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 8));
      var buf = await xlsx[i].arrayBuffer();
      newTpls.push({ id: id, name: name, kind: kind, carrier: carrier, status: 'active', isSeed: false, updatedAt: Date.now(), fileBuf: buf });
    }
    // 3. 上传每个模板文件 → templates/<id>.xlsx，userdata.json 只写元数据（v1.5.24 模板外置）
    for (var ni = 0; ni < newTpls.length; ni++) {
      var upOk = await _putTplFile(token, newTpls[ni]);
      if (!upOk) return { ok: false, error: '模板文件上传失败: ' + newTpls[ni].name };
    }
    remote.stores.templates = newTpls.map(function (t) { return _tplToMeta(t); });
    remote.updatedAt = Date.now();
    var payload = _preSerializeForSync(remote);
    var json = JSON.stringify(payload);
    if (json.length > 900 * 1024) return { ok: false, error: '数据超过 900KB（GitHub 接口上限 1MB），请减少模板数量' };
    var content = b64encodeUnicode(json);
    var res = await fetch(SYNC_API, {
      method: 'PUT',
      headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json', 'Content-Type': 'application/json' },
      body: JSON.stringify({ message: 'replace templates from local folder', content: content, branch: SYNC_BRANCH, sha: sha })
    });
    if (!res.ok) { var t = await res.text(); return { ok: false, error: '上传失败 HTTP ' + res.status + ' ' + t.slice(0, 200) }; }
    // 4. 重新同步本地（全量替换）
    var arr = _postDeserializeFromSync({ stores: { templates: newTpls } }).stores.templates;
    await db.clear('templates');
    await db.bulkPut('templates', arr);
    await saveSyncBase({ stores: { templates: newTpls } });
    return { ok: true, count: newTpls.length };
  }
  // 上传：用用户本地保存的细粒度 PAT。先拉当前服务端，与本地做三方合并（多人协同不互覆盖），
  // 再以服务端最新 sha 提交；若期间被他人并发提交（409/422）则重试最多 5 次。
  async function pushShared(token) {
    if (!token) return { ok: false, error: '请先在设置页填入「细粒度 PAT（仅本仓库 Contents 读写）」' };
    var local = {};
    for (var i = 0; i < SYNC_STORES.length; i++) local[SYNC_STORES[i]] = await db.all(SYNC_STORES[i]);
    // 本地剔除 seed.js 内置占位（isSeed:true），避免把 DEMO 数据推上团队库
    SYNC_STORES.forEach(function (s) { local[s] = (local[s] || []).filter(function (x) { return !x || !x.isSeed; }); });
      var base = await loadSyncBase();
      var lastErr = null;
      for (var attempt = 1; attempt <= 5; attempt++) {
        var sha = null, remote = { stores: {} };
        try {
          var head = await fetch(SYNC_API + '?ref=' + SYNC_BRANCH, { headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json' } });
          if (head.ok) { var hj = await head.json(); sha = hj.sha; remote = JSON.parse(b64decodeUnicode(hj.content)); }
        } catch (e) {}
        // v1.4.39 修复：pull 后还原 fileBuf（v1.4.38 之前 push 上去的可能是 {type:'Buffer', data:[...]}）
        remote = _postDeserializeFromSync(remote);
        var conflicts = 0, merged = {};
        for (var j = 0; j < SYNC_STORES.length; j++) {
          var s = SYNC_STORES[j];
          if (s === 'templates') {
            // v1.5.24 模板外置合并：云端=元数据(无 fileBuf)，本地=完整对象(有 fileBuf)。
            // 按 id 合并：本地新或本地独有 → 上传模板文件+写入元数据；云端新 → 保留云端元数据(文件已在云端)；仅云端 → 保留。
            var remoteTpls = (remote.stores && remote.stores.templates) || [];
            var remoteById = {};
            remoteTpls.forEach(function (x) { if (x && x.id) remoteById[x.id] = x; });
            var localTpls = local[s] || [];
            var mergedTpls = [], seen = {};
            for (var li = 0; li < localTpls.length; li++) {
              var lt = localTpls[li];
              if (!lt || !lt.id) continue;
              seen[lt.id] = true;
              var rt = remoteById[lt.id];
              if (!rt || (lt.updatedAt || 0) >= (rt.updatedAt || 0)) {
                // 文件未变（云端元数据 fileSize 与本地一致）则跳过上传，只写元数据，加快自动同步
                var _ltSize = lt.fileBuf ? _inspectFileBuf(lt.fileBuf).size : 0;
                if (rt && rt.fileSize && _ltSize === rt.fileSize) {
                  /* 文件已与云端一致，跳过 PUT */
                } else if (_ltSize > 0) {
                  var upOk = await _putTplFile(token, lt);
                  if (!upOk) return { ok: false, error: '模板文件上传失败: ' + lt.name };
                }
                mergedTpls.push(_tplToMeta(lt));
              } else {
                mergedTpls.push(rt);
              }
            }
            remoteTpls.forEach(function (x) { if (x && x.id && !seen[x.id]) mergedTpls.push(x); });
            merged[s] = mergedTpls;
          } else {
            merged[s] = threeWayMerge(base.stores ? base.stores[s] : null, local[s], remote.stores ? remote.stores[s] : null, function () { conflicts++; });
          }
        }
      var payload = { _meta: { app: 'trade-docs-system', ver: '1.4.35', updatedAt: new Date().toISOString(), stores: SYNC_STORES, mergedBy: 'client-3way' }, stores: merged };
      // v1.4.39 修复：push 前把 fileBuf(ArrayBuffer) 转成 base64 字符串，否则 JSON.stringify 会变成
      //   {type:'Buffer', data:[...]} 这种嵌套对象，pull 反序列化时丢失或变空。
      payload = _preSerializeForSync(payload);
      var json = JSON.stringify(payload);
      if (json.length > 900 * 1024) return { ok: false, error: '数据超过 900KB（GitHub 接口上限 1MB），请减少自定义模板数量后再上传' };
      var content = b64encodeUnicode(json);
      var body = { message: 'sync: 3-way merge parties+templates (conflicts=' + conflicts + ')', content: content, branch: SYNC_BRANCH };
      if (sha) body.sha = sha;
      var res = await fetch(SYNC_API, {
        method: 'PUT',
        headers: { 'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github+json', 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
      });
      if (res.ok) { await saveSyncBase({ stores: merged }); return { ok: true, merged: countRecords(merged), conflicts: conflicts }; }
      if (res.status === 409 || res.status === 422) { lastErr = '并发提交，自动重试 ' + attempt + '/5'; continue; }
      var t = await res.text();
      return { ok: false, error: '上传失败 HTTP ' + res.status + ' ' + t.slice(0, 200) };
    }
    return { ok: false, error: '上传失败（并发冲突重试超限）: ' + lastErr };
  }
  function countRecords(stores) { var n = 0; SYNC_STORES.forEach(function (s) { n += (stores[s] || []).length; }); return n; }
  // v1.5.22 自动团队同步：主数据/模板变更后自动 push 到团队库(GitHub userdata.json)，3s 防抖合并。
  // 需在设置页配置 token；无 token 静默跳过；失败仅 console 提示不阻塞操作。
  var _autoSyncTimer = null, _autoSyncBusy = false;
  async function autoSyncToTeam() {
    if (_autoSyncBusy) return;
    var c;
    try { c = await loadSyncCfg(); } catch (e) { return; }
    if (!c || !c.token) return; // 未配置 token：不自动同步
    if (_autoSyncTimer) clearTimeout(_autoSyncTimer);
    _autoSyncTimer = setTimeout(async function () {
      _autoSyncBusy = true;
      try {
        var r = await pushShared(c.token);
        if (r && r.ok) console.log('[auto-sync] 已自动同步到团队库 ' + r.merged + ' 条');
        else console.warn('[auto-sync] 自动同步失败: ' + (r && r.error || 'unknown'));
      } catch (e) { console.warn('[auto-sync] 自动同步异常: ' + e.message); }
      _autoSyncBusy = false;
    }, 3000);
  }
  // 导出自包含 HTML：把 7 个 store 全量烘焙进一个 html 副本下载（双击即用）
  async function exportSelfContainedHTML() {
    var data = {};
    for (var i = 0; i < db.STORES.length; i++) data[db.STORES[i]] = await db.all(db.STORES[i]);
    var inject = '<' + 'script>window.__USERDATA_B64__="' + b64encodeUnicode(JSON.stringify(data)) + '";<' + '/script>';
    var html;
    try {
      var res = await fetch(location.href.split('#')[0]);
      html = await res.text();
    } catch (e) { html = document.documentElement.outerHTML; }
    if (html.indexOf('<!--USERDATA_SLOT-->') >= 0) html = html.split('<!--USERDATA_SLOT-->').join(inject);
    else html = html.replace('</body>', inject + '\n</body>');
    var blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = '贸易单证系统-本地版.html';
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(a.href);
    toast('已导出自包含 HTML（数据已嵌入文件，双击即用）', 'ok');
  }
  async function exportJSON() {
    var data = {};
    for (var i = 0; i < db.STORES.length; i++) data[db.STORES[i]] = await db.all(db.STORES[i]);
    var blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = '贸易单证系统-数据备份-' + new Date().toISOString().slice(0, 10) + '.json';
    document.body.appendChild(a); a.click(); a.remove();
    URL.revokeObjectURL(a.href);
    toast('已导出 JSON 备份', 'ok');
  }
  function importJSON() {
    var inp = document.createElement('input');
    inp.type = 'file'; inp.accept = '.json,application/json';
    inp.onchange = async function () {
      var f = inp.files[0]; if (!f) return;
      try {
        var data = JSON.parse(await f.text());
        var total = 0;
        for (var i = 0; i < db.STORES.length; i++) {
          var s = db.STORES[i], list = data[s];
          if (list && list.length) { await db.bulkPut(s, list); total += list.length; }
        }
        toast('已导入 ' + total + ' 条记录（同名覆盖）', 'ok');
        render();
      } catch (e) { toast('导入失败: ' + e.message, 'err'); }
    };
    inp.click();
  }

  // ---------- 初始化 ----------
  // v1.4.49：init 顶层加 12s watchdog——超过 12s 还没到 render()（pullShared/seed 卡住）就显示红色提示卡
  // + 「跳过拉取继续」按钮（手动点立即跳过 pullShared 阶段）
  (async function init() {
    var watchdog = setTimeout(function () {
      if ($main && $main.innerHTML.indexOf('初始化') >= 0) {
        $main.innerHTML = '<div class="card vres block">' +
          '<b>初始化用时较长（>25s）</b>，通常卡在「从团队库拉取主数据」这一步（国内访问 GitHub 不稳）。' +
          '<p style="margin-top:10px"><button class="btn" id="init-skip-pull">跳过拉取，继续初始化</button> ' +
          '<button class="btn sm ghost" id="init-retry">整页重试</button></p></div>';
        var skipBtn = document.getElementById('init-skip-pull');
        if (skipBtn) skipBtn.onclick = function () { _initSkipPull = true; $main.innerHTML = '<div class="loading">正在初始化…</div>'; };
        var rBtn = document.getElementById('init-retry');
        if (rBtn) rBtn.onclick = function () { location.reload(); };
      }
    }, 25000);
    try {
      await db.open();
      await hydrateFromEmbedded();
      await _migrateBase64Templates();  // v1.4.42: 把残留 base64-string 模板还原为 ArrayBuffer
      await seed.run(db, engine, ExcelJS);
      // v1.5.42 启动自动拉取合并为一次（消除旧版三处串联拉取：首次强制/无模板自愈/auto 各拉一遍，最坏 60s+ 触发 watchdog 误跳过）。
      // 逻辑：auto 默认开（仅用户显式关过 auto===false 才停）→ 拉一次（pullShared 内部同源8s+raw15s 双通道兜底）；
      // 拉失败且本地无发票模板（新人首启）→ 可见提示引导手动重试。
      var autoFlag = (await loadSyncCfg()).auto;
      if (!_initSkipPull && autoFlag) {
        var pr = await pullShared();
        if (pr && pr.ok) {
          toast('已从团队库拉取 ' + pr.merged + ' 条主数据', 'ok');
          try { await db.put('config', { key: 'tplCleanV', value: 1459 }); } catch (e) {}
        } else {
          var tpls1 = await db.all('templates');
          var inv1 = tpls1.filter(function (t) { return t && t.kind === 'invoice' && t.status === 'active'; });
          if (inv1.length === 0) toast('团队库拉取失败（' + ((pr && pr.error) || '网络异常') + '），可在设置页点「重新同步」', 'warn');
        }
      }
      clearTimeout(watchdog);
      render();
    } catch (e) {
      clearTimeout(watchdog);
      $main.innerHTML = '<div class="card vres block">初始化失败: ' + esc(e.message) + '</div>';
      console.error(e);
    }
  })();
})();
