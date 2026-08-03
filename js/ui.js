/* L1 UI层：页签 + CRUD + 发票/订舱单双向导 */
(function () {
  'use strict';
  var db = TD.db, parser = TD.parser, validator = TD.validator, engine = TD.engine,
      exporter = TD.exporter, adapters = TD.adapters, seed = TD.seed;

  var $main = document.getElementById('main');
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
  async function loadWb(buf, fileName) {
    var fn = fileName || '';
    var ab = buf;
    var needConvert = /\.csv$/i.test(fn) || (/\.xls$/i.test(fn) && !/\.xlsx$/i.test(fn));
    if (needConvert) ab = await anyToXlsx(buf, fn);
    try {
      var wb = new ExcelJS.Workbook();
      await wb.xlsx.load(ab);
      return wb;
    } catch (err) {
      var msg = err.message || '';
      if (/zip|central directory|end of central/i.test(msg)) {
        throw new Error('文件不是有效的 .xlsx（可能是 .xls 旧格式、.csv 或文件已损坏），请用 Excel/WPS 另存为 .xlsx 后再试');
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
      '<span class="hint">解析后可直接「生成订单」——针对已在其他系统核对过的装箱清单，无需再导聚水潭订单</span></div>' +
      '<table class="grid"><tr><th>文件名</th><th>关联单号</th><th>箱数</th><th>总数量</th><th>SKU 汇总</th><th>总毛重(KG)</th><th>体积(CBM)</th><th>导入时间</th><th>操作</th></tr>' +
      (rows || '<tr><td colspan="9" class="empty">暂无装箱清单</td></tr>') + '</table></div>';
  };
  BINDERS.packings = function () {
    document.getElementById('pk-file').onchange = async function () {
      var f = this.files[0]; if (!f) return;
      try {
        var wb = await loadWb(await f.arrayBuffer(), f.name);
        var rows = parser.sheetToRows(wb.worksheets[0]);
        var hash = parser.hashRows(rows);
        var exist = await db.all('packings');
        if (exist.some(function (p) { return p.hash === hash; })) { toast('该装箱清单已导入过（内容完全相同），已拦截重复导入', 'err'); return; }
        var pk = parser.parsePacking(rows);
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
      '<label style="margin-top:8px">公司名（英文）</label><input id="pt-company" placeholder="如 JW PEI AP LIMITED">' +
      '<label style="margin-top:8px">地址（英文，可多行）</label><textarea id="pt-address" rows="2"></textarea>' +
      '<label style="margin-top:8px">城市</label><input id="pt-city">' +
      '<label style="margin-top:8px">省份/州</label><input id="pt-state">' +
      '<label style="margin-top:8px">邮编</label><input id="pt-zip">' +
      '<label style="margin-top:8px">电话</label><input id="pt-tel">' +
      '<label style="margin-top:8px">联系人</label><input id="pt-contact">' +
      '<label style="margin-top:8px">邮箱</label><input id="pt-email">' +
      '<label style="margin-top:8px">国家代码</label><input id="pt-country" placeholder="如 US / CN">' +
      '<label style="margin-top:8px">税号/EORI</label><input id="pt-taxno">' +
      '<div style="margin-top:12px;display:flex;gap:8px"><button class="btn" id="pt-save">保存</button><button class="btn ghost" id="pt-reset">清空</button></div></div></div>';
  };
  BINDERS.parties = function () {
    function resetForm() { ['pt-id', 'pt-name', 'pt-company', 'pt-address', 'pt-city', 'pt-state', 'pt-zip', 'pt-tel', 'pt-contact', 'pt-email', 'pt-country', 'pt-taxno'].forEach(function (i) { document.getElementById(i).value = ''; }); document.getElementById('pt-form-title').textContent = '新增'; }
    document.getElementById('pt-reset').onclick = resetForm;
    document.getElementById('pt-save').onclick = async function () {
      var name = val('pt-name');
      if (!name) { toast('名称必填', 'err'); return; }
      var obj = { type: val('pt-type'), name: name, company: val('pt-company'), address: val('pt-address'), city: val('pt-city'), state: val('pt-state'), zip: val('pt-zip'), tel: val('pt-tel'), contact: val('pt-contact'), email: val('pt-email'), country: val('pt-country'), taxNo: val('pt-taxno') };
      var id = val('pt-id');
      if (id) { var old = await db.get('parties', id); obj = Object.assign(old, obj); }
      await db.put('parties', obj);
      toast('已保存', 'ok'); render();
    };
    document.querySelectorAll('.pt-edit').forEach(function (b) {
      b.onclick = async function () {
        var p = await db.get('parties', b.dataset.id);
        document.getElementById('pt-id').value = p.id;
        document.getElementById('pt-type').value = p.type;
        document.getElementById('pt-name').value = p.name || '';
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
        document.getElementById('pt-form-title').textContent = '编辑: ' + p.name;
      };
    });
    document.querySelectorAll('.pt-del').forEach(function (b) {
      b.onclick = async function () {
        if (!(await confirmBox('确认删除该收发货人？', true))) return;
        await db.del('parties', b.dataset.id); toast('已删除', 'ok'); render();
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
    return '<tr data-sku="' + esc(d.sku) + '"><td class="mono">' + esc(d.sku) + '</td><td>' + esc(d.nameCn || '') + '</td><td>' + esc(d.nameEn || '') + '</td>' +
      '<td class="mono">' + esc(d.hsCode || '') + '</td>' +
      '<td class="num">' + (d.declarePrice || '') + '</td>' +
      '<td class="mono">' + esc(d.currency || 'USD') + (d.declarePriceRaw && d.currency && d.currency !== 'USD' ? ' <span class="hint" title="原币值 ' + d.declarePriceRaw + ' ' + d.currency + '">≈' + d.declarePriceRaw + '</span>' : '') + '</td>' +
      '<td>' + esc(d.material || '') + '</td>' +
      '<td>' + esc(d.brand || '') + '</td><td class="num">' + (d.nw || '') + '</td>' +
      '<td><button class="btn sm ghost dc-edit" data-sku="' + esc(d.sku) + '">编辑</button> <button class="btn sm danger dc-del" data-sku="' + esc(d.sku) + '">删除</button></td></tr>';
  }
  function declareMatches(d, f) {
    if (f.q) {
      var t = [d.sku, d.nameCn, d.nameEn, d.hsCode, d.material, d.brand].join(' ').toLowerCase();
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
      '<p class="hint">数据来源：飞书《申报信息》<b>主文档</b>（SKU主数据）+ <b>备用文档</b>（申报专用）。主文档优先，缺字段自动用备用文档同 SKU 补齐。点「☁️ 从飞书镜像同步」并入本地——仅填空字段，<b>不覆盖</b>手填项。</p>' +
      '<div class="card"><div class="toolbar">' +
      '<button class="btn" id="dc-add">＋新增SKU</button>' +
      '<label class="btn ghost" style="display:inline-block">📥 从 xlsx/xls/csv 导入<input type="file" id="dc-file" accept=".xlsx,.xls,.csv" style="display:none"></label>' +
      '<button class="btn warn" id="dc-pull">☁️ 从飞书镜像同步</button>' +
      '<input type="text" id="dc-search" class="input" placeholder="搜索 SKU / 品名 / HS编码 / 材质 / 品牌" style="width:260px;margin-left:auto">' +
      '</div>' +
      '<div id="dc-facets" class="facets"></div>' +
      '<div id="dc-count" class="hint"></div>' +
      '<table class="grid"><thead><tr><th>SKU</th><th>中文品名</th><th>英文品名</th><th>HS编码</th><th>申报价(USD)</th><th>币种</th><th>材质</th><th>品牌</th><th>单件净重</th><th>操作</th></tr></thead><tbody id="dc-tbody"></tbody></table>' +
      '<div id="dc-pager" class="pager"></div></div>';
  };
  function declareForm(d) {
    d = d || {};
    showModal('<h3>' + (d.sku ? '编辑' : '新增') + '申报要素</h3><div class="form-grid" style="margin-top:12px">' +
      '<div><label class="req">SKU</label><input id="dc-sku" value="' + esc(d.sku || '') + '"' + (d.sku ? ' readonly' : '') + '></div>' +
      '<div><label>中文品名</label><input id="dc-namecn" value="' + esc(d.nameCn || '') + '"></div>' +
      '<div><label class="req">英文品名</label><input id="dc-nameen" value="' + esc(d.nameEn || '') + '"></div>' +
      '<div><label class="req">HS编码</label><input id="dc-hs" value="' + esc(d.hsCode || '') + '"></div>' +
      '<div><label class="req">申报价(USD)</label><input id="dc-price" type="number" step="0.01" value="' + (d.declarePrice || '') + '" title="以USD计；若下方币种非USD，保存时自动折算为USD并记原币值"></div>' +
      '<div><label>币种</label><input id="dc-cur" value="' + esc(d.currency || 'USD') + '"></div>' +
      '<div><label class="req">材质</label><input id="dc-mat" value="' + esc(d.material || '') + '"></div>' +
      '<div><label>用途</label><input id="dc-usage" value="' + esc(d.usage || '') + '"></div>' +
      '<div><label>品牌</label><input id="dc-brand" value="' + esc(d.brand || 'NO BRAND') + '"></div>' +
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
        sku: sku, nameCn: val('dc-namecn'), nameEn: val('dc-nameen'), hsCode: val('dc-hs'),
        declarePrice: fxToUsd(dcRaw, dcCur), declarePriceRaw: dcRaw, currency: dcCur,
        material: val('dc-mat'), usage: val('dc-usage'), brand: val('dc-brand'),
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
        tbody.innerHTML = '<tr><td colspan="10" class="empty">无匹配记录</td></tr>';
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
      toast('正在从飞书《申报信息》镜像同步…');
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
    if (!data.length) return { ok: false, added: 0, merged: 0, error: '本地未找到飞书申报信息镜像（js/declare_data.js），请重新部署系统' };
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
      return '<tr><td>' + esc(t.name) + (t.builtin ? ' <span class="badge gray">内置</span>' : '') + '</td>' +
        '<td>' + kindBadge(t.kind) + '</td>' +
        '<td>' + esc(t.carrier || '') + '</td>' +
        '<td>' + (t.status === 'active' ? '<span class="badge green">启用</span>' : '<span class="badge red">停用</span>') + '</td>' +
        '<td>' + fmtTime(t.createdAt) + '</td>' +
        '<td><button class="btn sm ghost tp-prev" data-id="' + t.id + '">预览</button> ' +
        '<button class="btn sm ghost tp-map" data-id="' + t.id + '">字段</button> ' +
        '<button class="btn sm ' + (t.status === 'active' ? 'warn' : 'ok') + ' tp-toggle" data-id="' + t.id + '">' + (t.status === 'active' ? '停用' : '启用') + '</button> ' +
        '<button class="btn sm danger tp-del" data-id="' + t.id + '">删除</button></td></tr>';
    }).join('');
    var carrierOpts = CARRIERS.map(function (c) { return '<option>' + c + '</option>'; }).join('');
    return '<h2>模板库</h2><div class="card"><h3>上传新模板</h3><div class="form-grid">' +
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
        await db.put('templates', {
          name: name, kind: kind, carrier: carrier, status: 'active', fileBuf: buf,
          mapping: { required: engine.REQUIRED_FIELDS[kind] || [], scanned: scan }
        });
        toast('模板已上传：扫描到 ' + scan.fields.length + ' 个表头字段 + ' + scan.itemFields.length + ' 个明细字段', 'ok');
        render();
      } catch (e) { toast('模板解析失败: ' + e.message, 'err'); }
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
        showModal('<h3>字段映射 · ' + esc(t.name) + '</h3>' +
          '<h3>表头占位符 (' + scan.fields.length + ')</h3><p class="mono" style="word-break:break-all">' + scan.fields.map(esc).join('　') + '</p>' +
          '<h3>明细占位符 (' + scan.itemFields.length + ')</h3><p class="mono">' + scan.itemFields.map(esc).join('　') + '</p>' +
          '<h3>必填校验字段</h3><p>' + (t.mapping.required || []).map(function (r) { return '<span class="badge yellow">' + esc(r.label) + '</span> '; }).join('') + '</p>' +
          '<div style="text-align:right;margin-top:12px"><button class="btn" onclick="TDUI.closeModal()">关闭</button></div>');
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
    document.querySelectorAll('.tp-del').forEach(function (b) {
      b.onclick = async function () {
        var docs = await db.all('documents');
        var refs = docs.filter(function (d) { return d.templateId === b.dataset.id; });
        var msg = refs.length ? '⚠️ 该模板已被 ' + refs.length + ' 条单证记录引用，删除后这些记录将无法重新导出。确认删除？' : '确认删除该模板？';
        if (!(await confirmBox(msg, true))) return;
        if (refs.length && !(await confirmBox('二次确认：真的要删除被引用的模板吗？', true))) return;
        await db.del('templates', b.dataset.id); toast('已删除', 'ok'); render();
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
      var tpls = (await db.all('templates')).filter(function (t) { return t.kind === 'invoice' && t.status === 'active'; });
      var pkRows = pks.map(function (p) {
        var inter = (p.orderNos || []).filter(function (n) { return selNos.indexOf(n) >= 0; });
        var hint = inter.length === selNos.length && inter.length === (p.orderNos || []).length ? '<span class="badge green">完全匹配</span>' :
          inter.length ? '<span class="badge yellow">部分匹配(' + inter.length + '/' + selNos.length + ')</span>' : '<span class="badge gray">无交集</span>';
        var ck = w.packingId === p.id ? ' checked' : '';
        return '<tr class="checkrow"><td><input type="radio" name="wz-pk" value="' + p.id + '"' + ck + '></td>' +
          '<td>' + esc(p.fileName) + '</td><td class="mono">' + esc((p.orderNos || []).join(', ')) + '</td><td class="num">' + p.totals.boxCount + '</td><td class="num">' + p.totals.qty + '</td><td>' + hint + '</td></tr>';
      }).join('');
      var tplOpts = tpls.map(function (t) { return '<option value="' + t.id + '"' + (w.templateId === t.id ? ' selected' : '') + '>' + esc(t.name) + '（' + esc(t.carrier) + '）</option>'; }).join('');
      body.innerHTML = '<div class="card"><h3>① 选择装箱清单（用于单号/SKU/数量强校验及重量体积）</h3>' +
        '<table class="grid"><tr><th></th><th>文件名</th><th>关联单号</th><th>箱数</th><th>数量</th><th>与所选订单</th></tr>' +
        (pkRows || '<tr><td colspan="6" class="empty">暂无装箱清单，请先到「装箱清单」页上传</td></tr>') + '</table>' +
        '<h3 style="margin-top:16px">② 物流商 / 渠道 / 发票模板</h3><div class="form-grid">' +
        '<div><label>物流商</label><input id="wz-carrier" list="carrier-list2" value="' + esc(w.carrier) + '"><datalist id="carrier-list2">' + CARRIERS.map(function (c) { return '<option>' + c + '</option>'; }).join('') + '</datalist></div>' +
        '<div><label>渠道</label><input id="wz-channel" value="' + esc(w.channel) + '" placeholder="如 海运整柜/海快/美森"></div>' +
        '<div><label class="req">发票模板（仅显示启用中）</label><select id="wz-tpl"><option value="">请选择</option>' + tplOpts + '</select></div></div>' +
        '<div style="margin-top:14px;display:flex;gap:8px"><button class="btn ghost" id="wz-back1">← 上一步</button><button class="btn" id="wz-next1">下一步 →</button></div></div>';
      document.getElementById('wz-back1').onclick = function () { w.step = 0; render(); };
      document.getElementById('wz-next1').onclick = function () {
        var pk = document.querySelector('input[name="wz-pk"]:checked');
        if (!pk) { toast('必须选择装箱清单（发票强校验依赖装箱清单）', 'err'); return; }
        var tpl = val('wz-tpl');
        if (!tpl) { toast('请选择发票模板', 'err'); return; }
        w.packingId = pk.value; w.templateId = tpl;
        w.carrier = val('wz-carrier'); w.channel = val('wz-channel');
        w.step = 2; render();
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
        '<div><label class="req">SHIPPER 发货人</label><select id="wz-shipper">' + opts(shippers, w.shipperId) + '</select></div>' +
        '<div><label class="req">CONSIGNEE 收货人</label><select id="wz-consignee">' + opts(consignees, w.consigneeId) + '</select>' +
        (ordRecv ? '<p class="hint">📌 订单自带收货人「' + esc(ordRecv.receiver || ordRecv.buyer) + '」<button class="btn sm ghost" id="wz-use-ord-recv">直接抓取使用</button></p>' : '<p class="hint">订单无收发货人信息，请从主数据选择 <button class="btn sm ghost" id="wz-goto-parties">去「收发货人」页维护</button></p>') + '</div>' +
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
        var p = await db.put('parties', { type: 'consignee', name: ordRecv.receiver || ordRecv.buyer, address: ordRecv.address || '', tel: ordRecv.phone || '', country: ordRecv.country || '', remark: '从订单 ' + ordRecv.orderNo + ' 抓取' });
        w.consigneeId = p.id; toast('已抓取订单收货人并存入主数据', 'ok'); renderWizStep();
      };
      var gotoParties = document.getElementById('wz-goto-parties');
      if (gotoParties) gotoParties.onclick = function () {
        state.wiz = null; state.tab = 'parties';
        document.querySelectorAll('.tab').forEach(function (t) { t.classList.toggle('active', t.dataset.tab === 'parties'); });
        render();
      };
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
      var ctx = await buildWizContext(w, 'invoice');
      var tpl = await db.get('templates', w.templateId);
      var wb = await loadWb(tpl.fileBuf);
      var fillRes = engine.fillTemplate(wb, ctx.data, { logo: tpl.logo || null });
      w._wb = wb; w._data = ctx.data;
      var confirmed = w.doc && w.doc.status === 'confirmed';
      body.innerHTML = '<div class="card"><h3>发票预览（模板: ' + esc(tpl.name) + '）</h3>' +
        (fillRes.unresolved.length ? '<div class="vres warn">⚠️ 以下占位符无数据（已置空）: <span class="mono">' + fillRes.unresolved.filter(function (v, i, a) { return a.indexOf(v) === i; }).map(esc).join('　') + '</span></div>' : '') +
        '<div style="overflow:auto;border:1px solid #e3e8f0;border-radius:8px;padding:8px;background:#fafbfd">' + wbToHtml(wb) + '</div>' +
        '<div style="margin-top:14px;display:flex;gap:8px;align-items:center">' +
        '<button class="btn ghost" id="wz-back4">← 上一步</button>' +
        '<button class="btn warn" id="wz-confirm"' + (confirmed ? ' disabled' : '') + '>' + (confirmed ? '✓ 已确认' : '① 确认无误') + '</button>' +
        '<button class="btn ok" id="wz-export"' + (confirmed ? '' : ' disabled') + '>② 导出发票 xlsx</button>' +
        '<span class="hint">必须先「确认无误」才能导出（状态机: draft→validated→confirmed→exported）</span></div></div>';
      document.getElementById('wz-back4').onclick = function () { w.step = 3; w.doc = null; render(); };
      document.getElementById('wz-confirm').onclick = async function () {
        var doc = {
          kind: 'invoice', orderIds: w.orderIds.slice(), packingId: w.packingId, templateId: w.templateId,
          carrier: w.carrier, channel: w.channel, docNo: ctx.data.invoiceNo,
          data: ctx.data, status: 'draft'
        };
        if (!validator.canTransition('draft', 'validated')) { toast('状态机异常', 'err'); return; }
        doc.status = 'validated';
        if (!validator.canTransition('validated', 'confirmed')) { toast('状态机异常', 'err'); return; }
        doc.status = 'confirmed';
        w.doc = await db.put('documents', doc);
        toast('已确认，可导出', 'ok'); renderWizStep();
      };
      document.getElementById('wz-export').onclick = async function () {
        if (!w.doc || !validator.canTransition(w.doc.status, 'exported')) { toast('请先确认无误', 'err'); return; }
        var fname = exporter.safeName((w.carrier ? w.carrier + '_' : '') + ctx.data.invoiceNo) + '.xlsx';
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
    var boxMode = !!(packing && itemFields.some(function (f) { return /(boxNo|length|width|height)/.test(f); }));
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
      var tpls = (await db.all('templates')).filter(function (t) { return t.kind === 'booking' && t.status === 'active'; });
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
      document.getElementById('bw-next1').onclick = function () {
        var tpl = val('bw-tpl');
        if (!tpl) { toast('请选择订舱单模板', 'err'); return; }
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
      var ctx = await buildWizContext(w, 'booking');
      var tpl = await db.get('templates', w.templateId);
      var wb = await loadWb(tpl.fileBuf);
      var fillRes = engine.fillTemplate(wb, ctx.data, { logo: tpl.logo || null });
      w._wb = wb;
      var confirmed = w.doc && w.doc.status === 'confirmed';
      body.innerHTML = '<div class="card"><h3>订舱单预览（模板: ' + esc(tpl.name) + '）</h3>' +
        (fillRes.unresolved.length ? '<div class="vres warn">⚠️ 空占位符: <span class="mono">' + fillRes.unresolved.filter(function (v, i, a) { return a.indexOf(v) === i; }).map(esc).join('　') + '</span></div>' : '') +
        '<div style="overflow:auto;border:1px solid #e3e8f0;border-radius:8px;padding:8px;background:#fafbfd">' + wbToHtml(wb) + '</div>' +
        '<div style="margin-top:14px;display:flex;gap:8px;align-items:center">' +
        '<button class="btn ghost" id="bw-back4">← 上一步</button>' +
        '<button class="btn warn" id="bw-confirm"' + (confirmed ? ' disabled' : '') + '>' + (confirmed ? '✓ 已确认' : '① 确认无误') + '</button>' +
        '<button class="btn ok" id="bw-export"' + (confirmed ? '' : ' disabled') + '>② 导出订舱单 xlsx</button></div></div>';
      document.getElementById('bw-back4').onclick = function () { w.step = 3; w.doc = null; render(); };
      document.getElementById('bw-confirm').onclick = async function () {
        var doc = {
          kind: 'booking', orderIds: w.orderIds.slice(), packingId: w.packingId, templateId: w.templateId,
          carrier: w.carrier, docNo: ctx.data.invoiceNo, data: ctx.data, status: 'confirmed'
        };
        w.doc = await db.put('documents', doc);
        toast('已确认，可导出', 'ok'); renderBookingStep();
      };
      document.getElementById('bw-export').onclick = async function () {
        if (!w.doc || !validator.canTransition(w.doc.status, 'exported')) { toast('请先确认无误', 'err'); return; }
        var fname = exporter.safeName((w.carrier ? w.carrier + '_' : '') + 'BOOKING_' + ctx.data.invoiceNo) + '.xlsx';
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
      var wb = await loadWb(tpl.fileBuf);
      engine.fillTemplate(wb, d.data, { logo: tpl.logo || null });
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
      '<div class="card"><h3>⚠️ 数据管理</h3>' +
      '<button class="btn danger" id="st-wipe">清空全部本地数据（不可恢复）</button></div>' +
      '<footer class="note">贸易单证系统 · 纯前端本地存储(IndexedDB) · 七层解耦架构 · 模板占位符引擎</footer>';
  };
  BINDERS.settings = function () {
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
      toast('正在从飞书《申报信息》镜像同步…');
      var r = await syncDeclaresFromMirror();
      if (r.ok) toast('✅ 申报信息同步完成：新增 ' + r.added + ' 条、补齐 ' + r.merged + ' 条', 'ok');
      else toast('同步失败: ' + r.error, 'err');
    };
    document.getElementById('st-save-jst').onclick = async function () {
      await db.put('config', { key: 'jst', value: { appKey: val('st-jst-key'), appSecret: val('st-jst-secret') } });
      toast('聚水潭凭证已保存（API拉取功能待开通）', 'ok');
    };
    document.getElementById('st-wipe').onclick = async function () {
      if (!(await confirmBox('⚠️ 将清空订单/装箱清单/主数据/模板/单证记录全部本地数据，且不可恢复！确认？', true))) return;
      if (!(await confirmBox('二次确认：真的要清空全部数据吗？', true))) return;
      for (var i = 0; i < db.STORES.length; i++) await db.clear(db.STORES[i]);
      toast('已清空，正在重新初始化…', 'ok');
      await seed.run(db, engine, ExcelJS);
      render();
    };
  };

  // ---------- 初始化 ----------
  (async function init() {
    try {
      await db.open();
      await seed.run(db, engine, ExcelJS);
      render();
    } catch (e) {
      $main.innerHTML = '<div class="card vres block">初始化失败: ' + esc(e.message) + '</div>';
      console.error(e);
    }
  })();
})();
