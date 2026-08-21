/* L5 生成引擎：数据组装 + ExcelJS模板占位符填充（{{field}} / {{items.xxx}} 明细行复制） */
(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) module.exports = factory();
  else { root.TD = root.TD || {}; root.TD.engine = factory(); }
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var PH_RE = /\{\{\s*([A-Za-z0-9_.]+)\s*\}\}/g;

  // 取单元格纯文本（富文本合并为字符串；数字/其他类型返回 ''，调用方需自行处理数字）
  function cellString(cell) {
    var v = cell && cell.value;
    return (v && v.richText) ? v.richText.map(function (t) { return t.text; }).join('') : (typeof v === 'string' ? v : '');
  }

  /** 模板明细表头别名词典：把模板里真实列名映射到数据字段 */
  var HEADER_ALIASES = {
    no:      ['NO.', 'NO', '序号', '#', '项次'],
    boxNo:   ['CTNR NO', 'CTNR NO.', 'CTNRNO', 'CTNR', 'CARTON NO', 'CARTON#', '箱号', 'CARTON', 'CTN#', '货箱编号', 'FBA箱号'],
    boxCount:['CTNS', 'CARTONS', '箱数', '件数', 'TYPES OF PKG', 'TYPESOFPKG', 'NO. AND TYPES OF PKG'],
    sku:     ['SKU', '款号', '商品编码', '商家编码', '货号', '规格编码', 'MODEL NO', 'MODELNO'],
    model:   ['MODEL', '型号', 'MODEL NO.', 'MODEL NO'],
    nameEn:  ['DESCRIPTION', 'DESCRIPTION OF GOODS', 'GOODS DESCRIPTION', '英文品名', '品名(英文)', '品名'],
    nameCn:  ['中文品名', '品名(中文)', '商品名称', '中文名称'],
    hsCode:  ['HS CODE', 'HSCODE', 'HS编码', '海关编码', 'TARIFF CODE'],
    qty:     ['QTY', '数量', '件数', 'PCS', 'QUANTITY', 'TOTAL QUANTITY'],
    unit:    ['UNIT', '单位', 'UOM'],
    price:   ['UNIT PRICE', '单价', 'PRICE', 'DECLARED VALUE', '采购价格', '采购价'],
    amount:  ['AMOUNT', '金额', 'TOTAL', 'TOTAL AMOUNT', 'TOTAL PRICE', '总申报价值', '申报价值'],
    nw:      ['N.W', 'N.W.', 'NW', 'NET WEIGHT', '净重'],
    gw:      ['G.W', 'G.W.', 'GW', 'GROSS WEIGHT', '毛重', '货箱重量'], // v1.5.40 货箱重量=毛重（用户要求全模板通用）
    singleGw:['SINGLE GW', 'PER CTN G.W', 'WEIGHT/CTN', '单箱重量', '单箱毛重', 'WEIGHT PER CTN', '单箱重', '单箱重(KG)'],
    volume:  ['CBM', 'M3', 'VOL', 'MEAS', '体积', '尺码', 'MEAS\'T'],
    material:['MATERIAL', '材质', '质地'],
    usage:   ['用途', 'USE', 'USAGE', '用途说明', 'PURPOSE'],
    brand:   ['BRAND', '品牌'],
    origin:  ['ORIGIN', '原产地', '原产国', '产地', 'COUNTRY OF ORIGIN'],
    destCountry: ['目的国', '目的国家', 'DEST COUNTRY', 'DESTINATION', '目的港'], // v1.5.28 明细列头「目的国」→ items.destCountry（数据来自目的港国家）
    tradeCountry: ['贸易国', '贸易国（地区）', '贸易国家', 'TRADING COUNTRY', 'TRADE COUNTRY'], // v1.5.31 明细列头「贸易国（地区）」→ items.tradeCountry（=收货人国家中文名）
    tradeTerm: ['成交方式', '贸易条款', '成交条件', 'INCOTERMS', 'TRADE TERMS'], // v1.5.29 明细列头「成交方式 CIF/FOB」→ items.tradeTerm（=meta.incoterms）
    brandType: ['品牌类型', 'BRAND TYPE', '品牌类型（海关申报）'], // v1.5.32 明细列头「品牌类型」→ items.brandType（有品牌=4 境外品牌-其他）
    exportPrefer: ['出口享惠情况', '享惠情况', 'EXPORT PREFERENCE'], // v1.5.32 明细列头「出口享惠情况」→ items.exportPrefer（按目的国协定自动）
    dims: ['箱规', '装箱尺寸', '外箱尺寸', 'BOX SIZE', 'CARTON SIZE'], // 装箱尺寸（来自装箱单长宽高）
    productDims: ['产品尺寸', '尺寸', 'SIZE', 'DIMENSION', 'DIMS', '规格尺寸', '长×宽×高'], // v1.5.33 产品本身尺寸（SKU_DIMS，来源 JW PEI G Unit Q列「长高宽cm」）
    // v1.4.43 海运类模板「长/宽/高/产品图片/产品性质/备注」等 per-item 字段
    // 注意：alias 不再放单字母 'L'/'W'/'H'（normalizeHeader 去空格后会被任意含该字母的词误命中）
    lengthCm:['LENGTH', 'L(CM)', '长', '长CM'],
    widthCm: ['WIDTH',  'W(CM)', '宽', '宽CM'],
    heightCm:['HEIGHT', 'H(CM)', '高', '高CM'],
    imageUrl:['IMAGE', 'PHOTO', 'PICTURE', '图片', '产品图片', '商品图片'],
    // v1.4.48：中运通达等模板中英双语表头「Purpose(用途)/产品电池类型/是否带电(Y/N)/是否带磁(Y/N)/
    //          产品销售价格/图片链接/单个产品净重/币种」此前因 alias 缺失而整列空（已被 scanTemplate 兜底选中
    //          为明细表头，但列未映射任何字段 → 不写值）。补齐下列中英别名。
    productNature:['NATURE', 'PRODUCT NATURE', '性质', '产品性质'],
    batteryType:['BATTERY TYPE', 'POWERED', '电池类型', '产品电池类型', '电池'],
    electrified:['ELECTRIFIED', '是否带电', '带电(Y/N)', '带电'],
    magnetic:['MAGNETIC', '是否带磁', '带磁(Y/N)', '带磁'],
    productPrice:['产品价格', '产品销售价格', '销售价格', '零售价', 'RETAIL PRICE', 'SALE PRICE', 'UNIT PRICE'],
    productLink:['产品链接', '图片链接', '商品链接', 'LINK', 'URL', 'PRODUCT LINK'],
    singleNw:['单个产品净重', '单品净重', 'NET WT/PC', 'NET WEIGHT/PC'],
    currency:['CURRENCY', '币种', '货币'],
    remark:  ['REMARK', 'NOTES', 'NOTE', '备注', '说明']
  };

  function normalizeHeader(s) {
    return String(s || '').replace(/\s+/g, '').replace(/[（(].*?[)）]/g, '').replace(/[*※]+$/, '').toUpperCase();
  }
  function matchHeaderAlias(s) {
    var ns = normalizeHeader(s);
    if (!ns) return null;
    var keys = Object.keys(HEADER_ALIASES);
    var best = null, bestLen = 0;
    for (var i = 0; i < keys.length; i++) {
      var al = HEADER_ALIASES[keys[i]];
      for (var j = 0; j < al.length; j++) {
        var a = normalizeHeader(al[j]);
        if (!a) continue;
        // 完全匹配优先级最高，立刻返回
        if (ns === a) return keys[i];
        // 子串匹配时，只接受以单词边界开头/结尾的，避免 "UNIT PRICE" 被 "UNIT" 截胡
        var hit = ns.indexOf(a) >= 0 || a.indexOf(ns) >= 0;
        if (hit && a.length > bestLen) { best = keys[i]; bestLen = a.length; }
      }
    }
    return best;
  }

  function getPath(obj, path) {
    var parts = String(path).split('.');
    var cur = obj;
    for (var i = 0; i < parts.length; i++) {
      if (cur === null || cur === undefined) return undefined;
      cur = cur[parts[i]];
    }
    return cur;
  }
  function round(n, d) { var p = Math.pow(10, d); return Math.round((Number(n) || 0) * p) / p; }

  // v1.5.29 原产国显示：CN/CHN/CHINA → 「中国」（用户发票明细「原产国」列要求显示中文），空值也默认中国
  function _originLabel(v) {
    var s = String(v === undefined || v === null ? '' : v).trim().toUpperCase();
    if (!s || s === 'CN' || s === 'CHN' || s === 'CHINA' || s === 'CHINESE' || s === 'P.R.C') return '中国';
    return v;
  }

  // v1.5.31 目的港（pod）→ 目的国（中文国家名）。
  // 报关术语纠正：贸易国（地区）=收货人国家；目的国（运抵国）=目的港所在国家，与收货人无关。
  // pod 形式如 'US' / 'USLAX' / 'NEW YORK US' / '香港 HONG KONG HK'——取国家码（前2字母大写）查表，未命中原样返回。
  var _POD_COUNTRY = {
    US:'美国', USA:'美国', AMERICA:'美国',
    UK:'英国', GB:'英国', ENGLAND:'英国',
    CA:'加拿大', CANADA:'加拿大',
    AU:'澳大利亚', AUSTRALIA:'澳大利亚', NZ:'新西兰',
    HK:'中国香港', HONGKONG:'中国香港', 'HONG KONG':'中国香港', MO:'中国澳门', MACAU:'中国澳门', TW:'中国台湾', TAIWAN:'中国台湾', CHINA:'中国', CN:'中国', PRC:'中国',
    JP:'日本', JAPAN:'日本', KR:'韩国', KOREA:'韩国', KP:'朝鲜', MN:'蒙古', TH:'泰国', THAILAND:'泰国', VN:'越南', VIETNAM:'越南', MY:'马来西亚', MALAYSIA:'马来西亚',
    SG:'新加坡', SINGAPORE:'新加坡', ID:'印度尼西亚', INDONESIA:'印度尼西亚', PH:'菲律宾', PHILIPPINES:'菲律宾', MM:'缅甸', BURMA:'缅甸', KH:'柬埔寨', LA:'老挝', BN:'文莱',
    IN:'印度', INDIA:'印度', PK:'巴基斯坦', BD:'孟加拉', LK:'斯里兰卡', NP:'尼泊尔',
    DE:'德国', GERMANY:'德国', FR:'法国', FRANCE:'法国', IT:'意大利', ITALY:'意大利', ES:'西班牙', SPAIN:'西班牙', PT:'葡萄牙', NL:'荷兰', NETHERLANDS:'荷兰', BE:'比利时',
    AT:'奥地利', CH:'瑞士', SWITZERLAND:'瑞士', SE:'瑞典', NO:'挪威', DK:'丹麦', FI:'芬兰', PL:'波兰',
    CZ:'捷克', HU:'匈牙利', GR:'希腊', IE:'爱尔兰', RU:'俄罗斯', RUSSIA:'俄罗斯', UA:'乌克兰', TR:'土耳其', TURKEY:'土耳其',
    AE:'阿联酋', SA:'沙特阿拉伯', IL:'以色列', EG:'埃及', EGYPT:'埃及', ZA:'南非', KE:'肯尼亚',
    MX:'墨西哥', MEXICO:'墨西哥', BR:'巴西', BRAZIL:'巴西', AR:'阿根廷', CL:'智利', CO:'哥伦比亚', PE:'秘鲁',
    MA:'摩洛哥', NG:'尼日利亚', ET:'埃塞俄比亚'
  };
  function _podCountry(pod) {
    var p = String(pod === undefined || pod === null ? '' : pod).toUpperCase().trim();
    if (!p) return '';
    // v1.5.31：从 pod 中提取所有大写单词（如 'SHENZHEN CHINA' → ['SHENZHEN','CHINA']），按顺序查表（优先匹配末尾词）
    var words = p.match(/\b[A-Z]{2,}\b/g) || [];
    // 倒序匹配：优先匹配 pod 末尾的国家词（'SHENZHEN CHINA' → CHINA；'NEW YORK US' → US；'HONG KONG HK' → HK）
    for (var i = words.length - 1; i >= 0; i--) {
      var w = words[i];
      if (_POD_COUNTRY[w]) return _POD_COUNTRY[w];
      // 截前2字母作为国家码兜底（USLAX→US）
      var cc2 = w.slice(0, 2);
      if (_POD_COUNTRY[cc2]) return _POD_COUNTRY[cc2];
    }
    return pod || '';
  }

  // v1.5.32 出口享惠情况：按目的国（中文名）匹配中国签署的自贸协定。未命中默认「无」。
  // 说明：享惠需目的国与中国签有协定且货物符合原产地规则，系统仅按目的国给出申报建议，declareMap 显式值优先。
  var _EXPORT_PREFER = {
    '美国':'无','英国':'无','德国':'无','法国':'无','意大利':'无','西班牙':'无','葡萄牙':'无','荷兰':'无','比利时':'无',
    '奥地利':'无','瑞士':'无','瑞典':'无','挪威':'无','丹麦':'无','芬兰':'无','波兰':'无','捷克':'无','匈牙利':'无',
    '希腊':'无','爱尔兰':'无','俄罗斯':'无','乌克兰':'无','土耳其':'无','阿联酋':'无','沙特阿拉伯':'无','以色列':'无',
    '埃及':'无','南非':'无','肯尼亚':'无','墨西哥':'无','巴西':'无','阿根廷':'无','哥伦比亚':'无','印度':'无','尼泊尔':'无',
    '加拿大':'无','中国台湾':'无','朝鲜':'无','蒙古':'无','摩洛哥':'无','尼日利亚':'无','埃塞俄比亚':'无',
    '智利':'中国-智利自贸协定','秘鲁':'中国-秘鲁自贸协定','巴基斯坦':'中国-巴基斯坦自贸协定',
    '孟加拉':'亚太贸易协定','斯里兰卡':'亚太贸易协定','老挝':'亚太贸易协定',
    '澳大利亚':'中国-澳大利亚自贸协定','新西兰':'中国-新西兰自贸协定',
    '中国香港':'CEPA','中国澳门':'CEPA',
    '日本':'RCEP','韩国':'中国-韩国自贸协定',
    '泰国':'中国-东盟自贸区','越南':'中国-东盟自贸区','马来西亚':'中国-东盟自贸区','印度尼西亚':'中国-东盟自贸区',
    '菲律宾':'中国-东盟自贸区','缅甸':'中国-东盟自贸区','柬埔寨':'中国-东盟自贸区','文莱':'中国-东盟自贸区',
    '新加坡':'中国-新加坡自贸协定'
  };
  function _exportPreferLabel(destCountry) {
    var s = String(destCountry === undefined || destCountry === null ? '' : destCountry).trim();
    if (!s) return '';
    return _EXPORT_PREFER[s] !== undefined ? _EXPORT_PREFER[s] : '无';
  }

  /** 金额大写（美元英文） */
  function amountInWords(n, currency) {
    currency = currency || 'USD';
    var ones = ['', 'ONE', 'TWO', 'THREE', 'FOUR', 'FIVE', 'SIX', 'SEVEN', 'EIGHT', 'NINE', 'TEN', 'ELEVEN', 'TWELVE', 'THIRTEEN', 'FOURTEEN', 'FIFTEEN', 'SIXTEEN', 'SEVENTEEN', 'EIGHTEEN', 'NINETEEN'];
    var tens = ['', '', 'TWENTY', 'THIRTY', 'FORTY', 'FIFTY', 'SIXTY', 'SEVENTY', 'EIGHTY', 'NINETY'];
    function three(num) {
      var s = '';
      if (num >= 100) { s += ones[Math.floor(num / 100)] + ' HUNDRED'; num %= 100; if (num) s += ' AND '; }
      if (num >= 20) { s += tens[Math.floor(num / 10)]; num %= 10; if (num) s += '-' + ones[num]; }
      else if (num > 0) s += ones[num];
      return s;
    }
    var int = Math.floor(Math.abs(n)), cents = Math.round((Math.abs(n) - int) * 100);
    if (int === 0 && cents === 0) return 'ZERO ' + currency;
    var parts = [], units = ['', ' THOUSAND', ' MILLION', ' BILLION'], ui = 0;
    while (int > 0) { var seg = int % 1000; if (seg) parts.unshift(three(seg) + units[ui]); int = Math.floor(int / 1000); ui++; }
    var s = 'SAY ' + (currency === 'USD' ? 'US DOLLARS ' : currency + ' ') + (parts.join(' ') || 'ZERO');
    s += cents ? ' AND CENTS ' + three(cents) + ' ONLY' : ' ONLY';
    return s;
  }

  /**
   * 组装单证数据：orders(多单合并) + packing + meta(表头字段) + parties + declareMap
   * 返回统一 data：header字段 + items[] + totals
   */
  function buildDocData(opts) {
    var orders = opts.orders || [], packing = opts.packing || null;
    var meta = opts.meta || {}, shipper = opts.shipper || {}, consignee = opts.consignee || {}, notify = opts.notify || null;
    var declareMap = opts.declareMap || {};

    // 订单号列表
    var orderList = [];
    orders.forEach(function (o) { orderList.push(o.orderNo); });

    // 明细：方案A——模板需要箱号/尺寸(boxNo/length/width/height)且关联了装箱清单时，按「箱×SKU」逐行展开
    var boxMode = !!(opts.boxMode && packing && (packing.boxes || []).length);

    // 前置：按 SKU 聚合订单价格/币种/参考号
    var priceBySku = {}, currencyBySku = {}, orderSkuQty = {};
    orders.forEach(function (o) {
      (o.items || []).forEach(function (it) {
        var sku = String(it.sku).trim();
        var p = Number(it.price);
        if (p > 0 && priceBySku[sku] === undefined) priceBySku[sku] = p;
        if (it.currency && currencyBySku[sku] === undefined) currencyBySku[sku] = it.currency;
        if (!orderSkuQty[sku]) orderSkuQty[sku] = 0;
        orderSkuQty[sku] += Number(it.qty) || 0;
      });
    });
    var refId = meta.refId || (orderList.length ? orderList[0] : '');
    var destCountry = _podCountry(meta.pod) || (consignee && consignee.country) || '';
    // v1.5.31 贸易国（地区）=收货人所在国家（中文名）：与目的国（目的港国家）语义分离
    var tradeCountry = _podCountry(consignee && consignee.country) || (consignee && consignee.country) || '';
    var tradeTerm = meta.incoterms || '';

    // 按 SKU 聚合装箱信息（箱数、总净毛重、体积、第一箱尺寸）
    var boxAgg = {};
    if (packing) {
      (packing.boxes || []).forEach(function (b) {
        var sku = String(b.sku == null ? '' : b.sku).trim();
        if (!sku) return;
        if (!boxAgg[sku]) boxAgg[sku] = { count: 0, qty: 0, nw: 0, gw: 0, volume: 0, volumeWeight: 0, firstDims: '', firstBoxSpec: '' };
        var L = Number(b.length) || 0, W = Number(b.width) || 0, H = Number(b.height) || 0;
        var vol = (L && W && H) ? round(L * W * H / 1000000, 6) : 0;
        var volW = (L && W && H) ? round(L * W * H / 6000, 3) : 0;
        var dims = (L && W && H) ? (L + '×' + W + '×' + H) : '';
        boxAgg[sku].count += 1;
        boxAgg[sku].qty += Number(b.qty) || 0;
        boxAgg[sku].nw += Number(b.nw) || 0;
        boxAgg[sku].gw += Number(b.gw) || 0;
        boxAgg[sku].volume += vol;
        boxAgg[sku].volumeWeight += volW;
        if (!boxAgg[sku].firstDims && dims) { boxAgg[sku].firstDims = dims; boxAgg[sku].firstBoxSpec = dims; }
      });
      // v1.5.30 混箱箱数去重：构建「每箱首个 SKU」集合。
      // 混箱（一箱多 SKU）时 boxAgg[sku].count 会把同一箱同时计给每个 SKU（A/B 同箱各 count=1），
      // 若每行都写箱数 1 → 总箱数虚高。仅让「该箱首个 SKU」的行显示箱数，其余行留空。
      var _boxSkus = {};
      (packing.boxes || []).forEach(function (b2) {
        var sku2 = String(b2.sku == null ? '' : b2.sku).trim();
        var bk2 = String(b2.orderNo == null ? '' : b2.orderNo) + '/' + String(b2.boxNo == null ? '' : b2.boxNo).trim();
        if (!sku2 || bk2 === '/') return;
        var lst = _boxSkus[bk2] || (_boxSkus[bk2] = []);
        if (lst.indexOf(sku2) < 0) lst.push(sku2);
      });
      var _firstSkuSet = {};
      Object.keys(_boxSkus).forEach(function (k) { var f = _boxSkus[k][0]; if (f) _firstSkuSet[f] = true; });
    }

    function makeItemCore(sku, d, itR) {
      itR = itR || {};
      var price = (priceBySku[sku] !== undefined) ? priceBySku[sku] : (d.declarePrice || 0);
      var currency = currencyBySku[sku] || d.currency || meta.currency || 'USD';
      // v1.4.43: 回退到订单 item 的字段（declareMap 缺字段时不再丢品名/材质等）
      //          同时新增 lengthCm/widthCm/heightCm/singleGw/imageUrl/remark 出口供明细 per-item 列填充
      function _pick(k) { return (d[k] !== undefined && d[k] !== '') ? d[k] : (itR[k] !== undefined && itR[k] !== '' ? itR[k] : ''); }
      return {
        sku: sku,
        model: _pick('model') || sku,
        nameEn: _pick('nameEn'), nameCn: _pick('nameCn'),
        hsCode: _pick('hsCode'), material: _pick('material'), usage: _pick('usage'), brand: _pick('brand'),
        lengthCm: _pick('lengthCm') || _pick('length'),
        widthCm:  _pick('widthCm')  || _pick('width'),
        heightCm: _pick('heightCm') || _pick('height'),
        singleGw: _pick('singleGw'),
        imageUrl: _pick('imageUrl') || _pick('image') || _pick('photo'),
        remark:   _pick('remark')   || _pick('note'),
        productNature: _pick('productNature') || _pick('batteryType') || _pick('nature'),
        productPrice: d.productPrice || d.retailPrice || d.salePrice || itR.productPrice || '',
        productLink: d.productLink || d.link || d.url || itR.productLink || '',
        singleNw: d.singleNw || d.netWeightPerPc || d.nwPerPc || itR.singleNw || '',
        price: price, unit: d.unit || itR.unit || 'PCS', origin: _originLabel(d.origin || itR.origin || 'CN'),
        electrified: (d.electrified != null && d.electrified !== '') ? d.electrified : (itR.electrified != null ? itR.electrified : '否'),
        magnetic: (d.magnetic != null && d.magnetic !== '') ? d.magnetic : (itR.magnetic != null ? itR.magnetic : '否'),
        asin: d.asin || itR.asin || '', fnsku: d.fnsku || itR.fnsku || sku, note: d.note || itR.note || '', costPrice: d.costPrice || itR.costPrice || '',
        currency: currency, destCountry: destCountry, tradeCountry: tradeCountry, refId: refId,
        condition: d.condition || itR.condition || 'NEW', exportPrefer: d.exportPrefer || itR.exportPrefer || _exportPreferLabel(destCountry), tradeTerm: tradeTerm,
        // v1.5.32 品牌类型：有品牌(如 JW PEI)=境外品牌（其他）→4；无品牌→0。declareMap 显式 brandType 优先。
        brandType: d.brandType || itR.brandType || ((d.brand || itR.brand) ? '4' : '0'),
        // v1.5.33 产品本身尺寸（长*宽*高 cm）：declareMap 显式优先 → SKU_DIMS(源 JW PEI G Unit Q列) 兜底；与 dims(装箱尺寸) 语义分离
        productDims: d.productDims || itR.productDims || ((typeof window !== 'undefined' && window.SKU_DIMS && window.SKU_DIMS[sku]) || ''),
        batteryType: d.batteryType || itR.batteryType || '',
        taxNo: d.taxNo || itR.taxNo || '',
        shippingMarks: '', poNo: refId, manufacturer: d.manufacturer || itR.manufacturer || ''
      };
    }

    // ===== v1.4.60 明细排序：多订单合并时严格按「订单顺序 → 箱号数值顺序 → 装箱清单原始行序」=====
    // 业务规则：订单A(1..5箱) + 订单B(1..8箱) 合并成一张发票时，必须先完整列出 A 的 1-5 箱，
    // 再列 B 的 1-8 箱。旧实现用 boxNo 字符串比较且不分订单，会出现 A1,B1,A2,B2... 或 1,10,11,2 的乱序。
    var _ordSeq = {}, _ordN = 0;
    orderList.forEach(function (o) {
      var k = String(o == null ? '' : o).trim();
      if (_ordSeq[k] === undefined) _ordSeq[k] = _ordN++;
    });
    // 装箱清单里出现但不在订单列表中的订单号：按其在清单中首次出现顺序，排在已知订单之后
    (packing && packing.boxes || []).forEach(function (b) {
      var k = String(b.orderNo == null ? '' : b.orderNo).trim();
      if (_ordSeq[k] === undefined) _ordSeq[k] = 1000 + (_ordN++);
    });
    function _ordRank(orderNo) {
      var k = String(orderNo == null ? '' : orderNo).trim();
      var v = _ordSeq[k];
      return (v === undefined) ? 99999 : v;
    }
    /** 箱号数值化：'1'→1, 'B2'→2, '箱-10'→10, 'A-1'→1；无数字返回 NaN */
    function _boxNum(v) {
      var m = String(v == null ? '' : v).match(/(\d+(?:\.\d+)?)/);
      return m ? parseFloat(m[1]) : NaN;
    }
    /** 装箱行比较器：订单序 → 箱号数值 → 箱号原文 → 清单原始行序 */
    function _boxRowCmp(a, b) {
      var ao = _ordRank(a.orderNo), bo = _ordRank(b.orderNo);
      if (ao !== bo) return ao - bo;
      var an = _boxNum(a.boxNo), bn = _boxNum(b.boxNo);
      var aNaN = isNaN(an), bNaN = isNaN(bn);
      if (aNaN !== bNaN) return aNaN ? 1 : -1;       // 无数字箱号排最后
      if (!aNaN && an !== bn) return an - bn;         // 数值升序：1,2,...,9,10,11
      var as = String(a.boxNo == null ? '' : a.boxNo), bs = String(b.boxNo == null ? '' : b.boxNo);
      if (as !== bs) return as < bs ? -1 : 1;         // 数值相同（如 A1/B1）按原文
      return (a._ri || 0) - (b._ri || 0);             // 同订单同箱多 SKU：保持清单原始行顺序
    }

    var items = [];
    if (boxMode) {
      // 箱级规格归并：装箱清单只在「每箱首行」写 重量/长/宽/高，需按 订单号/箱号 聚合到该箱全部 SKU 行，
      // 否则非首行 SKU 的 单箱重量/长/宽/高 全空（模板「单箱重量(kg)」「长(cm)」等列按箱×SKU 展开后只有首行有值）。
      // 注意：仅用于明细「单箱重量/尺寸」展示列；it.gw/it.nw 仍保持「每行 per-SKU 装箱值」(首行有、其余0)，
      // 这样下方 totals 按行累加 = 各箱首行重量之和 = 总毛重，不会被箱重在多行重复放大。
      var boxSpecMap = {};
      (packing.boxes || []).forEach(function (b0) {
        var bk = String(b0.orderNo == null ? '' : b0.orderNo) + '/' + String(b0.boxNo == null ? '' : b0.boxNo).trim();
        if (bk === '/') bk = '@row' + Math.random();
        if (!boxSpecMap[bk]) boxSpecMap[bk] = { gw: 0, nw: 0, length: '', width: '', height: '', _set: false };
        if (!boxSpecMap[bk]._set && (Number(b0.gw) || Number(b0.length))) {
          boxSpecMap[bk].gw = Number(b0.gw) || 0;
          boxSpecMap[bk].nw = Number(b0.nw) || 0;
          boxSpecMap[bk].length = b0.length || '';
          boxSpecMap[bk].width = b0.width || '';
          boxSpecMap[bk].height = b0.height || '';
          boxSpecMap[bk]._set = true;
        }
      });
      // v1.4.65：boxMode 品名回退——装箱单只有 SKU，品名权威源=declareMap(申报表)；
      // 但 declareMap 缺该 SKU 时回退订单 items 的 name/nameEn/nameCn（聚水潭导入订单带 name），
      // 避免模板有品名列却留空（用户「模板里有的都得填」原则；与 v1.4.61 goodsSummary 回退一致）。
      var orderItemBySku = {};
      orders.forEach(function (o) {
        (o.items || []).forEach(function (it) {
          var sk = String(it.sku == null ? '' : it.sku).trim();
          if (!sk || orderItemBySku[sk]) return;
          var r = {};
          Object.keys(it).forEach(function (k) { r[k] = it[k]; });
          if (!r.nameEn && it.name) r.nameEn = it.name;
          if (!r.nameCn && it.name) r.nameCn = it.name;
          orderItemBySku[sk] = r;
        });
      });
      items = (packing.boxes || []).map(function (b, _ri) {
        var sku = String(b.sku == null ? '' : b.sku).trim();
        var d = declareMap[sku] || {};
        var bk = String(b.orderNo == null ? '' : b.orderNo) + '/' + String(b.boxNo == null ? '' : b.boxNo).trim();
        var spec = boxSpecMap[bk] || {};
        var L = Number(spec.length) || Number(b.length) || 0, W = Number(spec.width) || Number(b.width) || 0, H = Number(spec.height) || Number(b.height) || 0;
        var dims = (L && W && H) ? (L + '×' + W + '×' + H) : (b.boxSpec || '');
        var volW = (L && W && H) ? round(L * W * H / 6000, 3) : (b.volumeWeight || 0);
        var vol = (L && W && H) ? round(L * W * H / 1000000, 6) : 0;
        var it = makeItemCore(sku, d, orderItemBySku[sku]);
        it.boxNo = b.boxNo || ''; it.boxSpec = dims; it.dims = dims;
        it.length = L || ''; it.width = W || ''; it.height = H || '';
        // v1.4.47：boxMode 直接把箱规同步给 per-item 长/宽/高/单箱重字段（模板列映射的是 lengthCm/... 而非 length）
        // v1.4.56：箱级规格按 订单号/箱号 归并，非首行 SKU 也拿到整箱重量/尺寸
        it.lengthCm = L || ''; it.widthCm = W || ''; it.heightCm = H || '';
        it.singleGw = (spec.gw || Number(b.gw)) || 0;
        it.volume = vol; it.volumeWeight = volW;
        it.boxCount = 1; it.ctns = 1;
        it.qty = Number(b.qty) || 0; it.amount = round(it.qty * it.price, 2);
        // v1.4.58：总箱单个产品数量 = 单箱数量 × 箱数（boxMode 每行 = 1箱×SKU，boxCount=1）
        it.totalBoxQty = round(it.qty * (it.boxCount || 1), 0);
        // 保持 per-SKU 装箱值（首行有、其余0），供 totals 正确累加总毛重；不为单箱重量列污染
        it.nw = Number(b.nw) || 0; it.gw = Number(b.gw) || 0; it.boxNw = it.nw;
        it._weightSource = 'packing';
        // v1.4.60：保留排序键（订单号 + 装箱清单原始行序）
        it.orderNo = String(b.orderNo == null ? '' : b.orderNo).trim();
        it._ri = _ri;
        return it;
      });
      // v1.4.60：订单顺序 → 箱号数值顺序 → 清单原始行序（A的1..5箱全列完再列B的1..8箱）
      items.sort(_boxRowCmp);
      // v1.4.62：标记每箱首行（箱数仅首行显示，避免混装 N 行被误读为 N 箱）
      var _seenBox = {};
      items.forEach(function (it) {
        var bk = String(it.orderNo || '') + '/' + String(it.boxNo || '');
        it._firstOfBox = !_seenBox[bk];
        _seenBox[bk] = true;
      });
    } else {
      // 明细合并（跨订单同SKU同价合并）
      var agg = {};
      orders.forEach(function (o) {
        (o.items || []).forEach(function (it) {
          var sku = String(it.sku).trim();
          var d = declareMap[sku] || {};
          var price = (it.price !== undefined && it.price !== null && Number(it.price) > 0) ? Number(it.price) : (d.declarePrice || 0);
          var key = sku + '@' + price;
          if (!agg[key]) {
            agg[key] = makeItemCore(sku, d, it);
            agg[key].qty = 0; agg[key].amount = 0; agg[key].nw = 0; agg[key].gw = 0;
            agg[key]._priceSource = (it.price && Number(it.price) > 0) ? 'order' : (d.declarePrice ? 'master' : 'none');
          }
          agg[key].qty += Number(it.qty) || 0;
        });
      });
      items = Object.keys(agg).map(function (k) { return agg[k]; });
      // v1.4.60：非 boxMode（按 SKU 聚合）也不再用字典序，改为「SKU 在装箱清单箱号顺序中首次出现」；
      // 无装箱清单时退化为「订单录入顺序中首次出现」。保证多订单合并时明细顺序与装箱清单一致。
      var _skuSeq = {}, _sn = 0;
      if (packing && (packing.boxes || []).length) {
        (packing.boxes || []).map(function (b, i) { return { orderNo: b.orderNo, boxNo: b.boxNo, sku: b.sku, _ri: i }; })
          .sort(_boxRowCmp)
          .forEach(function (b) {
            var k = String(b.sku == null ? '' : b.sku).trim();
            if (k && _skuSeq[k] === undefined) _skuSeq[k] = _sn++;
          });
      }
      orders.forEach(function (o) {
        (o.items || []).forEach(function (it) {
          var k = String(it.sku == null ? '' : it.sku).trim();
          if (k && _skuSeq[k] === undefined) _skuSeq[k] = 1000 + (_sn++);
        });
      });
      items.sort(function (a, b) {
        var ax = _skuSeq[a.sku], bx = _skuSeq[b.sku];
        if (ax === undefined) ax = 99999;
        if (bx === undefined) bx = 99999;
        if (ax !== bx) return ax - bx;
        return a.sku < b.sku ? -1 : 1;
      });
    }

    // 非箱模式：从装箱清单按SKU分摊净毛重，并补充箱数/体积/箱规
    var pkAgg = {};
    if (packing && !boxMode) {
      (packing.boxes || []).forEach(function (b) {
        var k = String(b.sku).trim();
        if (!pkAgg[k]) pkAgg[k] = { qty: 0, nw: 0, gw: 0 };
        pkAgg[k].qty += b.qty; pkAgg[k].nw += b.nw || 0; pkAgg[k].gw += b.gw || 0;
      });
    }
    var totals = { qty: 0, amount: 0, nw: 0, gw: 0, boxCount: 0, volume: 0, volumeWeight: 0 };
    items.forEach(function (it, idx) {
      it.no = idx + 1;
      it.amount = round(it.qty * it.price, 2);
      var sku = String(it.sku).trim();
      var ba = boxAgg[sku];
      if (!boxMode) {
        var pk = pkAgg[sku];
        var d = declareMap[sku] || {};
        if (pk && pk.qty > 0) {
          // 装箱清单有该 SKU：优先用清单毛净重量；若清单无「毛重/净重」列（值为 0/空），回退到主数据（申报信息）重量，避免 G.W./N.W. 变 0
          it.nw = pk.nw ? round(pk.nw * (it.qty / pk.qty), 3) : round((d.nw || 0) * it.qty, 3);
          it.gw = pk.gw ? round(pk.gw * (it.qty / pk.qty), 3) : round((d.gw || 0) * it.qty, 3);
          it._weightSource = (pk.gw || pk.nw) ? 'packing' : 'master';
        } else {
          it.nw = round((d.nw || 0) * it.qty, 3);
          it.gw = round((d.gw || 0) * it.qty, 3);
          it._weightSource = d.nw ? 'master' : 'none';
        }
        if (ba) {
          it.boxCount = ba.count; it.ctns = ba.count;
          // v1.5.30 混箱箱数去重：该 SKU 从未作为任何箱的首行（只出现在混箱非首行）→ 箱数留空，避免每行都写 1 虚高
          if (packing && packing.boxes && packing.boxes.length && _firstSkuSet && !_firstSkuSet[sku]) { it.boxCount = ''; it.ctns = ''; }
          it.totalBoxQty = round(it.qty * (ba.count || 1), 0); // v1.4.58 总件数=单箱数量×箱数
          it.dims = ba.firstDims; it.boxSpec = ba.firstBoxSpec;
          it.volume = round(ba.volume, 6); it.volumeWeight = round(ba.volumeWeight, 3);
          it.boxNw = ba.count ? round(ba.nw / ba.count, 3) : 0;
          // v1.4.47：非 boxMode（纯标签驱动模板未触发 boxMode）时，从第一箱尺寸兜底补 per-item 长/宽/高/单箱重，
          // 让「长(cm)/宽(cm)/高(cm)/单箱重量」等列有值可填（避免整列空白）
          if (!boxMode) {
            var _m = String(ba.firstDims || '').match(/(\d+(?:\.\d+)?)\s*[×xX*]\s*(\d+(?:\.\d+)?)\s*[×xX*]\s*(\d+(?:\.\d+)?)/);
            if (_m) { it.lengthCm = Number(_m[1]); it.widthCm = Number(_m[2]); it.heightCm = Number(_m[3]); }
            it.singleGw = ba.count ? round(ba.gw / ba.count, 3) : 0;
          }
        }
      } else {
        it._weightSource = 'packing';
      }
      totals.qty += it.qty; totals.amount += it.amount; totals.nw += it.nw; totals.gw += it.gw; totals.volumeWeight += (Number(it.volumeWeight) || 0);
    });
    totals.amount = round(totals.amount, 2); totals.nw = round(totals.nw, 3); totals.gw = round(totals.gw, 3); totals.volumeWeight = round(totals.volumeWeight, 3);
    if (packing) {
      totals.boxCount = packing.totals.boxCount;
      // v1.4.52：混装去重——毛重/体积/体积重是「箱」级属性，按唯一箱号聚合（同箱多SKU只计一次），避免重复统计
      // 修正 v1.4.52 回归：仅当箱单确实含「毛重」列（箱级数据存在）时才用箱级聚合覆盖；
      // 否则（重量仅来自主数据 per-unit）保留上面已按明细汇总的值（本已正确处理混装，避免被清零成 0）
      if (packing.boxes && packing.boxes.length) {
        var _bt = { gw: 0, volume: 0, volumeWeight: 0 }, _seen = {}, _ri = 0, _hasBoxGw = false;
        (packing.boxes || []).forEach(function (b) {
          var k = String(b.orderNo == null ? '' : b.orderNo) + '/' + String(b.boxNo == null ? '' : b.boxNo).trim();
          if (k === '/') k = '@row' + (_ri++); // 无订单号也无箱号时按行独立，不跨行合并
          if (_seen[k]) return; _seen[k] = 1;
          var _g = Number(b.gw) || 0; if (_g > 0) _hasBoxGw = true;
          _bt.gw += _g;
          var L = Number(b.length) || 0, W = Number(b.width) || 0, H = Number(b.height) || 0;
          if (L && W && H) { _bt.volume += L * W * H / 1000000; _bt.volumeWeight += L * W * H / 6000; }
          else { _bt.volume += Number(b.volume) || 0; _bt.volumeWeight += Number(b.volumeWeight) || 0; }
        });
        if (_hasBoxGw) {
          totals.gw = round(_bt.gw, 3);
          totals.volume = round(_bt.volume, 4);
          totals.volumeWeight = round(_bt.volumeWeight, 3);
        }
      }
      if (packing.totals.nw) totals.nw = packing.totals.nw;
    }

    var currency = meta.currency || 'USD';
    // 给收发人补全常用字段默认值，避免模板占位符 unresolved
    // v1.4.62：地址保守解析——仅当结构化字段为空、地址有值时，抽取高置信度国家/邮编/城市回填。
    // 省份/州等自由文本易误判，一律留空不猜（避免海关单据编造错误数据）。
    function parseAddress(addr) {
      var out = { city: '', state: '', zip: '', country: '' };
      if (!addr || typeof addr !== 'string') return out;
      var s = addr.toUpperCase();
      // v1.4.64：国家匹配修正——①长名优先且「命中即停」：地址 "HONG KONG, CHINA" 只取 HK，不被后面的 CHINA 覆盖成 CN
      // ②短码（HK/US/CN/UK/FR…）必须单词边界 \b：杜绝 "CAUSEWAY"→US、"BUSAN"→US 的子串误判
      //   （用户真实发票 B7=US 的真正根因 = "CAUSEWAY BAY" 香港地名里的 "US" 被误判为国家）
      var longMap = [
        ['HONG KONG', 'HK'], ['HONGKONG', 'HK'], ['MACAU', 'MO'], ['MACAO', 'MO'], ['CHINA', 'CN'],
        ['UNITED STATES OF AMERICA', 'US'], ['UNITED STATES', 'US'], ['U.S.A.', 'US'], ['U.S.A', 'US'], ['USA', 'US'],
        ['UNITED KINGDOM', 'GB'], ['U.K.', 'GB'], ['U.K', 'GB'], ['ENGLAND', 'GB'], ['SCOTLAND', 'GB'], ['GREAT BRITAIN', 'GB'],
        ['JAPAN', 'JP'], ['GERMANY', 'DE'], ['DEUTSCHLAND', 'DE'], ['FRANCE', 'FR'], ['AUSTRALIA', 'AU'],
        ['CANADA', 'CA'], ['SINGAPORE', 'SG'], ['MALAYSIA', 'MY'], ['REPUBLIC OF KOREA', 'KR'], ['KOREA', 'KR'],
        ['TAIWAN', 'TW'], ['THAILAND', 'TH'], ['VIETNAM', 'VN'], ['INDIA', 'IN'], ['MEXICO', 'MX'], ['BRAZIL', 'BR']
      ];
      for (var i = 0; i < longMap.length; i++) {
        if (s.indexOf(longMap[i][0]) >= 0) { out.country = longMap[i][1]; break; }
      }
      if (!out.country) {
        var shortMap = [['HK', 'HK'], ['CN', 'CN'], ['US', 'US'], ['UK', 'GB'], ['GB', 'GB'], ['DE', 'DE'], ['FR', 'FR'], ['JP', 'JP'], ['AU', 'AU'], ['CA', 'CA'], ['SG', 'SG'], ['MY', 'MY'], ['KR', 'KR'], ['TW', 'TW'], ['TH', 'TH'], ['VN', 'VN'], ['IN', 'IN'], ['MX', 'MX'], ['BR', 'BR']];
        for (var j = 0; j < shortMap.length; j++) {
          if (new RegExp('\\b' + shortMap[j][0] + '\\b').test(s)) { out.country = shortMap[j][1]; break; }
        }
      }
      var zm = addr.match(/\b(\d{5,6}|\d{5}-\d{4})\b/);
      if (zm) out.zip = zm[1];
      var cityMap = ['HONG KONG', 'HONGKONG', 'SHENZHEN', 'GUANGZHOU', 'SHANGHAI', 'YIWU', 'NINGBO', 'TOKYO', 'OSAKA',
        'LOS ANGELES', 'NEW YORK', 'NEW JERSEY', 'NEWARK', 'CHICAGO', 'DALLAS', 'HOUSTON', 'MIAMI', 'SEATTLE', 'ATLANTA',
        'LONDON', 'MANCHESTER', 'SYDNEY', 'MELBOURNE', 'SINGAPORE', 'KUALA LUMPUR', 'BUSAN', 'SEOUL', 'TAIPEI', 'BANGKOK', 'HANOI'];
      cityMap.forEach(function (c) { if (s.indexOf(c) >= 0) out.city = c.replace('HONGKONG', 'HONG KONG'); });
      return out;
    }
    function fillPartyDefaults(p) {
      p = p || {};
      ['name', 'company', 'warehouseCode', 'address', 'city', 'state', 'zip', 'country', 'tel', 'email', 'contact', 'taxNo', 'vatNo', 'eori'].forEach(function (k) { if (p[k] === undefined) p[k] = ''; });
      // v1.4.62 地址保守解析；v1.4.64 用户选 B：地址里解析出的高置信国家「权威覆盖」已填值
      //（海关单据国家必须与地址一致，冲突时以地址为准，杜绝"地址 HONGKONG / 国家 US"打架退单）；
      // city/zip 仍仅填空字段（城市/邮编易多义，不覆盖用户已填）。
      if (p.address) {
        var pa = parseAddress(p.address);
        if (pa.country) p.country = pa.country;
        if (!p.zip && pa.zip) p.zip = pa.zip;
        if (!p.city && pa.city) p.city = pa.city;
      }
      // v1.4.60：VAT / EORI 拆成独立字段后，老档案只填了「税号/EORI」一个框 → 兼容回退，避免原本能填的格变空
      if (!p.vatNo && p.taxNo) p.vatNo = p.taxNo;
      if (!p.eori && p.taxNo) p.eori = p.taxNo;
      if (!p.taxNo && (p.vatNo || p.eori)) p.taxNo = p.vatNo || p.eori;
      return p;
    }
    shipper = fillPartyDefaults(shipper);
    consignee = fillPartyDefaults(consignee);
    notify = notify ? fillPartyDefaults(notify) : { name: 'SAME AS CONSIGNEE', address: '' };
    return {
      kind: opts.kind || 'invoice',
      invoiceNo: meta.invoiceNo || '',
      invoiceDate: meta.invoiceDate || new Date().toISOString().slice(0, 10),
      contractNo: meta.contractNo || '',
      orderNos: orderList.join(', '),
      refId: refId,
      shipper: shipper, consignee: consignee, notify: notify,
      incoterms: meta.incoterms || '',
      paymentTerms: meta.paymentTerms || '',
      transport: meta.transport || 'BY SEA',
      pol: meta.pol || '', pod: meta.pod || '',
      tradeCountry: tradeCountry, // v1.5.31 贸易国（地区）=收货人所在国家（中文名，报关术语）
      etd: meta.etd || '', vessel: meta.vessel || '',
      containerType: meta.containerType || '', containerQty: meta.containerQty || '',
      freightTerms: meta.freightTerms || '',
      shippingMarks: meta.shippingMarks || 'N/M',
      currency: currency,
      remark: meta.remark || '',
      dangerous: meta.dangerous || 'NON-DANGEROUS / GENERAL CARGO',
      customsType: meta.customsType || '',
      agent: meta.agent || '',
      // v1.4.61：品名汇总优先用明细 items；boxMode 下明细来自箱单（declareMap 缺品名时为空），
      //          回退到订单 items 的品名，确保「模板里有的字段都填进去」。
      goodsSummary: meta.goodsSummary || (function () {
        var _names = [];
        function _add(n) { if (n && _names.indexOf(n) < 0) _names.push(n); }
        (items || []).forEach(function (i) { _add(i.nameEn || i.nameCn); });
        if (!_names.length) { (orders || []).forEach(function (o) { (o.items || []).forEach(function (it) { _add(it.nameEn || it.nameCn); }); }); }
        return _names.slice(0, 3).join(', ');
      })(),
      items: items,
      totals: totals,
      amountInWords: amountInWords(totals.amount, currency)
    };
  }

  // ---------- 模板扫描 ----------
  /** 扫描明细行上方的表头，返回 {colNumber -> dataField} */
  function scanItemHeader(ws, itemsRow, mergedMaps) {
    var map = {};
    if (!ws || itemsRow <= 1) return map;
    var sub = (mergedMaps && mergedMaps.subordinate) || {};
    // 向上扫最多 3 行，取每个列「最靠近明细行」那一行的表头（低行覆盖高行）。
    // v1.4.57 修复：原先「首次命中即定」会让更靠上的表头标签（如 E17 投保币种→currency）
    // 抢占下方真正的明细表头（如 E18 货箱高度→heightCm），导致按表头兜底时把币种错填进高度列。
    // 改为「低行覆盖」——明细表头(紧邻明细行上方)优先，避免被更上方的单据表头标签劫持。
    for (var r = Math.max(1, itemsRow - 3); r < itemsRow; r++) {
      var row = ws.getRow(r);
      row.eachCell({ includeEmpty: false }, function (cell, colNumber) {
        if (sub[r + ',' + colNumber]) return; // 跳过合并从属格（B36:C36 合并的 C36 列等），避免重复映射字段
        var v = cell.value;
        var s = (v && v.richText) ? v.richText.map(function (t) { return t.text; }).join('') : (typeof v === 'string' ? v : '');
        var f = matchHeaderAlias(s);
        if (f) map[colNumber] = f; // 低行(更靠近明细)覆盖高行
      });
    }
    return map;
  }
  /** 扫描模板占位符，并尝试用表头识别兜底：返回 {fields, itemFields, itemsRow, itemHeaderMap, sheetName} */
  function scanTemplate(wb) {
    var result = { fields: [], itemFields: [], itemsRow: -1, itemHeaderMap: {}, sheetName: '' };
    var ws = wb.worksheets[0];
    if (!ws) return result;
    result.sheetName = ws.name;
    var seen = {}, seenItem = {};
    ws.eachRow({ includeEmpty: false }, function (row, rowNumber) {
      row.eachCell({ includeEmpty: false }, function (cell) {
        var v = cell.value;
        var s = (v && v.richText) ? v.richText.map(function (t) { return t.text; }).join('') : (typeof v === 'string' ? v : '');
        if (!s) return;
        var m; PH_RE.lastIndex = 0;
        while ((m = PH_RE.exec(s))) {
          var p = m[1];
          if (p.indexOf('items.') === 0) {
            if (result.itemsRow === -1) result.itemsRow = rowNumber;
            if (!seenItem[p]) { seenItem[p] = 1; result.itemFields.push(p); }
          } else if (!seen[p]) { seen[p] = 1; result.fields.push(p); }
        }
      });
    });
    // 兜底：没有占位符明细行时，识别真实表头区域（>=3 个可识别列头）
    if (result.itemsRow === -1) {
      var best = { row: -1, count: 0, map: {} };
      ws.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 30) return;
        var map = {}, count = 0;
        row.eachCell({ includeEmpty: false }, function (cell, colNumber) {
          var v = cell.value;
          var s = (v && v.richText) ? v.richText.map(function (t) { return t.text; }).join('') : (typeof v === 'string' ? v : '');
          var f = matchHeaderAlias(s);
          if (f && !map[colNumber]) { map[colNumber] = f; count++; }
        });
        if (count > best.count) best = { row: rowNumber, count: count, map: map };
      });
      if (best.count >= 3) {
        result.itemsRow = best.row + 1;
        result.itemHeaderMap = best.map;
      }
    } else {
      result.itemHeaderMap = scanItemHeader(ws, result.itemsRow);
    }
    return result;
  }

  /** 模板表头字符串读取（模块级，供表头识别填充使用） */
  function _cellStr(cell) {
    if (!cell) return '';
    var v = cell.value;
    if (!v) return '';
    if (v.richText) return v.richText.map(function (t) { return t.text; }).join('');
    if (typeof v === 'string') return v;
    if (v.text !== undefined) return v.text;
    if (v.formula !== undefined) return v.result !== undefined ? String(v.result) : '';
    if (v.sharedString !== undefined) return String(v.sharedString);
    return String(v);
  }

  /** 解析模板表头标签 -> {party, field, line}；非收发人字段标签返回 null */
  var PARTY_RE = [
    // v1.5.26 「我司/我方/开票方」=发货人(开票方)；「贵司」=收货人
    { re: /(发件人|发货人|托运人|我司|我方|开票方|开票人|开票单位|SHIPPER|SELLER)/i, party: 'shipper' },
    { re: /(收货人|收货|收件人|收件|贵司|CONSIGNEE|BUYER)/i, party: 'consignee' },
    { re: /(通知人|NOTIFY)/i, party: 'notify' }
  ];
  /** v1.4.60 排除词：这些是「商品/银行/注册/销售」类标签，绝不能被当成收发人或地址字段。
   *  典型误伤：中运通达 G20「Products Name中文品名」曾命中 /NAME/ → consignee.name；
   *           A18「VAT注册地址」、T20「销售地址」曾命中 /地址/ → consignee.address。 */
  var LABEL_EXCLUDE_RE = /(品名|货名|品目|商品名|产品名|货物名称|PRODUCTS?\s*NAME|NAME\s*OF\s*(GOODS|COMMODITY|PRODUCT)|销售|注册|开户|银行|BANK|规格|材质|型号|用途|品牌|海关编码|HS\s*CODE|申报要素)/i;

  function mapHeaderLabel(text) {
    if (!text || /\{\{/.test(text)) return null;
    if (LABEL_EXCLUDE_RE.test(text)) return null; // v1.4.60：商品/银行/注册类标签不进 party 通道
    var party = null;
    for (var i = 0; i < PARTY_RE.length; i++) if (PARTY_RE[i].re.test(text)) { party = PARTY_RE[i].party; break; }
    if (!party) return null;
    var field = null, line = 0;
    if (/(公司|COMPANY|企业)/i.test(text)) field = 'company';
    else if (/(邮箱|EMAIL|邮件|E-?MAIL)/i.test(text)) field = 'email';
    else if (/(电话|手机|TEL|PHONE|MOBILE)/i.test(text)) field = 'tel';
    // v1.4.60：EORI / VAT / 税号 拆成三个独立字段（原先全归 taxNo，导致 VAT号 与 EORI 两格填同一个值、eori 成死字段）
    else if (/EORI/i.test(text)) field = 'eori';
    else if (/(VAT|增值税)/i.test(text)) field = 'vatNo';
    else if (/(税号|TAX)/i.test(text)) field = 'taxNo';
    else if (/(国家|国别|COUNTRY)/i.test(text)) field = 'country';
    else if (/(邮编|ZIP|POSTAL?)/i.test(text)) field = 'zip';
    else if (/(省|州|STATE|PROVINCE)/i.test(text)) field = 'state';
    else if (/(城市|CITY)/i.test(text)) field = 'city';
    else if (/(地址编码|地址库编码)/i.test(text)) field = 'warehouseCode';
    else if (/地址/i.test(text)) {
      field = 'address';
      var m = text.match(/地址([一二三四五六七八九十\d])/);
      if (m) { var cn = { '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10 }; line = cn[m[1]] || parseInt(m[1], 10) || 1; }
    }
    else if (/(联系人|CONTACT)/i.test(text)) field = 'contact';
    else if (/全称/.test(text)) field = 'name'; // v1.5.26：「我司全称/单位全称」=名称（注：「公司全称」优先被 company 命中）
    else if (/(姓名|NAME)/i.test(text)) field = 'name';
    if (!field) return null;
    return { party: party, field: field, line: line };
  }

  /** 通用（非收发人）字段标签 -> 数据路径。用于真实承运商模板里「运输方式 / 起运国 / 客户单号」等
   *  未被 mapHeaderLabel(只认收发人三类) 覆盖的表头字段。path='' 表示已知但本系统无对应字段(跳过)。 */
  var GENERAL_LABEL_RULES = [
    // ---- v1.4.60 最前置：明细列标签误入表头区的排除（这些是 per-item 列，不是单据字段） ----
    { re: /总箱单?个?产品数量|单箱数量|每箱数量|单箱产品数量/, path: '' },
    // ---- v1.4.60 新增：原先规则表完全没有、导致用户录入被静默丢弃的单据字段 ----
    { re: /(发票号|發票號|INVOICE\s*(NO\b|NUMBER|#))/i, path: 'invoiceNo' },
    { re: /(发票日期|开票日期|INVOICE\s*DATE)/i, path: 'invoiceDate' },
    { re: /(合同号|合約號|CONTRACT\s*(NO\b|NUMBER))/i, path: 'contractNo' },
    { re: /(付款方式|支付方式|付款条款|PAYMENT\s*(TERMS?|METHOD))/i, path: 'paymentTerms' },
    { re: /(唛头|唛号|嘜頭|SHIPPING\s*MARKS?|MARKS?\s*(&|AND)\s*(NOS?|NUMBERS?))/i, path: 'shippingMarks' },
    { re: /(船名|航次|VESSEL|VOYAGE)/i, path: 'vessel' },
    { re: /(ETD|开船日期?|预计离港|离港日期)/i, path: 'etd' },
    { re: /(箱型|柜型|CONTAINER\s*TYPE)/i, path: 'containerType' },
    { re: /(箱量|柜量|柜数|CONTAINER\s*(QTY|QUANTITY))/i, path: 'containerQty' },
    { re: /(交货条款|贸易条款|价格条款|成交条款|INCOTERMS?|TRADE\s*TERMS?|DELIVERY\s*TERMS?)/i, path: 'incoterms' },
    // 进口商 = 收货人（云途 B2B 模板用语）。税号规则须在名称规则之前，否则「进口商税号」被名称截胡
    { re: /进口商\s*(税号|VAT|EORI)/i, path: 'consignee.vatNo' },
    { re: /进口商\s*(名称|公司)?/, path: 'consignee.company' },
    { re: /运输方式/, path: 'transport' },
    { re: /起运(国|港|地)/, path: 'pol' },
    { re: /目的(国|港|地区|地)/, path: 'pod' },
    { re: /贸易国/, path: 'tradeCountry' }, // v1.5.31 贸易国（地区）=收货人国家（固定 consignee.country 中文名，不随区块重定向）
    { re: /(报关|清关)方式/, path: 'customsType' },
    { re: /成交方式/, path: 'incoterms' },
    { re: /(客户订单号|客户单号|订单号|PO\s*NUMBER|PO号|P\.O\.)/i, path: 'orderNos' },
    { re: /参考号/, path: 'refId' },
    { re: /(预计总件数|总包装件数|预计件数|箱数)/, path: 'totals.boxCount' },
    { re: /(预计重量|总毛重|总重)/, path: 'totals.gw' },
    // v1.5.40 货箱重量=毛重：表头字段「货箱重量」也填 totals.gw（用户要求全模板通用；明细列头由 HEADER_ALIASES gw 处理）
    { re: /货箱重量/, path: 'totals.gw' },
    { re: /(预计体积|总体积)/, path: 'totals.volume' },
    { re: /(申报总价值|总申报价值)/, path: 'totals.amount' },
    { re: /(货物品名|商品品名|^品名$)/, path: 'goodsSummary' },
    { re: /备注/, path: 'remark' },
    // v1.4.60：EORI / VAT / 税号 三分（原先 VAT号 与 EORI 同映射到 taxNo，两格填同值）
    { re: /EORI/i, path: 'consignee.eori', partyScoped: true },
    { re: /(VAT|增值税)/i, path: 'consignee.vatNo', partyScoped: true },
    { re: /(税号|TAX\s*(ID|NO))/i, path: 'consignee.taxNo', partyScoped: true },
    { re: /仓库代码|WAREHOUSE\s*CODE|海外仓代码|仓库编号|FBA代码/, path: 'consignee.warehouseCode', partyScoped: true },
    // v1.4.57：亚丰模板「地址库编码」(FBA 货件地址库编码，无 party 前缀) → consignee.warehouseCode
    { re: /地址库编码/, path: 'consignee.warehouseCode', partyScoped: true },
    // v1.4.50：补"国家/省/城市/邮编/地址/电话/邮箱/联系人/姓名/公司"等纯字段标签（无"发件人/收货人"等 party 词时也能识别）
    // v1.4.60：加 partyScoped —— 由 buildLabelMap 按「所在区块最近的 party 词」重定向，
    //          否则带 NOTIFY 区块的模板会把通知人的国家/电话错填成收货人的值。
    { re: /国家|国别|COUNTRY/i, path: 'consignee.country', partyScoped: true },
    { re: /邮编|ZIP|POSTAL?/i, path: 'consignee.zip', partyScoped: true },
    { re: /省|州|STATE|PROVINCE/i, path: 'consignee.state', partyScoped: true },
    { re: /城市|CITY/i, path: 'consignee.city', partyScoped: true },
    { re: /地址/i, path: 'consignee.address', partyScoped: true },
    { re: /电话|手机|TEL|PHONE|MOBILE/i, path: 'consignee.tel', partyScoped: true },
    { re: /邮箱|EMAIL|邮件|E-?MAIL/i, path: 'consignee.email', partyScoped: true },
    { re: /联系人|CONTACT/i, path: 'consignee.contact', partyScoped: true },
    { re: /姓名|NAME/i, path: 'consignee.name', partyScoped: true },
    { re: /公司|COMPANY|企业/i, path: 'consignee.company', partyScoped: true },
    { re: /(揽货渠道|客户渠道|服务渠道|服务$)/, path: '' } // 无对应字段，跳过
  ];

  /** 统一字段标签识别：先试收发人三类(mapHeaderLabel)，再试通用字段词典。
   *  返回 {party, field, path, line, kind} 或 null。 */
  function mapFieldLabel(text, ctxParty) {
    if (!text || /\{\{/.test(text)) return null;
    var p = mapHeaderLabel(text);
    if (p) return { party: p.party, field: p.field, path: p.party + '.' + p.field, line: p.line, kind: 'party' };
    var t = String(text).trim();
    for (var i = 0; i < GENERAL_LABEL_RULES.length; i++) {
      var rule = GENERAL_LABEL_RULES[i];
      if (!rule.re.test(t)) continue;
      var path = rule.path;
      if (!path) return null; // 跳过规则：模板有此标签但系统无对应字段
      // v1.4.60 排除词：纯字段兜底通道（partyScoped）不得吃下「品名/销售地址/VAT注册地址/开户银行」等
      if (rule.partyScoped && LABEL_EXCLUDE_RE.test(t)) return null;
      // v1.4.60 区块就近归属：NOTIFY 区块内的「国家/电话/地址」归 notify，不再无脑归 consignee
      var party = null, field = null;
      if (rule.partyScoped && path.indexOf('consignee.') === 0) {
        var tgt = (ctxParty === 'shipper' || ctxParty === 'notify') ? ctxParty : 'consignee';
        field = path.slice('consignee.'.length);
        party = tgt;
        path = tgt + '.' + field;
      }
      return { party: party, field: field, path: path, line: 0, kind: 'general' };
    }
    return null;
  }

  /** 扫描模板表头区，生成标签->字段映射(labelMap)，供填充与 UI 编辑。
   *  itemsRowNum：明细表头行号，其之上的表头区才是字段标签所在；传 -1 时默认扫前 25 行。 */
  function buildLabelMap(wb, itemsRowNum) {
    var ws = wb.worksheets[0];
    if (!ws) return [];
    var mergedMaps = buildMergedMaps(ws);
    var map = [];
    var end = (itemsRowNum && itemsRowNum !== -1) ? itemsRowNum - 1 : 25;
    end = Math.min(end, 60);
    var seen = {};
    var curParty = '', curPartyRow = 0; // v1.4.60 区块就近归属状态
    for (var r = 1; r <= end; r++) {
      var row = ws.getRow(r);
      // v1.4.58：跳过「明细表头行」——该行有 >=3 个可 matchHeaderAlias 识别的明细列头
      //（如亚丰「货箱编号/中文品名/型号/数量/长宽高」），它们不是单据字段标签。
      // 若不跳过，「箱数*」「总箱单个产品数量*」「总申报价值（USD）」会被 GENERAL_LABEL_RULES
      // 误判为 totals.boxCount / totals.amount 标签，fillByFieldLabels 把 totals 值写进
      // 标签右侧第一个空格 → 覆盖相邻明细表头（亚丰 K27/M27/O27 被 13/13/2869.6 破坏）。
      var hdrCnt = 0;
      row.eachCell({ includeEmpty: false }, function (cellH) {
        if (matchHeaderAlias(_cellStr(cellH))) hdrCnt++;
      });
      if (hdrCnt >= 3) continue;
      // v1.4.60 区块归属：本行若出现「发件人/收件人/通知人」等 party 词，则其后（8 行内）的
      // 纯字段标签（国家/电话/邮编…）归属该 party。超过 8 行视为脱离区块，回落 consignee。
      row.eachCell({ includeEmpty: false }, function (cellP) {
        var sp = _cellStr(cellP);
        if (!sp || /\{\{/.test(sp) || LABEL_EXCLUDE_RE.test(sp)) return;
        for (var pi = 0; pi < PARTY_RE.length; pi++) {
          if (PARTY_RE[pi].re.test(sp)) { curParty = PARTY_RE[pi].party; curPartyRow = r; return; }
        }
      });
      if (curParty && (r - curPartyRow) > 8) { curParty = ''; curPartyRow = 0; }
      row.eachCell({ includeEmpty: false }, function (cell, c) {
        // 跳过合并从属格：同一合并块只在其主格处理一次，避免把值填进标签跨度破坏版式
        if (mergedMaps.masterOf[r + ',' + c]) return;
        var s = _cellStr(cell);
        if (!s || /\{\{/.test(s)) return;
        // v1.4.65：跳过装饰文字——真实字段标签（地址/邮箱/电话/公司电话等）通常 2-12 字；
        // 装饰/警告/提示文字（"因亚马逊地址时常变动..." "禁止私自修改..." 等）通常 20+ 字，
        // 后者若含"地址"等关键词会被 mapFieldLabel 误识别为字段标签，并因合并主格 mr.right 跨度大
        // 触发 fillByFieldLabels 把值写到合并区外的独立格子（V15 等模板装饰区）造成污染
        if (s.length > 20) return;
        var info = mapFieldLabel(s, curParty);
        if (!info) return;
        var key = r + ',' + c;
        if (seen[key]) return;
        seen[key] = 1;
        map.push({ r: r, c: c, label: s, path: info.path, party: info.party, field: info.field, line: info.line, kind: info.kind, resolved: !!info.path });
      });
    }
    return map;
  }

  /** 按标签映射填充表头区字段：对每个 resolved 标签，找到其右侧值格(跳过合并标签跨度/下一个标签)写入。
   *  与 fillHeaderByLabels 互补——它只处理收发人三类，此处处理全部(含收发人，幂等不冲突)。 */
  function fillByFieldLabels(ws, data, labelMap, mergedMaps) {
    if (!labelMap || !labelMap.length) return;
    var labelCols = {};
    labelMap.forEach(function (e) { (labelCols[e.r] = labelCols[e.r] || {})[e.c] = true; });
    var PH_TEST = /\{\{\s*([A-Za-z0-9_.]+)\s*\}\}/;
    labelMap.forEach(function (e) {
      if (!e.resolved || !e.path) return;
      var val = getPath(data, e.path);
      if (val === undefined || val === null || val === '') {
        if (e.party && e.field === 'address' && e.line > 0 && data[e.party] && data[e.party].address) {
          var lines = String(data[e.party].address).split(/\r?\n/);
          val = lines[e.line - 1] || '';
        }
        if (val === undefined || val === null || val === '') return;
      }
      var spanEnd = e.c;
      var mr = mergedMaps.masterRange[e.r + ',' + e.c];
      if (mr) spanEnd = Math.max(spanEnd, mr.right);
      var vStart = spanEnd + 1;
      var target = null;
      for (var cc = vStart; cc <= vStart + 12; cc++) {
        if (labelCols[e.r] && labelCols[e.r][cc]) break; // 撞到下一个标签，停止
        // v1.4.63：跳过合并从属格——避免 PASS1 清空合并区从属格后 fillByFieldLabels 误把值写进从属格，
        //          触发 ExcelJS 联动把整个合并区（H1:U14 等模板装饰大区）同步成该值
        if (mergedMaps && mergedMaps.subordinate && mergedMaps.subordinate[e.r + ',' + cc]) continue;
        var tc = ws.getRow(e.r).getCell(cc);
        var ts = _cellStr(tc);
        if (ts) { var m = ts.match(PH_TEST); if (m && m[1] === e.path) { target = tc; break; } }
      }
      // v1.4.63：fallback 写 vStart 也必须跳合并从属格（合并区只有主格可写）
      if (!target && !(labelCols[e.r] && labelCols[e.r][vStart]) && !(mergedMaps && mergedMaps.subordinate && mergedMaps.subordinate[e.r + ',' + vStart])) target = ws.getRow(e.r).getCell(vStart);
      if (!target) return;
      target.value = (typeof val === 'number') ? val : String(val);
    });
  }

  /** 模板表头识别填充：把发货人/收货人/地址/电话按模板自身表头标签写入对应单元格。
   *  两遍：先清理值列里的样本残留，再按标签写入（避免写入后被清理）。仅写入空单元格或同字段占位符。 */
  function fillHeaderByLabels(ws, data, headerEnd, itemsRowNumArg, writeEnd) {
    headerEnd = headerEnd || 25;
    // v1.5.27：PASS2 写值范围独立于 headerEnd。无占位符模板的 itemsRow 常被 scanTemplate 兜底
    // 提前判到 R4-6（明细表头行），headerEnd=itemsRow 导致 R19-23 的「我司全称/地址/联系人」等
    // 发货人信息区被跳过不填。写值只写右侧空格、不清任何东西，扩展到 writeEnd(默认25) 安全。
    writeEnd = writeEnd || headerEnd;
    // 值列：表头区（明细行之上）内包含 {{}} 占位符的列。
    // 关键：明细占位行(itemsRowNum)的 {{items.*}} 绝不能计入——否则其所在列(B/C/D…/I/K)
    // 被误判为「值列」，导致该列上方所有静态版式标签（承运商抬头/贸易术语/船名航次/货描表头…）
    // 在 PASS1 被误清，模板整体走样。仅扫描明细行之上区域。
    var valueCols = {};
    var valueEnd = (itemsRowNumArg && itemsRowNumArg !== -1) ? Math.min(headerEnd, itemsRowNumArg - 1) : headerEnd;
    for (var r = 1; r <= valueEnd; r++) {
      var row = ws.getRow(r);
      row.eachCell({ includeEmpty: true }, function (cell, col) {
        if (/\{\{/.test(_cellStr(cell))) valueCols[col] = true;
      });
    }
    // 合并单元格成员表：key = r + ',' + c；mergedMaster[r,c] = true 表示该单元格属于某个合并范围
    var mergedCell = {};
    // v1.4.50：mergedSubordinate 仅标记从属格（不含主格），供 PASS2 写值时跳过——避免合并从属格 _cellStr 返回主格值让 for 循环 c++ 越界
    var mergedSubordinate = {};
    (ws.model.merges || []).forEach(function (m) {
      var mm = (typeof m === 'string') ? m.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/) : null;
      if (!mm) return;
      var c1 = _colNum(mm[1]), r1 = parseInt(mm[2], 10), c2 = _colNum(mm[3]), r2 = parseInt(mm[4], 10);
      for (var rr = r1; rr <= r2; rr++) for (var cc = c1; cc <= c2; cc++) mergedCell[rr + ',' + cc] = true;
      // 主格 (r1, c1) 不算 subordinate；从属格 (rr != r1 || cc != c1) 才算
      for (var rr2 = r1; rr2 <= r2; rr2++) for (var cc2 = c1; cc2 <= c2; cc2++) {
        if (rr2 === r1 && cc2 === c1) continue;
        mergedSubordinate[rr2 + ',' + cc2] = true;
      }
    });
    // 这些词出现在单元格里时视为「标签/静态文本」，保留不清理。
    // 注意：「地址/电话/邮箱/联系人」等由 mapHeaderLabel 识别保护，不在这里兜底，
    // 否则会把 JW PEI、Noul LLC 等模板样本地址残留也保留下来。
    var KEEP = /(编码|编号|号|库|液体|粉末|危险品|清关|交税|交货|派送|参考|备注|保价|投保|箱数|商品|申报|材质|海关|用途|品牌|型号|英文|中文|品名|数量|重量|长|宽|高|带电|带磁|图片|链接|销售|价格|名称|发件人|收件人|发货人|通知人|收货|装货|卸货|船名|航次|委托|贸易术语|SHIPPER|CONSIGNEE|NOTIFY|BOOKING|INSTRUCTION|PO|NO\.|承运人|运费|吨位|航班|日期|文件|通知|声明|托运|到达|始发|体积|进仓|运单|预留)/i;

    // v1.4.50：扫描可识别 label cell，计算它们右侧"应该清的值 cell"（兼容无 {{}} 占位符的纯标签模板，
    //          例如中运通达-FBA订单(V3) 整张表没有任何 {{}} 占位符，valueCols 永远为空 → 样本数据全保留）
    // 同一行有多个 label 时，label 之间也属值 cell。
    // 识别范围：mapFieldLabel（涵盖 mapHeaderLabel 三类收发人 + GENERAL_LABEL_RULES 通用字段如仓库代码/运输方式/客户单号）
    var labelCellValueRanges = {}; // r -> [{c1,c2}]  闭区间
    var merged = ws.model.merges || [];
    function _parseMerge(m) {
      if (typeof m === 'string') {
        var mm = m.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
        if (!mm) return null;
        return { r1: parseInt(mm[2], 10), c1: _colNum(mm[1]), r2: parseInt(mm[4], 10), c2: _colNum(mm[3]) };
      }
      return null;
    }
    var labelCellList = []; // [{r, c, spanEnd}]  spanEnd = 标签合并区右边界
    for (var rL = 1; rL <= valueEnd; rL++) {
      var rowL = ws.getRow(rL);
      rowL.eachCell({ includeEmpty: false }, function (cellL, cL) {
        var sL = _cellStr(cellL);
        if (!sL || /\{\{/.test(sL)) return;
        if (!mapFieldLabel(sL)) return; // v1.4.50：用更宽的 mapFieldLabel（涵盖 GENERAL_LABEL_RULES，如"收件人(仓库代码)"/"运输方式"等）
        var spanEndL = cL;
        for (var mi = 0; mi < merged.length; mi++) {
          var m = _parseMerge(merged[mi]);
          if (m && m.r1 <= rL && rL <= m.r2 && m.c1 <= cL && cL <= m.c2) { spanEndL = Math.max(spanEndL, m.c2); break; }
        }
        labelCellList.push({ r: rL, c: cL, spanEnd: spanEndL });
      });
    }
    // 按行分组 label，按列排序
    var labelByRow = {};
    labelCellList.forEach(function (lc) { (labelByRow[lc.r] = labelByRow[lc.r] || []).push(lc); });
    Object.keys(labelByRow).forEach(function (rKey) {
      var arr = labelByRow[rKey].sort(function (a, b) { return a.c - b.c; });
      labelCellValueRanges[rKey] = [];
      for (var li = 0; li < arr.length; li++) {
        var start = arr[li].spanEnd + 1;
        var end = (li + 1 < arr.length) ? arr[li + 1].c - 1 : (start + 12); // 限 12 格防越界
        if (start <= end) labelCellValueRanges[rKey].push({ c1: start, c2: end });
      }
    });

    // PASS 1：清理值列里的样本残留（非占位符、非标签、非静态文本）。
    // 合并单元格成员只保留包含标签词的静态文本；具体样本数据（如旧地址、旧公司名）仍清空。
    // v1.4.50：兼容无 {{}} 占位符模板——按 mapFieldLabel label 右侧值格也清；label 右侧的值合并区（主格+从属格）也清（之前 mergedCell skip 会跳过整片合并区导致样本残留）
    for (var r1 = 1; r1 <= headerEnd; r1++) {
      var rowA = ws.getRow(r1);
      rowA.eachCell({ includeEmpty: true }, function (cell, col) {
        var s = _cellStr(cell);
        // v1.4.54：保护明细表头行（列标签 SKU/MODEL/DESCRIPTION/QTY 等），避免被当样本残留清空导致表头空白、数据列整体右移
        if (itemsRowNumArg && itemsRowNumArg !== -1 && r1 >= itemsRowNumArg - 3 && r1 <= itemsRowNumArg - 1) return;
        if (/\{\{/.test(s)) return;
        if (mapFieldLabel(s)) return; // v1.4.50：用更宽的 mapFieldLabel 保护"运输方式/客户单号/PO号"等 GENERAL_LABEL_RULES 也被识别为 label（避免被误清）
        var inLabelRange = false;
        if (labelCellValueRanges[r1]) {
          for (var ri = 0; ri < labelCellValueRanges[r1].length; ri++) {
            var rg = labelCellValueRanges[r1][ri];
            if (col >= rg.c1 && col <= rg.c2) { inLabelRange = true; break; }
          }
        }
        // v1.4.50：label 右侧值格（无论合并与否）都允许清——避免样本数据占据合并值格
        // v1.4.63：但合并从属格（subordinate cells）必须跳过——清空从属格会联动 ExcelJS 把整个合并区（H1:U14 警告区等）都设空字符串
        if (inLabelRange) {
          if (mergedSubordinate[r1 + ',' + col]) return; // 合并从属格不参与清理
          if (KEEP.test(s)) return; // 含保留词不动
          cell.value = '';
          return;
        }
        if (mergedCell[r1 + ',' + col]) return; // 非值格的合并锚点（版式标题/承运商抬头/贸易术语等）保留
        // v1.4.63：valueCols 清空也要跳合并从属格（同样的联动问题）
        if (mergedSubordinate[r1 + ',' + col]) return;
        if (valueCols[col] && !KEEP.test(s)) cell.value = '';
      });
    }

    // PASS 2：按模板表头标签写入收发人数据
    // v1.4.50：写值时跳过合并从属格（subordinate cells）——避免 PASS1 写合并主格后 for 循环看到从属格的 _cellStr 返回主格值继续 c++ 越界写到 col 11+（用户截图"地址"4 行重复就是因为这个越界）
    // v1.5.28：增加「区块归属」——R19「我司全称」之后 8 行内的「地址/联系人/电话/Email」等纯字段标签
    //   归属 shipper（无 party 词的标签行也能填），与 buildLabelMap 的 curParty 逻辑一致
    var curParty2 = '', curPartyRow2 = 0;
    for (var r2 = 1; r2 <= writeEnd; r2++) {
      var rowB = ws.getRow(r2);
      rowB.eachCell({ includeEmpty: false }, function (cellP, colP) {
        var sp = _cellStr(cellP);
        if (!sp || sp.length > 20) return;
        for (var pi2 = 0; pi2 < PARTY_RE.length; pi2++) {
          if (PARTY_RE[pi2].re.test(sp)) { curParty2 = PARTY_RE[pi2].party; curPartyRow2 = r2; return; }
        }
      });
      if (curParty2 && (r2 - curPartyRow2) > 8) { curParty2 = ''; curPartyRow2 = 0; }
      rowB.eachCell({ includeEmpty: true }, function (cell, col) {
        var s2 = _cellStr(cell);
        if (!s2) return;
        var info = mapHeaderLabel(s2);
        if (!info && curParty2) info = mapFieldLabel(s2, curParty2); // v1.5.28 区块归属兜底
        if (!info) return;
        var party = (data && data[info.party]) || {};
        var raw = party[info.field];
        var val = (info.field === 'address' && info.line > 0 && typeof raw === 'string')
          ? (raw.split(/\r?\n/)[info.line - 1] || '')
          : (raw === undefined || raw === null ? '' : raw);
        if (val === '') return;
        var target = null, isPh = false;
        for (var c = col + 1; c <= col + 12 && c <= ws.columnCount; c++) {
          if (mergedSubordinate[r2 + ',' + c]) continue; // 跳过合并从属格（主格保留，可写）——避免 _cellStr 返回主格值让 for 误以为非空继续 c++ 越界
          var tc = rowB.getCell(c); var ts = _cellStr(tc);
          if (ts && /\{\{/.test(ts)) {
            var mp = ts.match(/\{\{\s*([A-Za-z0-9_.]+)\s*\}\}/);
            if (mp && mp[1] === info.party + '.' + info.field) { target = tc; isPh = true; break; }
            continue; // 占位符不匹配，继续找空单元格
          }
          if (!ts) { target = tc; isPh = false; break; }
        }
        if (!target) return;
        if (!isPh && _cellStr(target)) return; // 已有内容则不改写
        target.value = (typeof val === 'number') ? val : String(val);
      });
    }
  }

  /** 解析 Excel 范围字符串，如 A1:B2 -> {top,left,bottom,right,range} */
  function _colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
  function _colLet(n) { var s = ''; while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  function parseMergeRef(m) {
    var mm = (typeof m === 'string') ? m.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/) : null;
    if (!mm) return null;
    return {
      top: parseInt(mm[2], 10), left: _colNum(mm[1]),
      bottom: parseInt(mm[4], 10), right: _colNum(mm[3]),
      range: m
    };
  }
  function makeMergeRef(r) { return _colLet(r.left) + r.top + ':' + _colLet(r.right) + r.bottom; }

  /** 构建合并单元格映射：subordinate[rn,cn] 表示该格是合并区域的从属格；
   *  masterOf[rn,cn] 指向其主格坐标字符串。用于避免重复填充/识别从属格。 */
  function buildMergedMaps(ws) {
    var subordinate = {}, masterOf = {}, masterRange = {};
    (ws.model.merges || []).forEach(function (m) {
      var ref = parseMergeRef(m);
      if (!ref) return;
      var masterKey = ref.top + ',' + ref.left;
      masterRange[masterKey] = ref;
      for (var r = ref.top; r <= ref.bottom; r++) {
        for (var c = ref.left; c <= ref.right; c++) {
          if (r === ref.top && c === ref.left) continue;
          subordinate[r + ',' + c] = true;
          masterOf[r + ',' + c] = masterKey;
        }
      }
    });
    return { subordinate: subordinate, masterOf: masterOf, masterRange: masterRange };
  }

  /** 把源模板 logo 贴回 workbook（预览与导出都能显示）。ws 为目标工作表。
   *  logo = { dataB64, ext, from:{col,row}, to:{col,row} } */
  function addLogo(wb, ws, logo) {
    if (!logo || !logo.dataB64 || !logo.ext) return;
    try {
      var logoBuf;
      if (typeof Buffer !== 'undefined') {
        logoBuf = Buffer.from(logo.dataB64, 'base64');
      } else if (typeof atob !== 'undefined') {
        var binary = atob(logo.dataB64);
        var arr = new Uint8Array(binary.length);
        for (var ii = 0; ii < binary.length; ii++) arr[ii] = binary.charCodeAt(ii);
        logoBuf = arr;
      }
      if (!logoBuf) return;
      var imgId = wb.addImage({ buffer: logoBuf, extension: logo.ext });
      var tl = logo.from || { col: 0, row: 0 };
      var br = logo.to || { col: (logo.from ? logo.from.col + 2 : 2), row: (logo.from ? logo.from.row + 2 : 2) };
      ws.addImage(imgId, { tl: tl, br: br });
    } catch (e) {}
  }

  /** 模板填充：wb已加载的模板workbook，data为buildDocData输出。原地填充。
   *  支持两种模式：① {{items.xxx}} 占位符（老模板） ② 表头识别（物流商真实模板，无占位符也可填）
   *  options.logo = { dataB64, ext, from:{col,row}, to:{col,row} } 可选，用于把源模板 logo 贴回输出 workbook */
  /** v1.4.60 —— 「本模板无法承载的字段」清单。
   *  原则：用户录入什么就该显示什么；模板确实没有对应栏位时，系统必须当场告知，
   *  而不是静默丢弃。返回 [{path,label,value}]，由 UI 在预览页列出。 */
  var UNCARRIED_PARTY_CN = { shipper: '发货人', consignee: '收货人', notify: '通知人' };
  var UNCARRIED_FIELD_CN = {
    name: '名称', company: '公司', warehouseCode: '仓库代码', address: '地址', city: '城市',
    state: '省/州', zip: '邮编', country: '国家', tel: '电话', email: '邮箱', contact: '联系人',
    taxNo: '税号', vatNo: 'VAT号', eori: 'EORI'
  };
  var UNCARRIED_META_CN = {
    invoiceNo: '发票号', invoiceDate: '发票日期', contractNo: '合同号', orderNos: '订单号',
    incoterms: '成交/交货条款', paymentTerms: '付款方式', transport: '运输方式',
    pol: '起运地', pod: '目的地', etd: 'ETD', vessel: '船名航次',
    containerType: '箱型', containerQty: '箱量', shippingMarks: '唛头',
    remark: '备注', customsType: '报关方式', agent: '代理'
  };
  // 系统默认值：不是用户录入的，不进"丢失"清单，避免噪音
  var UNCARRIED_DEFAULTS = { transport: 'BY SEA', shippingMarks: 'N/M', dangerous: 'NON-DANGEROUS / GENERAL CARGO' };
  function computeUncarried(data, labelMap, phPaths) {
    var carried = {};
    (labelMap || []).forEach(function (e) { if (e && e.resolved && e.path) carried[e.path] = 1; });
    (phPaths || []).forEach(function (p) { if (p) carried[p] = 1; });
    var out = [];
    ['shipper', 'consignee', 'notify'].forEach(function (pk) {
      var p = data && data[pk];
      if (!p || typeof p !== 'object') return;
      if (pk === 'notify' && String(p.name || '').toUpperCase() === 'SAME AS CONSIGNEE' && !p.address) return;
      Object.keys(UNCARRIED_FIELD_CN).forEach(function (f) {
        var v = p[f];
        if (v === undefined || v === null || String(v).trim() === '') return;
        var path = pk + '.' + f;
        if (carried[path]) return;
        // 地址多行占位（address1/address2…）也视为已承载
        if (f === 'address' && (carried[pk + '.address1'] || carried[pk + '.address2'])) return;
        out.push({ path: path, label: UNCARRIED_PARTY_CN[pk] + '·' + UNCARRIED_FIELD_CN[f], value: String(v) });
      });
    });
    Object.keys(UNCARRIED_META_CN).forEach(function (k) {
      var v = data && data[k];
      if (v === undefined || v === null || String(v).trim() === '') return;
      if (UNCARRIED_DEFAULTS[k] !== undefined && String(v) === UNCARRIED_DEFAULTS[k]) return;
      if (carried[k]) return;
      out.push({ path: k, label: UNCARRIED_META_CN[k], value: String(v) });
    });
    return out;
  }

  function fillTemplate(wb, data, options) {
    options = options || {};
    var filled = { replaced: [], unresolved: [], uncarried: [] };
    var ws = wb.worksheets[0];
    if (!ws) throw new Error('模板无工作表');
    // v1.5.14 重置"已嵌图"守卫：每次重新填充模板都应允许(且只应)嵌一次产品图；
    // 防止预览 step4 与导出复用同一 wb 时 embedProductImages 被重复调用导致每行图嵌两遍重叠。
    try { wb._productImagesEmbedded = false; } catch (e) {}

    // ① 抓取模板原始样式快照（边框/底色/字体/对齐），用于末尾 1:1 还原，
    //    避免 ExcelJS 写回时空值带样式的单元格丢边框。
    var origStyles = {};
    ws.eachRow({ includeEmpty: true }, function (row, rn) {
      row.eachCell({ includeEmpty: true }, function (cell, cn) {
        var st = cell.style;
        if (st && (st.border || st.fill || st.font || st.alignment)) {
          origStyles[rn] = origStyles[rn] || {};
          origStyles[rn][cn] = JSON.parse(JSON.stringify(st));
        }
      });
    });
    // 同时抓取原始合并单元格，插入明细行后需要按行位移重新计算并应用，
    // 否则下方合并区域会与新插入行重叠或错位，导致内容显示不全。
    var origMerges = (ws.model.merges || []).slice().map(parseMergeRef).filter(Boolean);
    // 同时构建合并从属格映射，避免对从属格重复识别表头/重复写值（如 B36:C36 合并的 C36）。
    var mergedMaps = buildMergedMaps(ws);

    // 行位移记录（仅第 2 段插入明细行会产生下移）；2.5 段改为清空值不删行，不产生位移。
    var splices = [];
    function mapRow(r) {
      var out = r;
      for (var i = 0; i < splices.length; i++) if (splices[i].at <= r) out += splices[i].delta;
      return out;
    }

    function replaceInString(s, ctx) {
      return s.replace(PH_RE, function (all, path) {
        var v = getPath(ctx, path);
        if (v === undefined || v === null) { filled.unresolved.push(path); return ''; }
        filled.replaced.push(path);
        return typeof v === 'number' ? String(v) : String(v);
      });
    }
    function isPureNumericPh(s, ctx) {
      PH_RE.lastIndex = 0;
      var m = PH_RE.exec(s);
      if (m && m[0] === s.trim()) {
        var v = getPath(ctx, m[1]);
        if (typeof v === 'number') return v;
      }
      return null;
    }
    // 1) 定位明细行：优先占位符，其次表头识别
    var itemsRowNum = -1, itemHeaderMap = {};
    ws.eachRow({ includeEmpty: false }, function (row, rowNumber) {
      if (itemsRowNum !== -1) return;
      row.eachCell({ includeEmpty: false }, function (cell) {
        var s = cellString(cell);
        if (s && /\{\{\s*items\./.test(s)) itemsRowNum = rowNumber;
      });
    });
    if (itemsRowNum !== -1) {
      itemHeaderMap = scanItemHeader(ws, itemsRowNum, mergedMaps);
    } else {
      // 兜底：无占位符时识别真实表头区域
      var best = { row: -1, count: 0, map: {} };
      ws.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber > 30) return;
        var map = {}, count = 0;
        row.eachCell({ includeEmpty: false }, function (cell, colNumber) {
          if (map[colNumber]) return;
          var f = matchHeaderAlias(cellString(cell));
          if (f) { map[colNumber] = f; count++; }
        });
        if (count > best.count) best = { row: rowNumber, count: count, map: map };
      });
      if (best.count >= 3) { itemsRowNum = best.row + 1; itemHeaderMap = best.map; }
    }
    // 记录实际明细起始行 + 列字段映射，供 embedProductImages 复用，避免两函数各自推算行/列号导致图片错位
    // （尤其用户上传的定制模板：{{items}}占位符行 / 表头行 / 产品图片表头行 可能不在同一行）
    try { wb._fillItemsRowNum = itemsRowNum; wb._itemHeaderMap = itemHeaderMap || {}; } catch (e) {}

    // 1.5) 模板表头识别填充（发货人/收货人/地址/电话等）：取模板自身表头标签填
    // v1.5.27 writeEnd=25：覆盖 itemsRow 之下的「我司/开票方」信息区（无占位符模板的 R19-23）
    fillHeaderByLabels(ws, data, itemsRowNum === -1 ? 25 : itemsRowNum, itemsRowNum, 25);

    // 1.6) 通用字段标签映射填充（运输方式/起运国/客户单号/总件数 等未被收发人三类覆盖的表头字段）
    //      优先用模板自带的 labelMap（可 UI 编辑），否则现场扫描；对纯 {{}} 占位符模板为空的，no-op。
    var _labelMap = (options && options.labelMap && options.labelMap.length)
      ? options.labelMap
      : buildLabelMap(wb, itemsRowNum);
    fillByFieldLabels(ws, data, _labelMap, mergedMaps);

    // 明细区下界（供 2.5 清理残留行与 3 跳过明细区使用，必须在前面算出）
    var itemEnd = itemsRowNum === -1 ? -1 : itemsRowNum + Math.max((data.items || []).length, 1) - 1;

    // 扫描模板实际明细槽位数：连续行「至少一格有 border」即视为同一明细槽
    //   v1.4.26 修复 KEAS 类「模板已预留 10 行只第 1 行有 items.* 又被插 9 行」bug
    //   - 仅靠 `{{items.*}}` 计数会漏掉「第 1 行有 items.* + 后续 N-1 行仅占位符+同款 border」的常见模板
    //   - 用 border 识别更稳：明细行必有边框（表头/页脚行一般无）
    function hasItemRowBorder(row) {
      var ok = false;
      row.eachCell({ includeEmpty: true }, function (cell) {
        if (cell.border && (cell.border.top || cell.border.bottom || cell.border.left || cell.border.right)) ok = true;
      });
      return ok;
    }
    var templateSlots = 1;
    if (itemsRowNum !== -1) {
      // v1.5.28：表尾特征词——即使这些行带边框也不计入明细槽位。
      // 买单要素模板 R17-R24 表尾（注意事项/我司/联系人/电话/Email）全部带边框，
      // 原逻辑把表尾误算进槽位 → items 不触发插入 → 后几个 SKU 直接覆盖表尾。
      var FOOTER_RE = /(注意事项|温馨提示|合计|TOTAL|我司|地址：|联系人：|联系电话|传真|开票|抬头|组织机构|签名|SIGNATURE|REMARK|备注|谨记|敬告)/i;
      while (true) {
        var nextRow = ws.getRow(itemsRowNum + templateSlots);
        if (!hasItemRowBorder(nextRow)) break;
        var isFooter = false;
        nextRow.eachCell({ includeEmpty: false }, function (cellF) {
          var sf = _cellStr(cellF);
          if (sf && FOOTER_RE.test(sf)) isFooter = true;
        });
        if (isFooter) break;
        templateSlots++;
        if (templateSlots > 50) break; // 防御
      }
    }

    // 2) 明细区展开：先复制模板行样式与占位内容 N-1 次
    var items = data.items || [];
    if (itemsRowNum !== -1 && items.length > 0) {
      var tplRow = ws.getRow(itemsRowNum);
      var tplCells = [];
      tplRow.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        tplCells.push({ col: colNumber, value: cell.value, style: cell.style });
      });
      var slotDelta = items.length - templateSlots; // 正=需插入；负=需删除
      if (slotDelta > 0) {
        // 模板槽位不足 → 插入 (items.length - templateSlots) 行
        var _insertArgs = [];
        for (var _k = 0; _k < slotDelta; _k++) _insertArgs.push([]);
        splices.push({ at: itemsRowNum + 1, delta: slotDelta });
        ws.spliceRows.apply(ws, [itemsRowNum + 1, 0].concat(_insertArgs));
        for (var i = 1; i < slotDelta + 1; i++) {
          var newRow = ws.getRow(itemsRowNum + i);
          newRow.height = tplRow.height;
          tplCells.forEach(function (tc) {
            var c = newRow.getCell(tc.col);
            c.value = tc.value;
            c.style = JSON.parse(JSON.stringify(tc.style || {}));
          });
        }
      } else if (slotDelta < 0) {
        // 模板槽位多于实际 items → 删除多余行（保留带边框整行→上移不影响：V1.4.15 实践已改为「保留行+清残留」，
        //   这里若直接 spliceRows 删 N 行会导致其下 VGM/footer 上移并破坏 2.5 段原样式，故暂保留「不减行」兜底）
        //   为避免插入/删除不对称导致版式错位，强制以 items.length 槽数为准（多删少插）—— 实测多数模板 items 数=模板槽数，无需走此分支。
        // 暂不实现：保持现状多槽位（多余槽位 2.5 段会清掉 {{items.*}} 占位符文本，整行仍带原样式）
      }
      // 逐行填充 items：占位符优先；无占位符但有表头映射的列，按表头写值
      // ⑤ 合并还原（必须前移到 pass2 之前：否则 mergedMaps 用的是插入前的旧合并坐标，
      //    会把页脚合并 A19:I19/B21:I21/F23:I23 误判到插入后的明细行 19/21/23，导致 B-I 列被当从属格跳过）
      if (itemsRowNum !== -1 && origMerges.length) {
        var delta = 0;
        for (var si = 0; si < splices.length; si++) if (splices[si].at === itemsRowNum + 1) { delta = splices[si].delta; break; }
        if (delta > 0) {
          var newMerges = [];
          var pattern = origMerges.filter(function (m) { return m.top === itemsRowNum; });
          origMerges.forEach(function (m) {
            if (m.bottom < itemsRowNum) {
              newMerges.push({ top: m.top, left: m.left, bottom: m.bottom, right: m.right });
            } else if (m.top > itemsRowNum) {
              // 修正：原用 m.top>itemEnd 会把紧贴明细区的页脚合并漏掉->丢合并；改用 m.top>itemsRowNum 判定「在明细区下方」
              newMerges.push({ top: m.top + delta, left: m.left, bottom: m.bottom + delta, right: m.right });
            }
          });
          for (var k = 0; k <= delta; k++) {
            pattern.forEach(function (p) {
              newMerges.push({ top: itemsRowNum + k, left: p.left, bottom: itemsRowNum + k, right: p.right });
            });
          }
          (ws.model.merges || []).slice().forEach(function (m) { try { ws.unMergeCells(m); } catch (e) {} });
          newMerges.forEach(function (m) {
            try { ws.mergeCells(makeMergeRef(m)); } catch (e) {}
          });
        }
      }
      // 合并还原后重算从属格映射，pass2 据此只跳过真正的明细行内合并（如 B:C 合并），不再误伤页脚合并错位的明细行
      mergedMaps = buildMergedMaps(ws);
      for (var r = 0; r < items.length; r++) {
        var rowObj = ws.getRow(itemsRowNum + r);
        var itR = items[r];
        // v1.4.27 选项A：整行无货物内容（无 nameEn/nameCn/description/sku）时整行留空，
        //   含 B(boxCount)/I(S/O NO)/K(boxNo)/M(SEAL NO) 等均不写，仅清掉残留 {{items.*}} 占位符
        var rowEmpty = !itR || (!itR.nameEn && !itR.nameCn && !itR.description && !itR.sku);
        if (rowEmpty) {
          (function () {
            var rn = itemsRowNum + r;
            rowObj.eachCell({ includeEmpty: true }, function (cell) {
              if (mergedMaps.subordinate[rn + ',' + cell.col]) return;
              var s = cellString(cell);
              if (s && s.indexOf('{{items.') >= 0) cell.value = '';
            });
          })();
          continue;
        }
        var ctx = Object.assign({}, data, { items: itR });
        rowObj.eachCell({ includeEmpty: true }, function (cell) {
          if (mergedMaps.subordinate[(itemsRowNum + r) + ',' + cell.col]) return; // 合并从属格由主格统一显示，不单独写值
          var s = cellString(cell);
          var hasPh = s && s.indexOf('{{') >= 0;
          if (hasPh) {
            var num = isPureNumericPh(s, ctx);
            cell.value = (num !== null) ? num : replaceInString(s, ctx);
          }
          // 表头兜底：该列上方有识别出的表头，且该单元格无占位符（或占位符解析后为空），则按表头写值
          // v1.4.53 修复：cellString 对数字返回 ''，故必须直接判断 cell.value，否则数字会被错映射字段覆盖
          var field = itemHeaderMap[cell.col];
          if (field) {
            var v = getPath(ctx, 'items.' + field);
            if (field === 'boxCount' && itR._firstOfBox === false) v = ''; // v1.4.62 箱数仅首行显示，避免混装被误读为总箱数
            if (v === undefined || v === null || v === '') v = getPath(ctx, field);
            if (v !== undefined && v !== null && v !== '') {
              // 如果占位符已经写了有效值（含数字），不覆盖；否则按表头写入
              // 注意：cellString 对数字返回 ''，故必须直接判断 cell.value，否则数字会被错映射字段覆盖
              var _cv = cell.value;
              var _filled = (_cv !== undefined && _cv !== null && _cv !== '');
              if (!hasPh || !_filled) cell.value = (typeof v === 'number') ? v : String(v);
            }
          }
        });
      }
      // 2.1.5) 兜底：模板预留下方明细行（如 KEAS 第 1 行有 items.* 后面 9 行仅 marks 占位符+border）
      //        rowObj.eachCell 不会遍历「无 value 仅有 style」的格，按表头映射显式补写
      if (Object.keys(itemHeaderMap).length > 0) {
        for (var r2 = 0; r2 < items.length; r2++) {
          var rowObj2 = ws.getRow(itemsRowNum + r2);
          var itR2 = items[r2];
          // v1.4.27 选项A：空内容行（无品名）整行留空，不写任何值（占位符已在上层清空）
          if (!itR2 || (!itR2.nameEn && !itR2.nameCn && !itR2.description && !itR2.sku)) continue;
          if (mergedMaps.subordinate[(itemsRowNum + r2) + ',' + 1]) continue; // 合并从属格主格统一显示
          var ctx2 = Object.assign({}, data, { items: itR2 });
          Object.keys(itemHeaderMap).forEach(function (colStr) {
            var col = parseInt(colStr, 10);
            if (mergedMaps.subordinate[(itemsRowNum + r2) + ',' + col]) return;
            var cell2 = rowObj2.getCell(col);
            var cv2 = cell2.value;
            if (cv2 !== undefined && cv2 !== null && cv2 !== '') return; // 已有内容（含数字）跳过，避免数字被 cellString 误判为空而覆盖
            var field2 = itemHeaderMap[col];
            if (!field2) return;
            if (field2 === 'boxCount' && itR2._firstOfBox === false) return; // v1.4.62 箱数仅首行显示，避免混装被误读为总箱数
            var v2 = getPath(ctx2, 'items.' + field2);
            if (v2 === undefined || v2 === null || v2 === '') v2 = getPath(ctx2, field2);
            if (v2 !== undefined && v2 !== null && v2 !== '') {
              cell2.value = (typeof v2 === 'number') ? v2 : String(v2);
            }
          });
        }
      }
      // 2.2) 列宽自适应：按 items 各字段最长值（中文字符按 2 倍宽估算）动态加宽
      //   itemHeaderMap 映射的列，避免长字符溢出覆盖相邻列；wrapText 关掉放末尾（避免 ④ 还原 alignment 被覆盖）。
      Object.keys(itemHeaderMap).forEach(function (colStr) {
        var col = parseInt(colStr, 10);
        var field = itemHeaderMap[col];
        if (!field) return;
        var sample = '';
        items.forEach(function (it) {
          var v = (it && it[field] !== undefined && it[field] !== null) ? it[field] : '';
          var s = String(v);
          var pxS = 0, pxSample = 0;
          for (var kk = 0; kk < s.length; kk++) {
            var ch = s.charCodeAt(kk);
            pxS += ((ch >= 0x4E00 && ch <= 0x9FFF) || (ch >= 0x3000 && ch <= 0x303F) || (ch >= 0xFF00 && ch <= 0xFFEF)) ? 14 : 7;
          }
          for (var kk2 = 0; kk2 < sample.length; kk2++) {
            var ch2 = sample.charCodeAt(kk2);
            pxSample += ((ch2 >= 0x4E00 && ch2 <= 0x9FFF) || (ch2 >= 0x3000 && ch2 <= 0x303F) || (ch2 >= 0xFF00 && ch2 <= 0xFFEF)) ? 14 : 7;
          }
          if (pxS > pxSample) sample = s;
        });
        if (!sample) return;
        var px = 0;
        for (var k = 0; k < sample.length; k++) {
          var ch3 = sample.charCodeAt(k);
          px += ((ch3 >= 0x4E00 && ch3 <= 0x9FFF) || (ch3 >= 0x3000 && ch3 <= 0x303F) || (ch3 >= 0xFF00 && ch3 <= 0xFFEF)) ? 14 : 7;
        }
        var wsCol = ws.getColumn(col);
        var oldW = wsCol.width || 8.43;
        var newW = Math.max(oldW, Math.ceil(px / 7) + 2);
        if (newW > oldW + 0.05) wsCol.width = newW;
      });
    }

    // 2.5) 清理模板多余的明细占位符：itemEnd 之后若仍有 {{items.*}} 残留（模板预设行数>实际条数），
    //      只清空占位符文本、保留带边框的整行（不再 spliceRows 删除行，避免底部行上移导致边框错位丢失）。
    if (itemsRowNum !== -1) {
      for (var rr = itemEnd + 1; rr <= ws.rowCount; rr++) {
        var rrow = ws.getRow(rr);
        rrow.eachCell({ includeEmpty: true }, function (cell) {
          var sv = cellString(cell);
          if (sv && /\{\{\s*items\./.test(sv)) cell.value = '';
        });
      }
    }

    // 3) 其余单元格占位符替换（{{items.*}} 已在第 2 段处理，此处跳过；
    //    但明细区内的非 items 占位符，如跨越明细行的 {{shippingMarks}} 合并主格，需要在此填充）
    ws.eachRow({ includeEmpty: false }, function (row, rowNumber) {
      row.eachCell({ includeEmpty: false }, function (cell) {
        var v = cell.value;
        var s = (v && v.richText) ? v.richText.map(function (t) { return t.text; }).join('') : (typeof v === 'string' ? v : null);
        if (s === null || s.indexOf('{{') < 0) return;
        if (/^\{\{\s*items\./.test(s)) return; // 明细占位符只在明细区展开，跳过
        var num = isPureNumericPh(s, data);
        cell.value = (num !== null) ? num : replaceInString(s, data);
      });
    });

    // ④ 1:1 还原模板样式：按行位移映射，把原始边框/底色/字体/对齐套回对应单元格，
    //    保证导出文件与模板「表格样式」完全一致（含底部边框行不被位移丢失）。
    Object.keys(origStyles).forEach(function (rn) {
      var tr = parseInt(rn, 10);
      var fr = mapRow(tr);
      Object.keys(origStyles[rn]).forEach(function (cn) {
        var cell = ws.getRow(fr).getCell(parseInt(cn, 10));
        cell.style = JSON.parse(JSON.stringify(origStyles[rn][cn]));
      });
    });



    // ⑤.5) 明细行 wrapText 关掉（放在 ④ 还原 origStyles 之后，避免被覆盖）：
    //   源模板 D/F/H 等列常 wrap=true，列宽刚好到边界时会强制把 "Handbag" 折成 "Handba\ng"
    //   显示错位；关 wrap 后长字符溢出右空白列更符合订舱单观感（且第 2.2 段已加宽列避免溢出到相邻内容列）。
    if (itemsRowNum !== -1) {
      for (var rrW = 0; rrW < items.length; rrW++) {
        var rrowW = ws.getRow(itemsRowNum + rrW);
        rrowW.eachCell({ includeEmpty: true }, function (cell) {
          if (cell.alignment && cell.alignment.wrapText) {
            cell.alignment = Object.assign({}, cell.alignment, { wrapText: false });
          }
        });
      }
    }
    // ⑥ 把源模板 logo 贴回输出 workbook（预览与导出都能显示）
    addLogo(wb, ws, options.logo);
    // ⑥.5) 合并从属格继承主格样式：ExcelJS 写回无值单元格时会丢弃其样式，导致合并块从属格
    //      （如 Aramex I21:K22 的 J21/K21/I22/J22/K22）在「写回→重读」后丢失对齐/字体等；
    //       把主格样式显式赋给从属格使其持久化。视觉不变（合并块本就按主格显示）。
    (ws.model.merges || []).forEach(function (mref) {
      var mm = parseMergeRef(mref); if (!mm) return;
      var master = ws.getCell(mm.top, mm.left);
      var ms = master.style; if (!ms) return;
      for (var rr = mm.top; rr <= mm.bottom; rr++) {
        for (var cc = mm.left; cc <= mm.right; cc++) {
          if (rr === mm.top && cc === mm.left) continue;
          var sc = ws.getCell(rr, cc);
          if (!sc.style || JSON.stringify(sc.style) !== JSON.stringify(ms)) {
            try { sc.style = JSON.parse(JSON.stringify(ms)); } catch (e) {}
          }
        }
      }
    });
    // v1.4.60：统计「用户录入了但本模板装不下」的字段，交 UI 显式提示（不静默丢弃）
    try {
      filled.uncarried = computeUncarried(data, _labelMap, filled.replaced.concat(filled.unresolved));
    } catch (e) { filled.uncarried = []; }
    return filled;
  }

  // ---------- 产品图嵌入（v1.5.1）：异步、防卡死、失败降级、内存缓存 ----------
  // 数据：window.PRODUCT_IMAGE_MAP { sku: "images/products/<sku>.jpg" }（同源静态文件，随 SPA 部署）
  // 匹配键：data.items[].sku（引擎实际用于生成发票的 SKU，不依赖 declare_data 是否有 imageUrl 字段）
  // 防卡死四道防线：①只拉本次发票涉及的 SKU ②并发限流6路 ③单张超时3s ④失败降级(不嵌,绝不阻断生成)
  var _imgCache = {}; // sku -> base64 dataURL（同会话不重复拉）

  function embedProductImages(wb, data, options) {
    var diag = { mapLen: 0, wsFound: false, imgCol: -1, headerRow: -1, itemsLen: 0, tasksLen: 0, err: null, done: false, missed: [] };
    if (typeof window !== 'undefined') window.__embDiag = diag;
    return new Promise(async function (resolve) {
      try {
        // v1.5.14 去重守卫：同一 wb 只嵌一次产品图（预览 step4 嵌过、导出复用同一 wb 时跳过，避免每行图嵌两遍重叠）。
        // fillTemplate 开头已把 wb._productImagesEmbedded 重置为 false，故"换数据重填→重嵌"仍正常。
        if (wb._productImagesEmbedded) { diag.done = true; resolve(); return; }
        // v1.5.20 fetch 模式: 图文件原封不动放 GitHub 仓库 images/sku_thumb/
        // v1.5.38 双镜像兜底：CloudStudio(国内快) 同源先试，GitHub raw 兜底——任何镜像的匿名用户都能拉到图
        // v1.5.44 唯一索引：SKU_IMAGE_INDEX(JW PEI G Unit 抽 153 张 SPU 真图，逐张核对内容正确，非占位图)
        //   + SPU 前缀兜底：J3120505X006-4-10 找不到时→ 自动用 J3120505X006-4(同款不同尺码共用产品图，正确行为非错图)
        var MAP = {};
        try {
          if (typeof window !== 'undefined') {
            // v1.5.50 修复：本地索引已是完整相对路径（如 "images/sku_thumb/104-1_4.jpeg"），禁止再拼前缀。
            var localIdx = window.SKU_IMAGE_INDEX || {};
            // v1.6.1 远程 URL 索引：飞书表 N 列商品图链接（Shopify/阿里 CDN 等，已实测 ACAO=* 可跨域 fetch）
            var urlIdx = window.SKU_IMAGE_URL_INDEX || {};
            var GH_PAGES = 'https://heryma99.github.io/trade-docs-system/';
            var RAW = 'https://raw.githubusercontent.com/heryma99/trade-docs-system/main/';
            var JSDELIVR = 'https://cdn.jsdelivr.net/gh/heryma99/trade-docs-system@main/';
            var CS_DOMAIN = (typeof location!=='undefined' && location.hostname.indexOf('app.workbuddy.link')!==-1)
              ? './'   // 当前就在 CloudStudio 镜像内, 用同域相对路径(必命中且最快)
              : 'https://a2012d426ebf40b8906cfe5f338c7516.app.workbuddy.link/';
            // v1.6.1 远程优先：候选链 = [远程URL] + [同 SKU 本地 sku_thumb 镜像](本地有则兜底)。
            //   少数 http://(https页混合内容)、需鉴权(飞书内部流401)、无 CORS(个别 host) 的远程链接，
            //   会在 fetch 失败后自然落到本地兜底候选，最大覆盖、绝不错图。
            for (var ku in urlIdx) {
              var ru = urlIdx[ku];
              if (!ru || !/^https?:/.test(ru)) continue;
              var lrel = localIdx[ku];
              var chain = [ru];
              if (lrel && lrel.indexOf('images/') === 0) {
                chain.push(lrel, GH_PAGES + lrel, JSDELIVR + lrel, RAW + lrel, CS_DOMAIN + lrel);
              }
              MAP[ku] = chain;
            }
            // 本地索引兜底：远程索引没有的 sku 沿用原有候选链
            // 候选链优先级(同域→GH Pages→jsDelivr CDN→Statically CDN→raw→CloudStudio),
            // 全部走完才记为取图失败; 每个候选内部还会重试, 不轻易放弃。
            for (var ka in localIdx) {
              var relA = localIdx[ka];                       // 已是相对路径，禁止再拼前缀
              if (!relA || relA.indexOf('images/') !== 0) continue;
              if (MAP[ka]) continue;
              MAP[ka] = [relA, GH_PAGES + relA, JSDELIVR + relA, RAW + relA, CS_DOMAIN + relA];
            }
          }
        } catch (e) {}
        diag.mapLen = MAP ? Object.keys(MAP).length : 0;
        if (!MAP || !Object.keys(MAP).length) { diag.done = true; resolve(); return; }
        // v1.5.15 严格精确匹配：图库按 SKU 全码建键，只认精确命中。
        // 用户铁律「没有就是没有，不要用错」——禁止任何前缀剥离/模糊兜底，
        // 否则缺失 SKU 会被塞进「另一款产品的图」造成张冠李戴。
        // v1.5.44 例外豁免：SPU 索引(JW PEI G Unit)天然只有 SPU 级图(例 J3120505X006-4)，同款尺码 SKU(J3120505X006-4-10)
        //   精确无图时, 允许按"末尾 -数字尺码"剥离一次降级到 SPU 主款(同款不同尺码共用产品图是正确行为, 非错图)。
        //   只降级一次, 防止跨款张冠李戴。
        function lookupImg(sku) {
          if (!sku) return null;
          var s = String(sku).trim();
          if (MAP[s]) return MAP[s];
          // SPU 前缀降级: 去掉末尾 "-数字" (尺码), 例 J3120505X006-4-10 → J3120505X006-4
          var m = /^(.+)-\d+(?:\.\d+)?$/.exec(s);
          if (m && MAP[m[1]]) return MAP[m[1]];
          return null;
        }
        var ws = (wb.worksheets && wb.worksheets[0]) || (wb.getWorksheet && wb.getWorksheet(1)) || (wb.getWorksheet && wb.getWorksheet('Sheet1'));
        diag.wsFound = !!ws;
        if (!ws) { diag.done = true; resolve(); return; }
        // 1) 扫描找"产品图片"列 + 表头行
        var imgCol = -1, headerRow = -1, sourceMap = 'scan';
        // 0) 优先复用 fillTemplate 已识别的 itemHeaderMap：直接把 imageUrl 字段对应的 colNumber 拿来用。
        //    itemHeaderMap 结构是 { 列号: 字段名 }（如 {"12":"imageUrl"}），故需找 value==='imageUrl' 的那个 key。
        //    解决：定制模板产品图片列不在 [5] 默认表头位置、扫描正则命中错位（如被模板装饰文字"因亚马逊..."中的
        //    其它字符误导到右侧很远空白列）、或明细表头行不在 R20+ 而在 R40+ 导致扫描行数不够；
        //    也覆盖英文表头（"Product Photo"）：scanItemHeader 经 HEADER_ALIASES 已映射成 imageUrl，扫描兜底却可能漏。
        try {
          var imgMap = null;
          if (wb._itemHeaderMap) {
            Object.keys(wb._itemHeaderMap).forEach(function (ck) {
              if (wb._itemHeaderMap[ck] === 'imageUrl') imgMap = parseInt(ck, 10);
            });
          }
          if (imgMap && imgMap > 0) {
            imgCol = imgMap;
            headerRow = -1; // 已知列号；起始行靠 _fillItemsRowNum
            sourceMap = 'itemHeaderMap';
          }
        } catch (_) {}
        var maxScanRow = Math.min((ws.rowCount || 10), 60);
        var maxScanCol = 40;
        // 1a) 优先精确表头（以 产品图片/商品图片 开头的单元格）。
        //     避免模板其他含"图片"二字的说明文字（如「图片请勿超出单元格」「详见图片」）
        //     出现在更靠前的行/更靠右的列时，被宽泛匹配抢先命中 → 图片错飞到模板最右空白列。
        if (imgCol < 0) {
        for (var rr = 1; rr <= maxScanRow && imgCol < 0; rr++) {
          var srow = ws.getRow(rr);
          for (var cc = 1; cc <= maxScanCol; cc++) {
            var scell = srow.getCell(cc);
            var stxt = (cellString(scell) || '').trim();
            var firstLine = stxt.split('\n')[0].trim();
            // 以 产品图片/商品图片 开头（兼容「产品图片(*)」「商品图片（不带电）」等多行/带后缀表头）
            if (/^(产品图片|商品图片)/.test(firstLine)) {
              imgCol = cc; headerRow = rr; sourceMap = 'scan-exact'; break;
            }
          }
        }
        // 1b) 兜底：宽泛匹配（图片/IMAGE/PHOTO/PICTURE），仅当精确表头未命中时
        if (imgCol < 0) {
          for (var rr2 = 1; rr2 <= maxScanRow && imgCol < 0; rr2++) {
            var srow2 = ws.getRow(rr2);
            for (var cc2 = 1; cc2 <= maxScanCol; cc2++) {
              var scell2 = srow2.getCell(cc2);
              var stxt2 = cellString(scell2) || '';
              if (/(产品图片|商品图片|图片|IMAGE|PHOTO|PICTURE)/i.test(String(stxt2).trim())) {
                imgCol = cc2; headerRow = rr2; sourceMap = 'scan-loose'; break;
              }
            }
          }
        }
        }
        diag.imgCol = imgCol; diag.headerRow = headerRow; diag.sourceMap = sourceMap;
        if (imgCol < 0) { diag.done = true; resolve(); return; } // 模板无产品图片列 → 跳过
        // v1.5.16 清模板预置图：模板在"产品图片列"自带/遗留的示例图(与 SKU 无一一对应关系)必须清除，
        // 否则会跟着导出造成"乱七八糟的图"（如用户截图里飘在表底的鞋图）。只清图片列的图，表头 logo 等不受影响。
        try {
          var imgsArr = (typeof ws._media !== 'undefined') ? ws._media : null;
          var before = (imgsArr && imgsArr.length) || 0;
          if (imgsArr && imgsArr.length) {
            var imgCol0 = imgCol - 1;                     // 0-based
            var headerRow0 = headerRow > 0 ? headerRow - 1 : -1;
            var keep = [];
            imgsArr.forEach(function (im) {
              var rng = im && im.model && im.model.range;
              var tl = rng && rng.tl;
              // 兼容 Node/浏览器压缩版字段名(nativeCol/nativeRow vs col/row)
              var c = tl && (typeof tl.nativeCol !== 'undefined') ? tl.nativeCol : (tl && tl.col);
              var r = tl && (typeof tl.nativeRow !== 'undefined') ? tl.nativeRow : (tl && tl.row);
              var inImgCol = (c === imgCol0);
              var inImgZone = inImgCol && (headerRow0 < 0 || r >= headerRow0);
              if (inImgZone) { diag.presetCleared = (diag.presetCleared || 0) + 1; }
              else keep.push(im);
            });
            imgsArr.length = 0;
            for (var ki = 0; ki < keep.length; ki++) imgsArr.push(keep[ki]);
            diag.presetBefore = before; diag.presetAfter = keep.length;
          }
        } catch (e) { diag.presetErr = String(e); }
        // 2) 收集嵌入任务
        var items = (data && data.items) || [];
        diag.itemsLen = items.length;
        // 优先复用 fillTemplate 写数据时的真实起始行；无记录时回退 headerRow+1
        var itemsRowNum = (wb._fillItemsRowNum && wb._fillItemsRowNum > 0) ? wb._fillItemsRowNum : (headerRow > 0 ? headerRow + 1 : 1);
        var tasks = [];
        var markCells = []; // v1.6.8 缺图标注：记录需写红/黄底「图片缺失」的单元格 {row,col,sku,type}
        for (var i = 0; i < items.length; i++) {
          var it = items[i];
          var sku = it && it.sku;
          if (!sku) continue;
          var rel = lookupImg(sku);
          if (!rel) { if (diag.missed.indexOf(String(sku)) < 0) diag.missed.push(String(sku)); markCells.push({ row: itemsRowNum + i, col: imgCol, sku: String(sku), type: 'missing' }); continue; }
          tasks.push({ row: itemsRowNum + i, col: imgCol, sku: String(sku), rel: rel });
        }
        // 导出单次嵌图上限（防浏览器卡死）：超过部分切片跳过，并记录溢出供 UI 警告
        var MAX_IMAGES_PER_EXPORT = 1000;
        diag.overflow = 0;
        if (tasks.length > MAX_IMAGES_PER_EXPORT) {
          diag.overflow = tasks.length - MAX_IMAGES_PER_EXPORT;
          tasks = tasks.slice(0, MAX_IMAGES_PER_EXPORT);
        }
        diag.tasksLen = tasks.length;
        if (!tasks.length) { diag.done = true; resolve(); return; }
        // 3) 并发限流 + 超时 + 降级 + 进度浮层（v1.5.50）
        // 国内网络下同源/GitHub 可能慢，放宽超时；并发仍限 3 防卡死。
        var CONC = 3, TIMEOUT_SAMEORIGIN = 12000, TIMEOUT_RAW = 20000, idx = 0;
        var doneCount = 0, failCount = 0, failedSkus = [];
        // v1.5.50 导出/嵌图进度浮层：图多（>5）时显示「已嵌 X/共 N 失败 Y」，避免用户以为卡死
        var progEl = null;
        if (tasks.length > 5 && typeof document !== 'undefined') {
          progEl = document.createElement('div');
          progEl.id = 'td-embed-progress';
          progEl.style.cssText = 'position:fixed;left:50%;top:20px;transform:translateX(-50%);z-index:99999;background:#0f172a;color:#e2e8f0;padding:10px 16px;border-radius:10px;font:13px/1.5 system-ui,sans-serif;box-shadow:0 6px 24px rgba(0,0,0,.35);min-width:260px;text-align:center';
          progEl.innerHTML = '🖼 正在嵌入产品图 <b>0/' + tasks.length + '</b> …';
          document.body.appendChild(progEl);
        }
        function bump(okImg) {
          doneCount++;
          if (!okImg) failCount++;
          if (progEl) progEl.innerHTML = '🖼 正在嵌入产品图 <b>' + doneCount + '/' + tasks.length + '</b>' + (failCount ? '　⚠️失败 ' + failCount : '') + ' …';
        }
        function worker() {
          if (idx >= tasks.length) return Promise.resolve();
          var t = tasks[idx++];
          return embedOne(wb, ws, t, TIMEOUT_SAMEORIGIN, TIMEOUT_RAW).then(function (okImg) {
            bump(okImg);
            if (!okImg) { failedSkus.push(t.sku); markCells.push({ row: t.row, col: t.col, sku: t.sku, type: 'network' }); }
            return worker();
          });
        }
        var workers = [];
        for (var c = 0; c < Math.min(CONC, tasks.length); c++) workers.push(worker());
        await Promise.all(workers);
        if (progEl) {
          progEl.innerHTML = (failCount ? '⚠️ 产品图嵌入完成（共 ' + tasks.length + '，失败 ' + failCount + '）' : '✅ 产品图嵌入完成（共 ' + tasks.length + '）');
          (function (el) { setTimeout(function () { if (el && el.parentNode) el.parentNode.removeChild(el); }, failCount ? 6000 : 1800); })(progEl);
        }
        // v1.6.8 缺图标注：嵌不上的单元格写入红/黄底「图片缺失」标记，预览与导出一致可见
        for (var mc = 0; mc < markCells.length; mc++) {
          var mk = markCells[mc];
          try {
            var mcell = ws.getCell(mk.row, mk.col);
            var isMiss = mk.type === 'missing';
            mcell.value = '图片缺失\n' + mk.sku + '\n(' + (isMiss ? '图库无图源' : '取图失败') + ')';
            mcell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isMiss ? 'FFFFC7CE' : 'FFFFE699' } };
            mcell.font = { bold: true, size: 9, color: { argb: isMiss ? 'FF9C0006' : 'FF9C5700' } };
            mcell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            var mrow = ws.getRow(mk.row); if (!mrow.height || mrow.height < 46) mrow.height = 46;
          } catch (e) {}
        }
        diag.marked = markCells.length;
        diag.done = true; diag.failed = failedSkus;
        // v1.5.53 诊断增强：把失败 SKU 按「图库缺图(missing)」vs「网络/CORS 拦截(network)」分类。
        // 依据：window.__embDiagFail[sku] 形如 "host@err → host@err → ..."，
        //   若所有候选均为 http4xx(典型 404) ⇒ 图库确实无此文件(missing)；
        //   否则存在 timeout/fetch-err/not-image ⇒ 图可能线上有、但浏览器取不到(network)。
        // UI 据此一眼区分"是没图还是网络卡"，不再笼统写"网络不可达"误导用户。
        try {
          diag.failDetail = {};
          for (var _fi = 0; _fi < failedSkus.length; _fi++) {
            var _fsku = String(failedSkus[_fi]);
            var _raw = (typeof window !== 'undefined' && window.__embDiagFail && window.__embDiagFail[_fsku]) || '';
            var _parts = _raw.split(' → ').filter(Boolean);
            var _errs = _parts.map(function (p) { var m = /@(.+)$/.exec(p); return m ? m[1] : 'fail'; });
            var _all404 = _errs.length > 0 && _errs.every(function (e) { return /^http4\d\d$/.test(e); });
            diag.failDetail[_fsku] = { type: _all404 ? 'missing' : 'network', codes: _errs.slice(0, 6) };
          }
        } catch (e) {}
        try { wb._productImagesEmbedded = true; } catch (e) {}
        resolve();
      } catch (e) { diag.err = String(e && e.stack || e); if (typeof console !== 'undefined') console.error('[emb-err]', e); diag.done = true; resolve(); }
    });
  }

  // v1.5.50 真图校验：拒绝把 HTML 错误页（如 CloudStudio 对缺图路径返回的 200 假页）当 JPEG 嵌进去
  function isRealImage(u8) {
    if (!u8 || u8.length < 4) return false;
    if (u8[0] === 0xFF && u8[1] === 0xD8) return true;            // JPEG
    if (u8[0] === 0x89 && u8[1] === 0x50 && u8[2] === 0x4E && u8[3] === 0x47) return true; // PNG
    if (u8[0] === 0x47 && u8[1] === 0x49 && u8[2] === 0x46) return true; // GIF
    return false;
  }

  function embedOne(wb, ws, t, TIMEOUT_SAMEORIGIN, TIMEOUT_RAW) {
    // v1.5.38 t.rel 可为 URL 数组（同源优先 + raw 兜底）；v1.5.37 前为单 URL
    var rels = (Array.isArray(t.rel) && t.rel.length) ? t.rel : [t.rel];
    if (_imgCache[t.sku]) { applyImg(wb, ws, t, _imgCache[t.sku]); return Promise.resolve(); }
    // v1.5.18 从 rel 扩展名识别格式(jpg/jpeg/png/webp)，兼容文件索引模式；data URL 旧模式按 png 判断
    if (!t.extension) {
      var mm = /\.(png|jpe?g|webp|gif)$/i.exec(String(t.rel || '').slice(0, 300));
      t.extension = mm ? (mm[1].toLowerCase() === 'jpg' ? 'jpeg' : mm[1].toLowerCase()) : ((/^data:image\/png/i.test(String(t.rel || ''))) ? 'png' : 'jpeg');
    }
    // v1.5.39 双超时策略：同源 8s 内必须返回（国内 GitHub Pages 通常 <3s；超时说明同源网络很差/被打），立刻切 raw 给 15s
    return new Promise(function (resolve) {
      function fetchOne(url, ms, i) {
        return new Promise(function (res) {
          var settled = false;
          var timer = setTimeout(function () {
            if (!settled) { settled = true; res({ ok: false, err: 'timeout', nextIdx: i + 1 }); }
          }, ms);
          // v1.6.6 防盗链：sursung/部分图床对带 Referer 的跨域请求返回 403 无 CORS 头 → 浏览器拦截。
          //   referrerPolicy:'no-referrer' 让浏览器不带 Referer，图床按无 Referer 放行(200+ACAO=*)。
          fetch(url, { referrerPolicy: 'no-referrer' }).then(function (r) {
            if (settled) return;
            if (!r.ok) { settled = true; clearTimeout(timer); res({ ok: false, err: 'http' + r.status, nextIdx: i + 1 }); return; }
            return r.arrayBuffer();
          }).then(async function (buf) {
            if (settled || buf === undefined) return;
            settled = true; clearTimeout(timer);
            var u8 = (buf instanceof Uint8Array) ? buf : new Uint8Array(buf);
            // v1.5.50 真图校验：假图（HTML 错误页）直接判失败，跳到下一个镜像候选，不嵌坏图
            if (!isRealImage(u8)) {
              // v1.6.2 兜底：webp 伪装文件（扩展名 .jpeg/.png 但内容实为 RIFF/WebP）→ 浏览器端 canvas 实时转 jpeg 再嵌
              //   根因：线上 sku_thumb 历史上混入 webp 伪装（恢复 git 历史 blob 时引入），engine 此前判 not-image 失败导致导出缺图。
              //   ExcelJS 不支持 webp extension，但浏览器 createImageBitmap 支持 webp 解码+canvas→jpeg，
              //   故在浏览器端即时转码，不动任何图片文件，零存储修复（如 8T117-31 即可恢复导出）。
              if (typeof createImageBitmap !== 'undefined' && u8[0] === 0x52 && u8[1] === 0x49 && u8[2] === 0x46 && u8[3] === 0x46) {
                try {
                  var blob = new Blob([u8]);
                  var bm = await createImageBitmap(blob);
                  var c = document.createElement('canvas');
                  c.width = bm.width; c.height = bm.height;
                  c.getContext('2d').drawImage(bm, 0, 0);
                  var jb = await new Promise(function (res2) { c.toBlob(res2, 'image/jpeg', 0.9); });
                  u8 = new Uint8Array(await jb.arrayBuffer());
                  t.extension = 'jpeg';
                } catch (e) { res({ ok: false, err: 'webp-decode-fail', nextIdx: i + 1 }); return; }
              } else {
                res({ ok: false, err: 'not-image', nextIdx: i + 1 }); return;
              }
            }
            // v1.6.5 大图压缩：>300KB 的图（如 sursung 3MB PNG）下载后 canvas 缩放最长边≤480px 转 jpeg。
            //   原因：飞书表里部分变体图是 ERP 大 PNG（1~3MB），国内访问慢 → 下载超时/嵌入 Excel 卡死。
            //   压缩后约 30~80KB，下载快、嵌入稳；压缩失败用原图继续（不阻断导出）。
            if (u8.length > 300 * 1024 && typeof createImageBitmap !== 'undefined' && typeof document !== 'undefined') {
              try {
                var blobL = new Blob([u8]);
                var bmL = await createImageBitmap(blobL);
                var maxS = 480;
                var sc = Math.min(1, maxS / Math.max(bmL.width, bmL.height));
                var cwL = Math.max(1, Math.round(bmL.width * sc));
                var chL = Math.max(1, Math.round(bmL.height * sc));
                var cL = document.createElement('canvas');
                cL.width = cwL; cL.height = chL;
                cL.getContext('2d').drawImage(bmL, 0, 0, cwL, chL);
                var jbL = await new Promise(function (res3) { cL.toBlob(res3, 'image/jpeg', 0.85); });
                if (jbL) { u8 = new Uint8Array(await jbL.arrayBuffer()); t.extension = 'jpeg'; }
              } catch (e) {}
            }
            _imgCache[t.sku] = u8;
            applyImg(wb, ws, t, u8);
            res({ ok: true });
          }).catch(function () {
            if (settled) return;
            settled = true; clearTimeout(timer);
            res({ ok: false, err: 'fetch-err', nextIdx: i + 1 });
          });
        });
      }
      // v1.5.52 候选链：每候选独立超时(按其在 MAP 中的优先级), 单候选失败重试 1 次(退避),
      // 全部候选穷尽才记失败。超时大幅放宽(同域/GH 30s, CDN 15s, raw 45s), 不再因限流/抖动轻易放弃。
      var PLAN_MS = [30000, 30000, 15000, 45000, 30000];
      var plan = [];
      for (var ri = 0; ri < rels.length; ri++) {
        plan.push({ url: rels[ri], ms: PLAN_MS[ri] || 20000 });
      }
      function step(i, p, attempt) {
        attempt = attempt || 0;
        if (i >= plan.length) { resolve(false); return; }
        fetchOne(p.url, p.ms, i).then(function (r) {
          if (r.ok) { resolve(true); return; }
          // 诊断：记录每个失败候选
          if (typeof window !== 'undefined') {
            try { if (!window.__embDiagFail) window.__embDiagFail = {}; window.__embDiagFail[t.sku] = (window.__embDiagFail[t.sku] || '') + (plan[i].url.replace(/^https?:\/\//, '') + '@' + (r.err || 'fail') + ' → '); } catch (e) {}
          }
          // 单候选重试 1 次(指数退避 1.5s), 仍失败再跳下一候选
          if (attempt < 1) {
            setTimeout(function () { step(i, p, attempt + 1); }, 1500 * (attempt + 1));
            return;
          }
          step(i + 1, plan[i + 1] || { url: rels[0], ms: 30000 }, 0);
        });
      }
      step(0, plan[0], 0);
    });
  }

  function applyImg(wb, ws, t, bytes) {
    try {
      // v1.5.14 进格：用 twoCellAnchor 让图片精确填满所在单元格（tl=单元格左上、br=单元格右下），
      // 图片边界=单元格边界，彻底解决"图飘在表格外/不在单元格内"。复用代码里 logo 已验证的 tl/br 写法。
      // （之前 oneCellAnchor 写死 92px 且 colOff/rowOff=0 钉在单元格左上角，单元格比图小时图溢出到外侧。）
      var imgId = wb.addImage({ buffer: bytes, extension: t.extension || 'jpeg' });
      var c0 = t.col - 1, r0 = t.row - 1;
      ws.addImage(imgId, { tl: { col: c0, row: r0 }, br: { col: c0 + 1, row: r0 + 1 } });
      // 抬高行高 + 产品图片列宽，使单元格接近正方、图片不变形且清晰（导出与预览一致）
      try { var rr = ws.getRow(t.row); if (!rr.height || rr.height < 60) rr.height = 60; } catch (e) {}
      try { var ic = ws.getColumn(t.col); if (!ic.width || ic.width < 12) ic.width = 12; } catch (e) {}
      var cell = ws.getCell(t.row, t.col);
      if (cell && cell.value !== null && cell.value !== undefined && cell.value !== '') {
        try { cell.value = ''; } catch (e) {}
      }
    } catch (e) {}
  }

  // ---------- 内置模板（骨架版，等真实样例后替换） ----------
  function _hdr(ws, cell, text, size, bold) {
    var c = ws.getCell(cell);
    c.value = text;
    c.font = { size: size || 10, bold: !!bold, name: 'Arial' };
  }
  var thin = { style: 'thin' };
  function _box(ws, range) {
    // range like 'A1:H1'
    ws.getCell(range.split(':')[0]).border = { top: thin, left: thin, bottom: thin, right: thin };
  }

  /** 通用商业发票模板（Commercial Invoice + 装箱信息） */
  function makeBuiltinInvoiceTemplate(ExcelJS) {
    var wb = new ExcelJS.Workbook();
    var ws = wb.addWorksheet('INVOICE');
    ws.columns = [
      { width: 6 }, { width: 18 }, { width: 30 }, { width: 14 }, { width: 10 }, { width: 12 }, { width: 14 }, { width: 12 }, { width: 12 }, { width: 11 }
    ];
    ws.mergeCells('A1:I1'); _hdr(ws, 'A1', 'COMMERCIAL INVOICE', 16, true);
    ws.getCell('A1').alignment = { horizontal: 'center' };
    _hdr(ws, 'A3', 'INVOICE NO.:', 10, true); _hdr(ws, 'C3', '{{invoiceNo}}');
    _hdr(ws, 'F3', 'DATE:', 10, true); _hdr(ws, 'G3', '{{invoiceDate}}');
    _hdr(ws, 'A4', 'CONTRACT NO.:', 10, true); _hdr(ws, 'C4', '{{contractNo}}');
    _hdr(ws, 'F4', 'ORDER NO.:', 10, true); _hdr(ws, 'G4', '{{orderNos}}');
    _hdr(ws, 'A6', 'SHIPPER:', 10, true);
    ws.mergeCells('B6:E6'); _hdr(ws, 'B6', '{{shipper.name}}');
    ws.mergeCells('B7:E7'); _hdr(ws, 'B7', '{{shipper.address}}');
    ws.mergeCells('B8:E8'); _hdr(ws, 'B8', 'TEL: {{shipper.tel}}');
    _hdr(ws, 'F6', 'CONSIGNEE:', 10, true);
    ws.mergeCells('G6:I6'); _hdr(ws, 'G6', '{{consignee.name}}');
    ws.mergeCells('G7:I7'); _hdr(ws, 'G7', '{{consignee.address}}');
    ws.mergeCells('G8:I8'); _hdr(ws, 'G8', 'TEL: {{consignee.tel}}');
    _hdr(ws, 'A9', 'NOTIFY PARTY:', 10, true);
    ws.mergeCells('B9:E9'); _hdr(ws, 'B9', '{{notify.name}}');
    _hdr(ws, 'A11', 'FROM:', 10, true); _hdr(ws, 'B11', '{{pol}}');
    _hdr(ws, 'D11', 'TO:', 10, true); _hdr(ws, 'E11', '{{pod}}');
    _hdr(ws, 'F11', 'BY:', 10, true); _hdr(ws, 'G11', '{{transport}}');
    _hdr(ws, 'A12', 'PRICE TERMS:', 10, true); _hdr(ws, 'B12', '{{incoterms}}');
    _hdr(ws, 'D12', 'PAYMENT:', 10, true); _hdr(ws, 'E12', '{{paymentTerms}}');
    _hdr(ws, 'F12', 'MARKS:', 10, true); _hdr(ws, 'G12', '{{shippingMarks}}');
    // 表头
    var headRow = 14;
    var heads = ['NO.', 'SKU / MODEL', 'DESCRIPTION OF GOODS', 'HS CODE', 'QTY', 'UNIT', 'UNIT PRICE ({{currency}})', 'AMOUNT ({{currency}})', 'N.W.(KG)', '产品图片'];
    heads.forEach(function (h, i) {
      var c = ws.getRow(headRow).getCell(i + 1);
      c.value = h; c.font = { size: 9, bold: true, name: 'Arial' };
      c.border = { top: thin, left: thin, bottom: thin, right: thin };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    // 明细占位行（含产品图片列，详情行加高以容纳缩略图）
    ws.getRow(15).height = 44;
    var vals = ['{{items.no}}', '{{items.model}}', '{{items.nameEn}}', '{{items.hsCode}}', '{{items.qty}}', '{{items.unit}}', '{{items.price}}', '{{items.amount}}', '{{items.nw}}', ''];
    vals.forEach(function (v, i) {
      var c = ws.getRow(15).getCell(i + 1);
      c.value = v; c.font = { size: 9, name: 'Arial' };
      c.border = { top: thin, left: thin, bottom: thin, right: thin };
      c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
    });
    // 合计
    _hdr(ws, 'A17', 'TOTAL:', 10, true);
    _hdr(ws, 'E17', '{{totals.qty}}', 10, true);
    _hdr(ws, 'H17', '{{totals.amount}}', 10, true);
    _hdr(ws, 'I17', '{{totals.nw}}', 10, true);
    _hdr(ws, 'A18', 'TOTAL PACKAGES:', 10, true); _hdr(ws, 'C18', '{{totals.boxCount}} CTNS');
    _hdr(ws, 'D18', 'G.W.:', 10, true); _hdr(ws, 'E18', '{{totals.gw}} KG');
    _hdr(ws, 'F18', 'VOLUME:', 10, true); _hdr(ws, 'G18', '{{totals.volume}} CBM');
    ws.mergeCells('A19:I19'); _hdr(ws, 'A19', '{{amountInWords}}', 10, true);
    _hdr(ws, 'A21', 'REMARK:', 10, true); ws.mergeCells('B21:I21'); _hdr(ws, 'B21', '{{remark}}');
    ws.mergeCells('F23:I23'); _hdr(ws, 'F23', 'SIGNATURE: ____________________', 10, true);
    return wb;
  }

  /** 通用订舱单模板（BOOKING FORM） */
  function makeBuiltinBookingTemplate(ExcelJS) {
    var wb = new ExcelJS.Workbook();
    var ws = wb.addWorksheet('BOOKING FORM');
    ws.columns = [{ width: 16 }, { width: 26 }, { width: 16 }, { width: 26 }, { width: 14 }, { width: 14 }];
    ws.mergeCells('A1:F1'); _hdr(ws, 'A1', 'SHIPPING BOOKING FORM 订舱委托书', 15, true);
    ws.getCell('A1').alignment = { horizontal: 'center' };
    _hdr(ws, 'A3', 'DATE 日期:', 10, true); _hdr(ws, 'B3', '{{invoiceDate}}');
    _hdr(ws, 'C3', 'REF NO. 委托号:', 10, true); _hdr(ws, 'D3', '{{invoiceNo}}');
    _hdr(ws, 'A5', 'SHIPPER 托运人:', 10, true);
    ws.mergeCells('B5:F5'); _hdr(ws, 'B5', '{{shipper.name}}');
    ws.mergeCells('B6:F6'); _hdr(ws, 'B6', '{{shipper.address}}  TEL: {{shipper.tel}}');
    _hdr(ws, 'A7', 'CONSIGNEE 收货人:', 10, true);
    ws.mergeCells('B7:F7'); _hdr(ws, 'B7', '{{consignee.name}}');
    ws.mergeCells('B8:F8'); _hdr(ws, 'B8', '{{consignee.address}}  TEL: {{consignee.tel}}');
    _hdr(ws, 'A9', 'NOTIFY 通知人:', 10, true);
    ws.mergeCells('B9:F9'); _hdr(ws, 'B9', '{{notify.name}}');
    _hdr(ws, 'A11', 'POL 起运港:', 10, true); _hdr(ws, 'B11', '{{pol}}');
    _hdr(ws, 'C11', 'POD 目的港:', 10, true); _hdr(ws, 'D11', '{{pod}}');
    _hdr(ws, 'A12', 'ETD 船期:', 10, true); _hdr(ws, 'B12', '{{etd}}');
    _hdr(ws, 'C12', 'VESSEL 船名航次:', 10, true); _hdr(ws, 'D12', '{{vessel}}');
    _hdr(ws, 'A13', 'CNTR 柜型柜量:', 10, true); _hdr(ws, 'B13', '{{containerType}} x {{containerQty}}');
    _hdr(ws, 'C13', 'TERMS 贸易条款:', 10, true); _hdr(ws, 'D13', '{{incoterms}}');
    _hdr(ws, 'A14', 'FREIGHT 运费条款:', 10, true); _hdr(ws, 'B14', '{{freightTerms}}');
    _hdr(ws, 'C14', 'AGENT 订舱代理:', 10, true); _hdr(ws, 'D14', '{{agent}}');
    _hdr(ws, 'A16', 'GOODS 品名概述:', 10, true);
    ws.mergeCells('B16:F16'); _hdr(ws, 'B16', '{{goodsSummary}}');
    _hdr(ws, 'A17', 'PACKAGES 总件数:', 10, true); _hdr(ws, 'B17', '{{totals.boxCount}} CTNS');
    _hdr(ws, 'C17', 'G.W. 总毛重:', 10, true); _hdr(ws, 'D17', '{{totals.gw}} KG');
    _hdr(ws, 'E17', 'CBM 总体积:', 10, true); _hdr(ws, 'F17', '{{totals.volume}}');
    _hdr(ws, 'A18', 'MARKS 唛头:', 10, true); ws.mergeCells('B18:F18'); _hdr(ws, 'B18', '{{shippingMarks}}');
    _hdr(ws, 'A19', 'DG 危险品声明:', 10, true); ws.mergeCells('B19:F19'); _hdr(ws, 'B19', '{{dangerous}}');
    _hdr(ws, 'A20', 'CUSTOMS 报关方式:', 10, true); _hdr(ws, 'B20', '{{customsType}}');
    _hdr(ws, 'A21', 'REMARK 备注:', 10, true); ws.mergeCells('B21:F21'); _hdr(ws, 'B21', '{{remark}}');
    ws.mergeCells('D23:F23'); _hdr(ws, 'D23', 'AUTHORIZED SIGNATURE: ____________', 10, true);
    return wb;
  }

  /** 各模板类型的默认必填字段（校验用） */
  var REQUIRED_FIELDS = {
    invoice: [
      { path: 'invoiceNo', label: '发票号' },
      { path: 'invoiceDate', label: '发票日期' },
      { path: 'shipper.name', label: 'SHIPPER名称' },
      { path: 'consignee.name', label: 'CONSIGNEE名称' },
      { path: 'incoterms', label: '贸易条款' },
      { path: 'pol', label: '起运港' },
      { path: 'pod', label: '目的港' },
      { path: 'currency', label: '币种' }
    ],
    booking: [
      { path: 'invoiceNo', label: '委托号' },
      { path: 'shipper.name', label: '托运人' },
      { path: 'consignee.name', label: '收货人' },
      { path: 'pol', label: '起运港' },
      { path: 'pod', label: '目的港' },
      { path: 'freightTerms', label: '运费条款' }
    ],
    packing: [
      { path: 'totals.boxCount', label: '总箱数' },
      { path: 'totals.gw', label: '总毛重' },
      { path: 'items.nameCn', label: '中文品名' }
    ],
    declare: [
      { path: 'items.nameCn', label: '中文品名' },
      { path: 'items.hsCode', label: '海关编码' },
      { path: 'items.qty', label: '申报数量' }
    ]
  };

  return {
    PH_RE: PH_RE,
    buildDocData: buildDocData,
    amountInWords: amountInWords,
    scanTemplate: scanTemplate,
    buildLabelMap: buildLabelMap,
    mapFieldLabel: mapFieldLabel,
    fillByFieldLabels: fillByFieldLabels,
    fillTemplate: fillTemplate,
    embedProductImages: embedProductImages,
    addLogo: addLogo,
    makeBuiltinInvoiceTemplate: makeBuiltinInvoiceTemplate,
    makeBuiltinBookingTemplate: makeBuiltinBookingTemplate,
    REQUIRED_FIELDS: REQUIRED_FIELDS
  };
});
