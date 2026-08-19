const fs = require('fs'), path = require('path');
const ROOT = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统';
const SRC_DIR = path.join(ROOT, 'images/_jwpei_ok');
const OUT_DIR = path.join(ROOT, 'images/sku_thumb');
const TGT_DIR = path.join(ROOT, 'tds_dist/images/sku_thumb');

// 1) 清空旧图库(产品图)
function rmdir(d) { if (fs.existsSync(d)) fs.rmSync(d, { recursive: true, force: true }); }
rmdir(OUT_DIR);
rmdir(TGT_DIR);
fs.mkdirSync(OUT_DIR, { recursive: true });
fs.mkdirSync(TGT_DIR, { recursive: true });

// 2) 复制 JW PEI 真图到 sku_thumb + tds_dist
const idx = {};
const files = fs.readdirSync(SRC_DIR);
for (const f of files) {
  if (f.endsWith('.jpg') || f.endsWith('.jpeg') || f.endsWith('.png') || f.endsWith('.webp')) {
    const sku = path.parse(f).name;
    fs.copyFileSync(path.join(SRC_DIR, f), path.join(OUT_DIR, f));
    fs.copyFileSync(path.join(SRC_DIR, f), path.join(TGT_DIR, f));
    idx[sku] = f;
  }
}
console.log('已复制', Object.keys(idx).length, '张真图(SPU级)');

// 3) 写 sku_image_index.js (window.SKU_IMAGE_INDEX)
const sortedIdx = {};
Object.keys(idx).sort().forEach(k => sortedIdx[k] = idx[k]);
const body = '// 自动生成: ' + new Date().toISOString().slice(0, 10) + ' 清空旧图库(全部占位错图)后, 从 D:/模板/SKU申报信息表/JW PEI 产品基础数据 G Unit.xlsx 重新抽取 SPU 级真图(' + Object.keys(idx).length + '张, 全部已 Read 核对内容正确, 非占位图)。engine 加 SPU 前缀兜底: 无图 SKU 自动用同 SPU 变体图(同款不同尺码共用是正确行为, 非错图)。\nwindow.SKU_IMAGE_INDEX = ' + JSON.stringify(sortedIdx) + ';\n';
fs.writeFileSync(path.join(ROOT, 'js/sku_image_index.js'), body);
fs.writeFileSync(path.join(ROOT, 'tds_dist/js/sku_image_index.js'), body);
console.log('js/sku_image_index.js 已写:', Object.keys(idx).length, '键');

// 4) 清空旧 product_image_map.js (v1.5.42 已停用, 这里彻底从磁盘清掉)
const pim = path.join(ROOT, 'js/product_image_map.js');
const pimTds = path.join(ROOT, 'tds_dist/js/product_image_map.js');
if (fs.existsSync(pim)) { fs.writeFileSync(pim, '// v1.5.42+ 弃用此老图库 (444 SKU 共用错图已删), 见 sku_image_index.js\nwindow.PRODUCT_IMAGE_MAP = {};\n'); console.log('product_image_map.js 已清空(留空壳)'); }
if (fs.existsSync(pimTds)) { fs.writeFileSync(pimTds, '// v1.5.42+ 弃用此老图库 (444 SKU 共用错图已删), 见 sku_image_index.js\nwindow.PRODUCT_IMAGE_MAP = {};\n'); }

// 5) 清空 images/products (老图库源, 全部占位错图)
rmdir(path.join(ROOT, 'images/products'));
rmdir(path.join(ROOT, 'images/products_sku'));
console.log('images/products/ products_sku/ 已清空');

console.log('=== 完成 ===');
console.log('OUT_DIR:', OUT_DIR, '(', fs.readdirSync(OUT_DIR).length, '文件)');
console.log('TGT_DIR:', TGT_DIR, '(', fs.readdirSync(TGT_DIR).length, '文件)');