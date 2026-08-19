// 生成图库审核相册 HTML: SKU + 图 + 商品名(declareData) + 源表系列，用户逐张核对
const fs = require('fs');
const ROOT = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统';

// 1. 索引
let raw = fs.readFileSync(ROOT + '/js/sku_image_index.js', 'utf8');
if (raw.charCodeAt(0) === 0xFEFF) raw = raw.slice(1);
const eqIdx = raw.indexOf('window.SKU_IMAGE_INDEX = ');
const scIdx = raw.indexOf('};', eqIdx);
const idx = JSON.parse(raw.substring(eqIdx + 'window.SKU_IMAGE_INDEX = '.length, scIdx + 1));

// 2. declareData 品名
const names = {};
for (let i = 1; i <= 16; i++) {
  try { eval(fs.readFileSync(ROOT + '/js/declare_data.part' + String(i).padStart(2, '0') + '.js', 'utf8')); } catch (e) {}
}
const dlist = global.window && window.TD && window.TD.declareData || [];
dlist.forEach(d => { if (d && d.sku && !names[d.sku]) names[d.sku] = (d.nameCn || d.shortName || d.goodsName || ''); });

// 3. 相册 HTML
const cards = Object.keys(idx).sort().map(sku => {
  const f = idx[sku];
  const nm = (names[sku] || '').replace(/"/g, '&quot;');
  const series = sku.replace(/-?\d+$/, '');
  return '<div class="card" data-series="' + series + '"><div class="imgbox"><img loading="lazy" src="images/sku_thumb/' + encodeURIComponent(f) + '" onerror="this.parentElement.classList.add(\'bad\');this.style.display=\'none\'"></div>' +
    '<div class="sku">' + sku + '</div><div class="nm">' + nm + '</div>' +
    '<label class="chk"><input type="checkbox" class="badck">错图</label></div>';
}).join('');

const html = `<!DOCTYPE html><html lang="zh"><head><meta charset="utf-8"><title>图库审核相册（${Object.keys(idx).length} 张）</title>
<style>
body{font-family:system-ui;background:#f5f6fa;margin:0;padding:16px}
h1{font-size:18px}
.toolbar{position:sticky;top:0;background:#fff;padding:8px 12px;border-bottom:1px solid #ddd;z-index:9;display:flex;gap:12px;align-items:center;flex-wrap:wrap}
.toolbar input{min-width:220px;padding:6px 10px}
.btn{padding:6px 14px;border:1px solid #ccc;background:#fff;cursor:pointer;border-radius:4px}
.btn.primary{background:#e74c3c;color:#fff;border-color:#e74c3c}
.gallery{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px;margin-top:12px}
.card{border:1px solid #e3e5ea;border-radius:6px;background:#fff;padding:8px;position:relative}
.card.bad{border-color:#e74c3c}
.imgbox{height:110px;display:flex;align-items:center;justify-content:center;background:#fafafa;border-radius:4px;overflow:hidden}
.imgbox img{max-width:100%;max-height:100%}
.sku{font-weight:600;font-size:13px;margin-top:6px;word-break:break-all}
.nm{font-size:11px;color:#888;height:28px;overflow:hidden;margin-top:2px}
.chk{font-size:12px;color:#e74c3c;display:flex;align-items:center;gap:4px;margin-top:4px}
.cnt{color:#666;font-size:13px}
</style></head><body>
<div class="toolbar">
  <h1>图库审核相册（<span class="cnt" id="total">${Object.keys(idx).length}</span> 张 · 勾选"错图"后点导出）</h1>
  <input id="search" placeholder="搜索 SKU / 系列（如 2T78 或 Thea）">
  <button class="btn" id="onlybad">只看错图</button>
  <button class="btn primary" id="export">导出勾选的错图SKU列表</button>
  <span class="cnt" id="selcount">已选 0</span>
</div>
<div class="gallery" id="g">${cards}</div>
<script>
const search=document.getElementById('search'), g=document.getElementById('g'), onlybad=document.getElementById('onlybad');
const cards=[...g.querySelectorAll('.card')];
const sel=()=>{document.getElementById('selcount').textContent='已选 '+cards.filter(c=>c.querySelector('.badck').checked).length;};
cards.forEach(c=>c.querySelector('.badck').addEventListener('change',sel));
search.addEventListener('input',()=>{const q=search.value.trim().toLowerCase();cards.forEach(c=>{c.style.display=(!q||c.dataset.series.includes(q)||c.querySelector('.sku').textContent.toLowerCase().includes(q)||c.querySelector('.nm').textContent.toLowerCase().includes(q))?'':'none';});});
onlybad.addEventListener('click',()=>{const b=onlybad.textContent.includes('只看错图');onlybad.textContent=b?'全部':'只看错图';cards.forEach(c=>c.style.display=(b?(!c.querySelector('.badck').checked):true)?'':'none');});
document.getElementById('export').addEventListener('click',()=>{const bad=cards.filter(c=>c.querySelector('.badck').checked).map(c=>c.querySelector('.sku').textContent);if(!bad.length){alert('未勾选任何错图');return;}const blob=new Blob([bad.join('\\n')],{type:'text/plain'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download='bad_skus.txt';a.click();});
</script></body></html>`;

fs.writeFileSync(ROOT + '/tests/_out/图库审核相册.html', html, 'utf8');
console.log('生成: tests/_out/图库审核相册.html |', Object.keys(idx).length, '张图');
