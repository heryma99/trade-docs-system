// v1.5.35 部署: 2T78-1 正确图(商品资料申报要素表 image1648) + sku_image_index 加回 + index bump
const fs = require('fs'), https = require('https');
const TOKEN = fs.readFileSync('C:/Users/cn/.workbuddy/connectors/3fe83c35-d7d3-4e71-869e-097580283ed4/tokens/github.txt', 'utf8').trim();
function req(m, p, b) {
  return new Promise((res, rej) => {
    const d = b ? JSON.stringify(b) : null;
    const r = https.request({ host: 'api.github.com', path: p, method: m, headers: { 'User-Agent': 'node', 'Authorization': 'token ' + TOKEN, 'Content-Type': 'application/json' } }, x => {
      let s = ''; x.on('data', c => s += c);
      x.on('end', () => { let j = null; try { j = JSON.parse(s); } catch (e) {} res({ s: x.statusCode, j }); });
    });
    r.on('error', rej); r.setTimeout(120000, () => r.destroy(new Error('TO'))); if (d) r.write(d); r.end();
  });
}
async function putFile(path, localPath, msg) {
  const g = await req('GET', '/repos/heryma99/trade-docs-system/contents/' + encodeURIComponent(path) + '?ref=main');
  const content = fs.readFileSync(localPath).toString('base64');
  const r = await req('PUT', '/repos/heryma99/trade-docs-system/contents/' + encodeURIComponent(path), { message: msg, content, sha: g.j && g.j.sha || null, branch: 'main' });
  console.log(path, '->', r.s);
}
(async () => {
  await putFile('images/sku_thumb/2T78-1.jpg', 'images/sku_thumb/2T78-1.jpg', 'v1.5.35 2T78-1正确图(商品资料申报要素表image1648, Thea黑手提包)');
  await putFile('js/sku_image_index.js', 'js/sku_image_index.js', 'v1.5.35 2T78-1加回正确图');
  await putFile('index.html', 'index.html', 'v1.5.35 bump sku_image_index?v=1.5.35');
})().catch(e => { console.error('ERR', e.message); process.exit(1); });
