const fs = require('fs'), https = require('https');
const TOKEN = fs.readFileSync('C:/Users/cn/.workbuddy/connectors/3fe83c35-d7d3-4e71-869e-097580283ed4/tokens/github.txt', 'utf8').trim();
const ROOT = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统';
const ORG = 'heryma99', REPO = 'trade-docs-system', BRANCH = 'main';
function req(method, path, body) {
  return new Promise((res, rej) => {
    const data = body ? JSON.stringify(body) : null;
    const r = https.request({ host: 'api.github.com', path, method, headers: { 'User-Agent': 'node', 'Authorization': 'token ' + TOKEN, 'Content-Type': 'application/json' } },
      x => { let s = ''; x.on('data', c => s += c); x.on('end', () => { let j = null; try { j = JSON.parse(s); } catch (e) {} res({ s: x.statusCode, j }); }); });
    r.on('error', rej); r.setTimeout(180000, () => r.destroy(new Error('TO'))); if (data) r.write(data); r.end();
  });
}
async function putFile(repoPath, localPath, msg) {
  const g = await req('GET', `/repos/${ORG}/${REPO}/contents/${encodeURIComponent(repoPath)}?ref=${BRANCH}`);
  const sha = (g.s === 200 && g.j) ? g.j.sha : null;
  const content = fs.readFileSync(localPath).toString('base64');
  const r = await req('PUT', `/repos/${ORG}/${REPO}/contents/${encodeURIComponent(repoPath)}`, { message: msg, content, sha, branch: BRANCH });
  console.log('  ' + repoPath.padEnd(42) + ' -> HTTP ' + r.s + (r.s === 200 || r.s === 201 ? ' ✅' : ' ❌ ' + JSON.stringify(r.j).slice(0, 100)));
  return r.s === 200 || r.s === 201;
}
(async () => {
  let idx = fs.readFileSync(ROOT + '/index.html', 'utf8');
  idx = idx.replace(/>v1\.5\.40</g, '>v1.5.42<');
  idx = idx.replace(/>v1\.5\.41</g, '>v1.5.42<');
  idx = idx.replace(/\?v=1\.5\.4[01]/g, '?v=1.5.42');
  fs.writeFileSync(ROOT + '/index.html', idx);
  console.log('index.html bump -> v1.5.42');
  console.log('=== PUT GitHub Pages ===');
  await putFile('js/engine.js', ROOT + '/js/engine.js', 'v1.5.42 弃用旧图库(444 SKU共用一张错图)只认SKU_IMAGE_INDEX 435键');
  await putFile('js/ui.js', ROOT + '/js/ui.js', 'v1.5.42 自动同步默认开(auto===false才停) + 启动拉取合并一次 + watchdog 25s');
  await putFile('js/sku_image_index.js', ROOT + '/js/sku_image_index.js', 'v1.5.42 740-7换正确图(河豚手拿包), DS1340-8删除(连衣裙无正确图) 435键');
  await putFile('images/sku_thumb/740-7.jpg', ROOT + '/images/sku_thumb/740-7.jpg', 'v1.5.42 740-7 正确图(商品资料表image747 河豚手拿包)');
  console.log('=== 同步 tds_dist ===');
  for (const f of ['js/engine.js', 'js/ui.js', 'js/sku_image_index.js', 'index.html']) {
    fs.copyFileSync(ROOT + '/' + f, ROOT + '/tds_dist/' + f);
  }
  fs.copyFileSync(ROOT + '/images/sku_thumb/740-7.jpg', ROOT + '/tds_dist/images/sku_thumb/740-7.jpg');
  // 拉最新 userdata.json 到 tds_dist
  const raw = await new Promise((res, rej) => {
    const r = https.request('https://raw.githubusercontent.com/heryma99/trade-docs-system/main/userdata.json', x => { let d = ''; x.on('data', c => d += c); x.on('end', () => res(d)); });
    r.on('error', rej); r.end();
  });
  fs.writeFileSync(ROOT + '/tds_dist/userdata.json', raw);
  console.log('tds_dist 同步完成(engine/ui/sku_image_index/index/740-7/userdata)');
})().catch(e => { console.error('FATAL', e); process.exit(1); });