const fs = require('fs'), path = require('path'), https = require('https');
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
  console.log('  ' + repoPath.padEnd(50) + ' -> HTTP ' + r.s + (r.s === 200 || r.s === 201 ? ' ✅' : ' ❌ ' + JSON.stringify(r.j).slice(0, 80)));
  return r.s === 200 || r.s === 201;
}
(async () => {
  let idx = fs.readFileSync(ROOT + '/index.html', 'utf8');
  idx = idx.replace(/>v1\.5\.4[0123]/g, '>v1.5.44<');
  idx = idx.replace(/\?v=1\.5\.4[0123]/g, '?v=1.5.44');
  fs.writeFileSync(ROOT + '/index.html', idx);
  console.log('index.html bump -> v1.5.44');
  console.log('=== PUT GitHub Pages ===');
  await putFile('js/engine.js', ROOT + '/js/engine.js', 'v1.5.44 清空旧图库(444错图+全部占位)+ SPU前缀降级(同款不同尺码共用主款图)');
  await putFile('js/sku_image_index.js', ROOT + '/js/sku_image_index.js', 'v1.5.44 清空后从JW PEI G Unit抽153张SPU真图');
  await putFile('js/product_image_map.js', ROOT + '/js/product_image_map.js', 'v1.5.44 老图库空壳(已停用, 兼容旧索引加载)');
  await putFile('index.html', ROOT + '/index.html', 'v1.5.44 bump 顶部版本');
  // 153 张图 → GitHub Contents API 不能批量(易429), 但用户已确认主要链路, 这次走 git push 路径
  // 改为: git add images/sku_thumb 然后提交, 但 git 损坏 → 用 Contents API 单文件逐上传
  const imgDir = ROOT + '/images/sku_thumb';
  const files = fs.readdirSync(imgDir).filter(f => /\.(jpe?g|png|webp)$/i.test(f));
  console.log('图片文件数:', files.length);
  // 先批量上传 5 个试, 剩余的写 git commit 脚本
  let ok = 0;
  for (const f of files.slice(0, 5)) {
    const r = await putFile('images/sku_thumb/' + f, path.join(imgDir, f), 'v1.5.44 ' + f);
    if (r) ok++;
  }
  console.log('首批 5 张上传 OK:', ok);
})().catch(e => { console.error('FATAL', e); process.exit(1); });