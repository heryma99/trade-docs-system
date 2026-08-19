// 部署 v1.5.20: 改 fetch 模式(SKU_IMAGE_INDEX→images/sku_thumb/), 删除 pack
const https = require('https');
const fs = require('fs');
const TOKEN = fs.readFileSync('C:/Users/cn/.workbuddy/connectors/3fe83c35-d7d3-4e71-869e-097580283ed4/tokens/github.txt', 'utf8').trim();
const OWNER = 'heryma99';
const REPO = 'trade-docs-system';
const BRANCH = 'main';
function req(method, path, body) {
  return new Promise((resolve, reject) => {
    const data = body ? JSON.stringify(body) : null;
    const r = https.request({ host: 'api.github.com', path, method, headers: { 'User-Agent': 'node', 'Authorization': 'token ' + TOKEN, 'Accept': 'application/vnd.github+json', 'Content-Type': 'application/json' } }, res => {
      let d = ''; res.on('data', c => d += c);
      res.on('end', () => { let json = null; try { json = JSON.parse(d); } catch (e) {} resolve({ status: res.statusCode, json, raw: d }); });
    });
    r.on('error', e => reject(e));
    r.setTimeout(60000, () => { r.destroy(new Error('TIMEOUT')); });
    if (data) r.write(data); r.end();
  });
}
async function getSha(path) {
  const { status, json } = await req('GET', `/repos/${OWNER}/${REPO}/contents/${path}?ref=${BRANCH}`);
  if (status === 200 && json && json.sha) return json.sha;
  if (status === 404) return null;
  throw new Error(`GET ${path} -> ${status} ${JSON.stringify(json).slice(0, 200)}`);
}
async function putFile(localPath, repoPath, message) {
  const b64 = fs.readFileSync(localPath).toString('base64');
  const sha = await getSha(repoPath);
  const body = { message, content: b64, branch: BRANCH };
  if (sha) body.sha = sha;
  const { status, json } = await req('PUT', `/repos/${OWNER}/${REPO}/contents/${repoPath}`, body);
  if (status !== 200 && status !== 201) throw new Error(`PUT ${repoPath} -> ${status} ${JSON.stringify(json).slice(0, 300)}`);
  console.log(`OK  ${repoPath}  commit=${json.commit.sha.slice(0, 7)}`);
}
async function delFile(repoPath, message) {
  const sha = await getSha(repoPath);
  if (!sha) { console.log('SKIP  ' + repoPath + ' (not found)'); return; }
  const { status, json } = await req('DELETE', `/repos/${OWNER}/${REPO}/contents/${repoPath}`, { message, sha, branch: BRANCH });
  if (status !== 200 && status !== 201) throw new Error('DEL ' + repoPath + ' -> ' + status);
  console.log('DEL  ' + repoPath);
}
(async () => {
  try {
    const eng = fs.readFileSync('js/engine.js', 'utf8');
    if (!eng.includes('images/sku_thumb/')) throw new Error('engine.js 未改 fetch 模式');
    const idx = fs.readFileSync('index.html', 'utf8');
    if (!idx.includes('sku_image_index.js?v=1.5.20')) throw new Error('index.html 未含 sku_image_index.js?v=1.5.20');
    const sk = fs.readFileSync('js/sku_image_index.js', 'utf8');
    if (!sk.includes('window.SKU_IMAGE_INDEX')) throw new Error('sku_image_index.js 未含 SKU_IMAGE_INDEX');
    await putFile('js/engine.js', 'js/engine.js', 'v1.5.20 engine 改 fetch 模式(SKU_IMAGE_INDEX→images/sku_thumb/)');
    await putFile('js/sku_image_index.js', 'js/sku_image_index.js', 'v1.5.20 SKU_IMAGE_INDEX 索引(420键,排除5错图)');
    await putFile('js/ui.js', 'js/ui.js', 'v1.5.20 同步版本号');
    await putFile('index.html', 'index.html', 'v1.5.20 改引用 sku_image_index.js');
    await delFile('js/sku_image_pack.js', 'v1.5.20 删除 sku_image_pack.js (改 fetch 模式)');
    const p = await req('GET', `/repos/${OWNER}/${REPO}/pages`);
    console.log('PAGES:', p.status, '| build_type:', p.json && p.json.build_type);
    console.log('DONE (图片批量上传由 upload_sku_images.js 后台执行)');
  } catch (e) { console.error('FAIL:', e.message); process.exit(1); }
})();