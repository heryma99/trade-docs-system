// 全量上传 images/sku_thumb/ 到 GitHub (v2: 422 处理+重试+进度+容错)
const https = require('https');
const fs = require('fs');
const path = require('path');
const TOKEN = fs.readFileSync('C:/Users/cn/.workbuddy/connectors/3fe83c35-d7d3-4e71-869e-097580283ed4/tokens/github.txt', 'utf8').trim();
const OWNER = 'heryma99', REPO = 'trade-docs-system', BRANCH = 'main';
const LOCAL_DIR = 'images/sku_thumb';
const REMOTE_DIR = 'images/sku_thumb';
const CONC = 6;

function apiReq(method, path, body) {
  return new Promise((resolve, reject) => {
    const data = body ? JSON.stringify(body) : null;
    const r = https.request({ host: 'api.github.com', path, method, headers: { 'User-Agent': 'node', 'Authorization': 'token ' + TOKEN, 'Content-Type': 'application/json' } }, res => {
      let d = ''; res.on('data', c => d += c);
      res.on('end', () => { let j = null; try { j = JSON.parse(d); } catch (e) {} resolve({ s: res.statusCode, j, d }); });
    });
    r.on('error', e => reject(e));
    r.setTimeout(45000, () => r.destroy(new Error('TIMEOUT')));
    if (data) r.write(data); r.end();
  });
}
async function uploadOne(f) {
  const rp = `${REMOTE_DIR}/${f}`;
  const enc = rp.split('/').map(encodeURIComponent).join('/');
  const b64 = fs.readFileSync(path.join(LOCAL_DIR, f)).toString('base64');
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      let { s, j } = await apiReq('PUT', `/repos/${OWNER}/${REPO}/contents/${enc}`, { message: 'v1.5.20 SKU图片按命名上传', content: b64, branch: BRANCH });
      if (s === 201 || s === 200) return 'ok';
      if (s === 422 && j && j.message && /sha|already|exists/i.test(j.message)) {
        // 已存在: 取 sha 覆盖
        const g = await apiReq('GET', `/repos/${OWNER}/${REPO}/contents/${enc}?ref=${BRANCH}`);
        if (g.s === 200 && g.j && g.j.sha) {
          ({ s, j } = await apiReq('PUT', `/repos/${OWNER}/${REPO}/contents/${enc}`, { message: 'v1.5.20 SKU图片按命名上传', content: b64, sha: g.j.sha, branch: BRANCH }));
          if (s === 200) return 'ok';
        }
      }
      return 'fail:' + s;
    } catch (e) { if (attempt === 2) return 'err:' + e.message; }
  }
  return 'fail';
}
(async () => {
  const files = fs.readdirSync(LOCAL_DIR).filter(f => f.endsWith('.jpg'));
  const total = files.length;
  const queue = files.slice();
  let done = 0, failList = [], t0 = Date.now();
  const workers = Array.from({ length: CONC }, () => (async () => {
    while (queue.length) {
      const f = queue.shift();
      const r = await uploadOne(f);
      if (r === 'ok') { done++; if (done % 100 === 0) console.log(`  ${done}/${total} (${((Date.now()-t0)/1000).toFixed(0)}s)`); }
      else { failList.push(f + ':' + r); }
    }
  })());
  await Promise.all(workers);
  console.log(`== 完成: ${done}/${total} OK, 失败 ${failList.length} ==`);
  console.log('失败样本:', failList.slice(0, 10));
  fs.writeFileSync('tests/_upload_report.txt', failList.join('\n'));
  process.exit(failList.length ? 1 : 0);
})().catch(e => { console.error('FATAL:', e); process.exit(1); });
