const https = require('https');
const fs = require('fs');
const TOKEN = fs.readFileSync('C:/Users/cn/.workbuddy/connectors/3fe83c35-d7d3-4e71-869e-097580283ed4/tokens/github.txt', 'utf8').trim();
const OWNER='heryma99', REPO='trade-docs-system', BRANCH='main';
function req(method, path, body) {
  return new Promise((resolve, reject) => {
    const data = body ? JSON.stringify(body) : null;
    const r = https.request({ host: 'api.github.com', path, method, headers: { 'User-Agent': 'node', 'Authorization': 'token ' + TOKEN, 'Content-Type': 'application/json' } }, res => {
      let d = ''; res.on('data', c => d += c);
      res.on('end', () => { let j = null; try { j = JSON.parse(d); } catch (e) {} resolve({ s: res.statusCode, j }); });
    });
    r.on('error', e => reject(e));
    r.setTimeout(120000, () => r.destroy(new Error('TIMEOUT')));
    if (data) r.write(data); r.end();
  });
}
async function putOne(f) {
  const rp = 'images/sku_thumb/' + f;
  const enc = rp.split('/').map(encodeURIComponent).join('/');
  const b64 = fs.readFileSync('images/sku_thumb/' + f).toString('base64');
  for (let i = 0; i < 6; i++) {
    let { s, j } = await req('PUT', `/repos/${OWNER}/${REPO}/contents/${enc}`, { message: 'v1.5.20 SKU图补传', content: b64, branch: BRANCH });
    if (s === 201 || s === 200) return 'ok';
    if (s === 409 || s === 422) {
      // 已存在: 取 sha 覆盖
      const g = await req('GET', `/repos/${OWNER}/${REPO}/contents/${enc}?ref=${BRANCH}`);
      if (g.s === 200 && g.j && g.j.sha) {
        ({ s, j } = await req('PUT', `/repos/${OWNER}/${REPO}/contents/${enc}`, { message: 'v1.5.20 SKU图补传', content: b64, sha: g.j.sha, branch: BRANCH }));
        if (s === 200) return 'ok';
      }
    }
    await new Promise(r => setTimeout(r, 800 * (i + 1)));
  }
  return 'fail:' + s;
}
(async () => {
  const files = fs.readFileSync('tests/_missing_imgs.txt', 'utf8').split('\n').filter(Boolean);
  console.log('待补传:', files.length);
  const queue = files.slice();
  let ok = 0, fail = [];
  const workers = Array.from({ length: 1 }, () => (async () => {
    while (queue.length) {
      const f = queue.shift();
      const r = await putOne(f);
      if (r === 'ok') { ok++; if (ok % 100 === 0) console.log('  ' + ok + '/' + files.length); }
      else fail.push(f + ':' + r);
    }
  })());
  await Promise.all(workers);
  console.log('== 补传完成: ' + ok + ' OK / ' + fail.length + ' 失败 ==');
  console.log('失败样本:', fail.slice(0, 8));
  process.exit(fail.length ? 1 : 0);
})().catch(e => { console.error('FATAL:', e); process.exit(1); });
