const https = require('https');
const fs = require('fs');
const TOKEN = fs.readFileSync('C:/Users/cn/.workbuddy/connectors/3fe83c35-d7d3-4e71-869e-097580283ed4/tokens/github.txt', 'utf8').trim();
const OWNER='heryma99', REPO='trade-docs-system', BRANCH='main';
function req(method, path, body) {
  return new Promise((resolve, reject) => {
    const data = body ? JSON.stringify(body) : null;
    const r = https.request({ host: 'api.github.com', path, method, headers: { 'User-Agent': 'node', 'Authorization': 'token ' + TOKEN, 'Content-Type': 'application/json' } }, res => {
      let d = ''; res.on('data', c => d += c);
      res.on('end', () => { let j = null; try { j = JSON.parse(d); } catch(e){} resolve({ s: res.statusCode, j, d }); });
    });
    r.on('error', e => reject(e));
    r.setTimeout(30000, () => r.destroy(new Error('TIMEOUT')));
    if (data) r.write(data); r.end();
  });
}
(async () => {
  const files = ['104-1.jpg', '104-11.jpg', '7C437-34.jpg'].filter(f => fs.existsSync('images/sku_thumb/' + f));
  for (const f of files) {
    const rp = 'images/sku_thumb/' + f;
    const enc = rp.split('/').map(encodeURIComponent).join('/');
    const b64 = fs.readFileSync('images/sku_thumb/' + f).toString('base64');
    const { s, j, d } = await req('PUT', `/repos/${OWNER}/${REPO}/contents/${enc}`, { message: 'test upload 3', content: b64, branch: BRANCH });
    console.log(f, '->', s, j && j.content ? 'OK sha=' + j.content.sha.slice(0, 7) : (j && j.message || d.slice(0, 150)));
  }
})();
