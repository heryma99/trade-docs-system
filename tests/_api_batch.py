import urllib.request, json, os, base64, re, sys, traceback
sys.stdout.reconfigure(line_buffering=True)

ROOT = r'D:\WB文件\2026-07-30-09-36-10\贸易单证系统'
TOKEN = open(r'C:/Users/cn/.workbuddy/connectors/3fe83c35-d7d3-4e71-869e-097580283ed4/tokens/github.txt', encoding='utf-8').read().strip()
API = 'https://api.github.com'
REPO = 'heryma99/trade-docs-system'
GH_PAT = re.compile(r'gh[pousr]_[A-Za-z0-9]{20,}')

START = int(sys.argv[1]) if len(sys.argv) > 1 else 0
END = int(sys.argv[2]) if len(sys.argv) > 2 else 10**9

def collect():
    out = []
    for f in ['index.html','styles.css','box_specs.js','userdata.json','README.md']:
        out.append(f)
    jsd = os.path.join(ROOT,'js')
    for f in sorted(os.listdir(jsd)):
        if not f.endswith('.js'): continue
        if f.endswith('.bak_test'): continue
        if f == 'sku_image_pack.js': continue
        out.append('js/'+f)
    for f in ['exceljs.min.js','xlsx.full.min.js']:
        out.append('vendor/'+f)
    idr = os.path.join(ROOT,'images','sku_thumb')
    for f in sorted(os.listdir(idr)):
        if f.lower().endswith(('.jpg','.jpeg','.png')):
            out.append('images/sku_thumb/'+f)
    return out

files = collect()
batch = files[START:END]
print(f'批次 [{START}:{END}] 共 {len(batch)} 文件 (总 {len(files)})')

def api(method, path, data=None):
    req = urllib.request.Request(API + path, headers={
        'User-Agent':'M','Authorization':f'Bearer {TOKEN}',
        'Accept':'application/vnd.github+json','Content-Type':'application/json'})
    req.method = method
    if data is not None: req.data = json.dumps(data).encode('utf-8')
    for attempt in range(4):
        try:
            r = urllib.request.urlopen(req, timeout=60)
            b = r.read()
            return json.loads(b.decode('utf-8')) if b else {}
        except urllib.error.HTTPError as e:
            err = e.read().decode('utf-8','ignore')
            print(f'  [HTTP {e.code}] {path}: {err[:200]}')
            if e.code == 422 and 'secret' in err.lower():
                raise SystemExit('SECRET_BLOCKED:'+path)
            if attempt < 3:
                import time; time.sleep(2); continue
            raise SystemExit(f'{method} {path} -> HTTP {e.code}')
        except Exception as e:
            if attempt < 3:
                import time; time.sleep(2); continue
            raise

main = api('GET', f'/repos/{REPO}/git/ref/heads/main')
base_sha = main['object']['sha']
base_commit = api('GET', f'/repos/{REPO}/git/commits/{base_sha}')
base_tree = base_commit['tree']['sha']

tree_entries = []
for i, rel in enumerate(batch):
    full = os.path.join(ROOT, rel)
    raw = open(full,'rb').read()
    if GH_PAT.search(raw.decode('utf-8','ignore')):
        print('⚠️ 跳过含token:', rel); continue
    b64 = base64.b64encode(raw).decode('ascii')
    blob = api('POST', f'/repos/{REPO}/git/blobs', {'content': b64, 'encoding':'base64'})
    tree_entries.append({'path': rel, 'mode':'100644', 'type':'blob', 'sha': blob['sha']})
    print(f'  blob {i+1}/{len(batch)} {rel} {len(raw)//1024}KB', flush=True)

tree = api('POST', f'/repos/{REPO}/git/trees', {'base_tree': base_tree, 'tree': tree_entries})
msg = f'deploy batch [{START}:{END}] ({len(tree_entries)} files)'
commit = api('POST', f'/repos/{REPO}/git/commits', {'message': msg, 'tree': tree['sha'], 'parents':[base_sha]})
api('PATCH', f'/repos/{REPO}/git/refs/heads/main', {'sha': commit['sha']})
print(f'✅ 批次提交 {commit["sha"][:7]}，main 已更新')
