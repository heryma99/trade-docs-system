import urllib.request, json, os, base64, re, sys, traceback
sys.stdout.reconfigure(line_buffering=True)

ROOT = r'D:\WB文件\2026-07-30-09-36-10\贸易单证系统'
TOKEN = open(r'C:/Users/cn/.workbuddy/connectors/3fe83c35-d7d3-4e71-869e-097580283ed4/tokens/github.txt', encoding='utf-8').read().strip()
API = 'https://api.github.com'
REPO = 'heryma99/trade-docs-system'
GH_PAT = re.compile(r'gh[pousr]_[A-Za-z0-9]{20,}')

def collect():
    out = []
    # 根级必需文件
    for f in ['index.html','styles.css','box_specs.js','userdata.json','README.md']:
        out.append(f)
    # js/ 下除备份/测试外的所有 .js
    jsd = os.path.join(ROOT,'js')
    for f in sorted(os.listdir(jsd)):
        if not f.endswith('.js'): continue
        if f.endswith('.bak_test'): continue
        if f == 'sku_image_pack.js': continue  # 未被 index.html 引用，跳过以减少体积
        out.append('js/'+f)
    # vendor
    for f in ['exceljs.min.js','xlsx.full.min.js']:
        out.append('vendor/'+f)
    # 线上图库（仅 sku_thumb）
    idr = os.path.join(ROOT,'images','sku_thumb')
    for f in sorted(os.listdir(idr)):
        if f.lower().endswith(('.jpg','.jpeg','.png')):
            out.append('images/sku_thumb/'+f)
    return out

files = collect()
print(f'待上传文件数: {len(files)}')
total = sum(os.path.getsize(os.path.join(ROOT,f)) for f in files)
print(f'总大小: {total/1024/1024:.1f} MB')

def api(method, path, data=None):
    url = API + path
    req = urllib.request.Request(url, headers={
        'User-Agent':'M','Authorization':f'Bearer {TOKEN}',
        'Accept':'application/vnd.github+json','Content-Type':'application/json'})
    req.method = method
    if data is not None: req.data = json.dumps(data).encode('utf-8')
    for attempt in range(3):
        try:
            r = urllib.request.urlopen(req, timeout=60)
            return json.loads(r.read().decode('utf-8')) if r.read else {}
        except urllib.error.HTTPError as e:
            err = e.read().decode('utf-8','ignore')
            if e.code == 422 and attempt < 2:
                import time; time.sleep(1); continue
            raise SystemExit(f'{method} {path} -> HTTP {e.code}: {err[:300]}')
        except Exception as e:
            if attempt < 2: continue
            raise

try:
    main = api('GET', f'/repos/{REPO}/git/ref/heads/main')
    base_sha = main['object']['sha']
    base_commit = api('GET', f'/repos/{REPO}/git/commits/{base_sha}')
    base_tree = base_commit['tree']['sha']
except Exception as e:
    print('PREP ERR', repr(e)); traceback.print_exc(); sys.exit(1)

tree_entries = []
n = 0
for rel in files:
    full = os.path.join(ROOT, rel)
    raw = open(full,'rb').read()
    # token 自检
    if GH_PAT.search(raw.decode('utf-8','ignore')):
        print('⚠️ 跳过含token文件:', rel); continue
    b64 = base64.b64encode(raw).decode('ascii')
    print(f'  -> [{n+1}/{len(files)}] {rel} {len(raw)/1024:.0f}KB', flush=True)
    try:
        blob = api('POST', f'/repos/{REPO}/git/blobs', {'content': b64, 'encoding':'base64'})
    except BaseException as e:
        print(f'❌ blob 失败 {rel} ({len(raw)/1024:.0f}KB): {repr(e)}');
        traceback.print_exc(); sys.exit(1)
    tree_entries.append({'path': rel, 'mode':'100644', 'type':'blob', 'sha': blob['sha']})
    n += 1
    if n % 20 == 0: print(f'  blobs {n}/{len(files)}')

try:
    tree = api('POST', f'/repos/{REPO}/git/trees', {'base_tree': base_tree, 'tree': tree_entries})
    print(f'新 tree: {tree["sha"][:7]} ({len(tree_entries)} entries)')
    msg = 'fix: 补全全部应用文件(js/vendor/styles/box_specs/images/sku_thumb)，修复404页面打不开'
    commit = api('POST', f'/repos/{REPO}/git/commits', {'message': msg, 'tree': tree['sha'], 'parents':[base_sha]})
    print(f'新 commit: {commit["sha"][:7]}')
    api('PATCH', f'/repos/{REPO}/git/refs/heads/main', {'sha': commit['sha']})
    print('✅ main 已更新')
except Exception as e:
    print('COMMIT ERR', repr(e)); traceback.print_exc(); sys.exit(1)
