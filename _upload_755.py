# -*- coding: utf-8 -*-
"""v1.5.48 补救：上传 755 正确 SKU 配图（线上 images 目录已空）。
索引格式："SKU": "path"（冒号后有空格），用 \s* 兼容。
"""
import urllib.request, ssl, json, base64, os, time, sys, re

ROOT = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统'
IMG_DIR = ROOT + '/tds_dist/images/sku_thumb'
IDX_LOCAL = ROOT + '/js/sku_image_index.js'
TOK = r'C:/Users/cn/.workbuddy/connectors/3fe83c35-d7d3-4e71-869e-097580283ed4/tokens/github.txt'
REPO = 'heryma99/trade-docs-system'
API = f'https://api.github.com/repos/{REPO}/contents/'

TOKEN = open(TOK, encoding='utf-8').read().strip()
ctx = ssl.create_default_context()

def h():
    return {'Authorization': 'Bearer ' + TOKEN, 'Accept': 'application/vnd.github.v3+json', 'User-Agent': 'deploy-bot'}

def get(path):
    return json.loads(urllib.request.urlopen(urllib.request.Request(API + path, headers=h()), timeout=30, context=ctx).read().decode('utf-8'))

def put(path, content_bytes, message):
    body = {'message': message, 'content': base64.b64encode(content_bytes).decode('ascii'), 'branch': 'main'}
    try:
        r = urllib.request.urlopen(urllib.request.Request(API + path, data=json.dumps(body).encode('utf-8'), headers={**h(), 'Content-Type': 'application/json'}, method='PUT'), timeout=60, context=ctx)
        return r.status
    except urllib.error.HTTPError as e:
        if e.code == 422:
            sha = get(path)['sha']
            body['sha'] = sha
            r = urllib.request.urlopen(urllib.request.Request(API + path, data=json.dumps(body).encode('utf-8'), headers={**h(), 'Content-Type': 'application/json'}, method='PUT'), timeout=60, context=ctx)
            return r.status
        raise

# 1) 解析索引：容忍冒号空格
idx_keys = set()
content = open(IDX_LOCAL, encoding='utf-8').read()
for m in re.finditer(r'"([^"]+)"\s*:\s*"images/sku_thumb/([^"]+)"', content):
    idx_keys.add(m.group(2))
print('解析索引键数:', len(idx_keys))

# 2) 线上现状
try:
    arr = get('images/sku_thumb')
    online = set(x['name'] for x in arr)
    print('线上现有:', len(online))
except urllib.error.HTTPError:
    online = set()
    print('线上 images 目录为空')

# 3) 上传缺失
to_upload = sorted(idx_keys - online)
print('需上传:', len(to_upload))
sys.stdout.flush()

ok = 0
fail = []
for i, fn in enumerate(to_upload, 1):
    p = os.path.join(IMG_DIR, fn)
    if not os.path.exists(p):
        print(f'  [{i}] 本地缺: {fn}')
        continue
    data = open(p, 'rb').read()
    path = 'images/sku_thumb/' + fn
    for attempt in range(3):
        try:
            st = put(path, data, 'chore: upload correct SKU image (v1.5.48, from 商品申报信息.xlsx)')
            if st in (200, 201):
                ok += 1
                break
        except Exception as e:
            if attempt == 2:
                fail.append((fn, str(e)[:100]))
            else:
                time.sleep(0.5 * (attempt + 1))
    if i % 100 == 0:
        print(f'  进度 {i}/{len(to_upload)} ok={ok} fail={len(fail)}')
        sys.stdout.flush()
    time.sleep(0.04)

print(f'上传完成 ok={ok} fail={len(fail)}')
for f, e in fail[:10]:
    print('  fail:', f, e)
