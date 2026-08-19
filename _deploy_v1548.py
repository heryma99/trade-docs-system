# -*- coding: utf-8 -*-
"""v1.5.48 部署：商品申报信息.xlsx → 755 正确 SKU 配图
步骤：清理本地旧残留 → 上传 755 新图 → 删线上 113 旧图 → 推新索引 → 升级版本
"""
import urllib.request, ssl, json, base64, os, time, sys, traceback

ROOT = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统'
IMG_DIR = ROOT + '/tds_dist/images/sku_thumb'
IDX_LOCAL = ROOT + '/js/sku_image_index.js'
TOK_FILE = r'C:/Users/cn/.workbuddy/connectors/3fe83c35-d7d3-4e71-869e-097580283ed4/tokens/github.txt'
REPO = 'heryma99/trade-docs-system'
BRANCH = 'main'
API = f'https://api.github.com/repos/{REPO}/contents/'
RAW = f'https://raw.githubusercontent.com/{REPO}/{BRANCH}/'

TOKEN = open(TOK_FILE, encoding='utf-8').read().strip()
ctx = ssl.create_default_context()

def h():
    return {'Authorization': 'Bearer ' + TOKEN, 'Accept': 'application/vnd.github.v3+json', 'User-Agent': 'deploy-bot'}

def get(path):
    return json.loads(urllib.request.urlopen(urllib.request.Request(API + path, headers=h()), timeout=30, context=ctx).read().decode('utf-8'))

def put(path, content_bytes, message):
    body = json.dumps({'message': message, 'content': base64.b64encode(content_bytes).decode('ascii'), 'branch': BRANCH}).encode('utf-8')
    # try PUT (create)
    try:
        r = urllib.request.urlopen(urllib.request.Request(API + path, data=body, headers={**h(), 'Content-Type': 'application/json'}, method='PUT'), timeout=60, context=ctx)
        return r.status
    except urllib.error.HTTPError as e:
        if e.code == 422:  # exists, need sha
            sha = get(path)['sha']
            body = json.dumps({'message': message, 'content': base64.b64encode(content_bytes).decode('ascii'), 'branch': BRANCH, 'sha': sha}).encode('utf-8')
            r = urllib.request.urlopen(urllib.request.Request(API + path, data=body, headers={**h(), 'Content-Type': 'application/json'}, method='PUT'), timeout=60, context=ctx)
            return r.status
        raise

def delete(path, sha, message):
    body = json.dumps({'message': message, 'sha': sha, 'branch': BRANCH}).encode('utf-8')
    r = urllib.request.urlopen(urllib.request.Request(API + path, data=body, headers={**h(), 'Content-Type': 'application/json'}, method='DELETE'), timeout=60, context=ctx)
    return r.status

# ============== 1) 清理本地残留 ==============
idx_keys = set()
for line in open(IDX_LOCAL, encoding='utf-8'):
    if '"' in line:
        # parse "SKU":"path"
        import re
        for m in re.finditer(r'"([^"]+)":"images/sku_thumb/([^"]+)"', line):
            idx_keys.add(m.group(2))
print('索引中文件数:', len(idx_keys))
local = os.listdir(IMG_DIR)
residue = [f for f in local if f not in idx_keys]
for f in residue:
    try:
        os.remove(os.path.join(IMG_DIR, f))
    except Exception as e:
        print('  本地清理失败:', f, e)
print('清理本地残留:', len(residue), '剩余:', len(os.listdir(IMG_DIR)))

# ============== 2) 获取线上图库清单 ==============
online = []
try:
    arr = get('images/sku_thumb')
    online = [(x['name'], x['sha']) for x in arr]
except urllib.error.HTTPError as e:
    if e.code == 404:
        online = []
    else:
        raise
print('线上现有:', len(online))

# ============== 3) 上传 755 新图 ==============
new_files = sorted(idx_keys)
exist_online = set(n for n, _ in online)
to_upload = [f for f in new_files if f not in exist_online]
print('需上传:', len(to_upload))
sys.stdout.flush()

ok = 0
fail = []
for i, fn in enumerate(to_upload, 1):
    p = os.path.join(IMG_DIR, fn)
    if not os.path.exists(p):
        print(f'  [{i}/{len(to_upload)}] 缺文件: {fn}')
        continue
    data = open(p, 'rb').read()
    path = 'images/sku_thumb/' + fn
    # retry
    for attempt in range(3):
        try:
            st = put(path, data, 'chore: upload correct SKU image (v1.5.48, from 商品申报信息.xlsx)')
            if st in (200, 201):
                ok += 1
                break
            else:
                print(f'  [{i}] 状态 {st}: {fn}')
        except Exception as e:
            if attempt == 2:
                fail.append((fn, str(e)))
            else:
                time.sleep(0.5 * (attempt + 1))
    if i % 50 == 0:
        print(f'  进度 {i}/{len(to_upload)} ok={ok} fail={len(fail)}')
        sys.stdout.flush()
    time.sleep(0.05)  # 轻微限流
print(f'上传完成 ok={ok} fail={len(fail)}')
for f, e in fail[:10]:
    print('  fail:', f, e[:80])

# ============== 4) 删除线上 113 旧图 ==============
# 重新获取（上传过程中 sha 不变，但保险起见重拉）
online = [(x['name'], x['sha']) for x in get('images/sku_thumb')]
to_del = [(n, s) for n, s in online if n not in idx_keys]
print('需删除:', len(to_del))
del_ok = 0
for n, s in to_del:
    try:
        delete('images/sku_thumb/' + n, s, 'chore: remove obsolete JW PEI image (v1.5.48)')
        del_ok += 1
    except Exception as e:
        print('  删除失败:', n, e)
    time.sleep(0.05)
print('删除完成:', del_ok)

# ============== 5) 推新索引 ==============
idx_data = open(IDX_LOCAL, 'rb').read()
put('js/sku_image_index.js', idx_data, 'chore: update SKU image index (v1.5.48)')
print('索引已推')

# ============== 6) 升级 index.html 版本 ==============
ih = open(ROOT + '/index.html', 'rb').read().decode('utf-8')
old = '>v1.5.47<'
new = '>v1.5.48<'
if old in ih:
    ih = ih.replace(old, new)
    ih = ih.replace('?v=1.5.47', '?v=1.5.48')
    open(ROOT + '/index.html', 'w', encoding='utf-8').write(ih)
    put('index.html', ih.encode('utf-8'), 'chore: bump version to v1.5.48')
    print('index.html 升级到 v1.5.48')
else:
    print('index.html 版本标签未找到:', old)

print('全部完成')
