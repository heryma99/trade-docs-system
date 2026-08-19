import os, urllib.request, urllib.error, json, base64, time, sys, traceback, re
print('start', flush=True)

TOKEN = open(r'C:/Users/cn/.workbuddy/connectors/3fe83c35-d7d3-4e71-869e-097580283ed4/tokens/github.txt', encoding='utf-8').read().strip()
GHPREFIX = 'https://api.github.com/repos/heryma99/trade-docs-system/contents/'
HEAD = {'Authorization': 'Bearer ' + TOKEN, 'Accept': 'application/vnd.github.v3+json', 'User-Agent': 'tds-deploy'}
OUT = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统'
IMG_DIR = OUT + '/tds_dist/images/sku_thumb'
ctx = __import__('ssl').create_default_context()

def gh_get(path):
    req = urllib.request.Request(GHPREFIX + path, headers=HEAD)
    return json.loads(urllib.request.urlopen(req, timeout=30, context=ctx).read().decode('utf-8'))

def gh_put(path, content_bytes, sha=None, message=''):
    body = {'message': message, 'branch': 'main', 'content': base64.b64encode(content_bytes).decode('ascii')}
    if sha: body['sha'] = sha
    for attempt in range(2):
        try:
            req = urllib.request.Request(GHPREFIX + path, data=json.dumps(body).encode('utf-8'),
                                          headers={**HEAD, 'Content-Type': 'application/json'}, method='PUT')
            r = urllib.request.urlopen(req, timeout=120, context=ctx)
            return json.loads(r.read().decode('utf-8'))
        except urllib.error.HTTPError as e:
            if attempt == 0 and e.code in (500, 502, 503, 504, 429):
                time.sleep(2); continue
            raise

def gh_delete(path, sha, message=''):
    body = {'message': message, 'sha': sha, 'branch': 'main'}
    req = urllib.request.Request(GHPREFIX + path, data=json.dumps(body).encode('utf-8'),
                                  headers={**HEAD, 'Content-Type': 'application/json'}, method='DELETE')
    r = urllib.request.urlopen(req, timeout=30, context=ctx)
    return r.status

# 仅传 113 新文件（从 sku_image_index.js 解析），并删线上旧文件
print('解析新集合 113', flush=True)
idx_txt = open(OUT + '/js/sku_image_index.js', encoding='utf-8').read()
new_files = sorted({m.group(2) for m in re.finditer(r'"([^"]+)":"([^"]+)"', idx_txt)})
print('新集合:', len(new_files), flush=True)

print('获取线上现状', flush=True)
existing = gh_get('images/sku_thumb?ref=main')
existing_map = {x['name']: x['sha'] for x in existing}
print('线上:', len(existing_map), flush=True)

to_upload = new_files
to_delete = [(n, existing_map[n]) for n in existing_map if n not in set(new_files)]
print('需上传:', len(to_upload), '| 需删除:', len(to_delete), flush=True)

print('=== 上传 113 新图 ===', flush=True)
ok = err = 0
for i, fn in enumerate(to_upload, 1):
    p = 'images/sku_thumb/' + fn
    sha = existing_map.get(fn)
    try:
        with open(os.path.join(IMG_DIR, fn), 'rb') as fp:
            data = fp.read()
        gh_put(p, data, sha, 'chore: re-upload ' + fn + ' (v1.5.47)')
        ok += 1
    except Exception as e:
        err += 1
        print('UPLOAD ERR', fn, repr(e)[:120], flush=True)
        print(traceback.format_exc(), flush=True)
    if i % 15 == 0 or i == len(to_upload):
        print(f'  [{i}/{len(to_upload)}] ok={ok} err={err}', flush=True)
    time.sleep(0.8)

print(f'上传完成: {ok}/{len(to_upload)} 成功', flush=True)

print('=== 删除旧图 ===', flush=True)
ok_d = err_d = 0
for fn, sha in to_delete:
    p = 'images/sku_thumb/' + fn
    try:
        st = gh_delete(p, sha, 'chore: remove obsolete ' + fn)
        if st in (200, 204): ok_d += 1
    except Exception as e:
        err_d += 1
        print('DEL ERR', fn, repr(e)[:80], flush=True)
    time.sleep(0.5)
    if ok_d % 20 == 0 and ok_d > 0:
        print(f'  删除进度 {ok_d}/{len(to_delete)}', flush=True)
print(f'删除完成: {ok_d}/{len(to_delete)}', flush=True)

print('=== 推送索引与版本 ===', flush=True)
try:
    m = gh_get('js/sku_image_index.js?ref=main')
    isha = m['sha']
except urllib.error.HTTPError:
    isha = None
with open(OUT + '/js/sku_image_index.js', 'rb') as fp:
    gh_put('js/sku_image_index.js', fp.read(), isha, 'chore: sku_image_index.js v1.5.47 (113 correct)')
print('  索引 OK', flush=True)

ih_sha = gh_get('index.html?ref=main')['sha']
gh_put('index.html', open(OUT + '/index.html', 'rb').read(), ih_sha, 'chore: bump version v1.5.47')
print('  index.html OK', flush=True)
print('DONE', flush=True)