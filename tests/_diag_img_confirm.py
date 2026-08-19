import urllib.request, os, json, re

def get_bytes(url, timeout=25):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    return urllib.request.urlopen(req, timeout=timeout).read()

def get_text(url, timeout=25):
    return get_bytes(url, timeout).decode('utf-8', 'ignore')

print('===== 1. 本地 tds_dist 是否带图片目录 =====')
td = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统/tds_dist'
if os.path.exists(td + '/images/sku_thumb'):
    n = len(os.listdir(td + '/images/sku_thumb'))
    print('tds_dist/images/sku_thumb 存在:', n, '个文件')
else:
    print('❌ tds_dist/images/sku_thumb 不存在！(CloudStudio 用户同源 fetch 图片必 404)')

print('\n===== 2. CloudStudio 线上图片可达性（匿名用户走同源 fetch）=====')
for f in ['images/sku_thumb/104-1.jpg', 'images/sku_thumb/JH30314-1.jpeg', 'images/sku_thumb/2T78-1.jpg']:
    try:
        b = get_bytes('https://fff0795e044547deab469fa3b01c1522.app.workbuddy.link/' + f, timeout=15)
        print('  ✅ CS', f, len(b), 'B', 'PK头(图片不该PK):', b[:2] == b'PK', '| jpg头:', b[:3] in (b'\xff\xd8\xff', b'\x89PN'))
    except Exception as e:
        print('  ❌ CS', f, 'ERR', str(e)[:70])

print('\n===== 3. GitHub Pages 线上图片可达性 =====')
for f in ['images/sku_thumb/104-1.jpg', 'images/sku_thumb/JH30314-1.jpeg']:
    try:
        b = get_bytes('https://heryma99.github.io/trade-docs-system/' + f, timeout=15)
        print('  ✅ GH', f, len(b), 'B', '| jpg头:', b[:3] in (b'\xff\xd8\xff', b'\x89PN'))
    except Exception as e:
        print('  ❌ GH', f, 'ERR', str(e)[:70])

print('\n===== 4. 团队库 userdata 里模板的图片相关配置 =====')
try:
    ud = json.loads(get_text('https://raw.githubusercontent.com/heryma99/trade-docs-system/main/userdata.json'))
    print('  模板数:', len(ud['stores']['templates']))
except Exception as e:
    print('  ERR', str(e)[:70])

print('\n===== 5. engine.js 图片 fetch 路径逻辑（线上）=====')
try:
    e = get_text('https://fff0795e044547deab469fa3b01c1522.app.workbuddy.link/js/engine.js')
    print('  含 images/sku_thumb 引用:', 'images/sku_thumb' in e)
    # 找 lookupImg 实现
    m = re.search(r'function lookupImg[\s\S]{0,400}', e)
    if m:
        print('  lookupImg 源码片段:')
        print('  ', m.group(0).replace('\n', ' ')[:300])
except Exception as e:
    print('  ERR', str(e)[:70])