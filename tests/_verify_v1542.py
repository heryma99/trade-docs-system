import urllib.request, re, json

def get_text(url, timeout=25):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    return urllib.request.urlopen(req, timeout=timeout).read().decode('utf-8', 'ignore')

def get_bytes(url, timeout=25):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    return urllib.request.urlopen(req, timeout=timeout).read()

print('===== v1.5.42 双镜像最终验证 =====')
for name, url in [('GitHub Pages', 'https://heryma99.github.io/trade-docs-system/'), ('CloudStudio', 'https://fff0795e044547deab469fa3b01c1522.app.workbuddy.link/')]:
    try:
        t = get_text(url + 'index.html')
        vs = re.findall(r'>v(\d+\.\d+\.\d+)<', t)
        e = get_text(url + 'js/engine.js')
        u = get_text(url + 'js/ui.js')
        idx = get_text(url + 'js/sku_image_index.js')
        m = re.search(r'window\.SKU_IMAGE_INDEX = (.*?);', idx, re.S)
        n_keys = len(json.loads(m.group(1)))
        print(f'  {name}: v{vs[:1]} | 弃用旧图库: {"PRODUCT_IMAGE_MAP || {}" not in e} | auto默认开: {"c.value.auto === false" in u} | 索引键: {n_keys}')
        # 740-7 新图
        b = get_bytes(url + 'images/sku_thumb/740-7.jpg', timeout=15)
        print(f'    740-7.jpg: {len(b)}B jpg头: {b[:3] == b"\\xff\\xd8\\xff"}')
        ud = json.loads(get_text(url + 'userdata.json', timeout=15))
        print(f'    同源userdata: {len(ud["stores"]["templates"])}模板 + {len(ud["stores"]["parties"])}主数据')
    except Exception as ex:
        print(f'  {name} ERR: {str(ex)[:80]}')