import urllib.request, re

def get_bytes(url, timeout=25):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    return urllib.request.urlopen(req, timeout=timeout).read()

t = get_bytes('https://heryma99.github.io/trade-docs-system/index.html').decode('utf-8', 'ignore')
vs = re.findall(r'>v(\d+\.\d+\.\d+)<', t)
print('GH index.html 顶部版本:', vs[:2])
print('GH 含 v1.5.42:', 'v1.5.42' in t)
print('GH index.html 里 ui.js?v:', re.findall(r'ui\.js\?v=([0-9.]+)', t)[:1])
print('GH index.html 里 engine.js?v:', re.findall(r'engine\.js\?v=([0-9.]+)', t)[:1])
print()
b = get_bytes('https://heryma99.github.io/trade-docs-system/images/sku_thumb/740-7.jpg')
print('740-7.jpg 字节:', len(b))
print('前 8 字节 hex:', b[:8].hex())
print('JFIF/JPEG 检查:', b[:2] == b'\xff\xd8', '| exif:', b[2:4] == b'\xff\xe1')