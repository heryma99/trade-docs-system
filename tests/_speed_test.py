import urllib.request, time

def probe(url, timeout=12):
    t1 = time.time()
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0', 'Range': 'bytes=0-0'})
    r = urllib.request.urlopen(req, timeout=timeout)
    r.read(1)
    return round(time.time() - t1, 2), r.status

print('=== GitHub Pages 图片响应速度（当前网络实测）===')
for f in ['104-1.jpg', 'JH30314-1.jpeg', '2T78-1.jpg', '7C437-34.jpg', '2T398-1.jpg']:
    try:
        dt, st = probe('https://heryma99.github.io/trade-docs-system/images/sku_thumb/' + f)
        print(f'  {f:20s} {dt}s HTTP{st}')
    except Exception as e:
        print(f'  {f:20s} ERR {str(e)[:40]}')

print('\n=== GitHub Pages 关键资源 ===')
for p in ['index.html', 'userdata.json', 'templates/tpl_msecxi8k_68l223.xlsx', 'js/engine.js']:
    try:
        dt, st = probe('https://heryma99.github.io/trade-docs-system/' + p)
        print(f'  {p:40s} {dt}s HTTP{st}')
    except Exception as e:
        print(f'  {p:40s} ERR {str(e)[:40]}')