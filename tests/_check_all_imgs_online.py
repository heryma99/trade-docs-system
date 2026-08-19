import json, re, time, urllib.request, random, sys

def probe(url, timeout=10):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0', 'Range': 'bytes=0-0'})
    r = urllib.request.urlopen(req, timeout=timeout)
    r.read(1)
    return r.status in (200, 206)

raw = open('D:/WB文件/2026-07-30-09-36-10/贸易单证系统/js/sku_image_index.js', encoding='utf-8').read()
if raw.startswith('\ufeff'):
    raw = raw[1:]
m = re.search(r'window\.SKU_IMAGE_INDEX = (.*?);', raw, re.S)
idx = json.loads(m.group(1))
items = list(idx.items())
print('索引键数:', len(items), flush=True)

# 已连续验证 0-120；现在随机抽 60 个（含末尾区间）
random.seed(42)
sample_idx = list(range(120, len(items)))
random.shuffle(sample_idx)
sample = sample_idx[:60]
# 加上关键发票 SKU（用户常用）
KEY = ['JH30314-1', 'JH30314-2', '2T273-1', '2T47-1', '2T321-35', '2T267-31', '2T68-1', '2T85-1', '5S61-1', '2T78-1', '104-1', '7C437-34', '2T398-1', '7C401-3', '2T398-30']
key_items = [(sku, idx[sku]) for sku in KEY if sku in idx]

base = 'https://heryma99.github.io/trade-docs-system/images/sku_thumb/'
missing = []
def check(items2, label):
    for sku, fname in items2:
        try:
            if not probe(base + fname):
                missing.append((sku, fname, 'HTTP非200'))
        except Exception as e:
            missing.append((sku, fname, str(e)[:30]))
    print(f'{label}: 检查{len(items2)}个, 累计缺失{len(missing)}', flush=True)

check([items[i] for i in sample], '随机抽样')
check(key_items, '关键发票SKU')
print('\n缺失/错误总数:', len(missing), flush=True)
for x in missing[:30]:
    print('  MISS', x, flush=True)