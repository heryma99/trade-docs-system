# -*- coding: utf-8 -*-
"""从 商品申报信息.xlsx 按文档顺序抽取 商品编码(SKU)->图片 映射。
WPS cellImages 格式：cellImage[i] 对应第 i+2 行的 商品编码(列B)。
1 SKU = 1 图（重复 SKU 保留体积最大的图）。
"""
import zipfile, re, os, json

SRC = 'D:/模板/SKU申报信息表/商品申报信息.xlsx'
OUT = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统'
IMG_DIR = OUT + '/tds_dist/images/sku_thumb'
IDX = OUT + '/js/sku_image_index.js'

z = zipfile.ZipFile(SRC)
ss = re.findall(r'<t[^>]*>(.*?)</t>', z.read('xl/sharedStrings.xml').decode('utf-8', 'ignore'), re.S)
sx = z.read('xl/worksheets/sheet1.xml').decode('utf-8', 'ignore')

# 1) 商品编码(列B) 按行读取
def read_B(row):
    m = re.search(r'<row [^>]*r="%d"[^>]*>(.*?)</row>' % row, sx, re.S)
    if not m:
        return None
    tm = re.search(r'<c r="B%d"([^>]*)>' % row, m.group(1))
    if not tm:
        return None
    cm = re.search(r'<v>([^<]*)</v>', m.group(1))
    if not cm:
        return None
    val = cm.group(1)
    if 't="s"' in tm.group(1) and val.isdigit():
        try:
            return ss[int(val)]
        except Exception:
            return val
    return val

# 2) cellImage 顺序 -> rId
ci = z.read('xl/cellimages.xml').decode('utf-8', 'ignore')
blocks = re.findall(r'<etc:cellImage>(.*?)</etc:cellImage>', ci, re.S)
emb_ids = []
for b in blocks:
    e = re.search(r'r:embed="(rId\d+)"', b)
    emb_ids.append(e.group(1) if e else None)

# 3) rId -> media
rels = z.read('xl/_rels/cellimages.xml.rels').decode('utf-8', 'ignore')
rel_map = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="(media/[^"]+)"', rels))

# 4) 构建 SKU -> (media, size)
best = {}   # sku -> (media_path, size_bytes)
order = []  # 保持首次出现顺序
for i, rid in enumerate(emb_ids):
    if not rid:
        continue
    sku = read_B(i + 2)
    if not sku:
        continue
    sku = str(sku).strip()
    if not re.match(r'^[A-Za-z0-9][A-Za-z0-9._/\-]{1,40}$', sku):
        continue
    media = rel_map.get(rid)
    if not media:
        continue
    media_full = 'xl/' + media  # rels Target 相对 xl/_rels/，实际在 xl/media/
    try:
        sz = z.getinfo(media_full).file_size
    except KeyError:
        continue
    if sku not in best or sz > best[sku][1]:
        if sku not in best:
            order.append(sku)
        best[sku] = (media, sz)

print('抽取 SKU 总数(去重):', len(best))
# 体积分布
from collections import Counter
bc = Counter()
for sku, (mp, sz) in best.items():
    if sz < 5000: bc['<5KB'] += 1
    elif sz < 20000: bc['5-20KB'] += 1
    elif sz < 100000: bc['20-100KB'] += 1
    elif sz < 500000: bc['100KB-500KB'] += 1
    else: bc['>500KB'] += 1
print('体积分布:', dict(bc))

# 5) 写文件 + 索引
os.makedirs(IMG_DIR, exist_ok=True)
index = {}
ext_cnt = Counter()
for sku in order:
    media, sz = best[sku]
    ext = os.path.splitext(media)[1].lower()
    if not ext:
        ext = '.jpg'
    fn = sku
    # 文件名非法字符处理
    safe = re.sub(r'[\\/:*?"<>|]', '_', fn)
    dest = os.path.join(IMG_DIR, safe + ext)
    with z.open(media_full) as fsrc, open(dest, 'wb') as fdst:
        fdst.write(fsrc.read())
    index[sku] = 'images/sku_thumb/%s%s' % (safe, ext)
    ext_cnt[ext] += 1

# 6) 写 sku_image_index.js
with open(IDX, 'w', encoding='utf-8') as f:
    f.write('// 自动生成：商品申报信息.xlsx 商品编码(列B) -> 图片(1 SKU=1 图)\n')
    f.write('// 抽取源: D:/模板/SKU申报信息表/商品申报信息.xlsx  (WPS cellImages, 文档顺序=产品行顺序)\n')
    f.write('window.SKU_IMAGE_INDEX = ')
    f.write(json.dumps(index, ensure_ascii=False, indent=0))
    f.write(';\n')

print('写文件:', len(order), ' 扩展名:', dict(ext_cnt))
print('索引键数:', len(index))
print('样例:', list(index.items())[:5])
