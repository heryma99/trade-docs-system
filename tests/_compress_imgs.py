# -*- coding: utf-8 -*-
"""压缩 832MB 原生图 → JPEG 300px(q80)，输出 images/sku_small/ + js/sku_image_index.js
排除已确认错图的 4 个 SKU(7C437-7/2T398-1/7C401-3/2T398-30) → 索引不含它们，导出仍留空
"""
import os, json, time
from PIL import Image

SRC = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统/images/products_sku'
OUT = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统/images/sku_thumb'
IDX = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统/js/sku_image_index.js'
EXCLUDE = {'7C437-7', '2T398-1', '7C401-3', '2T398-30'}  # 已人工确认错图
MAX_EDGE = 300
QUALITY = 80

os.makedirs(OUT, exist_ok=True)
files = [f for f in os.listdir(SRC) if not f.startswith('.')]
sku_main = {}   # sku -> 主图文件名
sku_all = {}    # sku -> [文件名...]
done = 0
for f in sorted(files):
    base = os.path.splitext(f)[0]
    # 拆 SKU 与序号: "7C401-3_2" -> sku=7C401-3, idx=2
    m = base.rsplit('_', 1)
    if len(m) == 2 and m[1].isdigit():
        sku, idx = m[0], int(m[1])
    else:
        sku, idx = base, 0
    if sku in EXCLUDE:
        continue
    out_name = f'{sku}.jpg' if idx == 0 else f'{sku}_{idx}.jpg'
    try:
        im = Image.open(os.path.join(SRC, f))
        im = im.convert('RGB')
        w, h = im.size
        if max(w, h) > MAX_EDGE:
            r = MAX_EDGE / max(w, h)
            im = im.resize((max(1, int(w * r)), max(1, int(h * r))), Image.LANCZOS)
        im.save(os.path.join(OUT, out_name), 'JPEG', quality=QUALITY)
        sku_all.setdefault(sku, []).append(out_name)
        if idx == 0:
            sku_main[sku] = out_name
        done += 1
    except Exception as e:
        print('SKIP', f, e)

total = sum(os.path.getsize(os.path.join(OUT, f)) for f in os.listdir(OUT))
print(f'压缩完成: {done} 张 | SKU {len(sku_main)} | 总大小 {total/1024/1024:.1f} MB | 耗时 {time.time()-t0:.0f}s')

# 生成前端索引 js（SKU -> 主图文件名，供 engine fetch）
body = '// 自动生成: 商品申报信息.xlsx 单元格图片按 SKU 命名(WebP 300px)。已排除人工确认错图SKU。\nwindow.SKU_IMAGE_INDEX = ' + json.dumps(sku_main, ensure_ascii=False, sort_keys=True) + ';\n'
open(IDX, 'w', encoding='utf-8').write(body)
print(f'索引: {IDX} ({os.path.getsize(IDX)//1024}KB, {len(sku_main)} SKU)')

# 验证: 16 发票 SKU
target = ['7C437-7','8T113-9','2T398-2','2T398-1','8T113-30','1C213-3','5S409-1','7C401-3','7C613-1','2T398-30','8T117-31','8T117-71','8T117-7','8T117-30','8T113-7','7C613-7']
print('\n16 发票 SKU 覆盖:')
for s in target:
    print(f'  {s}: {sku_main.get(s, "❌ 无(留空)")}')
