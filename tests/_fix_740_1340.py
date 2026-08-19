import shutil, os, re, json
from PIL import Image

ROOT = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统'

# 1) 740-7 正确图（商品资料表 image747.jpeg 38824B）→ 压缩到 300px → 覆盖 sku_thumb/740-7.jpg
src = ROOT + '/tests/_out/_correct_740/740-7.jpeg'
dst = ROOT + '/images/sku_thumb/740-7.jpg'
im = Image.open(src)
print('原图:', im.size, im.mode)
im.thumbnail((300, 300), Image.LANCZOS)
rgb = im.convert('RGB')
rgb.save(dst, 'JPEG', quality=80)
print('740-7 新图:', dst, os.path.getsize(dst), 'B')

# 2) DS1340-8 从 sku_image_index.js 删除（无正确图，留空）
p = ROOT + '/js/sku_image_index.js'
raw = open(p, encoding='utf-8').read()
if raw.startswith('\ufeff'): raw = raw[1:]
m = re.search(r'window\.SKU_IMAGE_INDEX = (.*?);', raw, re.S)
idx = json.loads(m.group(1))
before = len(idx)
for k in ['DS1340-8']:
    if k in idx:
        del idx[k]
        print(f'已删除索引键 {k}')
after = len(idx)
body = '// v1.5.42 全库MD5去重扫描修复: 740-7换商品资料表正确图(河豚手拿包), DS1340-8(连衣裙无正确图)删除留空\nwindow.SKU_IMAGE_INDEX = ' + json.dumps(idx, ensure_ascii=False, sort_keys=True) + ';\n'
open(p, 'w', encoding='utf-8').write(body)
print(f'索引键数: {before} -> {after}')
print('740-7 仍指向:', idx.get('740-7'))