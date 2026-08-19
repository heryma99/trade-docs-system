import json, zipfile, os, re

P = 'D:/模板/SKU申报信息表/JW PEI 产品基础数据 G Unit.xlsx'
OUT = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统'
IMG_DIR = OUT + '/tds_dist/images/sku_thumb'
Z = zipfile.ZipFile(P)

m = json.load(open(OUT+'/_jwpei_correct_map.json', encoding='utf-8'))

# 过滤低质图（<8KB 当无图）
MIN = 8000
keep = {}
for sku, media in m.items():
    sz = Z.getinfo(media).file_size
    if sz >= MIN:
        keep[sku] = media

print('过滤后可用图:', len(keep), '(过滤掉', len(m)-len(keep), '张<8KB低质图)')

# 直接覆盖写新图（沙箱拦截 os.remove，旧图残留不影响，部署时清 GitHub）
os.makedirs(IMG_DIR, exist_ok=True)
written = 0
for sku, media in keep.items():
    data = Z.read(media)
    # 扩展名按实际
    ext = '.jpeg' if media.lower().endswith('.jpeg') else ('.png' if media.lower().endswith('.png') else '.jpg')
    fn = sku + ext
    with open(os.path.join(IMG_DIR, fn), 'wb') as fp:
        fp.write(data)
    written += 1
print('已写入本地图片:', written)

# 生成 sku_image_index.js
pairs = ', '.join('"%s":"%s"' % (sku, sku + ('.jpeg' if keep[sku].lower().endswith('.jpeg') else '.png' if keep[sku].lower().endswith('.png') else '.jpg')) for sku in sorted(keep))
header = (
    "// 重建于 2026-08-15：v1.5.44 抽图 shared-string 转译 bug 导致 SKU<->图片整批错位；\n"
    "// 本次用 xlsx drawing anchor(row) 精确绑定 A 列 SKU，从 JW PEI 源表重新抽取 %d 张真图。\n"
    "// <8KB 缩略图按\"没有就是没有\"铁律过滤。engine.lookupImg 仍保留 SPU 前缀兜底。\n"
    "window.SKU_IMAGE_INDEX = {%s};\n"
) % (len(keep), pairs)
with open(OUT+'/js/sku_image_index.js', 'w', encoding='utf-8') as fp:
    fp.write(header)
print('已生成 js/sku_image_index.js (SKU 数 %d)' % len(keep))

# 根目录旧图不删（沙箱拦截），部署时清 GitHub
root_img = OUT + '/images/sku_thumb'
if os.path.isdir(root_img):
    print('注: 根目录 images/sku_thumb 旧图残留(沙箱拦截删除)，部署时由 GitHub 端清理')
