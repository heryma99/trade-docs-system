import re, json, base64, os, hashlib

print('========== 全图库 MD5 去重扫描（揪出"多SKU共用一张图"）==========')

# ===== 1. 新图库 sku_image_index.js (436键, 文件模式) =====
raw = open('D:/WB文件/2026-07-30-09-36-10/贸易单证系统/js/sku_image_index.js', encoding='utf-8').read()
if raw.startswith('\ufeff'): raw = raw[1:]
m = re.search(r'window\.SKU_IMAGE_INDEX = (.*?);', raw, re.S)
idx = json.loads(m.group(1))
img_dir = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统/images/sku_thumb'

md5_to_sku = {}  # md5 -> [skus]
missing = []
for sku, fname in idx.items():
    p = os.path.join(img_dir, fname)
    if not os.path.exists(p):
        missing.append((sku, fname))
        continue
    h = hashlib.md5(open(p, 'rb').read()).hexdigest()
    md5_to_sku.setdefault(h, []).append(sku)

dups_new = {h: s for h, s in md5_to_sku.items() if len(s) > 1}
print(f'\n【新图库 sku_image_index】{len(idx)} 键 | 缺失文件: {len(missing)}')
for sku, fname in missing[:10]:
    print(f'  ❌ MISS {sku} -> {fname}')
print(f'  共用图组(MD5相同多SKU): {len(dups_new)} 组')
for h, s in list(dups_new.items())[:20]:
    print(f'    {h[:10]} x{len(s)}: {s}')

# ===== 2. 老图库 product_image_map.js (865键, base64内联) =====
old_src = open('D:/WB文件/2026-07-30-09-36-10/贸易单证系统/js/product_image_map.js', encoding='utf-8').read()
old_md5 = {}  # sku -> md5
n_old = 0
for mm in re.finditer(r'"([^"]+)":"data:image/(\w+);base64,([^"]+)"', old_src):
    sku, ext, b64 = mm.group(1), mm.group(2), mm.group(3)
    n_old += 1
    try:
        old_md5[sku] = hashlib.md5(base64.b64decode(b64)).hexdigest()
    except Exception:
        old_md5[sku] = 'BAD:' + ext

md5_old = {}
for sku, h in old_md5.items():
    if h.startswith('BAD'):
        continue
    md5_old.setdefault(h, []).append(sku)
dups_old = {h: s for h, s in md5_old.items() if len(s) > 1}
print(f'\n【老图库 product_image_map】{n_old} 键 | 共用图组: {len(dups_old)} 组')
for h, s in list(dups_old.items()):
    if len(s) >= 3:
        print(f'    {h[:10]} x{len(s)}: {s[:10]}{"..." if len(s)>10 else ""}')

# ===== 3. 新老图库是否指向同一张图（重叠键）=====
overlap = set(idx.keys()) & set(old_md5.keys())
print(f'\n【新老图库重叠键】{len(overlap)} 个')
for sku in sorted(overlap)[:15]:
    same = idx[sku] and old_md5.get(sku)
    # 新图库文件md5 vs 老图库md5
    p = os.path.join(img_dir, idx[sku])
    if os.path.exists(p):
        h_new = hashlib.md5(open(p, 'rb').read()).hexdigest()
        same = (h_new == old_md5.get(sku))
    else:
        same = False
    print(f'    {sku}: 新库文件={idx[sku]} 老库md5={old_md5.get(sku,"")[:10]} 一致={same}')

# ===== 4. 输出清理建议 =====
print('\n========== 清理建议 ==========')
print('新图库缺失文件需补/删键; 共用图组需人工确认(同款不同色可能同图, 但大多数是错配)')
with open('D:/WB文件/2026-07-30-09-36-10/贸易单证系统/tests/_img_dup_report.json', 'w', encoding='utf-8') as f:
    json.dump({
        'new_dups': {h: s for h, s in dups_new.items()},
        'old_dups': {h: s for h, s in dups_old.items()},
        'new_missing': missing,
        'overlap': list(overlap),
    }, f, ensure_ascii=False, indent=1)
print('报告已写 tests/_img_dup_report.json')