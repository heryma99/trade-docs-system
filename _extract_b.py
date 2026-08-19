import zipfile, re, os, hashlib, json
from collections import Counter

SRC = "D:/模板/SKU申报信息表/商品申报信息.xlsx"
OUTDIR = "D:/WB文件/2026-07-30-09-36-10/贸易单证系统/images/sku_thumb"
os.makedirs(OUTDIR, exist_ok=True)

z = zipfile.ZipFile(SRC)
names = z.namelist()

ss = []
if 'xl/sharedStrings.xml' in names:
    s = z.read('xl/sharedStrings.xml').decode('utf-8', 'ignore')
    ss = re.findall(r'<t[^>]*>([\s\S]*?)</t>', s)

sx = z.read('xl/worksheets/sheet1.xml').decode('utf-8', 'ignore')
ci = z.read('xl/cellimages.xml').decode('utf-8', 'ignore')
rels = z.read('xl/_rels/cellimages.xml.rels').decode('utf-8', 'ignore')

# ID -> rId
id2rid = {}
for blk in re.findall(r'<etc:cellImage>.*?</etc:cellImage>', ci, re.S):
    name = re.search(r'name="(ID_[0-9A-F]{32})"', blk)
    emb = re.search(r'r:embed="(rId\d+)"', blk)
    if name and emb:
        id2rid[name.group(1)] = emb.group(1)

# rId -> media path
rid2media = {}
for m in re.findall(r'<Relationship Id="(rId\d+)"[^>]*Target="([^"]+)"', rels):
    rid2media[m[0]] = m[1] if m[1].startswith('xl/') else 'xl/' + m[1]

def get_cell(row, col):
    m = re.search(r'<row [^>]*r="%d"[^>]*>(.*?)</row>' % row, sx, re.S)
    if not m: return None
    cm = re.search(r'<c r="%s%d"([^>]*)>(.*?)</c>' % (col, row), m.group(1), re.S)
    if not cm: return None
    attr, inner = cm.group(1), cm.group(2)
    vm = re.search(r'<v>([^<]*)</v>', inner)
    if not vm: return None
    val = vm.group(1)
    if 't="s"' in attr and val.isdigit():
        try: return ss[int(val)]
        except: return val
    return val

def safe(fn):
    return re.sub(r'[\\/:*?"<>|\s]', '_', fn)

# Q 列 DISPIMG -> row -> ID
q_map = {}
for cm in re.finditer(r'<c r="Q(\d+)"[^>]*>(.*?)</c>', sx, re.S):
    rn = int(cm.group(1))
    if rn == 1: continue
    mid = re.search(r'DISPIMG\(&quot;(ID_[0-9A-F]{32})&quot;', cm.group(2))
    if mid:
        q_map[rn] = mid.group(1)

# 组装: sku -> media(去重读取)
sku_to_media = {}
for rn, imgid in q_map.items():
    sku = get_cell(rn, 'B')
    if not sku or not str(sku).strip(): continue
    if imgid not in id2rid: continue
    rid = id2rid[imgid]
    media = rid2media.get(rid)
    if not media: continue
    sku_to_media[str(sku).strip()] = media

print("SKU->media 映射数:", len(sku_to_media))

# 抽图: 每个 SKU 一个文件(同名重复也各存一份)
index = {}
media_cache = {}   # media_path -> bytes
written = 0
dup_sku = Counter()
for sku, media in sku_to_media.items():
    if media not in media_cache:
        media_cache[media] = z.read(media)
    data = media_cache[media]
    base = safe(sku)
    dup_sku[base] += 1
    if dup_sku[base] > 1:
        fname = "%s_%d.jpeg" % (base, dup_sku[base])
    else:
        fname = base + ".jpeg"
    # 防万一同 SKU 不同图但重名已被占
    while os.path.exists(os.path.join(OUTDIR, fname)):
        dup_sku[base] += 1
        fname = "%s_%d.jpeg" % (base, dup_sku[base])
    with open(os.path.join(OUTDIR, fname), 'wb') as f:
        f.write(data)
    index[sku] = "images/sku_thumb/" + fname
    written += 1

# 写索引
with open("D:/WB文件/2026-07-30-09-36-10/贸易单证系统/js/sku_image_index.js", 'w', encoding='utf-8') as f:
    f.write("window.SKU_IMAGE_INDEX=" + json.dumps(index, ensure_ascii=False, indent=0) + ";\n")

# 统计体积
total = sum(os.path.getsize(os.path.join(OUTDIR, fn)) for fn in os.listdir(OUTDIR))
print("写出文件数(本地):", written)
print("索引键数:", len(index))
print("distinct media 读取:", len(media_cache))
print("本地总体积: %.2f MB" % (total/1024/1024))
print("distinct 体积(去重后): %.2f MB" % (sum(len(v) for v in media_cache.values())/1024/1024))

# 清单(供部署用): sku -> 本地相对路径 + content sha256
manifest = []
for sku, rel in index.items():
    lpath = "D:/WB文件/2026-07-30-09-36-10/贸易单证系统/" + rel
    with open(lpath, 'rb') as f:
        b = f.read()
    manifest.append({"sku": sku, "rel": rel, "sha256": hashlib.sha256(b).hexdigest(), "size": len(b)})
with open("D:/WB文件/2026-07-30-09-36-10/贸易单证系统/_manifest_b.json", 'w') as f:
    json.dump(manifest, f)
print("manifest 条目:", len(manifest))
