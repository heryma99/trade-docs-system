import zipfile, re, os, hashlib, json, glob

P = 'D:/模板/SKU申报信息表/JW PEI 产品基础数据 G Unit.xlsx'
OUT = 'D:/WB文件/2026-07-30-09-36-10/贸易单证系统'
Z = zipfile.ZipFile(P)
ss = []
try:
    ssc = Z.read('xl/sharedStrings.xml').decode('utf-8', 'ignore')
    ss = re.findall(r'<t[^>]*>(.*?)</t>', ssc, re.S)
except KeyError:
    pass

media = {n: hashlib.md5(Z.read(n)).hexdigest() for n in Z.namelist() if n.startswith('xl/media/')}

def sheet_cells_translated(sx, row):
    m=re.search(r'<row [^>]*r="%d"[^>]*>(.*?)</row>'%row, sx, re.S)
    if not m: return {}
    out={}
    for cm in re.finditer(r'<c r="([A-Z]+)%d"([^>]*)>(?:<f[^>]*>.*?</f>)?<v>([^<]*)</v>'%row, m.group(1)):
        ref,attrs,val=cm.group(1),cm.group(2),cm.group(3)
        if 't="s"' in attrs and val and val.isdigit():
            try: val=ss[int(val)]
            except: pass
        if val: out[ref]=val.strip()
    return out

def sku_pattern(v):
    return bool(re.match(r'^[A-Za-z0-9][A-Za-z0-9._-]{3,19}$', v)) and ('-' in v or '_' in v)

drawing_to_sheet={}
for n in Z.namelist():
    m=re.search(r'xl/worksheets/_rels/(sheet\d+)\.xml\.rels$', n)
    if not m: continue
    c=Z.read(n).decode('utf-8','ignore')
    dm=re.search(r'Target="\.\./drawings/(drawing\d+)\.xml"', c)
    if dm: drawing_to_sheet[dm.group(1)]=m.group(1)

drawing_rels={}
for n in Z.namelist():
    m=re.search(r'xl/drawings/_rels/(drawing\d+)\.xml\.rels$', n)
    if not m: continue
    c=Z.read(n).decode('utf-8','ignore')
    mp={}
    for rm in re.finditer(r'Id="([^"]+)"[^>]*Target="([^"]+)"', c):
        rid,tgt=rm.group(1),rm.group(2)
        if 'media/' in tgt: mp[rid]='xl/media/'+os.path.basename(tgt)
    drawing_rels[m.group(1)]=mp

print('drawing_to_sheet:', drawing_to_sheet)
correct={}
for dfile, sid in drawing_to_sheet.items():
    dxml=Z.read('xl/drawings/%s.xml'%dfile).decode('utf-8','ignore')
    sx=Z.read('xl/worksheets/%s.xml'%sid).decode('utf-8','ignore')
    rels=drawing_rels.get(dfile,{})
    anchors=re.findall(r'<xdr:(?:oneCell|twoCell)Anchor>(.*?)</xdr:(?:oneCell|twoCell)Anchor>', dxml, re.S)
    print('  ', dfile, 'anchors', len(anchors), 'rels', len(rels))
    for a in anchors:
        cm=re.search(r'<xdr:from>.*?<xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>', a, re.S)
        if not cm: continue
        col0,row0=int(cm.group(1)),int(cm.group(2))
        r1=row0+1
        bm=re.search(r'r:embed="([^"]+)"', a)
        if not bm: continue
        mn=rels.get(bm.group(1))
        if not mn: continue
        sku=sheet_cells_translated(sx, r1).get('A')
        if not sku or not sku_pattern(sku):
            for dr in (-2,-1,1,2):
                cc=sheet_cells_translated(sx, r1+dr)
                if cc.get('A') and sku_pattern(cc['A']):
                    sku=cc['A']; break
        if sku and sku_pattern(sku):
            correct[sku]=(mn, media.get(mn))

print('=== 正确映射 SKU 数:', len(correct))
for k in list(correct.keys())[:15]:
    print(' ', k, '->', correct[k][0])

# 比对现有 sku_image_index.js
idx_txt=open(OUT+'/js/sku_image_index.js', encoding='utf-8').read()
idx_map=dict(re.findall(r'"([^"]+)":"([^"]+)"', idx_txt))
print('现有索引 SKU 数:', len(idx_map))

local_md5={}
for f in glob.glob(os.path.join(OUT,'tds_dist/images/sku_thumb','*.jp*')):
    local_md5[os.path.basename(f)]=hashlib.md5(open(f,'rb').read()).hexdigest()

matched=mismatch=missing=0
for sku,(mn,md5v) in correct.items():
    if sku in idx_map:
        em=local_md5.get(idx_map[sku])
        if em==md5v: matched+=1
        else:
            mismatch+=1
            if mismatch<=12: print('  错配:', sku, '现有', idx_map[sku], '应为', os.path.basename(mn))
    else:
        missing+=1
print('命中且一致:', matched, '| 命中但错配:', mismatch, '| 正确映射不在现有索引:', missing)

with open(OUT+'/_jwpei_correct_map.json','w',encoding='utf-8') as f:
    json.dump({k:v[0] for k,v in correct.items()}, f, ensure_ascii=False)
print('saved _jwpei_correct_map.json')
