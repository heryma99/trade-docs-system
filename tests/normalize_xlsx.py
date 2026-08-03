# -*- coding: utf-8 -*-
# 纯 zip 层手术：彻底移除内嵌图片/图形/批注等可选部件及其在 sheet、_rels、[Content_Types] 中的引用，
# 使 ExcelJS 读回时不再因悬空引用在 reconcile 阶段报错。用于清洗 ExcelJS 写出的模板。
import sys, zipfile, re

src, dst = sys.argv[1], sys.argv[2]
DROP_PREFIXES = ('xl/media/', 'xl/drawings/', 'xl/charts/', 'xl/comments/',
                'xl/calcChain.xml', 'xl/tables/', 'xl/pivotTables/',
                'xl/queryTables/', 'xl/printerSettings/', 'xl/vmlDrawings/')
DROP_TYPES = ['/drawing', '/comments', '/vmlDrawing', '/chart', '/media',
              '/table', '/pivotTable', '/queryTable', '/calcChain']

with zipfile.ZipFile(src, 'r') as z:
    names = z.namelist()
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for n in names:
            if any(n.startswith(p) for p in DROP_PREFIXES):
                continue
            data = z.read(n)
            if re.match(r'xl/worksheets/sheet\d+\.xml$', n):
                s = data.decode('utf-8', 'ignore')
                s = re.sub(r'<drawing[^>]*>.*?</drawing>', '', s, flags=re.S)
                s = re.sub(r'<drawing[^>]*/>', '', s)
                s = re.sub(r'<legacyDrawing[^>]*/>', '', s)
                s = re.sub(r'<comments[^>]*>.*?</comments>', '', s, flags=re.S)
                s = re.sub(r'<comments[^>]*/>', '', s)
                s = re.sub(r'<picture[^>]*/>', '', s)
                data = s.encode('utf-8')
            elif n.startswith('xl/worksheets/_rels/') or n == 'xl/_rels/workbook.xml.rels':
                s = data.decode('utf-8', 'ignore')
                for t in DROP_TYPES:
                    s = re.sub(r'<Relationship[^>]*Type="[^"]*' + re.escape(t) + r'[^"]*"[^>]*/>', '', s)
                data = s.encode('utf-8')
            elif n == '[Content_Types].xml':
                s = data.decode('utf-8', 'ignore')
                s = re.sub(r'<Override[^>]*PartName="[^"]*(media|drawings|charts|comments|calcChain|tables|pivotTables|queryTables|printerSettings|vmlDrawings)[^"]*"[^>]*/>', '', s)
                data = s.encode('utf-8')
            zo.writestr(n, data)
print('normalized', src, '->', dst)
