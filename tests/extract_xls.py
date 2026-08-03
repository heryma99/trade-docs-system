# -*- coding: utf-8 -*-
# 用 xlrd 提取 .xls 模板的明细表头行 + 1 行样本（xlrd 1.2.0 读 xls）
import sys, xlrd

HINTS = ['品名','海关','HS','品牌','型号','材质','用途','数量','单价','金额','箱号','毛重','净重','长宽高','申报',
         'QTY','SKU','Reference','FBA','Ref','ASIN','FNSKU','箱数','件数','体积','尺寸','规格','CTN','GW','NW',
         'PURPOSE','MATERIAL','MODEL','BRAND','DESCRIPTION','EORI','VAT','牌','款','原产','目的','成交']

def score(s):
    if s is None: return 0
    s = str(s)
    return sum(1 for h in HINTS if h in s)

def cellval(c):
    if c.ctype == xlrd.XL_CELL_DATE:
        try: return xlrd.xldate.xldate_as_datetime(c.value, 0).strftime('%Y-%m-%d')
        except Exception: return str(c.value)
    if c.ctype == xlrd.XL_CELL_EMPTY: return None
    return c.value

for f in sys.argv[1:]:
    print('\n########## XLS FILE: ' + f.split('/')[-1])
    wb = xlrd.open_workbook(f)
    for sh in wb.sheets():
        best, bestscore = -1, 0
        for r in range(min(sh.nrows, 60)):
            sc = 0
            for c in range(sh.ncols):
                sc += score(cellval(sh.cell(r, c)))
            if sc > bestscore:
                bestscore, best = sc, r
        if best < 0 or bestscore < 3:
            print('  SHEET "%s": 无明细表头' % sh.name); continue
        print('  SHEET "%s" 明细表头行=R%d (score=%d, ncols=%d)' % (sh.name, best+1, bestscore, sh.ncols))
        cols = []
        for c in range(sh.ncols):
            v = cellval(sh.cell(best, c))
            if v is not None and str(v).strip() != '':
                cols.append('%d:%s' % (c+1, str(v)[:30]))
        print('    表头: ' + ' | '.join(cols))
        sc2 = []
        for c in range(sh.ncols):
            v = cellval(sh.cell(best+1, c)) if best+1 < sh.nrows else None
            if v is not None and str(v).strip() != '':
                sc2.append('%d:%s' % (c+1, str(v)[:24]))
        if sc2: print('    样本: ' + ' | '.join(sc2))
