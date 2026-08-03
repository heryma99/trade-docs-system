import openpyxl
import os

path = os.path.join(os.path.dirname(__file__), 'sample_packing.xlsx')
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Sheet1'
headers = ['订单号', '箱号', '店铺号', '款号', '数量', '纸箱规格', '毛重(kg)', '净重(kg)', '体积重(kg)', '卡板号', '卡板尺寸', '打板高度(cm)']
ws.append(headers)
rows = [
    ['wh14172', 'CARTON-1', 'A店', '1C131-2', 5, '58*38*37', 7.93, 7.33, 13.59, 'P1', '120*100', 120],
    ['wh14172', 'CARTON-1', 'A店', '1C131-4', 2, '58*38*37', 7.93, 7.33, 13.59, 'P1', '120*100', 120],
    ['wh14172', 'CARTON-2', 'A店', '1C131-1', 10, '60*30*40.5', 12.54, 11.99, 12.15, 'P2', '120*100', 120],
    ['wh14172', 'CARTON-2', 'A店', '1C131-4', 3, '60*30*40.5', 12.54, 11.99, 12.15, 'P2', '120*100', 120],
]
for r in rows:
    ws.append(r)
wb.save(path)
print('saved', path)
