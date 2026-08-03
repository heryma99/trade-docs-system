# -*- coding: utf-8 -*-
# 生成 UTF-8 编码（带 BOM）的 CSV 测试文件，模拟现代导出
lines = ['订单号,SKU,数量', 'SO001,ABC123,5', 'SO002,中文品名XYZ,10']
with open('tests/sample_utf8.csv', 'w', encoding='utf-8-sig', newline='') as f:
    f.write('\n'.join(lines))
print('written tests/sample_utf8.csv (utf-8 BOM)')
