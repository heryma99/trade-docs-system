# -*- coding: utf-8 -*-
# 生成 GBK 编码的 CSV 测试文件（模拟中文 Windows 导出的非 UTF-8 文本）
lines = ['订单号,SKU,数量', 'SO001,ABC123,5', 'SO002,中文品名XYZ,10']
with open('tests/sample.csv', 'w', encoding='gbk', newline='') as f:
    f.write('\n'.join(lines))
print('written tests/sample.csv (gbk)')
