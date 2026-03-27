import sys
sys.path.insert(0, 'F:/ship-statics')
from core.hydrostatics_full import calc_hydrostatics_table
from core.exporter_full import export_hydrostatics_excel
import os

# 示例型值表（11站，17水线，模拟典型货船）
stations   = [-5, -4, -3, -2, -1, 0, 1, 2, 3, 4, 5]
waterlines = [0.0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0]

import math
def y_val(x, z):
    x_n = x / 5.0
    z_n = z / 8.0
    if abs(x_n) > 1.0: return 0.0
    bf = 1.0 - 0.55 * (abs(x_n) ** 2.8)
    df = 0.25 + 0.75 * (z_n ** 0.35) if z_n > 0 else 0.0
    y = 7.6 * bf * df
    if abs(x) == 5: y = 0.0
    return max(0.0, round(y, 3))

offsets = [[y_val(x, z) for x in stations] for z in waterlines]

offsets_data = {
    "stations":   stations,
    "waterlines": waterlines,
    "offsets":    offsets,
}

# 多吃水计算
table = calc_hydrostatics_table(offsets_data, d_min=1.0, d_max=8.0, d_step=1.0, method='trapz')

print('=== 多吃水静水力计算结果 ===')
print('%6s %8s %10s %8s %8s %8s %8s %8s %8s' % (
    'd', 'V', 'Delta(海)', 'xb', 'KB', 'BM', 'KM', 'Cb', 'TPC'))
for row in table['rows']:
    print('%6.1f %8.2f %10.2f %8.4f %8.4f %8.4f %8.4f %8.4f %8.4f' % (
        row['d'], row['V'], row['delta_sea'],
        row['xb'], row['zb'], row['BM'], row['zM'],
        row['Cb'], row['TPC']))

print()
print('计算方法:', table['method'])
print('吃水数量:', table['n_drafts'])
print('中横剖面面积 Am =', table['Am'], 'm2')

# 导出 Excel
out = 'F:/ship-statics/outputs/test_hydrostatics.xlsx'
export_hydrostatics_excel(table, out)
print('Excel 导出 OK:', os.path.getsize(out), 'bytes')
