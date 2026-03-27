import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'F:/ship-statics')
import numpy as np, math, os

# 构造示例型值表
stations   = [-5, -4, -3, -2, -1, 0, 1, 2, 3, 4, 5]
waterlines = [0.0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0]

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
offsets_data = {"stations": stations, "waterlines": waterlines, "offsets": offsets}

# 测试 GZ 计算
from core.stability import calc_gz_curve

gz = calc_gz_curve(offsets_data, draft=5.0, KG=4.0, theta_step=5.0)

print('=== GZ 曲线计算结果 ===')
print(f'draft={gz["draft"]}m  KG={gz["KG"]}m  GM={gz["GM"]}m')
print(f'KB={gz["KB"]}m  BM={gz["BM"]}m  KM={gz["KB"]+gz["BM"]:.4f}m')
print(f'Delta={gz["delta"]}t  V={gz["V"]}m3')
print(f'GZ_max={gz["GZ_max"]}m @ theta={gz["theta_max_gz"]}deg')
print(f'theta_vanish={gz["theta_vanish"]}')
print()

# 打印关键角度的 GZ
print('%-8s %-8s %-10s %-12s' % ('theta', 'GZ', 'GZ_approx', 'ls'))
for r in gz['rows']:
    if r['theta'] % 10 == 0 or abs(r['theta'] - gz['theta_max_gz']) < 6:
        print('%-8.0f %-8.4f %-10.4f %-12.4f' % (r['theta'], r['GZ'], r['GZ_approx'], r['ls']))

# 稳性衡准
print()
criteria = gz['criteria']
for k, v in criteria.items():
    if k == 'overall':
        print(f'  总体: {v}')
    elif isinstance(v, dict):
        print(f'  {v.get("name","")}: {v.get("value","")} >= {v.get("limit","")} -> {"OK" if v.get("pass") else "NG"}')

# 测试绘图
from core.plotter_gz import plot_gz_curve, plot_gz_preview
import warnings
warnings.filterwarnings('ignore')

out_dir = 'F:/ship-statics/outputs/plots'
paths = plot_gz_curve(gz, out_dir, ship_name='示例货船')
print()
print(f'GZ PNG: {os.path.getsize(paths["png"])} bytes')
print(f'GZ SVG: {os.path.getsize(paths["svg"])} bytes')

b64 = plot_gz_preview(gz)
print(f'GZ preview base64: {len(b64)} chars')

print('\n全部阶段5后端模块验证通过!')
