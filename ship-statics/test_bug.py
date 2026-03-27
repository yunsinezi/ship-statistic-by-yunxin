import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'F:/ship-statics')
import numpy as np, math

# 构造示例型值表
stations   = [-5, -4, -3, -2, -1, 0, 1, 2, 3, 4, 5]
waterlines = [0.0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0]
def y_val(x, z):
    x_n = x / 5.0; z_n = z / 8.0
    if abs(x_n) > 1.0: return 0.0
    bf = 1.0 - 0.55 * (abs(x_n) ** 2.8)
    df = 0.25 + 0.75 * (z_n ** 0.35) if z_n > 0 else 0.0
    y = 7.6 * bf * df
    if abs(x) == 5: y = 0.0
    return max(0.0, round(y, 3))
offsets = [[y_val(x, z) for x in stations] for z in waterlines]
offsets_data = {"stations": stations, "waterlines": waterlines, "offsets": offsets}

# 测试当前 stability 模块
from core.stability import calc_gz_curve
gz = calc_gz_curve(offsets_data, draft=5.0, KG=4.0, theta_step=5.0)

print('=== 当前 GZ 计算（有 bug）===')
print('theta  GZ(m)     GZ_approx')
for r in gz['rows']:
    if r['theta'] % 15 == 0 or r['theta'] >= 70:
        print('%-6.0f %-10.4f %-10.4f' % (r['theta'], r['GZ'], r['GZ_approx']))
print()
print('问题: theta=90 时 GZ =', gz['rows'][-1]['GZ'], '(应为约2m以内)')
