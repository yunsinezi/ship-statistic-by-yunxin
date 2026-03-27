import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'F:/ship-statics')

import math, os
import numpy as np

# 构造示例型值表（11站，17水线）
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

# ── 测试1：邦戎曲线计算 ──────────────────────────────────
print("=== 邦戎曲线计算 ===")
from core.bonjean import calc_bonjean_table
bonjean = calc_bonjean_table(offsets_data)
print(f"站数: {bonjean['n_stations']}, 水线数: {bonjean['n_waterlines']}")
print("中站(x=0)各水线横剖面面积 A:")
mid_idx = 5  # x=0 是第6站（索引5）
for j, z in enumerate(bonjean['waterlines'][::4]):  # 每4条水线取一个
    wl_j = bonjean['waterlines'].index(z)
    print(f"  z={z:.1f}m: A={bonjean['A'][mid_idx][wl_j]:.3f} m², S={bonjean['S'][mid_idx][wl_j]:.3f} m³")

# ── 测试2：静水力曲线绘制 ────────────────────────────────
print("\n=== 静水力曲线绘制 ===")
from core.hydrostatics_full import calc_hydrostatics_table
from core.plotter import plot_hydrostatics, plot_bonjean, plot_hydrostatics_preview, plot_bonjean_preview

table = calc_hydrostatics_table(offsets_data, d_min=1.0, d_max=8.0, d_step=1.0)
out_dir = 'F:/ship-statics/outputs/plots'

paths = plot_hydrostatics(table, out_dir, ship_name='示例货船')
print(f"静水力曲线 PNG: {os.path.getsize(paths['png'])} bytes")
print(f"静水力曲线 SVG: {os.path.getsize(paths['svg'])} bytes")

# ── 测试3：邦戎曲线绘制 ──────────────────────────────────
print("\n=== 邦戎曲线绘制 ===")
paths2 = plot_bonjean(bonjean, out_dir, ship_name='示例货船')
print(f"邦戎曲线 PNG: {os.path.getsize(paths2['png'])} bytes")
print(f"邦戎曲线 SVG: {os.path.getsize(paths2['svg'])} bytes")

# ── 测试4：预览图（base64）───────────────────────────────
print("\n=== 预览图生成 ===")
b64_hydro = plot_hydrostatics_preview(table)
b64_bonj  = plot_bonjean_preview(bonjean)
print(f"静水力预览 base64 长度: {len(b64_hydro)} chars")
print(f"邦戎预览 base64 长度: {len(b64_bonj)} chars")

print("\n全部阶段4模块验证通过！")
