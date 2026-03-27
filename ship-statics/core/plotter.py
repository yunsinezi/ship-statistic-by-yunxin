# ============================================================
# core/plotter.py — 静水力曲线 & 邦戎曲线绘制模块
#
# 阶段4 新增模块
#
# 绘图规范：
#   - 图幅：A3（420×297mm），横向放置
#   - 分辨率：300 DPI（高清打印级）
#   - 字体：中文用 SimHei（黑体），英文/数字用 DejaVu Sans
#   - 线宽：主曲线 1.5pt，辅助线 0.5pt
#   - 颜色：每条曲线用不同颜色，色盲友好配色
#   - 坐标：吃水为纵轴（Y轴），参数为横轴（X轴）
#   - 标注：每条曲线末端标注参数名和单位
#
# 导出格式：
#   - PNG：300 DPI，适合课设报告插图
#   - SVG：矢量格式，可导入 AutoCAD / Visio
# ============================================================

import os
import io
import numpy as np
import matplotlib
matplotlib.use('Agg')   # 非交互后端，适合服务器环境
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib.lines import Line2D
from matplotlib import rcParams

# ── 字体配置 ──────────────────────────────────────────────
# 优先使用系统中文字体，按优先级尝试
_CN_FONTS = ['SimHei', 'Microsoft YaHei', 'WenQuanYi Micro Hei',
             'Noto Sans CJK SC', 'Arial Unicode MS', 'DejaVu Sans']

def _setup_font():
    """配置 matplotlib 中文字体"""
    import matplotlib.font_manager as fm
    available = {f.name for f in fm.fontManager.ttflist}
    for font in _CN_FONTS:
        if font in available:
            rcParams['font.family'] = font
            break
    else:
        # 找不到中文字体，使用英文标注
        rcParams['font.family'] = 'DejaVu Sans'
    rcParams['axes.unicode_minus'] = False   # 负号正常显示

_setup_font()

# ── 色盲友好配色（10色）────────────────────────────────────
COLORS = [
    '#1f77b4',  # 蓝
    '#d62728',  # 红
    '#2ca02c',  # 绿
    '#ff7f0e',  # 橙
    '#9467bd',  # 紫
    '#8c564b',  # 棕
    '#e377c2',  # 粉
    '#17becf',  # 青
    '#bcbd22',  # 黄绿
    '#7f7f7f',  # 灰
]

# ── A3 图幅尺寸（英寸）────────────────────────────────────
A3_W = 16.54   # 420mm
A3_H = 11.69   # 297mm


# ══════════════════════════════════════════════════════════
# 静水力曲线绘制
# ══════════════════════════════════════════════════════════

# 14条静水力曲线的定义
# key: 对应 hydrostatics_full.py 返回的字段名
# label: 图例标签（中文）
# unit: 单位
# group: 分组（用于多子图布局）
HYDRO_CURVES = [
    # 组1：体积/排水量类
    {"key": "V",           "label": "排水体积 V",    "unit": "m3",      "group": 1},
    {"key": "delta_sea",   "label": "排水量 D(海)",  "unit": "t",       "group": 1},
    {"key": "delta_fresh", "label": "排水量 D(淡)",  "unit": "t",       "group": 1},
    # 组2：浮心/漂心坐标
    {"key": "xb",          "label": "浮心纵坐标 xb", "unit": "m",       "group": 2},
    {"key": "zb",          "label": "浮心高度 KB",   "unit": "m",       "group": 2},
    {"key": "xf",          "label": "漂心纵坐标 xf", "unit": "m",       "group": 2},
    # 组3：水线面/TPC/MTC
    {"key": "Aw",          "label": "水线面积 Aw",   "unit": "m2",      "group": 3},
    {"key": "TPC",         "label": "TPC",           "unit": "t/cm",    "group": 3},
    {"key": "MTC",         "label": "MTC",           "unit": "t.m/cm",  "group": 3},
    # 组4：稳心
    {"key": "BM",          "label": "横稳心半径 BM", "unit": "m",       "group": 4},
    {"key": "BML",         "label": "纵稳心半径 BML","unit": "m",       "group": 4},
    {"key": "zM",          "label": "稳心高度 KM",   "unit": "m",       "group": 4},
    # 组5：船型系数
    {"key": "Cb",          "label": "方形系数 Cb",   "unit": "-",       "group": 5},
    {"key": "Cwp",         "label": "水线面系数 Cwp","unit": "-",       "group": 5},
]


def plot_hydrostatics(table: dict,
                      output_dir: str,
                      fmt: str = 'png',
                      ship_name: str = '') -> dict:
    """
    绘制完整静水力曲线图（A3，300DPI）。

    布局：5组子图，每组2~3条曲线，共14条。
    坐标：吃水为纵轴（Y轴），参数为横轴（X轴）。

    参数：
        table      : dict  calc_hydrostatics_table() 返回的结果
        output_dir : str   输出目录
        fmt        : str   'png' 或 'svg'
        ship_name  : str   船名（用于标题）

    返回：
        dict { 'png': path, 'svg': path }  实际保存路径
    """
    rows    = table["rows"]
    drafts  = [r["d"] for r in rows]
    method  = "梯形法" if table.get("method") == "trapz" else "辛普森法"

    # ── 创建图幅（A3横向，5列子图）────────────────────────
    fig, axes = plt.subplots(1, 5, figsize=(A3_W, A3_H),
                              gridspec_kw={'wspace': 0.45})
    fig.patch.set_facecolor('#fafbff')

    # ── 大标题 ────────────────────────────────────────────
    title = f"静水力曲线图"
    if ship_name:
        title = f"{ship_name}  静水力曲线图"
    fig.suptitle(title,
                 fontsize=14, fontweight='bold', y=0.98,
                 color='#1a3a6a')

    # 副标题
    L = table.get("L", "—")
    fig.text(0.5, 0.945,
             f"计算方法：{method}  |  船长 L={L}m  |  "
             f"依据：盛振邦《船舶原理》上册  |  武汉理工大学船海学院",
             ha='center', fontsize=7.5, color='#555', style='italic')

    # ── 按组绘制 ──────────────────────────────────────────
    group_titles = {
        1: "排水量 / 体积",
        2: "浮心 / 漂心坐标",
        3: "水线面积 / TPC / MTC",
        4: "稳心参数",
        5: "船型系数",
    }

    for g_idx, (ax, g_num) in enumerate(zip(axes, range(1, 6))):
        curves = [c for c in HYDRO_CURVES if c["group"] == g_num]
        _draw_hydro_group(ax, rows, drafts, curves, group_titles[g_num])

    # ── 图框装饰 ──────────────────────────────────────────
    _add_border(fig)

    # ── 保存 ──────────────────────────────────────────────
    os.makedirs(output_dir, exist_ok=True)
    paths = {}

    png_path = os.path.join(output_dir, "静水力曲线图.png")
    fig.savefig(png_path, dpi=300, bbox_inches='tight',
                facecolor=fig.get_facecolor())
    paths['png'] = png_path

    svg_path = os.path.join(output_dir, "静水力曲线图.svg")
    fig.savefig(svg_path, format='svg', bbox_inches='tight',
                facecolor=fig.get_facecolor())
    paths['svg'] = svg_path

    plt.close(fig)
    return paths


def _draw_hydro_group(ax, rows, drafts, curves, group_title):
    """
    在单个子图上绘制一组静水力曲线。

    坐标规范：
        X轴：参数值（各曲线共享，用双X轴或归一化）
        Y轴：吃水 d（m），从小到大向上
    """
    ax.set_facecolor('#f8faff')
    ax.grid(True, linestyle='--', linewidth=0.4, color='#c8d8f0', alpha=0.8)
    ax.set_ylabel("吃水 d (m)", fontsize=8, color='#1a3a6a')
    ax.set_title(group_title, fontsize=8.5, fontweight='bold',
                 color='#1a3a6a', pad=6)

    # 吃水轴（Y轴）范围
    d_arr = np.array(drafts)
    ax.set_ylim(d_arr.min() - 0.1, d_arr.max() + 0.1)
    ax.yaxis.set_major_locator(ticker.MultipleLocator(1.0))
    ax.yaxis.set_minor_locator(ticker.MultipleLocator(0.5))
    ax.tick_params(axis='y', labelsize=7)

    # 每条曲线归一化到 [0,1] 后绘制，X轴显示归一化刻度
    # 在曲线末端标注实际值范围
    legend_handles = []
    legend_labels  = []

    for c_idx, curve in enumerate(curves):
        key    = curve["key"]
        label  = curve["label"]
        unit   = curve["unit"]
        color  = COLORS[c_idx % len(COLORS)]

        vals = np.array([r.get(key, 0.0) for r in rows], dtype=float)

        # 归一化到 [0.05, 0.95]（留边距）
        v_min, v_max = vals.min(), vals.max()
        if abs(v_max - v_min) < 1e-8:
            v_norm = np.full_like(vals, 0.5)
        else:
            v_norm = 0.05 + 0.9 * (vals - v_min) / (v_max - v_min)

        # 绘制曲线（吃水为Y轴）
        line, = ax.plot(v_norm, d_arr,
                        color=color, linewidth=1.5,
                        marker='o', markersize=2.5,
                        label=f"{label} ({unit})")

        # 在曲线末端（最大吃水处）标注参数名
        ax.annotate(
            f"{label}\n[{v_min:.2f}~{v_max:.2f} {unit}]",
            xy=(v_norm[-1], d_arr[-1]),
            xytext=(v_norm[-1] + 0.02, d_arr[-1] - 0.3),
            fontsize=5.5, color=color,
            arrowprops=dict(arrowstyle='-', color=color, lw=0.5),
            bbox=dict(boxstyle='round,pad=0.2', facecolor='white',
                      edgecolor=color, alpha=0.85, linewidth=0.5),
        )

        legend_handles.append(line)
        legend_labels.append(f"{label} ({unit})")

    # X轴：归一化刻度，不显示数字（数字在标注中）
    ax.set_xlim(-0.05, 1.1)
    ax.xaxis.set_major_locator(ticker.MultipleLocator(0.25))
    ax.xaxis.set_ticklabels([])
    ax.tick_params(axis='x', length=3)

    # 图例
    ax.legend(handles=legend_handles, labels=legend_labels,
              loc='lower right', fontsize=5.5,
              framealpha=0.9, edgecolor='#c8d8f0',
              handlelength=1.5, handletextpad=0.4)

    # 边框
    for spine in ax.spines.values():
        spine.set_edgecolor('#2e7dd1')
        spine.set_linewidth(0.8)


# ══════════════════════════════════════════════════════════
# 邦戎曲线绘制
# ══════════════════════════════════════════════════════════

def plot_bonjean(bonjean: dict,
                 output_dir: str,
                 fmt: str = 'png',
                 ship_name: str = '') -> dict:
    """
    绘制邦戎曲线图（A3，300DPI）。

    布局：
        左图：各站横剖面面积 A(x,z) — 吃水曲线
        右图：各站面积静矩 S(x,z) — 吃水曲线

    坐标规范（符合船舶工程绘图标准）：
        X轴：面积 A（m²）或静矩 S（m³）
        Y轴：吃水 z（m），从下到上增大

    参数：
        bonjean    : dict  calc_bonjean_table() 返回的结果
        output_dir : str   输出目录
        fmt        : str   'png' 或 'svg'
        ship_name  : str   船名

    返回：
        dict { 'png': path, 'svg': path }
    """
    stations   = bonjean["stations"]
    waterlines = bonjean["waterlines"]
    A_mat      = np.array(bonjean["A"])   # [n_sta, n_wl]
    S_mat      = np.array(bonjean["S"])   # [n_sta, n_wl]
    n_sta      = bonjean["n_stations"]

    # ── 创建图幅（A3横向，2列子图）────────────────────────
    fig, (ax_A, ax_S) = plt.subplots(1, 2, figsize=(A3_W, A3_H),
                                      gridspec_kw={'wspace': 0.35})
    fig.patch.set_facecolor('#fafbff')

    # ── 大标题 ────────────────────────────────────────────
    title = "邦戎曲线图（Bonjean Curves）"
    if ship_name:
        title = f"{ship_name}  邦戎曲线图（Bonjean Curves）"
    fig.suptitle(title, fontsize=14, fontweight='bold', y=0.98,
                 color='#1a3a6a')
    fig.text(0.5, 0.945,
             "依据：盛振邦《船舶原理》上册 §2.3  |  武汉理工大学船海学院",
             ha='center', fontsize=8, color='#555', style='italic')

    # ── 左图：横剖面面积曲线 ──────────────────────────────
    _draw_bonjean_subplot(
        ax=ax_A,
        stations=stations,
        waterlines=waterlines,
        data_mat=A_mat,
        xlabel="横剖面面积 A (m2)",
        title="各站横剖面面积 A(x, z)",
        subtitle="A(x,z) = 2*int(0,z) y(x,z) dz",
    )

    # ── 右图：面积静矩曲线 ────────────────────────────────
    _draw_bonjean_subplot(
        ax=ax_S,
        stations=stations,
        waterlines=waterlines,
        data_mat=S_mat,
        xlabel="面积静矩 S (m3)",
        title="各站面积静矩 S(x, z)",
        subtitle="S(x,z) = 2*int(0,z) y(x,z)*z dz",
    )

    # ── 图框装饰 ──────────────────────────────────────────
    _add_border(fig)

    # ── 保存 ──────────────────────────────────────────────
    os.makedirs(output_dir, exist_ok=True)
    paths = {}

    png_path = os.path.join(output_dir, "邦戎曲线图.png")
    fig.savefig(png_path, dpi=300, bbox_inches='tight',
                facecolor=fig.get_facecolor())
    paths['png'] = png_path

    svg_path = os.path.join(output_dir, "邦戎曲线图.svg")
    fig.savefig(svg_path, format='svg', bbox_inches='tight',
                facecolor=fig.get_facecolor())
    paths['svg'] = svg_path

    plt.close(fig)
    return paths


def _draw_bonjean_subplot(ax, stations, waterlines, data_mat,
                           xlabel, title, subtitle):
    """
    绘制单个邦戎曲线子图（面积或静矩）。

    每条曲线对应一个站，X轴为参数值，Y轴为吃水。
    """
    ax.set_facecolor('#f8faff')
    ax.grid(True, linestyle='--', linewidth=0.4, color='#c8d8f0', alpha=0.8)

    wl_arr = np.array(waterlines)
    n_sta  = len(stations)

    # 绘制每个站的曲线
    legend_handles = []
    legend_labels  = []

    for i, x_sta in enumerate(stations):
        color = COLORS[i % len(COLORS)]
        vals  = data_mat[i]   # 该站在各水线处的值

        line, = ax.plot(vals, wl_arr,
                        color=color, linewidth=1.4,
                        marker='o', markersize=2.5,
                        label=f"x={x_sta:.1f}m")

        # 在曲线末端（最大吃水处）标注站号
        ax.annotate(
            f"x={x_sta:.1f}",
            xy=(vals[-1], wl_arr[-1]),
            xytext=(vals[-1] * 1.01, wl_arr[-1] - 0.15),
            fontsize=5.5, color=color,
            va='top',
        )

        legend_handles.append(line)
        legend_labels.append(f"x={x_sta:.1f}m")

    # 坐标轴设置
    ax.set_xlabel(xlabel, fontsize=8.5, color='#1a3a6a')
    ax.set_ylabel("吃水 z (m)", fontsize=8.5, color='#1a3a6a')
    ax.set_title(f"{title}\n{subtitle}",
                 fontsize=9, fontweight='bold', color='#1a3a6a', pad=8)

    ax.set_ylim(wl_arr.min() - 0.1, wl_arr.max() + 0.3)
    ax.yaxis.set_major_locator(ticker.MultipleLocator(1.0))
    ax.yaxis.set_minor_locator(ticker.MultipleLocator(0.5))
    ax.tick_params(labelsize=7)

    # X轴从0开始
    ax.set_xlim(left=0)
    ax.xaxis.set_major_formatter(ticker.FormatStrFormatter('%.1f'))

    # 图例（分两列）
    ncol = max(1, n_sta // 8)
    ax.legend(handles=legend_handles, labels=legend_labels,
              loc='lower right', fontsize=5.5, ncol=ncol,
              framealpha=0.9, edgecolor='#c8d8f0',
              handlelength=1.5, handletextpad=0.4)

    # 边框
    for spine in ax.spines.values():
        spine.set_edgecolor('#2e7dd1')
        spine.set_linewidth(0.8)


# ══════════════════════════════════════════════════════════
# 工具函数
# ══════════════════════════════════════════════════════════

def _add_border(fig):
    """在图幅四周添加装饰边框（模拟工程图纸边框）"""
    fig.add_artist(
        plt.Rectangle((0.01, 0.01), 0.98, 0.97,
                       fill=False, edgecolor='#2e7dd1',
                       linewidth=1.5, transform=fig.transFigure)
    )


def fig_to_base64(fig) -> str:
    """将 matplotlib Figure 转为 base64 字符串（用于前端预览）"""
    import base64
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=120, bbox_inches='tight',
                facecolor=fig.get_facecolor())
    buf.seek(0)
    return base64.b64encode(buf.read()).decode('utf-8')


def plot_hydrostatics_preview(table: dict) -> str:
    """
    生成静水力曲线预览图（低分辨率，用于前端展示）。
    返回 base64 编码的 PNG 字符串。
    """
    rows   = table["rows"]
    drafts = [r["d"] for r in rows]

    # 预览图：3列子图，精简版
    fig, axes = plt.subplots(1, 3, figsize=(12, 6),
                              gridspec_kw={'wspace': 0.4})
    fig.patch.set_facecolor('#fafbff')
    fig.suptitle("静水力曲线预览", fontsize=11, fontweight='bold',
                 color='#1a3a6a', y=0.98)

    preview_groups = [
        [{"key": "V",         "label": "V(m3)",    "unit": "m3"},
         {"key": "delta_sea", "label": "D海(t)",   "unit": "t"}],
        [{"key": "zb",        "label": "KB(m)",    "unit": "m"},
         {"key": "BM",        "label": "BM(m)",    "unit": "m"},
         {"key": "zM",        "label": "KM(m)",    "unit": "m"}],
        [{"key": "Cb",        "label": "Cb",       "unit": "-"},
         {"key": "Cwp",       "label": "Cwp",      "unit": "-"},
         {"key": "TPC",       "label": "TPC",      "unit": "t/cm"}],
    ]

    for ax, curves in zip(axes, preview_groups):
        _draw_hydro_group(ax, rows, drafts, curves, "")

    b64 = fig_to_base64(fig)
    plt.close(fig)
    return b64


def plot_bonjean_preview(bonjean: dict) -> str:
    """
    生成邦戎曲线预览图（低分辨率，用于前端展示）。
    返回 base64 编码的 PNG 字符串。
    """
    stations   = bonjean["stations"]
    waterlines = bonjean["waterlines"]
    A_mat      = np.array(bonjean["A"])

    fig, ax = plt.subplots(figsize=(7, 5))
    fig.patch.set_facecolor('#fafbff')
    ax.set_facecolor('#f8faff')
    ax.grid(True, linestyle='--', linewidth=0.4, color='#c8d8f0', alpha=0.8)

    wl_arr = np.array(waterlines)
    for i, x_sta in enumerate(stations):
        color = COLORS[i % len(COLORS)]
        ax.plot(A_mat[i], wl_arr, color=color, linewidth=1.2,
                marker='o', markersize=2, label=f"x={x_sta:.1f}")

    ax.set_xlabel("横剖面面积 A (m²)", fontsize=8)
    ax.set_ylabel("吃水 z (m)", fontsize=8)
    ax.set_title("邦戎曲线预览（各站横剖面面积）", fontsize=9,
                 fontweight='bold', color='#1a3a6a')
    ax.set_xlim(left=0)
    ax.legend(fontsize=5.5, ncol=2, loc='lower right',
              framealpha=0.9, edgecolor='#c8d8f0')
    ax.tick_params(labelsize=7)

    b64 = fig_to_base64(fig)
    plt.close(fig)
    return b64
