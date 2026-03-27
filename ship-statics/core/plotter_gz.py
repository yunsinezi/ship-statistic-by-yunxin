import os
import math
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from matplotlib import rcParams

# 复用 plotter.py 中的配置
from core.plotter import (COLORS, A3_W, A3_H, _add_border,
                           fig_to_base64, _setup_font)
_setup_font()




def plot_gz_curve(gz_result: dict,
                  output_dir: str,
                  ship_name: str = '') -> dict:
    """
    绘制单工况 GZ 曲线图（A3，300DPI）。

    图面内容：
        主图：GZ-θ 曲线（复原力臂 vs 横倾角）
        辅助线：初稳性近似线（GZ = GM·sinθ）
        标注：GZ_max、θ_max、θ_vanish、GM 等特征值
        衡准线：规范要求的最小 GZ 值（0.20m）
        面积阴影：0°~30°、30°~40° 稳性面积

    参数：
        gz_result  : dict  calc_gz_curve() 返回的结果
        output_dir : str   输出目录
        ship_name  : str   船名

    返回：
        dict { 'png': path, 'svg': path }
    """
    rows         = gz_result["rows"]
    GM           = gz_result["GM"]
    KG           = gz_result["KG"]
    KB           = gz_result["KB"]
    BM           = gz_result["BM"]
    draft        = gz_result["draft"]
    delta        = gz_result["delta"]
    GZ_max       = gz_result["GZ_max"]
    theta_max_gz = gz_result["theta_max_gz"]
    theta_vanish = gz_result.get("theta_vanish")
    criteria     = gz_result["criteria"]

    thetas    = np.array([r['theta']     for r in rows])
    gz_arr    = np.array([r['GZ']        for r in rows])
    gz_approx = np.array([r['GZ_approx'] for r in rows])

    # ── 创建图幅（A3横向，主图+衡准表）──────────────────
    fig = plt.figure(figsize=(A3_W, A3_H))
    fig.patch.set_facecolor('#fafbff')

    gs = fig.add_gridspec(1, 2, width_ratios=[7, 3], wspace=0.08,
                          left=0.06, right=0.97, top=0.88, bottom=0.10)
    ax_main  = fig.add_subplot(gs[0])
    ax_table = fig.add_subplot(gs[1])

    # ── 大标题 ────────────────────────────────────────────
    title = "稳性横截曲线图（GZ 曲线）"
    if ship_name:
        title = f"{ship_name}  稳性横截曲线图（GZ 曲线）"
    fig.suptitle(title, fontsize=13, fontweight='bold', y=0.96,
                 color='#1a3a6a')
    fig.text(0.5, 0.915,
             f"d={draft:.2f}m  KG={KG:.3f}m  GM={GM:.3f}m  "
             f"D={delta:.1f}t  |  盛振邦《船舶原理》上册 §3.2  |  武汉理工大学船海学院",
             ha='center', fontsize=8, color='#555', style='italic')

    # ── 主图：GZ 曲线 ─────────────────────────────────────
    ax = ax_main
    ax.set_facecolor('#f8faff')
    ax.grid(True, linestyle='--', linewidth=0.5, color='#c8d8f0', alpha=0.8)

    # 稳性面积阴影（0°~30°）
    mask_30 = thetas <= 30
    ax.fill_between(thetas[mask_30], 0, gz_arr[mask_30],
                    alpha=0.15, color='#2e7dd1', label='GZ面积(0°~30°)')

    # 稳性面积阴影（30°~40°）
    mask_40 = (thetas >= 30) & (thetas <= 40)
    ax.fill_between(thetas[mask_40], 0, gz_arr[mask_40],
                    alpha=0.20, color='#059669', label='GZ面积(30°~40°)')

    # 初稳性近似线（虚线）
    ax.plot(thetas, gz_approx, '--', color='#aaa', linewidth=1.0,
            label=f'初稳性近似 GM·sin(θ)={GM:.3f}m')

    # 主 GZ 曲线
    ax.plot(thetas, gz_arr, '-o', color='#d62728', linewidth=2.0,
            markersize=3.5, label='GZ 曲线（变排水量法）', zorder=5)

    # 零线
    ax.axhline(0, color='#333', linewidth=0.8, linestyle='-')

    # 规范最小 GZ 衡准线（0.20m）
    ax.axhline(0.20, color='#ff7f0e', linewidth=1.0, linestyle=':',
               label='规范最小 GZ=0.20m')

    # 标注 GZ_max
    ax.annotate(
        f'GZ_max={GZ_max:.3f}m\n@{theta_max_gz:.0f}deg',
        xy=(theta_max_gz, GZ_max),
        xytext=(theta_max_gz + 5, GZ_max + 0.02),
        fontsize=8, color='#d62728', fontweight='bold',
        arrowprops=dict(arrowstyle='->', color='#d62728', lw=1.2),
        bbox=dict(boxstyle='round,pad=0.3', facecolor='#fff0f0',
                  edgecolor='#d62728', alpha=0.9),
    )

    # 标注稳性消失角
    if theta_vanish is not None:
        ax.axvline(theta_vanish, color='#9467bd', linewidth=1.0,
                   linestyle='--', alpha=0.7)
        y_lim = ax.get_ylim()
        ax.text(theta_vanish + 1, y_lim[0] + 0.01,
                f'theta_v={theta_vanish:.1f}deg',
                fontsize=7.5, color='#9467bd', va='bottom')

    # 标注 GM
    ax.text(5, GM * math.sin(math.radians(5)) + 0.01,
            f'GM={GM:.3f}m', fontsize=7.5, color='#555')

    ax.set_xlabel("横倾角 θ (°)", fontsize=10, color='#1a3a6a')
    ax.set_ylabel("复原力臂 GZ (m)", fontsize=10, color='#1a3a6a')
    ax.set_title("GZ-θ 曲线（稳性横截曲线）", fontsize=10,
                 fontweight='bold', color='#1a3a6a', pad=8)
    ax.set_xlim(0, max(thetas) + 2)
    ax.xaxis.set_major_locator(ticker.MultipleLocator(10))
    ax.xaxis.set_minor_locator(ticker.MultipleLocator(5))
    ax.tick_params(labelsize=8)
    ax.legend(loc='upper right', fontsize=7.5, framealpha=0.9,
              edgecolor='#c8d8f0')

    for spine in ax.spines.values():
        spine.set_edgecolor('#2e7dd1')
        spine.set_linewidth(0.8)

    # ── 右侧：稳性衡准表 ──────────────────────────────────
    ax_table.axis('off')
    ax_table.set_facecolor('#f8faff')

    ax_table.text(0.5, 0.97, "稳性衡准校核",
                  ha='center', va='top', fontsize=9, fontweight='bold',
                  color='#1a3a6a', transform=ax_table.transAxes)
    ax_table.text(0.5, 0.93, "《国内航行海船法定检验技术规则》",
                  ha='center', va='top', fontsize=6.5, color='#555',
                  style='italic', transform=ax_table.transAxes)

    params = [
        ("吃水 d",   f"{draft:.3f} m"),
        ("KG",       f"{KG:.3f} m"),
        ("KB",       f"{KB:.3f} m"),
        ("BM",       f"{BM:.3f} m"),
        ("KM",       f"{KB+BM:.3f} m"),
        ("GM",       f"{GM:.3f} m"),
        ("排水量 D", f"{delta:.1f} t"),
    ]
    y_pos = 0.88
    for name, val in params:
        ax_table.text(0.05, y_pos, name, fontsize=7.5, color='#333',
                      transform=ax_table.transAxes)
        ax_table.text(0.95, y_pos, val, fontsize=7.5, color='#1a4a8a',
                      ha='right', fontweight='bold',
                      transform=ax_table.transAxes)
        y_pos -= 0.055

    y_pos -= 0.03
    crit_items = [
        ("GM",           criteria.get("GM")),
        ("GZ_max",       criteria.get("GZ_max")),
        ("theta_max",    criteria.get("theta_max")),
        ("theta_vanish", criteria.get("theta_vanish")),
        ("area_0_30",    criteria.get("area_0_30")),
        ("area_0_40",    criteria.get("area_0_40")),
        ("area_30_40",   criteria.get("area_30_40")),
    ]

    for _, crit in crit_items:
        if crit is None:
            continue
        ok    = crit.get("pass", False)
        color = '#059669' if ok else '#d62728'
        mark  = "OK" if ok else "NG"
        name  = crit.get("name", "")
        val   = crit.get("value")
        limit = crit.get("limit")
        unit  = crit.get("unit", "")

        val_str = f"{val:.3f}" if isinstance(val, float) else str(val)
        lim_str = f">={limit}"

        ax_table.text(0.05, y_pos, f"[{mark}] {name}",
                      fontsize=6.5, color=color,
                      transform=ax_table.transAxes)
        ax_table.text(0.95, y_pos, f"{val_str} {unit} ({lim_str})",
                      fontsize=6.0, color=color, ha='right',
                      transform=ax_table.transAxes)
        y_pos -= 0.065

    overall = criteria.get("overall", {})
    ok_all  = overall.get("pass", False)
    conclusion = "满足稳性衡准" if ok_all else "不满足稳性衡准"
    ax_table.text(0.5, max(y_pos - 0.02, 0.03),
                  conclusion,
                  ha='center', fontsize=9, fontweight='bold',
                  color='#059669' if ok_all else '#d62728',
                  transform=ax_table.transAxes,
                  bbox=dict(boxstyle='round,pad=0.4',
                            facecolor='#f0fff4' if ok_all else '#fff0f0',
                            edgecolor='#059669' if ok_all else '#d62728',
                            alpha=0.9))

    _add_border(fig)

    os.makedirs(output_dir, exist_ok=True)
    paths = {}

    png_path = os.path.join(output_dir, "GZ曲线图.png")
    fig.savefig(png_path, dpi=300, bbox_inches='tight',
                facecolor=fig.get_facecolor())
    paths['png'] = png_path

    svg_path = os.path.join(output_dir, "GZ曲线图.svg")
    fig.savefig(svg_path, format='svg', bbox_inches='tight',
                facecolor=fig.get_facecolor())
    paths['svg'] = svg_path

    plt.close(fig)
    return paths


def plot_gz_family(gz_family: list,
                   output_dir: str,
                   ship_name: str = '') -> dict:
    """
    绘制多排水量 GZ 曲线族图（A3，300DPI）。
    """
    fig, ax = plt.subplots(figsize=(A3_W, A3_H))
    fig.patch.set_facecolor('#fafbff')
    ax.set_facecolor('#f8faff')
    ax.grid(True, linestyle='--', linewidth=0.5, color='#c8d8f0', alpha=0.8)

    title = "GZ 曲线族（多排水量）"
    if ship_name:
        title = f"{ship_name}  GZ 曲线族（多排水量）"
    fig.suptitle(title, fontsize=13, fontweight='bold', y=0.97,
                 color='#1a3a6a')

    for i, res in enumerate(gz_family):
        if "error" in res:
            continue
        rows  = res["rows"]
        draft = res["draft"]
        delta = res["delta"]
        GM    = res["GM"]
        color = COLORS[i % len(COLORS)]

        thetas = [r['theta'] for r in rows]
        gz_arr = [r['GZ']    for r in rows]

        ax.plot(thetas, gz_arr, '-o', color=color, linewidth=1.5,
                markersize=2.5,
                label=f"d={draft:.1f}m  D={delta:.0f}t  GM={GM:.3f}m")

    ax.axhline(0,    color='#333', linewidth=0.8)
    ax.axhline(0.20, color='#ff7f0e', linewidth=1.0, linestyle=':',
               label='规范最小 GZ=0.20m')

    ax.set_xlabel("横倾角 θ (°)", fontsize=10, color='#1a3a6a')
    ax.set_ylabel("复原力臂 GZ (m)", fontsize=10, color='#1a3a6a')
    ax.set_title("GZ 曲线族（不同排水量）", fontsize=10,
                 fontweight='bold', color='#1a3a6a', pad=8)
    ax.xaxis.set_major_locator(ticker.MultipleLocator(10))
    ax.tick_params(labelsize=8)
    ax.legend(loc='upper right', fontsize=7, framealpha=0.9,
              edgecolor='#c8d8f0')

    for spine in ax.spines.values():
        spine.set_edgecolor('#2e7dd1')
        spine.set_linewidth(0.8)

    _add_border(fig)

    os.makedirs(output_dir, exist_ok=True)
    paths = {}

    png_path = os.path.join(output_dir, "GZ曲线族.png")
    fig.savefig(png_path, dpi=300, bbox_inches='tight',
                facecolor=fig.get_facecolor())
    paths['png'] = png_path

    svg_path = os.path.join(output_dir, "GZ曲线族.svg")
    fig.savefig(svg_path, format='svg', bbox_inches='tight',
                facecolor=fig.get_facecolor())
    paths['svg'] = svg_path

    plt.close(fig)
    return paths


def plot_gz_preview(gz_result: dict) -> str:
    """
    生成 GZ 曲线预览图（低分辨率，用于前端展示）。
    返回 base64 编码的 PNG 字符串。
    """
    rows         = gz_result["rows"]
    GM           = gz_result["GM"]
    GZ_max       = gz_result["GZ_max"]
    theta_max_gz = gz_result["theta_max_gz"]
    draft        = gz_result["draft"]
    KG           = gz_result["KG"]

    thetas    = [r['theta']     for r in rows]
    gz_arr    = [r['GZ']        for r in rows]
    gz_approx = [r['GZ_approx'] for r in rows]

    fig, ax = plt.subplots(figsize=(9, 5.5))
    fig.patch.set_facecolor('#fafbff')
    ax.set_facecolor('#f8faff')
    ax.grid(True, linestyle='--', linewidth=0.4, color='#c8d8f0', alpha=0.8)

    t_arr = np.array(thetas)
    g_arr = np.array(gz_arr)
    mask_30 = t_arr <= 30
    ax.fill_between(t_arr[mask_30], 0, g_arr[mask_30],
                    alpha=0.15, color='#2e7dd1')
    mask_40 = (t_arr >= 30) & (t_arr <= 40)
    ax.fill_between(t_arr[mask_40], 0, g_arr[mask_40],
                    alpha=0.20, color='#059669')

    ax.plot(thetas, gz_approx, '--', color='#aaa', linewidth=1.0,
            label=f'GM·sin(θ)={GM:.3f}m')
    ax.plot(thetas, gz_arr, '-o', color='#d62728', linewidth=2.0,
            markersize=3, label='GZ 曲线', zorder=5)
    ax.axhline(0,    color='#333', linewidth=0.8)
    ax.axhline(0.20, color='#ff7f0e', linewidth=1.0, linestyle=':',
               label='规范最小 0.20m')

    ax.annotate(f'GZ_max={GZ_max:.3f}m',
                xy=(theta_max_gz, GZ_max),
                xytext=(theta_max_gz + 4, GZ_max + 0.01),
                fontsize=8, color='#d62728', fontweight='bold',
                arrowprops=dict(arrowstyle='->', color='#d62728', lw=1.0))

    ax.set_xlabel("横倾角 θ (°)", fontsize=9)
    ax.set_ylabel("GZ (m)", fontsize=9)
    ax.set_title(f"GZ 曲线预览  d={draft:.2f}m  KG={KG:.3f}m  GM={GM:.3f}m",
                 fontsize=9, fontweight='bold', color='#1a3a6a')
    ax.legend(fontsize=7.5, loc='upper right', framealpha=0.9)
    ax.tick_params(labelsize=7.5)

    b64 = fig_to_base64(fig)
    plt.close(fig)
    return b64
