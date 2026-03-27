# ============================================================
# core/bonjean.py — 邦戎曲线计算模块
#
# 阶段4 新增模块
#
# 严格遵循：盛振邦《船舶原理》上册 §2.3 邦戎曲线
#           武汉理工大学船海学院课程设计规范
#
# 邦戎曲线（Bonjean Curves）定义：
#   以各站横剖面面积 A(x, z) 和面积静矩 S(x, z) 为纵坐标，
#   以吃水 z 为横坐标，对每个站绘制的曲线族。
#
# 用途：
#   1. 快速查取任意吃水下各站横剖面面积
#   2. 计算不均匀装载下的排水量和浮心位置
#   3. 为静水力曲线计算提供基础数据
#
# 计算公式（梯形法）：
#   A(x, z) = 2 · ∫₀ᶻ y(x, ζ) dζ    （横剖面面积，m²）
#   S(x, z) = 2 · ∫₀ᶻ y(x, ζ)·ζ dζ  （面积对基线的静矩，m³）
#
# 其中 y(x, ζ) 为站 x 处、水线高度 ζ 处的半宽（m）
# ============================================================

import numpy as np


def _trapz(y: np.ndarray, x: np.ndarray) -> float:
    """梯形法积分（兼容 numpy 1.x / 2.x）"""
    fn = getattr(np, 'trapezoid', None) or getattr(np, 'trapz')
    return float(fn(y, x))


def calc_bonjean_table(offsets_data: dict,
                       method: str = 'trapz') -> dict:
    """
    计算完整邦戎曲线数据表。

    对每个站（x 坐标），在每条水线高度 z 处，
    用梯形法沿 Z 方向积分，计算：
        A(x, z) = 2·∫₀ᶻ y(x,ζ) dζ   横剖面面积（m²）
        S(x, z) = 2·∫₀ᶻ y(x,ζ)·ζ dζ  面积静矩（m³）

    对应教材：盛振邦《船舶原理》上册 §2.3

    参数：
        offsets_data : dict  型值表数据，包含：
                             stations   : list[float]  各站 x 坐标
                             waterlines : list[float]  各水线 z 坐标（从小到大）
                             offsets    : list[list]   半宽矩阵 [wl_idx][sta_idx]
        method       : str   积分方法（'trapz' 默认）

    返回：
        dict，包含：
            stations   : list[float]   各站 x 坐标
            waterlines : list[float]   各水线 z 坐标
            A          : list[list]    横剖面面积 A[sta_idx][wl_idx]（m²）
            S          : list[list]    面积静矩   S[sta_idx][wl_idx]（m³）
            n_stations : int
            n_waterlines: int
    """
    stations   = np.array(offsets_data["stations"],   dtype=float)
    waterlines = np.array(offsets_data["waterlines"], dtype=float)
    offsets    = np.array(offsets_data["offsets"],    dtype=float)
    # offsets 形状：[n_waterlines, n_stations]

    n_sta = len(stations)
    n_wl  = len(waterlines)

    # 结果矩阵：A[sta_idx][wl_idx]，S[sta_idx][wl_idx]
    A_mat = np.zeros((n_sta, n_wl))
    S_mat = np.zeros((n_sta, n_wl))

    for i in range(n_sta):
        # 取该站在各水线处的半宽序列
        y_col = offsets[:, i]   # shape: [n_wl]

        for j in range(n_wl):
            if j == 0:
                # 基线处面积为零
                A_mat[i, j] = 0.0
                S_mat[i, j] = 0.0
                continue

            # 取从基线（j=0）到当前水线（j）的切片
            z_slice = waterlines[:j+1]   # z 坐标切片
            y_slice = y_col[:j+1]        # 半宽切片

            # ── 横剖面面积 A(x, z) = 2·∫₀ᶻ y(x,ζ) dζ ──────
            # 物理意义：该站在吃水 z 处的横剖面面积（双侧）
            A_mat[i, j] = 2.0 * _trapz(y_slice, z_slice)

            # ── 面积静矩 S(x, z) = 2·∫₀ᶻ y(x,ζ)·ζ dζ ──────
            # 物理意义：横剖面面积对基线的静矩
            # 用于计算浮心高度：zb = S_total / V_total
            S_mat[i, j] = 2.0 * _trapz(y_slice * z_slice, z_slice)

    return {
        "stations":    [round(float(x), 4) for x in stations],
        "waterlines":  [round(float(z), 4) for z in waterlines],
        "A":           [[round(float(A_mat[i, j]), 4) for j in range(n_wl)]
                        for i in range(n_sta)],
        "S":           [[round(float(S_mat[i, j]), 4) for j in range(n_wl)]
                        for i in range(n_sta)],
        "n_stations":  n_sta,
        "n_waterlines": n_wl,
    }


def bonjean_at_draft(bonjean: dict, draft: float) -> dict:
    """
    从邦戎曲线数据中插值，获取指定吃水下各站的横剖面面积和静矩。

    参数：
        bonjean : dict   calc_bonjean_table() 的返回值
        draft   : float  目标吃水（m）

    返回：
        dict，包含：
            stations : list[float]  各站 x 坐标
            A        : list[float]  各站横剖面面积（m²）
            S        : list[float]  各站面积静矩（m³）
    """
    stations   = np.array(bonjean["stations"])
    waterlines = np.array(bonjean["waterlines"])
    A_mat      = np.array(bonjean["A"])   # [n_sta, n_wl]
    S_mat      = np.array(bonjean["S"])   # [n_sta, n_wl]

    # 线性插值
    A_at_draft = np.interp(draft, waterlines, A_mat.T).tolist()
    # np.interp 对每列插值：A_mat.T 形状 [n_wl, n_sta]
    # 实际上需要对每个站分别插值
    A_result = []
    S_result = []
    for i in range(len(stations)):
        A_result.append(round(float(np.interp(draft, waterlines, A_mat[i])), 4))
        S_result.append(round(float(np.interp(draft, waterlines, S_mat[i])), 4))

    return {
        "stations": bonjean["stations"],
        "A":        A_result,
        "S":        S_result,
        "draft":    round(float(draft), 3),
    }
