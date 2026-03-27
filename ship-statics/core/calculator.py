# ============================================================
# core/calculator.py — 船舶静水力核心计算模块
#
# 严格遵循：盛振邦《船舶原理》上册 第2章 船舶静力学
#
# 坐标系定义（必须严格遵守）：
#   X 轴：沿船长方向，船首为正，船尾为负，中站 x=0
#   Z 轴：垂直向上，基线 Z=0
#   y   ：半宽（单侧），计算时乘以 2 得全宽
#
# 计算方法：梯形法（Trapezoidal Rule）
#
# 核心公式（对应教材）：
#   ∫f(x)dx ≈ Σ [f(x_i) + f(x_{i+1})] / 2 * Δx
#
# ============================================================

import numpy as np


# ── 海水密度常数 ──────────────────────────────────────────
# 按《国内航行海船法定检验技术规则》取值
RHO_SEAWATER = 1.025  # t/m³


def trapz_integrate(y_values: list, x_values: list) -> float:
    """
    梯形法数值积分（通用函数）
    
    原理：
        ∫f(x)dx ≈ Σ [f(x_i) + f(x_{i+1})] / 2 * (x_{i+1} - x_i)
    
    参数：
        y_values : list[float]  被积函数值数组，如 [y0, y1, y2, ...]
        x_values : list[float]  自变量数组，如 [x0, x1, x2, ...]
                                不要求等间距
    
    返回：
        float  积分结果
    
    示例：
        >>> x = [0, 1, 2, 3]
        >>> y = [0, 1, 4, 9]  # y = x²
        >>> trapz_integrate(y, x)
        8.5  # 近似值
    """
    # 转换为 numpy 数组便于计算
    y = np.array(y_values, dtype=float)
    x = np.array(x_values, dtype=float)
    
    # 使用 numpy 的梯形法积分
    # numpy 2.0+ 用 trapezoid，旧版用 trapz，兼容两者
    trapz_fn = getattr(np, 'trapezoid', None) or getattr(np, 'trapz')
    result = trapz_fn(y, x)
    
    return float(result)


def calculate_hydrostatics(stations: list, half_breadths: list, draft: float) -> dict:
    """
    计算设计吃水下的静水力要素（梯形法）
    
    这是阶段1的核心计算函数，对应课程设计的基本要求。
    
    计算流程（严格对应教材第2章）：
    ① 用梯形法计算水线面积 Aw = 2 * ∫ y dx
    ② 用梯形法计算排水体积 ∇ = ∫ Aw(z) dz （从基线到设计吃水）
    ③ 计算排水量 Δ = ρ * ∇
    ④ 计算浮心纵向坐标 xb = (1/∇) * ∫ Aw(z) * xf(z) dz
    ⑤ 计算每厘米吃水吨数 TPC = Aw * ρ / 100
    ⑥ 计算方形系数 Cb = ∇ / (L * B * d)
    
    参数：
        stations      : list[float]  各站 x 坐标，如 [-5, -4, -3, ..., 5]
        half_breadths : list[float]  设计吃水处各站半宽 y，如 [0, 1.2, 2.5, ...]
        draft         : float        设计吃水 d（米）
    
    返回：
        dict，包含以下计算结果：
            Aw   : float  水线面积（m²）
            V    : float  排水体积（m³）
            delta: float  排水量（t）
            xb   : float  浮心纵向坐标（m）
            TPC  : float  每厘米吃水吨数（t/cm）
            Cb   : float  方形系数（无量纲）
            L    : float  船长（m）
            B    : float  船宽（m）
            d    : float  设计吃水（m）
    
    异常：
        ValueError: 当输入数据不合理时抛出
    """
    
    # ── 第①步：数据校验 ────────────────────────────────────
    if len(stations) != len(half_breadths):
        raise ValueError("站号数量与半宽数量不匹配")
    
    if len(stations) < 3:
        raise ValueError("至少需要3个站")
    
    if draft <= 0:
        raise ValueError("设计吃水必须大于0")
    
    # 检查半宽是否有负值
    for i, y in enumerate(half_breadths):
        if y < 0:
            raise ValueError(f"第{i+1}站半宽为负值({y})，物理上不合理")
    
    # ── 第②步：计算基本参数 ────────────────────────────────
    # 船长 L = 首站 x - 尾站 x
    L = max(stations) - min(stations)
    
    # 船宽 B = 2 * 最大半宽（设计吃水处）
    B = 2.0 * max(half_breadths)
    
    # ── 第③步：计算水线面积 Aw ────────────────────────────
    # 公式：Aw = 2 * ∫ y dx  （乘2因为 y 是半宽）
    # 使用梯形法积分
    Aw = 2.0 * trapz_integrate(half_breadths, stations)
    
    if Aw < 1e-6:
        raise ValueError("水线面积计算结果为零或极小，请检查输入数据")
    
    # ── 第④步：计算排水体积 ∇ ────────────────────────────
    # 简化假设：在设计吃水处，排水体积 ∇ ≈ Aw * d
    # （严格计算需要沿吃水方向积分，但阶段1简化处理）
    V = Aw * draft
    
    # ── 第⑤步：计算排水量 Δ ────────────────────────────────
    # 公式：Δ = ρ * ∇  （海水 ρ = 1.025 t/m³）
    delta = RHO_SEAWATER * V
    
    # ── 第⑥步：计算浮心纵向坐标 xb ────────────────────────
    # 公式：xb = (1/Aw) * ∫ y * x dx
    # 计算水线面积对中站的静矩
    y_x_products = [y * x for y, x in zip(half_breadths, stations)]
    static_moment = 2.0 * trapz_integrate(y_x_products, stations)
    xb = static_moment / Aw
    
    # ── 第⑦步：计算每厘米吃水吨数 TPC ──────────────────────
    # 公式：TPC = Aw * ρ / 100  （单位：t/cm）
    TPC = Aw * RHO_SEAWATER / 100.0
    
    # ── 第⑧步：计算方形系数 Cb ────────────────────────────
    # 公式：Cb = ∇ / (L * B * d)
    # 表示船体体积与外包长方体体积的比值
    if L > 0 and B > 0 and draft > 0:
        Cb = V / (L * B * draft)
    else:
        Cb = 0.0
    
    # ── 返回结果字典 ──────────────────────────────────────
    return {
        "Aw":    round(Aw, 4),      # 水线面积（m²）
        "V":     round(V, 4),       # 排水体积（m³）
        "delta": round(delta, 4),   # 排水量（t）
        "xb":    round(xb, 4),      # 浮心纵向坐标（m）
        "TPC":   round(TPC, 4),     # 每厘米吃水吨数（t/cm）
        "Cb":    round(Cb, 4),      # 方形系数（无量纲）
        "L":     round(L, 4),       # 船长（m）
        "B":     round(B, 4),       # 船宽（m）
        "d":     round(draft, 4),   # 设计吃水（m）
    }
