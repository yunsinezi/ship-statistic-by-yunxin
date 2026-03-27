# ============================================================
# core/hydrostatics_full.py — 完整静水力全参数计算模块
#
# 阶段3 新增模块
#
# 严格遵循：盛振邦《船舶原理》上册 第2章 船舶静力学
#           武汉理工大学船海学院《船舶静力学》课程设计规范
#
# 坐标系（严格遵守）：
#   X 轴：沿船长方向，船首为正，船尾为负，中站 x=0
#   Z 轴：垂直向上，基线 Z=0
#   y   ：半宽（单侧），计算时乘以 2 得全宽
#
# 计算方法：
#   主方法：梯形法（Trapezoidal Rule）
#   可选方法：辛普森 1/3 法（Simpson's 1/3 Rule）
#
# 全部 14 项静水力参数（对应教材）：
#   ∇    排水体积（m³）
#   Δ    排水量（t）—— 海水 / 淡水双版本
#   xb   浮心纵向坐标（m）
#   zb   浮心垂向坐标（m），即 KB
#   Aw   水线面面积（m²）
#   xf   漂心纵向坐标（m）
#   TPC  每厘米吃水吨数（t/cm）
#   MTC  每厘米纵倾力矩（t·m/cm）
#   BM   横稳心半径（m）
#   BML  纵稳心半径（m）
#   zM   横稳心高度（m），即 KM = KB + BM
#   Cb   方形系数
#   Cp   棱形系数
#   Cwp  水线面系数
#   Cm   中横剖面系数
# ============================================================

import numpy as np

# ── 密度常数 ──────────────────────────────────────────────
RHO_SEA   = 1.025   # 海水密度（t/m³），按规范取值
RHO_FRESH = 1.000   # 淡水密度（t/m³）


# ══════════════════════════════════════════════════════════
# 积分工具函数
# ══════════════════════════════════════════════════════════

def _trapz(y: np.ndarray, x: np.ndarray) -> float:
    """
    梯形法数值积分（内部工具函数）

    公式：∫f(x)dx ≈ Σ [f(x_i)+f(x_{i+1})]/2 · (x_{i+1}-x_i)

    参数：
        y : np.ndarray  被积函数值
        x : np.ndarray  自变量（不要求等间距）
    返回：
        float  积分值
    """
    fn = getattr(np, 'trapezoid', None) or getattr(np, 'trapz')
    return float(fn(y, x))


def _simpson(y: np.ndarray, x: np.ndarray) -> float:
    """
    辛普森 1/3 法数值积分（可选方法）

    要求：等间距，且点数为奇数（偶数段）
    若不满足条件，自动退化为梯形法。

    公式（等间距 h）：
        ∫f(x)dx ≈ h/3 · [f(x0) + 4f(x1) + 2f(x2) + 4f(x3) + ... + f(xn)]

    参数：
        y : np.ndarray  被积函数值
        x : np.ndarray  自变量（要求等间距）
    返回：
        float  积分值
    """
    n = len(y)
    if n < 3 or (n - 1) % 2 != 0:
        # 不满足辛普森条件，退化为梯形法
        return _trapz(y, x)

    h = (x[-1] - x[0]) / (n - 1)
    # 辛普森系数：1, 4, 2, 4, 2, ..., 4, 1
    coeffs = np.ones(n)
    coeffs[1:-1:2] = 4   # 奇数位系数为 4
    coeffs[2:-2:2] = 2   # 偶数位系数为 2

    return float(h / 3.0 * np.dot(coeffs, y))


def integrate(y: np.ndarray, x: np.ndarray, method: str = 'trapz') -> float:
    """
    统一积分接口，支持梯形法和辛普森法。

    参数：
        y      : np.ndarray  被积函数值
        x      : np.ndarray  自变量
        method : str         'trapz'（梯形法，默认）或 'simpson'（辛普森法）
    返回：
        float  积分值
    """
    if method == 'simpson':
        return _simpson(y, x)
    return _trapz(y, x)


# ══════════════════════════════════════════════════════════
# 单水线计算：水线面要素
# ══════════════════════════════════════════════════════════

def calc_waterplane(stations: np.ndarray, half_breadths: np.ndarray,
                    method: str = 'trapz') -> dict:
    """
    计算单条水线的水线面要素。

    对应教材：盛振邦《船舶原理》上册 §2.2 水线面要素

    计算量：
        Aw   水线面面积（m²）
        xf   漂心纵向坐标（m）—— 水线面积形心
        IL   水线面对过漂心纵轴的惯性矩（m⁴）—— 用于计算 BM
        IT   水线面对过漂心横轴的惯性矩（m⁴）—— 用于计算 BML

    公式（梯形法）：
        Aw = 2 · ∫ y dx                          （水线面积）
        Sf = 2 · ∫ y·x dx                        （对中站的静矩）
        xf = Sf / Aw                              （漂心纵坐标）
        IL0 = (2/3) · ∫ y³ dx                    （对中纵轴的惯性矩）
        IL  = IL0 - Aw · xf²                     （平行轴定理修正到漂心）
        IT0 = 2 · ∫ y·x² dx                      （对中横轴的惯性矩）
        IT  = IT0 - Aw · xf²  （注：此处 IT 用于 BML，对纵轴）

    参数：
        stations     : np.ndarray  各站 x 坐标（m）
        half_breadths: np.ndarray  各站半宽 y（m）
        method       : str         积分方法

    返回：
        dict { Aw, xf, IL, IT }
    """
    x = stations
    y = half_breadths

    # ── 水线面积 Aw = 2·∫y dx ────────────────────────────
    Aw = 2.0 * integrate(y, x, method)

    if Aw < 1e-8:
        return {"Aw": 0.0, "xf": 0.0, "IL": 0.0, "IT": 0.0}

    # ── 漂心纵坐标 xf = (2·∫y·x dx) / Aw ────────────────
    Sf = 2.0 * integrate(y * x, x, method)
    xf = Sf / Aw

    # ── 横稳心半径用惯性矩 IL（对过漂心纵轴）────────────
    # 先算对中纵轴（x=0）的惯性矩：IL0 = (2/3)·∫y³ dx
    IL0 = (2.0 / 3.0) * integrate(y ** 3, x, method)
    # 平行轴定理：IL = IL0 - Aw·xf²
    IL = IL0 - Aw * xf ** 2

    # ── 纵稳心半径用惯性矩 IT（对过漂心横轴）────────────
    # IT0 = 2·∫y·x² dx（对中横轴的惯性矩）
    IT0 = 2.0 * integrate(y * x ** 2, x, method)
    # 平行轴定理：IT = IT0 - Aw·xf²
    # 注：严格应为 IT = IT0 - Aw·xf²，此处 xf 为漂心到中站的距离
    IT = IT0 - Aw * xf ** 2

    return {
        "Aw": float(Aw),
        "xf": float(xf),
        "IL": float(IL),
        "IT": float(IT),
    }


# ══════════════════════════════════════════════════════════
# 单吃水全参数计算
# ══════════════════════════════════════════════════════════

def calc_one_draft(draft: float,
                   stations: np.ndarray,
                   waterlines: np.ndarray,
                   offsets: np.ndarray,
                   Am: float,
                   method: str = 'trapz') -> dict:
    """
    计算指定吃水下的完整静水力参数（14项）。

    对应教材：盛振邦《船舶原理》上册 第2章

    计算流程：
    ① 找到 draft 对应的水线层（插值或精确匹配）
    ② 对该吃水处的水线，计算水线面要素（Aw, xf, IL, IT）
    ③ 沿 Z 方向积分（从基线到 draft），计算体积要素（∇, xb, zb）
    ④ 由体积要素计算排水量、稳心等

    参数：
        draft      : float        目标吃水（m）
        stations   : np.ndarray   各站 x 坐标（m）
        waterlines : np.ndarray   各水线 z 坐标（m），从小到大
        offsets    : np.ndarray   半宽矩阵 [wl_idx, sta_idx]
        Am         : float        中横剖面面积（m²），用于计算 Cm
        method     : str          积分方法

    返回：
        dict，包含全部 14 项静水力参数
    """

    # ── 第①步：对 draft 处的水线进行插值 ────────────────
    # 找到 draft 在 waterlines 中的位置（线性插值）
    if draft <= waterlines[0]:
        # 吃水小于等于基线，所有参数为零
        return _zero_result(draft)

    if draft > waterlines[-1]:
        draft = waterlines[-1]  # 限制在最大水线

    # 找到 draft 所在的水线区间 [wl_lo, wl_hi]
    idx_hi = np.searchsorted(waterlines, draft, side='right')
    idx_hi = min(idx_hi, len(waterlines) - 1)
    idx_lo = max(idx_hi - 1, 0)

    z_lo = waterlines[idx_lo]
    z_hi = waterlines[idx_hi]

    # 线性插值系数
    if abs(z_hi - z_lo) < 1e-10:
        alpha = 1.0
    else:
        alpha = (draft - z_lo) / (z_hi - z_lo)

    # 插值得到 draft 处的半宽
    y_lo = offsets[idx_lo]
    y_hi = offsets[idx_hi]
    y_draft = y_lo + alpha * (y_hi - y_lo)  # 线性插值

    # ── 第②步：计算 draft 处水线面要素 ──────────────────
    wp = calc_waterplane(stations, y_draft, method)
    Aw = wp["Aw"]
    xf = wp["xf"]
    IL = wp["IL"]   # 横稳心用惯性矩（m⁴）
    IT = wp["IT"]   # 纵稳心用惯性矩（m⁴）

    if Aw < 1e-8:
        return _zero_result(draft)

    # ── 第③步：沿 Z 方向积分（从基线到 draft）────────────
    # 构建从 z=0 到 z=draft 的水线切片
    # 取 waterlines 中 <= draft 的部分，再加上 draft 处的插值点

    # 找到所有 z <= draft 的水线
    mask = waterlines <= draft
    z_slice = waterlines[mask].copy()
    Aw_slice = np.zeros(len(z_slice))
    xf_slice = np.zeros(len(z_slice))

    for i, wl_idx in enumerate(np.where(mask)[0]):
        wp_i = calc_waterplane(stations, offsets[wl_idx], method)
        Aw_slice[i] = wp_i["Aw"]
        xf_slice[i] = wp_i["xf"]

    # 如果 draft 不在 waterlines 中，追加插值点
    if abs(draft - z_slice[-1]) > 1e-6:
        z_slice  = np.append(z_slice,  draft)
        Aw_slice = np.append(Aw_slice, Aw)
        xf_slice = np.append(xf_slice, xf)

    # ── 第④步：计算排水体积 ∇ ────────────────────────────
    # 公式：∇ = ∫₀ᵈ Aw(z) dz
    # 物理意义：各水线面积沿吃水方向的积分
    V = integrate(Aw_slice, z_slice, method)

    if V < 1e-6:
        return _zero_result(draft)

    # ── 第⑤步：计算浮心纵向坐标 xb ──────────────────────
    # 公式：xb = (1/∇) · ∫₀ᵈ Aw(z)·xf(z) dz
    # 物理意义：各水线面积形心的加权平均
    xb = integrate(Aw_slice * xf_slice, z_slice, method) / V

    # ── 第⑥步：计算浮心垂向坐标 zb（即 KB）─────────────
    # 公式：zb = (1/∇) · ∫₀ᵈ Aw(z)·z dz
    # 物理意义：各水线面积对基线的静矩 / 体积
    zb = integrate(Aw_slice * z_slice, z_slice, method) / V

    # ── 第⑦步：计算排水量 ────────────────────────────────
    # 海水排水量：Δ_sea = ρ_sea · ∇
    delta_sea   = RHO_SEA   * V
    # 淡水排水量：Δ_fresh = ρ_fresh · ∇
    delta_fresh = RHO_FRESH * V

    # ── 第⑧步：计算每厘米吃水吨数 TPC ───────────────────
    # 公式：TPC = Aw · ρ / 100
    # 物理意义：吃水增加 1cm 时排水量的增量（海水）
    TPC = Aw * RHO_SEA / 100.0

    # ── 第⑨步：计算横稳心半径 BM ────────────────────────
    # 公式：BM = IL / ∇
    # 物理意义：水线面对过漂心纵轴的惯性矩 / 排水体积
    # 对应教材：§2.4 初稳性
    BM = IL / V

    # ── 第⑩步：计算纵稳心半径 BML ───────────────────────
    # 公式：BML = IT / ∇
    # 物理意义：水线面对过漂心横轴的惯性矩 / 排水体积
    BML = IT / V

    # ── 第⑪步：计算横稳心高度 zM（即 KM）───────────────
    # 公式：KM = KB + BM = zb + BM
    zM = zb + BM

    # ── 第⑫步：计算每厘米纵倾力矩 MTC ───────────────────
    # 公式：MTC = Δ · GML / (100 · L)
    # 简化：MTC = ρ · ∇ · BML / (100 · L)
    # 物理意义：使船纵倾 1cm 所需的力矩（t·m/cm）
    L = float(stations[-1] - stations[0])
    if L > 1e-6:
        MTC = RHO_SEA * V * BML / (100.0 * L)
    else:
        MTC = 0.0

    # ── 第⑬步：计算船型系数 ─────────────────────────────
    B = 2.0 * float(np.max(y_draft))   # 型宽（设计吃水处最大全宽）

    # 方形系数 Cb = ∇ / (L·B·d)
    # 物理意义：船体体积与外包长方体之比
    if L > 1e-6 and B > 1e-6 and draft > 1e-6:
        Cb = V / (L * B * draft)
    else:
        Cb = 0.0

    # 棱形系数 Cp = ∇ / (Am·L)
    # 物理意义：船体体积与以中横剖面为底的棱柱体之比
    # Am 为中横剖面面积（由调用方传入）
    if Am > 1e-6 and L > 1e-6:
        Cp = V / (Am * L)
    else:
        Cp = 0.0

    # 水线面系数 Cwp = Aw / (L·B)
    # 物理意义：水线面积与外包矩形之比
    if L > 1e-6 and B > 1e-6:
        Cwp = Aw / (L * B)
    else:
        Cwp = 0.0

    # 中横剖面系数 Cm = Am / (B·d)
    # 物理意义：中横剖面面积与外包矩形之比
    if B > 1e-6 and draft > 1e-6:
        Cm = Am / (B * draft)
    else:
        Cm = 0.0

    # ── 返回结果 ─────────────────────────────────────────
    def r(v, n=4):
        """四舍五入到 n 位小数"""
        return round(float(v), n)

    return {
        "d":           r(draft, 3),    # 吃水（m）
        "V":           r(V, 3),        # 排水体积（m³）
        "delta_sea":   r(delta_sea, 3),   # 海水排水量（t）
        "delta_fresh": r(delta_fresh, 3), # 淡水排水量（t）
        "xb":          r(xb, 4),       # 浮心纵向坐标（m）
        "zb":          r(zb, 4),       # 浮心垂向坐标 KB（m）
        "Aw":          r(Aw, 3),       # 水线面面积（m²）
        "xf":          r(xf, 4),       # 漂心纵向坐标（m）
        "TPC":         r(TPC, 4),      # 每厘米吃水吨数（t/cm）
        "MTC":         r(MTC, 4),      # 每厘米纵倾力矩（t·m/cm）
        "BM":          r(BM, 4),       # 横稳心半径（m）
        "BML":         r(BML, 4),      # 纵稳心半径（m）
        "zM":          r(zM, 4),       # 横稳心高度 KM（m）
        "Cb":          r(Cb, 4),       # 方形系数
        "Cp":          r(Cp, 4),       # 棱形系数
        "Cwp":         r(Cwp, 4),      # 水线面系数
        "Cm":          r(Cm, 4),       # 中横剖面系数
        "B":           r(B, 3),        # 型宽（m）
        "L":           r(L, 3),        # 船长（m）
    }


def _zero_result(draft: float) -> dict:
    """返回全零结果（吃水为零或极小时）"""
    return {
        "d": round(draft, 3),
        "V": 0.0, "delta_sea": 0.0, "delta_fresh": 0.0,
        "xb": 0.0, "zb": 0.0,
        "Aw": 0.0, "xf": 0.0,
        "TPC": 0.0, "MTC": 0.0,
        "BM": 0.0, "BML": 0.0, "zM": 0.0,
        "Cb": 0.0, "Cp": 0.0, "Cwp": 0.0, "Cm": 0.0,
        "B": 0.0, "L": 0.0,
    }


# ══════════════════════════════════════════════════════════
# 中横剖面面积计算
# ══════════════════════════════════════════════════════════

def calc_midship_area(stations: np.ndarray,
                      waterlines: np.ndarray,
                      offsets: np.ndarray,
                      draft: float,
                      method: str = 'trapz') -> float:
    """
    计算中横剖面（x=0 处）在设计吃水下的面积 Am。

    公式：Am = 2 · ∫₀ᵈ y_mid(z) dz
    其中 y_mid(z) 为中站（x=0）处各水线的半宽。

    参数：
        stations   : np.ndarray  各站 x 坐标
        waterlines : np.ndarray  各水线 z 坐标
        offsets    : np.ndarray  半宽矩阵 [wl_idx, sta_idx]
        draft      : float       设计吃水（m）
        method     : str         积分方法

    返回：
        float  中横剖面面积（m²）
    """
    # 找到中站（x=0）的索引
    mid_idx = int(np.argmin(np.abs(stations)))

    # 取从基线到 draft 的水线切片
    mask = waterlines <= draft
    z_slice = waterlines[mask]
    y_mid   = offsets[mask, mid_idx]  # 中站各水线半宽

    if len(z_slice) < 2:
        return 0.0

    # Am = 2 · ∫ y_mid(z) dz
    Am = 2.0 * integrate(y_mid, z_slice, method)
    return float(Am)


# ══════════════════════════════════════════════════════════
# 多吃水遍历计算（主函数）
# ══════════════════════════════════════════════════════════

def calc_hydrostatics_table(offsets_data: dict,
                             d_min: float = None,
                             d_max: float = None,
                             d_step: float = 0.5,
                             method: str = 'trapz') -> dict:
    """
    多吃水遍历计算，生成完整静水力计算表。

    对每个吃水值计算全部 14 项静水力参数，
    结果以列表形式返回，可直接用于绘图和导出。

    参数：
        offsets_data : dict   parse_offsets_excel() 或手动输入解析后的数据
                              必须包含：stations, waterlines, offsets
        d_min        : float  最小吃水（m），默认为最小水线高度
        d_max        : float  最大吃水（m），默认为最大水线高度
        d_step       : float  吃水步长（m），默认 0.5m
        method       : str    积分方法（'trapz' 或 'simpson'）

    返回：
        dict，包含：
            rows    : list[dict]  每个吃水的计算结果（14项参数）
            method  : str         使用的积分方法
            n_drafts: int         计算的吃水数量
            d_min   : float       实际最小吃水
            d_max   : float       实际最大吃水
            d_step  : float       吃水步长
            L       : float       船长
            columns : list[dict]  列定义（用于前端表格渲染）
    """
    # ── 解析输入数据 ──────────────────────────────────────
    stations   = np.array(offsets_data["stations"],   dtype=float)
    waterlines = np.array(offsets_data["waterlines"], dtype=float)
    offsets    = np.array(offsets_data["offsets"],    dtype=float)
    # offsets 形状：[n_waterlines, n_stations]

    # ── 确定吃水范围 ──────────────────────────────────────
    wl_min = float(waterlines[0])
    wl_max = float(waterlines[-1])

    if d_min is None:
        d_min = wl_min
    if d_max is None:
        d_max = wl_max

    # 限制在水线范围内
    d_min = max(d_min, wl_min)
    d_max = min(d_max, wl_max)

    if d_min >= d_max:
        raise ValueError(f"吃水范围无效：d_min={d_min} >= d_max={d_max}")

    # ── 生成吃水序列 ──────────────────────────────────────
    # 从 d_min 到 d_max，步长 d_step
    drafts = []
    d = d_min
    while d <= d_max + 1e-9:
        drafts.append(round(d, 4))
        d += d_step
    # 确保包含 d_max
    if abs(drafts[-1] - d_max) > 1e-6:
        drafts.append(round(d_max, 4))

    # ── 计算中横剖面面积 Am（只算一次，用设计吃水 d_max）──
    Am = calc_midship_area(stations, waterlines, offsets, d_max, method)

    # ── 逐吃水计算 ────────────────────────────────────────
    rows = []
    for d in drafts:
        row = calc_one_draft(d, stations, waterlines, offsets, Am, method)
        rows.append(row)

    # ── 列定义（用于前端表格渲染和 Excel 导出）────────────
    columns = [
        {"key": "d",           "label": "吃水 d",     "unit": "m",      "fmt": ".3f"},
        {"key": "V",           "label": "排水体积 ∇",  "unit": "m³",     "fmt": ".3f"},
        {"key": "delta_sea",   "label": "排水量 Δ(海)", "unit": "t",     "fmt": ".3f"},
        {"key": "delta_fresh", "label": "排水量 Δ(淡)", "unit": "t",     "fmt": ".3f"},
        {"key": "xb",          "label": "浮心纵坐标 xb","unit": "m",     "fmt": ".4f"},
        {"key": "zb",          "label": "浮心高度 KB",  "unit": "m",     "fmt": ".4f"},
        {"key": "Aw",          "label": "水线面积 Aw",  "unit": "m²",    "fmt": ".3f"},
        {"key": "xf",          "label": "漂心纵坐标 xf","unit": "m",     "fmt": ".4f"},
        {"key": "TPC",         "label": "TPC",          "unit": "t/cm",  "fmt": ".4f"},
        {"key": "MTC",         "label": "MTC",          "unit": "t·m/cm","fmt": ".4f"},
        {"key": "BM",          "label": "横稳心半径 BM","unit": "m",     "fmt": ".4f"},
        {"key": "BML",         "label": "纵稳心半径 BML","unit": "m",    "fmt": ".4f"},
        {"key": "zM",          "label": "稳心高度 KM",  "unit": "m",     "fmt": ".4f"},
        {"key": "Cb",          "label": "方形系数 Cb",  "unit": "—",     "fmt": ".4f"},
        {"key": "Cp",          "label": "棱形系数 Cp",  "unit": "—",     "fmt": ".4f"},
        {"key": "Cwp",         "label": "水线面系数 Cwp","unit": "—",    "fmt": ".4f"},
        {"key": "Cm",          "label": "中剖面系数 Cm","unit": "—",     "fmt": ".4f"},
    ]

    return {
        "rows":     rows,
        "method":   method,
        "n_drafts": len(rows),
        "d_min":    d_min,
        "d_max":    d_max,
        "d_step":   d_step,
        "L":        float(stations[-1] - stations[0]),
        "Am":       round(Am, 4),
        "columns":  columns,
    }
