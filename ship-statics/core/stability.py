# ============================================================
# core/stability.py — 稳性横截曲线（GZ 曲线）计算模块
#
# 阶段5 新增模块
#
# 严格遵循：
#   盛振邦《船舶原理》上册 第3章 船舶稳性
#   《国内航行海船法定检验技术规则》稳性衡准
#   武汉理工大学船海学院《船舶静力学》课程设计规范
#
# ── 计算方法：变排水量法（Constant Displacement Method）──
#
# 基本思路：
#   船舶横倾角为 θ 时，保持排水量 Δ 不变，
#   求倾斜后浮心位置 B'，进而计算复原力臂 GZ。
#
# 坐标系（与前各阶段一致）：
#   X 轴：沿船长，船首正，船尾负，中站 x=0
#   Z 轴：垂直向上，基线 Z=0
#   Y 轴：横向，右舷正
#   θ   ：横倾角，右倾为正（度）
#
# ── 核心公式（对应教材 §3.2）────────────────────────────
#
# 1. 倾斜水线面积 Aw(θ)：
#    船倾斜 θ 后，新水线与船体的交线面积
#    用梯形法对倾斜后各站半宽积分
#
# 2. 倾斜后浮心坐标 B'：
#    B'y = (1/∇) · ∫∫ y·dA·dz  （横向浮心）
#    B'z = (1/∇) · ∫∫ z·dA·dz  （垂向浮心）
#
# 3. 复原力臂 GZ（Righting Lever Arm）：
#    GZ = B'y·cosθ + B'z·sinθ - KG·sinθ
#       = (B'y - KG·sinθ)·cosθ + (B'z - KG·cosθ)·sinθ
#    简化为：
#    GZ = (KB + BM·cosθ - KG)·sinθ  （小角度近似，仅供参考）
#
#    精确公式（变排水量法）：
#    GZ = l_s - KG·sinθ
#    其中 l_s 为稳性横截臂（Righting Arm），由倾斜体积积分得到
#
# 4. 初稳性高度 GM：
#    GM = KB + BM - KG = KM - KG
#    GZ ≈ GM·sinθ  （小角度，θ < 10°）
#
# ── 变排水量法实现说明 ────────────────────────────────────
#
# 对于每个横倾角 θ：
#   ① 将型值表坐标系旋转 θ 角
#   ② 在旋转坐标系中，找到使排水量等于目标值的水线高度
#   ③ 计算旋转后浮心坐标 (y_B', z_B')
#   ④ 将浮心坐标转回原坐标系
#   ⑤ GZ = y_B'·cosθ + z_B'·sinθ - KG·sinθ
#
# ── 简化实现（适用于课程设计）────────────────────────────
#
# 课程设计中常用的简化变排水量法：
#   利用已计算的静水力参数（KB、BM、BML）和
#   Bonjean 曲线数据，通过以下步骤计算 GZ：
#
#   对每个横倾角 θ：
#   ① 计算楔形体体积 v（从直立水线到倾斜水线的楔形）
#   ② 楔形体重心坐标 (y_v, z_v)
#   ③ 倾斜后浮心：
#      y_B' = (∇·y_B + v·y_v) / ∇  （y_B 直立时浮心横坐标=0）
#      z_B' = (∇·z_B + v·z_v) / ∇
#   ④ GZ = y_B'·cosθ + z_B'·sinθ - KG·sinθ
#
# ============================================================

import numpy as np
import math

# ── 密度常数 ──────────────────────────────────────────────
RHO_SEA   = 1.025
RHO_FRESH = 1.000


def _trapz(y: np.ndarray, x: np.ndarray) -> float:
    """梯形法积分（兼容 numpy 1.x / 2.x）"""
    fn = getattr(np, 'trapezoid', None) or getattr(np, 'trapz')
    return float(fn(y, x))


# ══════════════════════════════════════════════════════════
# 倾斜水线计算
# ══════════════════════════════════════════════════════════

def _get_inclined_waterplane(stations: np.ndarray,
                              waterlines: np.ndarray,
                              offsets: np.ndarray,
                              theta_deg: float,
                              draft_center: float) -> dict:
    """
    计算船舶横倾 theta_deg 度时，各站处的倾斜水线高度和半宽。

    原理：
        船绕龙骨中心线（y=0, z=draft_center）旋转 θ 角后，
        原水线面（z=draft_center）变为倾斜平面。
        对于站 x 处，倾斜水线在原坐标系中的高度为：
            z_wl(x) = draft_center  （绕中心旋转，水线高度不变）
        但各站的有效吃水（从基线到倾斜水线的垂直距离）变化：
            对于右倾 θ，右舷（y>0）水线升高，左舷降低
            有效吃水 d_eff(y) = draft_center + y·tan(θ)

    简化处理（课程设计适用）：
        假设船体左右对称，对每个站，
        倾斜后该站的有效吃水为：
            d_eff = draft_center + y_half·sin(θ)  （近似）
        其中 y_half 为该站在 draft_center 处的半宽

    参数：
        stations      : np.ndarray  各站 x 坐标
        waterlines    : np.ndarray  各水线 z 坐标
        offsets       : np.ndarray  半宽矩阵 [wl_idx, sta_idx]
        theta_deg     : float       横倾角（度），右倾为正
        draft_center  : float       旋转中心吃水（m），通常取设计吃水

    返回：
        dict { 'y_half': array, 'z_wl': array }
        y_half: 各站在倾斜水线处的半宽（m）
        z_wl:   各站倾斜水线高度（m）
    """
    theta = math.radians(theta_deg)
    n_sta = len(stations)

    # 在 draft_center 处插值各站半宽
    y_at_draft = np.zeros(n_sta)
    for i in range(n_sta):
        y_at_draft[i] = float(np.interp(draft_center, waterlines, offsets[:, i]))

    # 倾斜水线高度（各站处）
    # 右倾 θ 时，站 x 处的水线高度：
    # z_wl(x) = draft_center + y_at_draft(x) * sin(θ)
    # 注：这是近似公式，精确计算需要迭代
    z_wl = draft_center + y_at_draft * math.sin(theta)

    # 在倾斜水线处插值各站半宽
    y_inclined = np.zeros(n_sta)
    for i in range(n_sta):
        z_i = float(np.clip(z_wl[i], waterlines[0], waterlines[-1]))
        y_inclined[i] = float(np.interp(z_i, waterlines, offsets[:, i]))

    return {
        'y_half': y_inclined,
        'z_wl':   z_wl,
        'y_at_draft': y_at_draft,
    }


# ══════════════════════════════════════════════════════════
# 楔形体积分（变排水量法核心）
# ══════════════════════════════════════════════════════════

def _calc_wedge(stations: np.ndarray,
                waterlines: np.ndarray,
                offsets: np.ndarray,
                theta_deg: float,
                draft: float) -> dict:
    """
    计算横倾 θ 时，从直立水线到倾斜水线之间的楔形体参数。

    楔形体定义：
        右倾 θ 时，右舷水线升高，左舷水线降低。
        楔形体 = 右舷新增浸水体积 - 左舷露出水面体积
        （对于对称船体，两者相等，楔形体净体积=0，但重心位置不同）

    楔形体重心坐标（对应教材 §3.2 楔形体法）：
        对于右倾 θ，楔形体的横向重心：
        y_v = (2/3) · y_half · cos(θ)  （近似）
        z_v = draft + (1/3) · y_half · sin(θ)  （近似）

    精确公式（梯形法）：
        楔形体体积微元：dv = y² · tan(θ) · dx  （对称船体）
        楔形体体积：v = ∫ y² · tan(θ) dx = tan(θ) · ∫ y² dx
        楔形体横向静矩：Sv_y = (2/3) · tan(θ) · ∫ y³ dx
        楔形体重心横坐标：y_v = Sv_y / v = (2/3) · (∫y³dx) / (∫y²dx)

    参数：
        stations   : np.ndarray  各站 x 坐标
        waterlines : np.ndarray  各水线 z 坐标
        offsets    : np.ndarray  半宽矩阵 [wl_idx, sta_idx]
        theta_deg  : float       横倾角（度）
        draft      : float       直立吃水（m）

    返回：
        dict { 'v': 楔形体体积, 'y_v': 横向重心, 'z_v': 垂向重心 }
    """
    theta = math.radians(theta_deg)
    tan_t = math.tan(theta)

    # 在 draft 处插值各站半宽
    n_sta = len(stations)
    y_at_draft = np.zeros(n_sta)
    for i in range(n_sta):
        y_at_draft[i] = float(np.interp(draft, waterlines, offsets[:, i]))

    # ── 楔形体体积 v = tan(θ) · ∫ y² dx ─────────────────
    # 对应教材公式：v = tan(θ) · IL_wl（水线面对纵轴的惯性矩）
    # IL_wl = (1/3) · ∫ (2y)³/4 dx = (2/3) · ∫ y³ dx  （双侧）
    # 注：此处 y 为半宽，双侧楔形体体积：
    # v = 2 · ∫ (1/2) · y² · tan(θ) dx = tan(θ) · ∫ y² dx
    int_y2 = _trapz(y_at_draft ** 2, stations)
    int_y3 = _trapz(y_at_draft ** 3, stations)

    v = tan_t * int_y2   # 楔形体体积（m³）

    # ── 楔形体横向重心 y_v ────────────────────────────────
    # 公式：y_v = (2/3) · (∫y³dx) / (∫y²dx)
    # 对应教材：楔形体重心在水线面以上 (2/3)·y 处
    if abs(int_y2) > 1e-10:
        y_v = (2.0 / 3.0) * int_y3 / int_y2
    else:
        y_v = 0.0

    # ── 楔形体垂向重心 z_v ────────────────────────────────
    # 楔形体重心高度 ≈ 直立水线高度（近似）
    z_v = draft

    return {
        'v':   float(v),
        'y_v': float(y_v),
        'z_v': float(z_v),
    }


# ══════════════════════════════════════════════════════════
# 单横倾角 GZ 计算
# ══════════════════════════════════════════════════════════

def calc_gz_one_angle(theta_deg: float,
                      draft: float,
                      KB: float,
                      BM: float,
                      V: float,
                      stations: np.ndarray,
                      waterlines: np.ndarray,
                      offsets: np.ndarray,
                      KG: float) -> dict:
    """
    计算指定横倾角 θ 下的复原力臂 GZ。

    采用变排水量法（楔形体法），对应教材 §3.2。

    计算步骤：
    ① 计算楔形体参数（体积 v，重心 y_v, z_v）
    ② 计算倾斜后浮心坐标：
       y_B' = v · y_v / V   （楔形体横向静矩 / 总体积）
       z_B' = KB + v · (z_v - KB) / V  （近似）
    ③ 计算稳性横截臂 l_s：
       l_s = y_B' · cos(θ) + z_B' · sin(θ)
    ④ 复原力臂 GZ：
       GZ = l_s - KG · sin(θ)

    参数：
        theta_deg  : float  横倾角（度）
        draft      : float  直立吃水（m）
        KB         : float  浮心高度（m）
        BM         : float  横稳心半径（m）
        V          : float  排水体积（m³）
        stations   : np.ndarray  各站 x 坐标
        waterlines : np.ndarray  各水线 z 坐标
        offsets    : np.ndarray  半宽矩阵
        KG         : float  重心高度（m）

    返回：
        dict { 'theta': θ, 'GZ': GZ值, 'ls': 稳性横截臂, 'y_B': 浮心横坐标, 'z_B': 浮心垂坐标 }
    """
    theta = math.radians(theta_deg)
    cos_t = math.cos(theta)
    sin_t = math.sin(theta)

    if abs(theta_deg) < 1e-6:
        # θ=0，GZ=0
        return {
            'theta': 0.0, 'GZ': 0.0, 'ls': 0.0,
            'y_B': 0.0, 'z_B': KB,
        }

    # ── 第①步：楔形体参数 ────────────────────────────────
    wedge = _calc_wedge(stations, waterlines, offsets, theta_deg, draft)
    v   = wedge['v']
    y_v = wedge['y_v']
    z_v = wedge['z_v']

    # ── 第②步：倾斜后浮心坐标 ────────────────────────────
    # 直立时浮心：y_B=0（对称船体），z_B=KB
    # 倾斜后浮心（楔形体法）：
    #   y_B' = (V·0 + v·y_v) / V = v·y_v / V
    #   z_B' = (V·KB + v·z_v) / V
    # 注：此处假设楔形体体积相对总体积较小（小角度近似）
    # 对于大角度，需要迭代求解使排水量守恒的水线
    if V > 1e-6:
        y_B_prime = v * y_v / V
        z_B_prime = (V * KB + v * (z_v - KB)) / V
    else:
        y_B_prime = 0.0
        z_B_prime = KB

    # ── 第③步：稳性横截臂 l_s ────────────────────────────
    # 公式：l_s = y_B'·cos(θ) + z_B'·sin(θ)
    # 物理意义：倾斜后浮力作用线到重心垂线的距离
    ls = y_B_prime * cos_t + z_B_prime * sin_t

    # ── 第④步：复原力臂 GZ ───────────────────────────────
    # 公式：GZ = l_s - KG·sin(θ)
    # 对应教材 §3.2 式(3-8)
    GZ = ls - KG * sin_t

    return {
        'theta': float(theta_deg),
        'GZ':    round(float(GZ), 4),
        'ls':    round(float(ls), 4),
        'y_B':   round(float(y_B_prime), 4),
        'z_B':   round(float(z_B_prime), 4),
    }


# ══════════════════════════════════════════════════════════
# 完整 GZ 曲线计算（多角度）
# ══════════════════════════════════════════════════════════

def calc_gz_curve(offsets_data: dict,
                  draft: float,
                  KG: float,
                  theta_min: float = 0.0,
                  theta_max: float = 90.0,
                  theta_step: float = 5.0,
                  rho: float = RHO_SEA) -> dict:
    """
    计算完整的 GZ 曲线（稳性横截曲线）。

    对应教材：盛振邦《船舶原理》上册 §3.2 大倾角稳性

    参数：
        offsets_data : dict   型值表数据（stations, waterlines, offsets）
        draft        : float  设计吃水（m）
        KG           : float  重心高度（m）
        theta_min    : float  最小横倾角（度），默认 0°
        theta_max    : float  最大横倾角（度），默认 90°
        theta_step   : float  横倾角步长（度），默认 5°
        rho          : float  水密度（t/m³），默认海水 1.025

    返回：
        dict，包含：
            rows       : list[dict]  每个角度的计算结果
            GM         : float       初稳性高度（m）
            KM         : float       稳心高度（m）
            KB         : float       浮心高度（m）
            BM         : float       横稳心半径（m）
            KG         : float       重心高度（m）
            delta      : float       排水量（t）
            V          : float       排水体积（m³）
            draft      : float       吃水（m）
            theta_max_gz: float      最大 GZ 对应的角度（度）
            GZ_max     : float       最大复原力臂（m）
            theta_vanish: float      稳性消失角（GZ=0 的角度，度）
            criteria   : dict        稳性衡准校核结果
    """
    stations   = np.array(offsets_data["stations"],   dtype=float)
    waterlines = np.array(offsets_data["waterlines"], dtype=float)
    offsets    = np.array(offsets_data["offsets"],    dtype=float)

    # ── 第①步：计算直立状态静水力参数 ───────────────────
    from core.hydrostatics_full import calc_one_draft, calc_midship_area
    Am = calc_midship_area(stations, waterlines, offsets, draft)
    hydro = calc_one_draft(draft, stations, waterlines, offsets, Am)

    KB = hydro['zb']
    BM = hydro['BM']
    V  = hydro['V']
    KM = hydro['zM']
    GM = KM - KG   # 初稳性高度

    delta = rho * V   # 排水量（t）

    # ── 第②步：生成横倾角序列 ────────────────────────────
    thetas = []
    t = theta_min
    while t <= theta_max + 1e-9:
        thetas.append(round(t, 2))
        t += theta_step
    if abs(thetas[-1] - theta_max) > 1e-6:
        thetas.append(round(theta_max, 2))

    # ── 第③步：逐角度计算 GZ ─────────────────────────────
    rows = []
    for theta in thetas:
        result = calc_gz_one_angle(
            theta_deg=theta,
            draft=draft,
            KB=KB,
            BM=BM,
            V=V,
            stations=stations,
            waterlines=waterlines,
            offsets=offsets,
            KG=KG,
        )
        # 附加初稳性近似值（用于对比）
        result['GZ_approx'] = round(GM * math.sin(math.radians(theta)), 4)
        rows.append(result)

    # ── 第④步：统计特征值 ────────────────────────────────
    gz_vals = [r['GZ'] for r in rows]
    thetas_arr = np.array(thetas)
    gz_arr     = np.array(gz_vals)

    # 最大 GZ 及对应角度
    idx_max = int(np.argmax(gz_arr))
    GZ_max       = float(gz_arr[idx_max])
    theta_max_gz = float(thetas_arr[idx_max])

    # 稳性消失角（GZ 从正变负的角度）
    theta_vanish = None
    for i in range(len(gz_vals) - 1):
        if gz_vals[i] >= 0 and gz_vals[i+1] < 0:
            # 线性插值
            t1, t2 = thetas[i], thetas[i+1]
            g1, g2 = gz_vals[i], gz_vals[i+1]
            if abs(g2 - g1) > 1e-10:
                theta_vanish = round(t1 + (t2 - t1) * (-g1) / (g2 - g1), 2)
            break

    # ── 第⑤步：稳性衡准校核（《国内航行海船法定检验技术规则》）──
    criteria = _check_stability_criteria(rows, GM, GZ_max, theta_max_gz, theta_vanish)

    return {
        "rows":          rows,
        "GM":            round(float(GM), 4),
        "KM":            round(float(KM), 4),
        "KB":            round(float(KB), 4),
        "BM":            round(float(BM), 4),
        "KG":            round(float(KG), 4),
        "delta":         round(float(delta), 3),
        "V":             round(float(V), 3),
        "draft":         round(float(draft), 3),
        "theta_max_gz":  theta_max_gz,
        "GZ_max":        round(float(GZ_max), 4),
        "theta_vanish":  theta_vanish,
        "criteria":      criteria,
        "theta_step":    theta_step,
        "rho":           rho,
    }


# ══════════════════════════════════════════════════════════
# 稳性衡准校核
# ══════════════════════════════════════════════════════════

def _check_stability_criteria(rows: list,
                               GM: float,
                               GZ_max: float,
                               theta_max_gz: float,
                               theta_vanish) -> dict:
    """
    按《国内航行海船法定检验技术规则》校核稳性衡准。

    主要衡准（对应规则 §3.2）：
    1. 初稳性高度 GM ≥ 0.15 m
    2. 最大复原力臂 GZ_max ≥ 0.20 m
    3. 最大 GZ 对应角度 θ_max ≥ 25°
    4. 稳性消失角 θ_v ≥ 60°（一般货船）
    5. 30°~40° 范围内的 GZ 面积（稳性面积）≥ 0.055 m·rad
    6. 0°~30° 范围内的 GZ 面积 ≥ 0.055 m·rad
    7. 0°~40° 范围内的 GZ 面积 ≥ 0.090 m·rad

    参数：
        rows         : list[dict]  GZ 曲线数据
        GM           : float       初稳性高度（m）
        GZ_max       : float       最大复原力臂（m）
        theta_max_gz : float       最大 GZ 对应角度（度）
        theta_vanish : float       稳性消失角（度）

    返回：
        dict，每项衡准的校核结果
    """
    # 提取 GZ 数据
    thetas = np.array([r['theta'] for r in rows])
    gz_arr = np.array([r['GZ']    for r in rows])

    def gz_area(t1, t2):
        """计算 [t1, t2] 范围内的 GZ 面积（m·rad）"""
        mask = (thetas >= t1) & (thetas <= t2)
        if mask.sum() < 2:
            return 0.0
        t_slice = np.radians(thetas[mask])
        g_slice = gz_arr[mask]
        fn = getattr(np, 'trapezoid', None) or getattr(np, 'trapz')
        return float(fn(g_slice, t_slice))

    area_0_30  = gz_area(0, 30)
    area_0_40  = gz_area(0, 40)
    area_30_40 = gz_area(30, 40)

    def check(val, limit, name, unit=""):
        ok = val >= limit if val is not None else False
        return {
            "name":  name,
            "value": round(float(val), 4) if val is not None else None,
            "limit": limit,
            "unit":  unit,
            "pass":  ok,
            "mark":  "✓" if ok else "✗",
        }

    criteria = {
        "GM":          check(GM,           0.15,  "初稳性高度 GM",          "m"),
        "GZ_max":      check(GZ_max,       0.20,  "最大复原力臂 GZ_max",    "m"),
        "theta_max":   check(theta_max_gz, 25.0,  "最大GZ对应角度 θ_max",   "°"),
        "theta_vanish":check(theta_vanish, 60.0,  "稳性消失角 θ_v",         "°"),
        "area_0_30":   check(area_0_30,    0.055, "GZ面积(0°~30°)",         "m·rad"),
        "area_0_40":   check(area_0_40,    0.090, "GZ面积(0°~40°)",         "m·rad"),
        "area_30_40":  check(area_30_40,   0.030, "GZ面积(30°~40°)",        "m·rad"),
    }

    # 总体判断
    all_pass = all(v["pass"] for v in criteria.values())
    criteria["overall"] = {"pass": all_pass, "mark": "✓ 满足稳性衡准" if all_pass else "✗ 不满足稳性衡准"}

    return criteria


# ══════════════════════════════════════════════════════════
# 多排水量 GZ 曲线族计算
# ══════════════════════════════════════════════════════════

def calc_gz_family(offsets_data: dict,
                   drafts: list,
                   KG: float,
                   theta_step: float = 5.0,
                   rho: float = RHO_SEA) -> list:
    """
    计算多个吃水（排水量）下的 GZ 曲线族。

    参数：
        offsets_data : dict   型值表数据
        drafts       : list   吃水列表（m）
        KG           : float  重心高度（m，假设各工况相同）
        theta_step   : float  横倾角步长（度）
        rho          : float  水密度

    返回：
        list[dict]  每个吃水的 GZ 曲线计算结果
    """
    results = []
    for d in drafts:
        try:
            result = calc_gz_curve(
                offsets_data=offsets_data,
                draft=d,
                KG=KG,
                theta_step=theta_step,
                rho=rho,
            )
            results.append(result)
        except Exception as e:
            results.append({"draft": d, "error": str(e)})
    return results
