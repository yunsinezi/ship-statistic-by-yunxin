# ============================================================
# core/excel_parser.py — 型值表 Excel 导入与数据校验模块
#
# 阶段2 新增模块，不影响原有 calculator.py / exporter.py
#
# 支持的 Excel 格式（标准型值表）：
#
#   第1行：标题行（忽略）
#   第2行：列标题行，格式如下：
#     A列：站号（如 0, 1, 2...）或直接是 x 坐标
#     B列：站号 X 坐标 (m)  ← 如果 A 列是站号编号
#     C列起：各水线半宽，列标题为水线高度，如 "WL0.5" 或 "0.5"
#
#   简化格式（推荐，与模板一致）：
#     第1行：标题
#     第2行：列标题 → [站号, X坐标(m), WL0.0, WL0.5, WL1.0, ...]
#     第3行起：数据行
#
# 坐标系（严格遵守）：
#   X 轴：船首为正，船尾为负，中站 x=0
#   y   ：半宽（单侧）
# ============================================================

import openpyxl
import re


# ── 校验规则常量 ──────────────────────────────────────────
ENDPOINT_TOLERANCE = 0.01   # 端点半宽允许的最大值（超过则警告）
MIN_STATIONS       = 3      # 最少站数
MIN_WATERLINES     = 2      # 最少水线数


def parse_offsets_excel(filepath: str) -> dict:
    """
    解析标准格式的型值表 Excel 文件。

    参数：
        filepath : str  Excel 文件路径（.xlsx）

    返回：
        dict，包含：
            stations      : list[float]  各站 x 坐标（米）
            waterlines    : list[float]  各水线高度 z（米）
            offsets       : list[list]   半宽矩阵 offsets[wl_idx][sta_idx]
            n_stations    : int          站数
            n_waterlines  : int          水线数
            L             : float        船长（m）
            station_spacing: float       平均站距（m）
            validation    : dict         校验报告（见 validate_offsets）

    异常：
        ValueError : 格式严重错误时抛出
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 3:
        raise ValueError("Excel 行数不足，至少需要：标题行 + 列标题行 + 1行数据")

    # ── 第1步：解析列标题行（第2行，index=1）────────────────
    header_row = rows[1]

    # 找到 X 坐标列（第2列，index=1）和水线列（第3列起）
    # 列标题示例：["站号", "X坐标(m)", "WL0.0", "WL0.5", "WL1.0", ...]
    waterline_cols = []   # [(列索引, 水线高度z), ...]

    for col_idx, cell_val in enumerate(header_row):
        if col_idx < 2:
            continue  # 跳过站号列和X坐标列
        if cell_val is None:
            continue

        # 尝试从列标题中提取水线高度
        # 支持格式：WL0.5、WL 0.5、0.5、z=0.5 等
        z = _extract_waterline_height(str(cell_val))
        if z is not None:
            waterline_cols.append((col_idx, z))

    if len(waterline_cols) < MIN_WATERLINES:
        raise ValueError(
            f"未能识别到足够的水线列（识别到 {len(waterline_cols)} 列，"
            f"至少需要 {MIN_WATERLINES} 列）。\n"
            f"请确保列标题格式为：WL0.0, WL0.5, WL1.0 ... 或直接写数字"
        )

    # ── 第2步：解析数据行（第3行起）────────────────────────
    stations   = []   # 各站 x 坐标
    offsets_by_station = []  # offsets_by_station[sta_idx][wl_idx] = 半宽

    for row_idx, row in enumerate(rows[2:], start=3):
        # 跳过空行
        if row[0] is None and row[1] is None:
            continue

        # 读取 X 坐标（第2列，index=1）
        x_val = row[1] if len(row) > 1 else None
        if x_val is None:
            # 尝试第1列
            x_val = row[0]
        if x_val is None:
            continue

        try:
            x = float(x_val)
        except (ValueError, TypeError):
            continue  # 跳过非数字行（可能是说明行）

        # 读取各水线半宽
        half_breadths = []
        for col_idx, z in waterline_cols:
            if col_idx < len(row) and row[col_idx] is not None:
                try:
                    y = float(row[col_idx])
                except (ValueError, TypeError):
                    y = 0.0
            else:
                y = 0.0
            half_breadths.append(y)

        stations.append(x)
        offsets_by_station.append(half_breadths)

    if len(stations) < MIN_STATIONS:
        raise ValueError(
            f"有效站数不足（读到 {len(stations)} 站，至少需要 {MIN_STATIONS} 站）"
        )

    # ── 第3步：转置矩阵 ──────────────────────────────────
    # offsets_by_station[sta][wl] → offsets[wl][sta]
    n_wl  = len(waterline_cols)
    n_sta = len(stations)
    waterlines = [z for _, z in waterline_cols]

    offsets = []
    for wl_idx in range(n_wl):
        row_data = [offsets_by_station[sta_idx][wl_idx] for sta_idx in range(n_sta)]
        offsets.append(row_data)

    # ── 第4步：计算船长和站距 ────────────────────────────
    L = max(stations) - min(stations)
    import numpy as np
    diffs = [stations[i+1] - stations[i] for i in range(len(stations)-1)]
    station_spacing = sum(diffs) / len(diffs) if diffs else 0.0

    # ── 第5步：数据校验 ──────────────────────────────────
    validation = validate_offsets(stations, waterlines, offsets)

    # 应用自动修正
    offsets = validation["corrected_offsets"]

    return {
        "stations":        stations,
        "waterlines":      waterlines,
        "offsets":         offsets,
        "n_stations":      n_sta,
        "n_waterlines":    n_wl,
        "L":               round(L, 4),
        "station_spacing": round(station_spacing, 4),
        "validation":      validation,
    }


def _extract_waterline_height(text: str):
    """
    从列标题文本中提取水线高度数值。

    支持格式：
        "WL0.5"  → 0.5
        "WL 1.0" → 1.0
        "0.5"    → 0.5
        "z=1.5"  → 1.5
        "1.5m"   → 1.5
        "水线1.0" → 1.0

    参数：
        text : str  列标题文本

    返回：
        float 或 None（无法识别时返回 None）
    """
    text = text.strip()

    # 直接是数字
    try:
        return float(text)
    except ValueError:
        pass

    # 匹配 WL0.5、WL 0.5、WL1、wl0.5 等
    m = re.match(r'(?:WL|wl|Wl)\s*([\d.]+)', text)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass

    # 匹配 z=0.5、Z=1.0
    m = re.match(r'[zZ]\s*=\s*([\d.]+)', text)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass

    # 匹配末尾数字（如 "水线1.0"、"1.0m"）
    m = re.search(r'([\d.]+)\s*m?$', text)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass

    return None


def validate_offsets(stations: list, waterlines: list, offsets: list) -> dict:
    """
    型值数据自动校验与修正。

    校验项目：
        1. 端点半宽校验：首尾站半宽应为 0（允许误差 ENDPOINT_TOLERANCE）
        2. 坐标系校验：中站应为 x=0，船首为正，船尾为负
        3. 负值校验：半宽不能为负
        4. 缺失值校验：NaN 或 None 值
        5. 单调性校验：水线高度应单调递增

    参数：
        stations   : list[float]  各站 x 坐标
        waterlines : list[float]  各水线高度
        offsets    : list[list]   半宽矩阵 offsets[wl][sta]

    返回：
        dict，包含：
            warnings          : list[str]  警告信息列表
            errors            : list[str]  错误信息列表
            corrections       : list[str]  自动修正记录
            corrected_offsets : list[list] 修正后的半宽矩阵
            is_valid          : bool       是否通过校验（无错误）
            coord_check       : dict       坐标系检查结果
    """
    import copy
    corrected = copy.deepcopy(offsets)

    warnings    = []
    errors      = []
    corrections = []

    n_wl  = len(waterlines)
    n_sta = len(stations)

    # ── 校验1：坐标系检查 ────────────────────────────────
    coord_check = _check_coordinate_system(stations)
    if coord_check["has_issue"]:
        warnings.extend(coord_check["messages"])

    # ── 校验2：水线高度单调性 ────────────────────────────
    for i in range(len(waterlines) - 1):
        if waterlines[i+1] <= waterlines[i]:
            errors.append(
                f"水线高度不单调：WL{waterlines[i]} 之后是 WL{waterlines[i+1]}，"
                f"水线高度应从小到大排列"
            )

    # ── 校验3：端点半宽校验与修正 ────────────────────────
    # 首站（x 最小，即船尾）和尾站（x 最大，即船首）的半宽应为 0
    first_sta_idx = 0
    last_sta_idx  = n_sta - 1

    for wl_idx in range(n_wl):
        z = waterlines[wl_idx]

        # 检查首站（船尾端）
        y_first = corrected[wl_idx][first_sta_idx]
        if abs(y_first) > ENDPOINT_TOLERANCE:
            corrections.append(
                f"WL{z}：船尾端（x={stations[first_sta_idx]}）半宽 {y_first:.3f} ≠ 0，"
                f"已自动修正为 0"
            )
            corrected[wl_idx][first_sta_idx] = 0.0

        # 检查末站（船首端）
        y_last = corrected[wl_idx][last_sta_idx]
        if abs(y_last) > ENDPOINT_TOLERANCE:
            corrections.append(
                f"WL{z}：船首端（x={stations[last_sta_idx]}）半宽 {y_last:.3f} ≠ 0，"
                f"已自动修正为 0"
            )
            corrected[wl_idx][last_sta_idx] = 0.0

    # ── 校验4：负值检查与修正 ────────────────────────────
    for wl_idx in range(n_wl):
        for sta_idx in range(n_sta):
            y = corrected[wl_idx][sta_idx]
            if y < -ENDPOINT_TOLERANCE:
                corrections.append(
                    f"WL{waterlines[wl_idx]}，站x={stations[sta_idx]}："
                    f"半宽为负值 {y:.3f}，已修正为 0"
                )
                corrected[wl_idx][sta_idx] = 0.0
            elif y < 0:
                # 极小负值（浮点误差），静默修正
                corrected[wl_idx][sta_idx] = 0.0

    # ── 校验5：缺失值检查 ────────────────────────────────
    missing_count = 0
    for wl_idx in range(n_wl):
        for sta_idx in range(n_sta):
            y = corrected[wl_idx][sta_idx]
            if y is None or (isinstance(y, float) and y != y):  # NaN check
                missing_count += 1
                corrected[wl_idx][sta_idx] = 0.0

    if missing_count > 0:
        warnings.append(f"发现 {missing_count} 个缺失值，已自动填充为 0")

    # ── 校验6：基线（WL0）半宽应全为0 ───────────────────
    if waterlines[0] == 0.0:
        non_zero = [corrected[0][i] for i in range(n_sta) if corrected[0][i] > ENDPOINT_TOLERANCE]
        if non_zero:
            warnings.append(
                f"基线（WL0.0）存在非零半宽，最大值 {max(non_zero):.3f}，请确认是否正确"
            )

    return {
        "warnings":          warnings,
        "errors":            errors,
        "corrections":       corrections,
        "corrected_offsets": corrected,
        "is_valid":          len(errors) == 0,
        "coord_check":       coord_check,
    }


def _check_coordinate_system(stations: list) -> dict:
    """
    检查坐标系是否符合规范：
        - 中站 x=0（允许误差 0.01m）
        - 船首（最大x）为正
        - 船尾（最小x）为负

    参数：
        stations : list[float]  各站 x 坐标

    返回：
        dict，包含 has_issue, messages
    """
    messages  = []
    has_issue = False

    x_min = min(stations)
    x_max = max(stations)

    # 检查是否有 x=0 的站（中站）
    has_zero = any(abs(x) < 0.01 for x in stations)
    if not has_zero:
        messages.append(
            f"未找到中站（x=0），当前站号范围 [{x_min}, {x_max}]。"
            f"请确认坐标系：中站应为 x=0"
        )
        has_issue = True

    # 检查船首为正
    if x_max <= 0:
        messages.append(
            f"船首端最大 x={x_max}，应为正值。"
            f"请确认坐标系：船首为正，船尾为负"
        )
        has_issue = True

    # 检查船尾为负
    if x_min >= 0:
        messages.append(
            f"船尾端最小 x={x_min}，应为负值。"
            f"请确认坐标系：船首为正，船尾为负"
        )
        has_issue = True

    return {"has_issue": has_issue, "messages": messages}
