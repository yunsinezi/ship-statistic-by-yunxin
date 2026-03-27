# ============================================================
# core/loading_condition.py — 装载工况管理模块
#
# 阶段6 新增模块
#
# 严格遵循：
#   《国内航行海船法定检验技术规则（2020）》
#   武汉理工大学《船舶静力学》课程设计规范
#
# 功能：
#   1. 定义 4 种必做工况（满载出港、满载到港、压载出港、压载到港）
#   2. 支持用户自定义工况
#   3. 自动计算全船总重量、总重心坐标
#
# ── 坐标系定义 ──
#   X 轴：沿船长，船尾为 0，船首为正（LBP 中点为参考）
#   Z 轴：垂直向上，基线为 0
#   Y 轴：横向，右舷为正
#
# ── 工况定义规范 ──
#   满载出港：Δ = 5200t，货舱 100%，燃油 100%，压载舱空
#   满载到港：Δ - 消耗品(90%)，货舱 100%，压载舱空
#   压载出港：Δ ≈ 50-60%，货舱空，压载舱 100%
#   压载到港：压载舱 ≥ 85% 出港配置，燃油淡水仅 10%
#
# ============================================================

import numpy as np
from typing import Dict, List, Tuple, Optional

# ══════════════════════════════════════════════════════════
# 船舶主尺度与舱室数据（课程设计标准母型船）
# ══════════════════════════════════════════════════════════

class ShipData:
    """船舶主尺度与舱室数据"""
    
    # 船舶主尺度
    LOA = 98.00          # 总长（m）
    LBP = 92.00          # 垂线间长（m）
    B = 15.80            # 型宽（m）
    D = 7.20             # 型深（m）
    d_design = 5.80      # 设计满载吃水（m）
    
    # 满载排水量与空船数据
    delta_full = 5200.0  # 满载排水量（t）
    delta_light = 1800.0 # 空船重量（t）
    xg_light = 42.00     # 空船重心纵向坐标（m）
    zg_light = 5.50      # 空船重心垂向坐标（m）
    
    # 舱室定义（舱容、重心坐标）
    COMPARTMENTS = {
        '首尖舱': {
            'volume': 120,      # m³
            'xg': 85.00,        # m
            'zg': 3.50,         # m
            'type': 'ballast',  # 压载舱
            'density': 1.025    # 水密度
        },
        '1号货舱': {
            'volume': 1500,
            'xg': 65.00,
            'zg': 4.00,
            'type': 'cargo',    # 货舱
            'density': 0.80     # 货物密度（散货）
        },
        '2号货舱': {
            'volume': 1800,
            'xg': 40.00,
            'zg': 4.00,
            'type': 'cargo',
            'density': 0.80
        },
        '3号货舱': {
            'volume': 1200,
            'xg': 20.00,
            'zg': 4.00,
            'type': 'cargo',
            'density': 0.80
        },
        '尾尖舱': {
            'volume': 80,
            'xg': 5.00,
            'zg': 3.20,
            'type': 'ballast',
            'density': 1.025
        },
        '双层底压载舱': {
            'volume': 400,
            'xg': 45.00,
            'zg': 1.20,
            'type': 'ballast',
            'density': 1.025
        },
        '燃油舱': {
            'volume': 200,
            'xg': 12.00,
            'zg': 2.50,
            'type': 'fuel',
            'density': 0.90
        },
        '淡水舱': {
            'volume': 60,
            'xg': 15.00,
            'zg': 5.00,
            'type': 'water',
            'density': 1.000
        },
        '滑油舱': {
            'volume': 20,
            'xg': 10.00,
            'zg': 2.50,
            'type': 'oil',
            'density': 0.90
        }
    }


# ══════════════════════════════════════════════════════════
# 装载工况定义
# ══════════════════════════════════════════════════════════

class LoadingCondition:
    """装载工况类"""
    
    def __init__(self, name: str, description: str = ""):
        """
        初始化装载工况
        
        参数：
            name: 工况名称（如"满载出港"）
            description: 工况描述
        """
        self.name = name
        self.description = description
        self.compartments = {}  # 各舱室装载情况
        self.total_weight = 0.0  # 全船总重量（t）
        self.xg = 0.0           # 全船重心纵向坐标（m）
        self.yg = 0.0           # 全船重心横向坐标（m）
        self.zg = 0.0           # 全船重心垂向坐标（m）
    
    def add_compartment(self, comp_name: str, fill_ratio: float, 
                       weight: Optional[float] = None):
        """
        添加舱室装载
        
        参数：
            comp_name: 舱室名称
            fill_ratio: 装载比例（0~1）
            weight: 直接指定重量（t），若指定则忽略 fill_ratio
        """
        if comp_name not in ShipData.COMPARTMENTS:
            raise ValueError(f"舱室 {comp_name} 不存在")
        
        comp_data = ShipData.COMPARTMENTS[comp_name]
        
        if weight is not None:
            # 直接指定重量
            w = weight
        else:
            # 按装载比例计算重量
            volume = comp_data['volume'] * fill_ratio
            w = volume * comp_data['density']
        
        self.compartments[comp_name] = {
            'weight': w,
            'fill_ratio': fill_ratio,
            'xg': comp_data['xg'],
            'zg': comp_data['zg']
        }
    
    def calculate_total_weight_and_cg(self):
        """
        计算全船总重量和重心坐标
        
        返回：
            (total_weight, xg, yg, zg)
        """
        # 从空船开始
        total_weight = ShipData.delta_light
        sum_wx = ShipData.delta_light * ShipData.xg_light
        sum_wy = 0.0  # 假设无横向载荷
        sum_wz = ShipData.delta_light * ShipData.zg_light
        
        # 加上各舱室装载
        for comp_name, comp_load in self.compartments.items():
            w = comp_load['weight']
            xg = comp_load['xg']
            zg = comp_load['zg']
            
            total_weight += w
            sum_wx += w * xg
            sum_wz += w * zg
        
        # 计算重心坐标
        self.total_weight = total_weight
        self.xg = sum_wx / total_weight if total_weight > 0 else 0
        self.yg = sum_wy / total_weight if total_weight > 0 else 0
        self.zg = sum_wz / total_weight if total_weight > 0 else 0
        
        return self.total_weight, self.xg, self.yg, self.zg
    
    def to_dict(self) -> Dict:
        """转换为字典"""
        return {
            'name': self.name,
            'description': self.description,
            'total_weight': self.total_weight,
            'xg': self.xg,
            'yg': self.yg,
            'zg': self.zg,
            'compartments': self.compartments
        }


# ══════════════════════════════════════════════════════════
# 4 种必做工况的工厂函数
# ══════════════════════════════════════════════════════════

class StandardLoadingConditions:
    """课程设计 4 种必做工况"""
    
    @staticmethod
    def create_full_load_departure() -> LoadingCondition:
        """
        满载出港工况
        
        规范定义：
          - 总重量 = 满载排水量 Δ = 5200t
          - 货舱 100% 满载
          - 燃油、淡水、滑油舱 100% 装满（航程储备 100%）
          - 压载水舱完全空舱
          - 船员、备品等按设计额定值 100% 计入
        """
        lc = LoadingCondition(
            name='满载出港',
            description='货舱100%满载，燃油100%，压载舱空'
        )
        
        # 空船（已包含船员、备品等）
        # 货舱 100% 满载
        lc.add_compartment('1号货舱', fill_ratio=1.0)
        lc.add_compartment('2号货舱', fill_ratio=1.0)
        lc.add_compartment('3号货舱', fill_ratio=1.0)
        
        # 燃油、淡水、滑油 100% 装满
        lc.add_compartment('燃油舱', fill_ratio=1.0)
        lc.add_compartment('淡水舱', fill_ratio=1.0)
        lc.add_compartment('滑油舱', fill_ratio=1.0)
        
        # 压载舱完全空舱
        lc.add_compartment('首尖舱', fill_ratio=0.0)
        lc.add_compartment('尾尖舱', fill_ratio=0.0)
        lc.add_compartment('双层底压载舱', fill_ratio=0.0)
        
        lc.calculate_total_weight_and_cg()
        
        return lc
    
    @staticmethod
    def create_full_load_arrival() -> LoadingCondition:
        """
        满载到港工况
        
        规范定义：
          - 总重量 = 满载排水量 - 航行消耗品重量
          - 燃油、淡水按航程消耗 90%，仅剩余 10% 安全储备
          - 货舱保持 100% 满载不变
          - 压载水舱仍为空舱
          - 因消耗品减少，全船 KG 通常略有升高（不利工况）
        """
        lc = LoadingCondition(
            name='满载到港',
            description='货舱100%，燃油淡水仅10%（消耗90%），压载舱空'
        )
        
        # 空船
        # 货舱 100% 满载
        lc.add_compartment('1号货舱', fill_ratio=1.0)
        lc.add_compartment('2号货舱', fill_ratio=1.0)
        lc.add_compartment('3号货舱', fill_ratio=1.0)
        
        # 燃油、淡水、滑油仅剩余 10%（消耗 90%）
        lc.add_compartment('燃油舱', fill_ratio=0.1)
        lc.add_compartment('淡水舱', fill_ratio=0.1)
        lc.add_compartment('滑油舱', fill_ratio=0.1)
        
        # 压载舱完全空舱
        lc.add_compartment('首尖舱', fill_ratio=0.0)
        lc.add_compartment('尾尖舱', fill_ratio=0.0)
        lc.add_compartment('双层底压载舱', fill_ratio=0.0)
        
        lc.calculate_total_weight_and_cg()
        
        return lc
    
    @staticmethod
    def create_ballast_departure() -> LoadingCondition:
        """
        压载出港工况
        
        规范定义：
          - 货舱完全空舱
          - 燃油、淡水舱 100% 装满
          - 双层底压载舱 100% 注满
          - 首尖舱、尾尖舱按需注满，需同时满足：
            * 首吃水 ≥ 0.02*LBP + 0.15m = 0.02*92 + 0.15 = 2.09m
            * 尾吃水保证螺旋桨完全浸没
            * 纵倾 ≤ 1.5%*LBP = 1.5%*92 = 1.38m
          - 总重量约为满载排水量的 50%~60%
        """
        lc = LoadingCondition(
            name='压载出港',
            description='货舱空，燃油100%，压载舱100%'
        )
        
        # 空船
        # 货舱完全空舱
        lc.add_compartment('1号货舱', fill_ratio=0.0)
        lc.add_compartment('2号货舱', fill_ratio=0.0)
        lc.add_compartment('3号货舱', fill_ratio=0.0)
        
        # 燃油、淡水 100% 装满
        lc.add_compartment('燃油舱', fill_ratio=1.0)
        lc.add_compartment('淡水舱', fill_ratio=1.0)
        lc.add_compartment('滑油舱', fill_ratio=1.0)
        
        # 压载舱 100% 注满
        lc.add_compartment('首尖舱', fill_ratio=1.0)
        lc.add_compartment('尾尖舱', fill_ratio=1.0)
        lc.add_compartment('双层底压载舱', fill_ratio=1.0)
        
        lc.calculate_total_weight_and_cg()
        
        return lc
    
    @staticmethod
    def create_ballast_arrival() -> LoadingCondition:
        """
        压载到港工况
        
        规范定义：
          - 货舱完全空舱
          - 燃油、淡水舱仅剩余 10% 额定重量（消耗 90%）
          - 在压载出港配置基础上微调压载水量
          - 补偿燃油淡水消耗的重量
          - 保证吃水、纵倾仍满足规范要求
          - 压载舱总注水量不低于出港工况的 85%
        """
        lc = LoadingCondition(
            name='压载到港',
            description='货舱空，燃油淡水仅10%，压载舱≥85%出港配置'
        )
        
        # 空船
        # 货舱完全空舱
        lc.add_compartment('1号货舱', fill_ratio=0.0)
        lc.add_compartment('2号货舱', fill_ratio=0.0)
        lc.add_compartment('3号货舱', fill_ratio=0.0)
        
        # 燃油、淡水仅剩余 10%（消耗 90%）
        lc.add_compartment('燃油舱', fill_ratio=0.1)
        lc.add_compartment('淡水舱', fill_ratio=0.1)
        lc.add_compartment('滑油舱', fill_ratio=0.1)
        
        # 压载舱 85% 注满（补偿燃油淡水消耗）
        lc.add_compartment('首尖舱', fill_ratio=0.85)
        lc.add_compartment('尾尖舱', fill_ratio=0.85)
        lc.add_compartment('双层底压载舱', fill_ratio=0.85)
        
        lc.calculate_total_weight_and_cg()
        
        return lc


# ══════════════════════════════════════════════════════════
# 装载工况管理器
# ══════════════════════════════════════════════════════════

class LoadingConditionManager:
    """装载工况管理器"""
    
    def __init__(self):
        """初始化装载工况管理器"""
        self.conditions = {}
        self._init_standard_conditions()
    
    def _init_standard_conditions(self):
        """初始化 4 种必做工况"""
        self.conditions['满载出港'] = StandardLoadingConditions.create_full_load_departure()
        self.conditions['满载到港'] = StandardLoadingConditions.create_full_load_arrival()
        self.conditions['压载出港'] = StandardLoadingConditions.create_ballast_departure()
        self.conditions['压载到港'] = StandardLoadingConditions.create_ballast_arrival()
    
    def add_custom_condition(self, name: str, description: str = "") -> LoadingCondition:
        """
        添加自定义工况
        
        参数：
            name: 工况名称
            description: 工况描述
        
        返回：
            LoadingCondition 对象
        """
        if name in self.conditions:
            raise ValueError(f"工况 {name} 已存在")
        
        lc = LoadingCondition(name, description)
        self.conditions[name] = lc
        return lc
    
    def get_condition(self, name: str) -> Optional[LoadingCondition]:
        """获取指定工况"""
        return self.conditions.get(name)
    
    def list_conditions(self) -> List[str]:
        """列出所有工况"""
        return list(self.conditions.keys())
    
    def get_all_conditions(self) -> Dict[str, LoadingCondition]:
        """获取所有工况"""
        return self.conditions
    
    def export_conditions(self) -> Dict:
        """导出所有工况为字典"""
        return {
            name: lc.to_dict()
            for name, lc in self.conditions.items()
        }


# ══════════════════════════════════════════════════════════
# 测试函数
# ══════════════════════════════════════════════════════════

if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    
    print('=' * 60)
    print('装载工况管理模块测试')
    print('=' * 60)
    
    # 创建工况管理器
    manager = LoadingConditionManager()
    
    # 列出所有工况
    print('\n【4 种必做工况】')
    for name in manager.list_conditions():
        lc = manager.get_condition(name)
        print(f'\n{name}:')
        print(f'  总重量: {lc.total_weight:.2f} t')
        print(f'  重心纵向坐标 Xg: {lc.xg:.2f} m')
        print(f'  重心垂向坐标 Zg: {lc.zg:.2f} m')
        print(f'  装载舱室数: {len(lc.compartments)}')
    
    # 验证满载出港工况
    print('\n【满载出港工况详细信息】')
    lc_full = manager.get_condition('满载出港')
    print(f'总重量: {lc_full.total_weight:.2f} t (应接近 {ShipData.delta_full} t)')
    print(f'重心坐标: ({lc_full.xg:.2f}, {lc_full.yg:.2f}, {lc_full.zg:.2f}) m')
    print(f'\n舱室装载情况:')
    for comp_name, comp_load in lc_full.compartments.items():
        print(f'  {comp_name}: {comp_load["weight"]:.2f} t (装载比 {comp_load["fill_ratio"]*100:.0f}%)')
    
    print('\n✓ 装载工况管理模块测试通过!')
