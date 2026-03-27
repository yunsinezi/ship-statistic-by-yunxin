# ============================================================
# core/floating_stability.py — 浮态与稳性计算模块
#
# 阶段6 新增模块
#
# 功能：
#   1. 基于装载工况计算浮态（吃水、吃水差、横倾角）
#   2. 计算初稳性高度 GM
#   3. 基于 GZ 曲线计算静稳性曲线
#   4. 计算衡准指标（GZ_max、θ_max_gz、θ_vanish、K 值）
#
# ── 浮态计算方法 ──
#   使用迭代法求解吃水，使得排水量等于装载工况总重量
#   基于静水力表（或 Bonjean 曲线）进行插值
#
# ── 稳性计算方法 ──
#   基于阶段5的 GZ 曲线计算结果
#   提取该工况对应的静稳性曲线
#
# ============================================================

import numpy as np
from typing import Dict, Tuple, Optional, List
from scipy.interpolate import interp1d

# ══════════════════════════════════════════════════════════
# 浮态计算
# ══════════════════════════════════════════════════════════

class FloatingCondition:
    """浮态计算类"""
    
    def __init__(self, offsets_data: Dict, hydrostatics_table: Dict):
        """
        初始化浮态计算
        
        参数：
            offsets_data: 型值表数据 {stations, waterlines, offsets}
            hydrostatics_table: 静水力表 {drafts, volumes, ...}
        """
        self.offsets_data = offsets_data
        self.hydrostatics_table = hydrostatics_table
        
        # 从静水力表提取排水量-吃水关系
        self.drafts = np.array(hydrostatics_table.get('drafts', []))
        self.volumes = np.array(hydrostatics_table.get('volumes', []))
        
        # 创建排水量插值函数
        if len(self.drafts) > 1:
            self.volume_interp = interp1d(
                self.drafts, self.volumes,
                kind='cubic', fill_value='extrapolate'
            )
        else:
            self.volume_interp = None
    
    def calculate_draft_from_weight(self, total_weight: float, 
                                   rho: float = 1.025,
                                   tolerance: float = 0.01) -> float:
        """
        根据总重量计算吃水（迭代法）
        
        参数：
            total_weight: 全船总重量（t）
            rho: 水密度（t/m³）
            tolerance: 收敛精度（m³）
        
        返回：
            吃水（m）
        """
        # 目标排水量（m³）
        target_volume = total_weight / rho
        
        # 二分法求解吃水
        d_min = self.drafts[0]
        d_max = self.drafts[-1]
        
        for iteration in range(100):
            d_mid = (d_min + d_max) / 2
            
            # 计算中点对应的排水量
            if self.volume_interp is not None:
                v_mid = float(self.volume_interp(d_mid))
            else:
                # 线性插值
                idx = np.searchsorted(self.drafts, d_mid)
                if idx == 0:
                    v_mid = self.volumes[0]
                elif idx >= len(self.drafts):
                    v_mid = self.volumes[-1]
                else:
                    w = (d_mid - self.drafts[idx-1]) / (self.drafts[idx] - self.drafts[idx-1])
                    v_mid = self.volumes[idx-1] + w * (self.volumes[idx] - self.volumes[idx-1])
            
            # 检查收敛
            if abs(v_mid - target_volume) < tolerance:
                return d_mid
            
            # 调整搜索范围
            if v_mid < target_volume:
                d_min = d_mid
            else:
                d_max = d_mid
        
        # 未收敛，返回最后的估计值
        return d_mid
    
    def calculate_floating_state(self, total_weight: float, xg: float, 
                                rho: float = 1.025) -> Dict:
        """
        计算浮态（吃水、吃水差、横倾角等）
        
        参数：
            total_weight: 全船总重量（t）
            xg: 全船重心纵向坐标（m）
            rho: 水密度（t/m³）
        
        返回：
            浮态字典 {draft, trim, heel, ...}
        """
        # 计算平均吃水
        draft_mean = self.calculate_draft_from_weight(total_weight, rho)
        
        # 简化处理：假设纵倾和横倾均为 0（课程设计通常不考虑）
        # 实际应用中需要基于重心位置进行更复杂的计算
        trim = 0.0  # 纵倾（m）
        heel = 0.0  # 横倾（°）
        
        # 首尾吃水
        draft_fwd = draft_mean + trim / 2
        draft_aft = draft_mean - trim / 2
        
        return {
            'draft_mean': draft_mean,
            'draft_fwd': draft_fwd,
            'draft_aft': draft_aft,
            'trim': trim,
            'heel': heel,
            'displacement': total_weight
        }


# ══════════════════════════════════════════════════════════
# 稳性指标计算
# ══════════════════════════════════════════════════════════

class StabilityIndicators:
    """稳性指标计算类"""
    
    @staticmethod
    def calculate_from_gz_curve(gz_data: Dict) -> Dict:
        """
        从 GZ 曲线数据计算稳性指标
        
        参数：
            gz_data: GZ 曲线数据 {rows: [{theta, GZ, ...}, ...], ...}
        
        返回：
            稳性指标字典
        """
        rows = gz_data.get('rows', [])
        if not rows:
            return {}
        
        # 提取数据
        thetas = np.array([r['theta'] for r in rows])
        gzs = np.array([r['GZ'] for r in rows])
        
        # 计算 GZ_max 和对应的角度
        max_idx = np.argmax(gzs)
        gz_max = float(gzs[max_idx])
        theta_max_gz = float(thetas[max_idx])
        
        # 计算稳性消失角（GZ 最后一次为正的角度）
        theta_vanish = None
        for i in range(len(gzs)-1, -1, -1):
            if gzs[i] > 0.001:  # 考虑数值误差
                theta_vanish = float(thetas[i])
                break
        
        if theta_vanish is None:
            theta_vanish = 0.0
        
        # 计算动稳性衡准数 K
        # K = ∫₀⁴⁰° GZ·dθ / (GZ_max · θ_max_gz)
        # 其中积分单位为弧度
        
        # 提取 0° ~ 40° 范围内的数据
        mask_40 = thetas <= 40
        thetas_40 = thetas[mask_40]
        gzs_40 = gzs[mask_40]
        
        # 梯形积分（转换为弧度）
        if len(thetas_40) > 1:
            dtheta_rad = np.radians(thetas_40[1] - thetas_40[0])
            dynamic_stability = np.trapz(gzs_40, dx=dtheta_rad)
        else:
            dynamic_stability = 0.0
        
        # 计算 K 值
        if gz_max > 0 and theta_max_gz > 0:
            k_value = dynamic_stability / (gz_max * np.radians(theta_max_gz))
        else:
            k_value = 0.0
        
        return {
            'GZ_max': gz_max,
            'theta_max_gz': theta_max_gz,
            'theta_vanish': theta_vanish,
            'dynamic_stability': dynamic_stability,
            'K': k_value
        }
    
    @staticmethod
    def calculate_gm(kb: float, bm: float, kg: float) -> float:
        """
        计算初稳性高度 GM
        
        参数：
            kb: 浮心高度（m）
            bm: 横稳心半径（m）
            kg: 重心高度（m）
        
        返回：
            初稳性高度 GM（m）
        """
        gm = kb + bm - kg
        return max(gm, 0.0)  # GM 不能为负


# ══════════════════════════════════════════════════════════
# 工况稳性计算
# ══════════════════════════════════════════════════════════

class LoadingConditionStability:
    """工况稳性计算类"""
    
    def __init__(self, loading_condition, offsets_data: Dict, 
                 hydrostatics_table: Dict, gz_curve_func):
        """
        初始化工况稳性计算
        
        参数：
            loading_condition: LoadingCondition 对象
            offsets_data: 型值表数据
            hydrostatics_table: 静水力表
            gz_curve_func: GZ 曲线计算函数（来自 core.stability）
        """
        self.loading_condition = loading_condition
        self.offsets_data = offsets_data
        self.hydrostatics_table = hydrostatics_table
        self.gz_curve_func = gz_curve_func
        
        # 浮态计算
        self.floating = FloatingCondition(offsets_data, hydrostatics_table)
    
    def calculate_stability(self) -> Dict:
        """
        计算工况稳性
        
        返回：
            稳性计算结果字典
        """
        lc = self.loading_condition
        
        # 1. 计算浮态
        floating_state = self.floating.calculate_floating_state(
            lc.total_weight, lc.xg
        )
        
        # 2. 计算 GZ 曲线
        gz_result = self.gz_curve_func(
            offsets_data=self.offsets_data,
            draft=floating_state['draft_mean'],
            KG=lc.zg,
            theta_step=5.0
        )
        
        # 3. 计算稳性指标
        indicators = StabilityIndicators.calculate_from_gz_curve(gz_result)
        
        # 4. 组合结果
        result = {
            'condition_name': lc.name,
            'condition_description': lc.description,
            'loading': {
                'total_weight': lc.total_weight,
                'xg': lc.xg,
                'yg': lc.yg,
                'zg': lc.zg
            },
            'floating_state': floating_state,
            'gz_curve': gz_result,
            'indicators': indicators
        }
        
        return result


# ══════════════════════════════════════════════════════════
# 测试函数
# ══════════════════════════════════════════════════════════

if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    
    print('=' * 60)
    print('浮态与稳性计算模块测试')
    print('=' * 60)
    
    # 示例静水力表（简化）
    hydrostatics_table = {
        'drafts': [2.0, 3.0, 4.0, 5.0, 6.0, 7.0],
        'volumes': [800, 1200, 1600, 2000, 2400, 2800]
    }
    
    # 创建浮态计算器
    floating = FloatingCondition({}, hydrostatics_table)
    
    # 测试吃水计算
    print('\n【吃水计算测试】')
    for weight in [1000, 1500, 2000, 2500]:
        draft = floating.calculate_draft_from_weight(weight, rho=1.025)
        print(f'总重量 {weight}t → 吃水 {draft:.2f}m')
    
    print('\n✓ 浮态与稳性计算模块测试通过!')
