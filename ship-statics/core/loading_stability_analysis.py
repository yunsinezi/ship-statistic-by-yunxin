# ============================================================
# core/loading_stability_analysis.py — 装载工况稳性分析集成模块
#
# 阶段6 集成模块
#
# 功能：
#   1. 集成装载工况管理、浮态计算、稳性计算、规范判定
#   2. 提供统一的分析接口
#   3. 生成完整的稳性分析报告
#
# ============================================================

import sys
sys.path.insert(0, 'F:/ship-statics')

from core.loading_condition import LoadingConditionManager, StandardLoadingConditions
from core.floating_stability import LoadingConditionStability, StabilityIndicators
from core.stability_criteria import StabilityJudgment
from typing import Dict, List, Optional

# ══════════════════════════════════════════════════════════
# 装载工况稳性分析
# ══════════════════════════════════════════════════════════

class LoadingStabilityAnalysis:
    """装载工况稳性分析类"""
    
    def __init__(self, offsets_data: Dict, hydrostatics_table: Dict, 
                 gz_curve_func):
        """
        初始化分析
        
        参数：
            offsets_data: 型值表数据
            hydrostatics_table: 静水力表
            gz_curve_func: GZ 曲线计算函数
        """
        self.offsets_data = offsets_data
        self.hydrostatics_table = hydrostatics_table
        self.gz_curve_func = gz_curve_func
        
        # 装载工况管理器
        self.condition_manager = LoadingConditionManager()
        
        # 分析结果
        self.analysis_results = {}  # {工况名: 分析结果}
        self.judgment_results = {}  # {工况名: 判定结果}
    
    def analyze_condition(self, condition_name: str) -> Dict:
        """
        分析单个工况
        
        参数：
            condition_name: 工况名称
        
        返回：
            分析结果字典
        """
        # 获取工况
        lc = self.condition_manager.get_condition(condition_name)
        if lc is None:
            raise ValueError(f"工况 {condition_name} 不存在")
        
        # 计算稳性
        stability_calc = LoadingConditionStability(
            lc, self.offsets_data, self.hydrostatics_table, self.gz_curve_func
        )
        result = stability_calc.calculate_stability()
        
        # 规范判定
        judgment = StabilityJudgment(result)
        judgment_result = judgment.judge_all()
        
        # 保存结果
        self.analysis_results[condition_name] = result
        self.judgment_results[condition_name] = judgment_result
        
        return result
    
    def analyze_all_standard_conditions(self) -> Dict:
        """
        分析所有 4 种必做工况
        
        返回：
            所有工况的分析结果
        """
        results = {}
        for condition_name in ['满载出港', '满载到港', '压载出港', '压载到港']:
            try:
                results[condition_name] = self.analyze_condition(condition_name)
            except Exception as e:
                results[condition_name] = {'error': str(e)}
        
        return results
    
    def get_judgment_summary(self) -> Dict:
        """
        获取所有工况的判定总结
        
        返回：
            判定总结字典
        """
        summary = {
            'total_conditions': len(self.judgment_results),
            'passed_conditions': 0,
            'failed_conditions': 0,
            'conditions': {}
        }
        
        for condition_name, judgment in self.judgment_results.items():
            passed = judgment.get('overall_pass', False)
            summary['conditions'][condition_name] = {
                'passed': passed,
                'failed_items': judgment.get('failed_items', [])
            }
            
            if passed:
                summary['passed_conditions'] += 1
            else:
                summary['failed_conditions'] += 1
        
        return summary
    
    def generate_report(self) -> str:
        """
        生成完整的稳性分析报告
        
        返回：
            报告文本
        """
        report = []
        report.append('=' * 80)
        report.append('【装载工况稳性分析报告】')
        report.append('=' * 80)
        report.append('')
        
        # 总体判定
        summary = self.get_judgment_summary()
        report.append('【总体判定】')
        report.append(f'分析工况数: {summary["total_conditions"]}')
        report.append(f'合格工况: {summary["passed_conditions"]}')
        report.append(f'不合格工况: {summary["failed_conditions"]}')
        report.append('')
        
        # 各工况详细结果
        report.append('【各工况详细结果】')
        report.append('-' * 80)
        
        for condition_name in ['满载出港', '满载到港', '压载出港', '压载到港']:
            if condition_name in self.judgment_results:
                judgment = self.judgment_results[condition_name]
                passed = judgment.get('overall_pass', False)
                status = '✓ 合格' if passed else '✗ 不合格'
                
                report.append(f'\n{condition_name}: {status}')
                
                if not passed:
                    failed_items = judgment.get('failed_items', [])
                    report.append(f'  不满足项: {", ".join(failed_items)}')
                
                # 关键指标
                indicators = judgment.get('indicators', {})
                report.append(f'  GM: {indicators.get("GM", 0):.4f}m')
                report.append(f'  GZ_max: {indicators.get("GZ_max", 0):.4f}m')
                report.append(f'  θ_max_gz: {indicators.get("theta_max_gz", 0):.1f}°')
                report.append(f'  θ_vanish: {indicators.get("theta_vanish", 0):.1f}°')
                report.append(f'  K: {indicators.get("K", 0):.4f}')
        
        report.append('')
        report.append('=' * 80)
        
        return '\n'.join(report)


# ══════════════════════════════════════════════════════════
# 测试函数
# ══════════════════════════════════════════════════════════

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    
    print('=' * 80)
    print('装载工况稳性分析集成模块测试')
    print('=' * 80)
    
    # 模拟数据
    offsets_data = {
        'stations': [-5, -4, -3, -2, -1, 0, 1, 2, 3, 4, 5],
        'waterlines': [0.0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0],
        'offsets': [[0]*11 for _ in range(17)]  # 简化的型值表
    }
    
    hydrostatics_table = {
        'drafts': [2.0, 3.0, 4.0, 5.0, 6.0, 7.0],
        'volumes': [800, 1200, 1600, 2000, 2400, 2800]
    }
    
    # 模拟 GZ 曲线计算函数
    def mock_gz_curve(offsets_data, draft, KG, theta_step=5.0):
        return {
            'GM': 0.8234,
            'rows': [
                {'theta': i*theta_step, 'GZ': 0.42*np.sin(np.radians(i*theta_step))}
                for i in range(int(90/theta_step)+1)
            ]
        }
    
    import numpy as np
    
    # 创建分析器
    analysis = LoadingStabilityAnalysis(offsets_data, hydrostatics_table, mock_gz_curve)
    
    # 分析单个工况
    print('\n【分析满载出港工况】')
    try:
        result = analysis.analyze_condition('满载出港')
        print(f'总重量: {result["loading"]["total_weight"]:.2f}t')
        print(f'重心坐标: ({result["loading"]["xg"]:.2f}, {result["loading"]["yg"]:.2f}, {result["loading"]["zg"]:.2f})m')
    except Exception as e:
        print(f'错误: {e}')
    
    # 生成报告
    print('\n【生成分析报告】')
    report = analysis.generate_report()
    print(report)
    
    print('\n✓ 装载工况稳性分析集成模块测试通过!')
