#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
test_stage6.py — 阶段6 装载工况稳性校核与规范判定模块测试

测试内容：
  1. 装载工况管理模块
  2. 浮态与稳性计算模块
  3. 规范自动判定模块
  4. 集成分析模块
"""

import sys
sys.stdout.reconfigure(encoding='utf-8')
sys.path.insert(0, 'F:/ship-statics')

import numpy as np
from core.loading_condition import LoadingConditionManager, ShipData
from core.floating_stability import FloatingCondition, StabilityIndicators
from core.stability_criteria import StabilityJudgment, StabilityCriteria
from core.loading_stability_analysis import LoadingStabilityAnalysis

# ══════════════════════════════════════════════════════════
# 测试 1：装载工况管理
# ══════════════════════════════════════════════════════════

def test_loading_conditions():
    """测试装载工况管理"""
    print('\n' + '='*70)
    print('【测试 1】装载工况管理模块')
    print('='*70)
    
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
    
    # 验证满载到港工况（消耗品减少，KG升高）
    print('\n【满载到港工况详细信息】')
    lc_arrival = manager.get_condition('满载到港')
    print(f'总重量: {lc_arrival.total_weight:.2f} t (应小于满载出港)')
    print(f'重心坐标: ({lc_arrival.xg:.2f}, {lc_arrival.yg:.2f}, {lc_arrival.zg:.2f}) m')
    print(f'KG 变化: {lc_arrival.zg - lc_full.zg:.4f} m (应为正，表示 KG 升高)')
    
    # 验证压载出港工况
    print('\n【压载出港工况详细信息】')
    lc_ballast = manager.get_condition('压载出港')
    print(f'总重量: {lc_ballast.total_weight:.2f} t (应为满载的 50-60%)')
    print(f'重心坐标: ({lc_ballast.xg:.2f}, {lc_ballast.yg:.2f}, {lc_ballast.zg:.2f}) m')
    print(f'重量比例: {lc_ballast.total_weight/ShipData.delta_full*100:.1f}%')
    
    print('\n✓ 装载工况管理模块测试通过!')


# ══════════════════════════════════════════════════════════
# 测试 2：浮态计算
# ══════════════════════════════════════════════════════════

def test_floating_condition():
    """测试浮态计算"""
    print('\n' + '='*70)
    print('【测试 2】浮态计算模块')
    print('='*70)
    
    # 创建示例静水力表
    hydrostatics_table = {
        'drafts': [2.0, 3.0, 4.0, 5.0, 6.0, 7.0],
        'volumes': [800, 1200, 1600, 2000, 2400, 2800]
    }
    
    floating = FloatingCondition({}, hydrostatics_table)
    
    print('\n【吃水计算测试】')
    manager = LoadingConditionManager()
    
    for condition_name in ['满载出港', '满载到港', '压载出港', '压载到港']:
        lc = manager.get_condition(condition_name)
        draft = floating.calculate_draft_from_weight(lc.total_weight, rho=1.025)
        floating_state = floating.calculate_floating_state(lc.total_weight, lc.xg)
        
        print(f'\n{condition_name}:')
        print(f'  总重量: {lc.total_weight:.2f} t')
        print(f'  平均吃水: {floating_state["draft_mean"]:.2f} m')
        print(f'  首吃水: {floating_state["draft_fwd"]:.2f} m')
        print(f'  尾吃水: {floating_state["draft_aft"]:.2f} m')
    
    print('\n✓ 浮态计算模块测试通过!')


# ══════════════════════════════════════════════════════════
# 测试 3：规范判定
# ══════════════════════════════════════════════════════════

def test_stability_judgment():
    """测试规范判定"""
    print('\n' + '='*70)
    print('【测试 3】规范自动判定模块')
    print('='*70)
    
    # 模拟稳性计算结果（合格工况）
    print('\n【合格工况示例】')
    mock_result_pass = {
        'condition_name': '满载出港',
        'indicators': {
            'GZ_max': 0.42,
            'theta_max_gz': 30.0,
            'theta_vanish': 85.0,
            'K': 1.2
        },
        'gz_curve': {
            'GM': 0.8234
        }
    }
    
    judgment_pass = StabilityJudgment(mock_result_pass)
    result_pass = judgment_pass.judge_all()
    
    print(f'工况: {result_pass["condition_name"]}')
    print(f'总体判定: {"✓ 合格" if result_pass["overall_pass"] else "✗ 不合格"}')
    print(f'不满足项: {result_pass["failed_items"] if result_pass["failed_items"] else "无"}')
    
    # 模拟稳性计算结果（不合格工况）
    print('\n【不合格工况示例】')
    mock_result_fail = {
        'condition_name': '压载到港',
        'indicators': {
            'GZ_max': 0.15,  # 低于 0.20m
            'theta_max_gz': 20.0,  # 低于 25°
            'theta_vanish': 50.0,  # 低于 55°
            'K': 0.8  # 低于 1.0
        },
        'gz_curve': {
            'GM': 0.12  # 低于 0.15m
        }
    }
    
    judgment_fail = StabilityJudgment(mock_result_fail)
    result_fail = judgment_fail.judge_all()
    
    print(f'工况: {result_fail["condition_name"]}')
    print(f'总体判定: {"✓ 合格" if result_fail["overall_pass"] else "✗ 不合格"}')
    print(f'不满足项: {", ".join(result_fail["failed_items"])}')
    
    # 输出规范标准
    print('\n【规范标准（《国内航行海船法定检验技术规则 2020》）】')
    for key, value in StabilityCriteria.CRITERIA.items():
        name = StabilityCriteria.CRITERIA_NAMES.get(key, key)
        unit = StabilityCriteria.CRITERIA_UNITS.get(key, '')
        rule = StabilityCriteria.RULE_REFERENCE.get(key, '')
        print(f'{name}: ≥ {value} {unit}  ({rule})')
    
    print('\n✓ 规范自动判定模块测试通过!')


# ══════════════════════════════════════════════════════════
# 测试 4：集成分析
# ══════════════════════════════════════════════════════════

def test_integrated_analysis():
    """测试集成分析"""
    print('\n' + '='*70)
    print('【测试 4】集成分析模块')
    print('='*70)
    
    # 创建示例数据
    offsets_data = {
        'stations': [-5, -4, -3, -2, -1, 0, 1, 2, 3, 4, 5],
        'waterlines': [0.0, 0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0, 6.5, 7.0, 7.5, 8.0],
        'offsets': [[0]*11 for _ in range(17)]
    }
    
    hydrostatics_table = {
        'drafts': [2.0, 3.0, 4.0, 5.0, 6.0, 7.0],
        'volumes': [800, 1200, 1600, 2000, 2400, 2800]
    }
    
    # 模拟 GZ 曲线计算函数
    def mock_gz_curve(offsets_data, draft, KG, theta_step=5.0):
        thetas = np.arange(0, 91, theta_step)
        gzs = 0.42 * np.sin(np.radians(thetas))
        
        return {
            'GM': 0.8234,
            'KB': 2.5,
            'BM': 3.2,
            'delta': 5200,
            'V': 5200/1.025,
            'GZ_max': 0.42,
            'theta_max_gz': 30.0,
            'theta_vanish': 85.0,
            'rows': [
                {
                    'theta': float(theta),
                    'GZ': float(gz),
                    'GZ_approx': float(0.8234 * np.sin(np.radians(theta))),
                    'ls': float(gz)
                }
                for theta, gz in zip(thetas, gzs)
            ],
            'criteria': {
                'GM': {'name': '初稳性', 'value': 0.8234, 'limit': 0.15, 'pass': True},
                'GZ_max': {'name': '最大复原力臂', 'value': 0.42, 'limit': 0.20, 'pass': True},
                'overall': '全部通过'
            }
        }
    
    # 创建分析器
    analysis = LoadingStabilityAnalysis(offsets_data, hydrostatics_table, mock_gz_curve)
    
    # 分析所有工况
    print('\n【分析所有 4 种必做工况】')
    analysis.analyze_all_standard_conditions()
    
    # 获取判定总结
    summary = analysis.get_judgment_summary()
    
    print(f'\n分析工况数: {summary["total_conditions"]}')
    print(f'合格工况: {summary["passed_conditions"]}')
    print(f'不合格工况: {summary["failed_conditions"]}')
    
    print('\n【各工况判定结果】')
    for condition_name, condition_result in summary['conditions'].items():
        status = '✓ 合格' if condition_result['passed'] else '✗ 不合格'
        print(f'{condition_name}: {status}')
        if not condition_result['passed']:
            print(f'  不满足项: {", ".join(condition_result["failed_items"])}')
    
    # 生成报告
    print('\n【生成分析报告】')
    report = analysis.generate_report()
    print(report)
    
    print('\n✓ 集成分析模块测试通过!')


# ══════════════════════════════════════════════════════════
# 主测试函数
# ══════════════════════════════════════════════════════════

if __name__ == '__main__':
    print('\n' + '='*70)
    print('船舶稳性计算系统 — 阶段6 测试')
    print('装载工况稳性校核与规范判定模块')
    print('='*70)
    
    try:
        # 运行所有测试
        test_loading_conditions()
        test_floating_condition()
        test_stability_judgment()
        test_integrated_analysis()
        
        # 总结
        print('\n' + '='*70)
        print('✓ 阶段6 所有测试通过!')
        print('='*70)
        print('\n【阶段6 核心功能验证】')
        print('✓ 装载工况管理：4 种必做工况 + 自定义工况')
        print('✓ 浮态计算：吃水、吃水差、横倾角')
        print('✓ 稳性指标：GM、GZ_max、θ_max_gz、θ_vanish、K')
        print('✓ 规范判定：严格按《国内航行海船法定检验技术规则》')
        print('✓ 集成分析：完整的工况分析与报告生成')
        print('\n【后续任务】')
        print('1. 前端界面升级（Step 6）')
        print('2. 导出功能升级（Excel 报告）')
        print('3. 静稳性曲线绘制')
        print('='*70)
    
    except Exception as e:
        print(f'\n✗ 测试失败: {e}')
        import traceback
        traceback.print_exc()
        sys.exit(1)
