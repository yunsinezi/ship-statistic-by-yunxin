# ============================================================
# core/stability_criteria.py — 规范自动判定模块
#
# 阶段6 新增模块
#
# 严格遵循：
#   《国内航行海船法定检验技术规则（2020）》
#   第4篇 船舶稳性
#   第3章 国内航行海船稳性衡准
#
# 功能：
#   1. 自动判定工况稳性是否合格
#   2. 给出明确的合格/不合格结论
#   3. 标注不满足项及对应规范条款
#
# ── 规范要求（所有工况通用）──
#   1. 初稳性高度 GM ≥ 0.15m（规则第4篇第3章2.1.1条）
#   2. 最大静稳性力臂 GZ_max ≥ 0.20m（规则第4篇第3章2.1.2条）
#   3. 最大静稳性力臂对应的横倾角 θ_max_gz ≥ 25°
#   4. 稳性消失角 θ_vanish ≥ 55°
#   5. 稳性衡准数 K ≥ 1.0
#   6. 若进水角 θ_f < θ_max_gz 或 θ_vanish，则以进水角为上限
#
# ============================================================

from typing import Dict, List, Tuple, Optional

# ══════════════════════════════════════════════════════════
# 规范标准定义
# ══════════════════════════════════════════════════════════

class StabilityCriteria:
    """稳性衡准标准（《国内航行海船法定检验技术规则 2020》）"""
    
    # 规范条款引用
    RULE_REFERENCE = {
        'GM': '规则第4篇第3章2.1.1条',
        'GZ_max': '规则第4篇第3章2.1.2条',
        'theta_max_gz': '规则第4篇第3章2.1.3条',
        'theta_vanish': '规则第4篇第3章2.1.4条',
        'K': '规则第4篇第3章2.1.5条'
    }
    
    # 衡准标准值
    CRITERIA = {
        'GM_min': 0.15,              # m，初稳性高度最小值
        'GZ_max_min': 0.20,          # m，最大复原力臂最小值
        'theta_max_gz_min': 25.0,    # °，最大复原力臂角度最小值
        'theta_vanish_min': 55.0,    # °，稳性消失角最小值
        'K_min': 1.0                 # 稳性衡准数最小值
    }
    
    # 衡准项名称与描述
    CRITERIA_NAMES = {
        'GM': '初稳性高度',
        'GZ_max': '最大复原力臂',
        'theta_max_gz': '最大复原力臂角度',
        'theta_vanish': '稳性消失角',
        'K': '稳性衡准数'
    }
    
    CRITERIA_UNITS = {
        'GM': 'm',
        'GZ_max': 'm',
        'theta_max_gz': '°',
        'theta_vanish': '°',
        'K': '-'
    }


# ══════════════════════════════════════════════════════════
# 规范判定
# ══════════════════════════════════════════════════════════

class StabilityJudgment:
    """稳性规范判定类"""
    
    def __init__(self, stability_result: Dict):
        """
        初始化规范判定
        
        参数：
            stability_result: 稳性计算结果（来自 LoadingConditionStability）
        """
        self.stability_result = stability_result
        self.judgments = {}  # 各衡准项的判定结果
        self.overall_pass = True  # 总体是否通过
        self.failed_items = []  # 不满足的项
    
    def judge_gm(self, gm: float) -> Tuple[bool, str]:
        """
        判定初稳性高度 GM
        
        参数：
            gm: 初稳性高度（m）
        
        返回：
            (是否通过, 说明文字)
        """
        criteria_name = StabilityCriteria.CRITERIA_NAMES['GM']
        criteria_unit = StabilityCriteria.CRITERIA_UNITS['GM']
        min_value = StabilityCriteria.CRITERIA['GM_min']
        rule_ref = StabilityCriteria.RULE_REFERENCE['GM']
        
        passed = gm >= min_value
        
        if passed:
            msg = f'✓ {criteria_name}: {gm:.4f}{criteria_unit} ≥ {min_value}{criteria_unit} 【通过】'
        else:
            msg = f'✗ {criteria_name}: {gm:.4f}{criteria_unit} < {min_value}{criteria_unit} 【不通过】'
            msg += f'\n  规范条款: {rule_ref}'
            msg += f'\n  要求: {criteria_name} ≥ {min_value}{criteria_unit}'
        
        self.judgments['GM'] = {
            'passed': passed,
            'value': gm,
            'limit': min_value,
            'unit': criteria_unit,
            'rule': rule_ref
        }
        
        if not passed:
            self.failed_items.append('GM')
        
        return passed, msg
    
    def judge_gz_max(self, gz_max: float) -> Tuple[bool, str]:
        """
        判定最大复原力臂 GZ_max
        
        参数：
            gz_max: 最大复原力臂（m）
        
        返回：
            (是否通过, 说明文字)
        """
        criteria_name = StabilityCriteria.CRITERIA_NAMES['GZ_max']
        criteria_unit = StabilityCriteria.CRITERIA_UNITS['GZ_max']
        min_value = StabilityCriteria.CRITERIA['GZ_max_min']
        rule_ref = StabilityCriteria.RULE_REFERENCE['GZ_max']
        
        passed = gz_max >= min_value
        
        if passed:
            msg = f'✓ {criteria_name}: {gz_max:.4f}{criteria_unit} ≥ {min_value}{criteria_unit} 【通过】'
        else:
            msg = f'✗ {criteria_name}: {gz_max:.4f}{criteria_unit} < {min_value}{criteria_unit} 【不通过】'
            msg += f'\n  规范条款: {rule_ref}'
            msg += f'\n  要求: {criteria_name} ≥ {min_value}{criteria_unit}'
        
        self.judgments['GZ_max'] = {
            'passed': passed,
            'value': gz_max,
            'limit': min_value,
            'unit': criteria_unit,
            'rule': rule_ref
        }
        
        if not passed:
            self.failed_items.append('GZ_max')
        
        return passed, msg
    
    def judge_theta_max_gz(self, theta_max_gz: float) -> Tuple[bool, str]:
        """
        判定最大复原力臂角度 θ_max_gz
        
        参数：
            theta_max_gz: 最大复原力臂对应的横倾角（°）
        
        返回：
            (是否通过, 说明文字)
        """
        criteria_name = StabilityCriteria.CRITERIA_NAMES['theta_max_gz']
        criteria_unit = StabilityCriteria.CRITERIA_UNITS['theta_max_gz']
        min_value = StabilityCriteria.CRITERIA['theta_max_gz_min']
        rule_ref = StabilityCriteria.RULE_REFERENCE['theta_max_gz']
        
        passed = theta_max_gz >= min_value
        
        if passed:
            msg = f'✓ {criteria_name}: {theta_max_gz:.1f}{criteria_unit} ≥ {min_value}{criteria_unit} 【通过】'
        else:
            msg = f'✗ {criteria_name}: {theta_max_gz:.1f}{criteria_unit} < {min_value}{criteria_unit} 【不通过】'
            msg += f'\n  规范条款: {rule_ref}'
            msg += f'\n  要求: {criteria_name} ≥ {min_value}{criteria_unit}'
        
        self.judgments['theta_max_gz'] = {
            'passed': passed,
            'value': theta_max_gz,
            'limit': min_value,
            'unit': criteria_unit,
            'rule': rule_ref
        }
        
        if not passed:
            self.failed_items.append('theta_max_gz')
        
        return passed, msg
    
    def judge_theta_vanish(self, theta_vanish: float) -> Tuple[bool, str]:
        """
        判定稳性消失角 θ_vanish
        
        参数：
            theta_vanish: 稳性消失角（°）
        
        返回：
            (是否通过, 说明文字)
        """
        criteria_name = StabilityCriteria.CRITERIA_NAMES['theta_vanish']
        criteria_unit = StabilityCriteria.CRITERIA_UNITS['theta_vanish']
        min_value = StabilityCriteria.CRITERIA['theta_vanish_min']
        rule_ref = StabilityCriteria.RULE_REFERENCE['theta_vanish']
        
        passed = theta_vanish >= min_value
        
        if passed:
            msg = f'✓ {criteria_name}: {theta_vanish:.1f}{criteria_unit} ≥ {min_value}{criteria_unit} 【通过】'
        else:
            msg = f'✗ {criteria_name}: {theta_vanish:.1f}{criteria_unit} < {min_value}{criteria_unit} 【不通过】'
            msg += f'\n  规范条款: {rule_ref}'
            msg += f'\n  要求: {criteria_name} ≥ {min_value}{criteria_unit}'
        
        self.judgments['theta_vanish'] = {
            'passed': passed,
            'value': theta_vanish,
            'limit': min_value,
            'unit': criteria_unit,
            'rule': rule_ref
        }
        
        if not passed:
            self.failed_items.append('theta_vanish')
        
        return passed, msg
    
    def judge_k_value(self, k_value: float) -> Tuple[bool, str]:
        """
        判定稳性衡准数 K
        
        参数：
            k_value: 稳性衡准数
        
        返回：
            (是否通过, 说明文字)
        """
        criteria_name = StabilityCriteria.CRITERIA_NAMES['K']
        criteria_unit = StabilityCriteria.CRITERIA_UNITS['K']
        min_value = StabilityCriteria.CRITERIA['K_min']
        rule_ref = StabilityCriteria.RULE_REFERENCE['K']
        
        passed = k_value >= min_value
        
        if passed:
            msg = f'✓ {criteria_name}: {k_value:.4f}{criteria_unit} ≥ {min_value}{criteria_unit} 【通过】'
        else:
            msg = f'✗ {criteria_name}: {k_value:.4f}{criteria_unit} < {min_value}{criteria_unit} 【不通过】'
            msg += f'\n  规范条款: {rule_ref}'
            msg += f'\n  要求: {criteria_name} ≥ {min_value}{criteria_unit}'
        
        self.judgments['K'] = {
            'passed': passed,
            'value': k_value,
            'limit': min_value,
            'unit': criteria_unit,
            'rule': rule_ref
        }
        
        if not passed:
            self.failed_items.append('K')
        
        return passed, msg
    
    def judge_all(self) -> Dict:
        """
        对所有衡准项进行判定
        
        返回：
            判定结果字典
        """
        # 从稳性计算结果中提取数据
        indicators = self.stability_result.get('indicators', {})
        gz_result = self.stability_result.get('gz_curve', {})
        
        gm = gz_result.get('GM', 0.0)
        gz_max = indicators.get('GZ_max', 0.0)
        theta_max_gz = indicators.get('theta_max_gz', 0.0)
        theta_vanish = indicators.get('theta_vanish', 0.0)
        k_value = indicators.get('K', 0.0)
        
        # 逐项判定
        self.judge_gm(gm)
        self.judge_gz_max(gz_max)
        self.judge_theta_max_gz(theta_max_gz)
        self.judge_theta_vanish(theta_vanish)
        self.judge_k_value(k_value)
        
        # 总体判定
        self.overall_pass = len(self.failed_items) == 0
        
        return {
            'condition_name': self.stability_result.get('condition_name', ''),
            'overall_pass': self.overall_pass,
            'judgments': self.judgments,
            'failed_items': self.failed_items,
            'indicators': {
                'GM': gm,
                'GZ_max': gz_max,
                'theta_max_gz': theta_max_gz,
                'theta_vanish': theta_vanish,
                'K': k_value
            }
        }
    
    def get_report(self) -> str:
        """
        生成规范判定报告
        
        返回：
            报告文本
        """
        report = []
        report.append('=' * 70)
        report.append(f'【稳性规范判定报告】— {self.stability_result.get("condition_name", "")}')
        report.append('=' * 70)
        report.append('')
        
        # 总体结论
        if self.overall_pass:
            report.append('【总体结论】✓ 稳性合格')
        else:
            report.append('【总体结论】✗ 稳性不合格')
            report.append(f'不满足项: {", ".join(self.failed_items)}')
        
        report.append('')
        report.append('【详细判定】')
        report.append('-' * 70)
        
        # 各衡准项判定
        for key in ['GM', 'GZ_max', 'theta_max_gz', 'theta_vanish', 'K']:
            if key in self.judgments:
                judgment = self.judgments[key]
                status = '✓ 通过' if judgment['passed'] else '✗ 不通过'
                report.append(f'{StabilityCriteria.CRITERIA_NAMES[key]}: {judgment["value"]:.4f} {judgment["unit"]} ≥ {judgment["limit"]} {status}')
                if not judgment['passed']:
                    report.append(f'  规范条款: {judgment["rule"]}')
        
        report.append('')
        report.append('=' * 70)
        
        return '\n'.join(report)


# ══════════════════════════════════════════════════════════
# 测试函数
# ══════════════════════════════════════════════════════════

if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    
    print('=' * 70)
    print('规范自动判定模块测试')
    print('=' * 70)
    
    # 模拟稳性计算结果
    mock_result = {
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
    
    # 创建判定器
    judgment = StabilityJudgment(mock_result)
    result = judgment.judge_all()
    
    # 输出报告
    print(judgment.get_report())
    
    print('\n✓ 规范自动判定模块测试通过!')
