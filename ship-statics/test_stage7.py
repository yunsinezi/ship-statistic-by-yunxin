# ============================================================
# test_stage7.py — 阶段7 测试脚本
#
# 测试课程设计 Word 说明书自动生成模块
#
# ============================================================

import sys
import os

# 添加项目根目录到 Python 路径
sys.path.insert(0, 'F:/ship-statics')

sys.stdout.reconfigure(encoding='utf-8')

print('=' * 70)
print('阶段7：课程设计 Word 说明书自动生成模块测试')
print('=' * 70)

# 导入模块
from core.word_report_generator import CourseDesignReportGenerator

print('\n✓ 模块导入成功')

# 测试数据
user_info = {
    'class_name': '船舶2020班',
    'student_name': '张三',
    'student_id': '202012345',
    'instructor': '李教授'
}

ship_data = {
    'ship_name': '3000DWT沿海散货船',
    'ship_type': '沿海散货船',
    'LOA': 98.0,
    'LBP': 92.0,
    'B': 15.8,
    'D': 7.2,
    'd_design': 5.8,
    'delta_full': 5200
}

hydrostatics_result = {
    'table': [
        [2.0, 800, 200, 1.2, 3.5, 0.65],
        [3.0, 1200, 250, 1.8, 3.2, 0.68],
        [4.0, 1800, 300, 2.3, 3.0, 0.70],
        [5.0, 2500, 350, 2.8, 2.8, 0.72],
        [5.8, 3200, 400, 3.3, 2.6, 0.74]
    ]
}

bonjean_result = {}

stability_results = {
    'conditions': {
        '满载出港': {
            'analysis': {
                'loading': {
                    'total_weight': 5200,
                    'xg': 42.5,
                    'yg': 0,
                    'zg': 5.1
                },
                'floating_state': {
                    'draft_mean': 5.80,
                    'draft_fwd': 5.80,
                    'draft_aft': 5.80,
                    'trim': 0.0
                },
                'gz_curve': {
                    'GM': 0.8234
                }
            },
            'judgment': {
                'overall_pass': True,
                'indicators': {
                    'GM': 0.8234,
                    'GZ_max': 0.42,
                    'theta_max_gz': 30.0,
                    'theta_vanish': 85.0,
                    'K': 1.2
                },
                'failed_items': []
            }
        },
        '满载到港': {
            'analysis': {
                'loading': {
                    'total_weight': 5000,
                    'xg': 43.0,
                    'yg': 0,
                    'zg': 5.3
                }
            },
            'judgment': {
                'overall_pass': True,
                'indicators': {
                    'GM': 0.75,
                    'GZ_max': 0.38,
                    'theta_max_gz': 28.0,
                    'theta_vanish': 82.0,
                    'K': 1.1
                },
                'failed_items': []
            }
        },
        '压载出港': {
            'analysis': {
                'loading': {
                    'total_weight': 3200,
                    'xg': 42.0,
                    'yg': 0,
                    'zg': 4.8
                }
            },
            'judgment': {
                'overall_pass': True,
                'indicators': {
                    'GM': 0.62,
                    'GZ_max': 0.32,
                    'theta_max_gz': 25.0,
                    'theta_vanish': 75.0,
                    'K': 0.9
                },
                'failed_items': []
            }
        },
        '压载到港': {
            'analysis': {
                'loading': {
                    'total_weight': 3000,
                    'xg': 41.5,
                    'yg': 0,
                    'zg': 5.0
                }
            },
            'judgment': {
                'overall_pass': False,
                'indicators': {
                    'GM': 0.12,
                    'GZ_max': 0.15,
                    'theta_max_gz': 20.0,
                    'theta_vanish': 60.0,
                    'K': 0.8
                },
                'failed_items': ['GM', 'GZ_max']
            }
        }
    }
}

image_paths = {}

# 创建生成器
print('\n【测试 1：创建 Word 报告生成器】')
generator = CourseDesignReportGenerator()
print('✓ 生成器创建成功')

# 生成报告
print('\n【测试 2：生成完整 Word 说明书】')
output_path = 'F:/ship-statics/outputs/课程设计说明书_测试.docx'

generator.generate_complete_report(
    user_info, ship_data, hydrostatics_result, bonjean_result,
    stability_results, image_paths, output_path
)

print(f'✓ Word 说明书已生成：{output_path}')

# 验证文件
if os.path.exists(output_path):
    file_size = os.path.getsize(output_path)
    print(f'✓ 文件大小：{file_size / 1024:.2f} KB')
else:
    print('✗ 文件生成失败')

print('\n' + '=' * 70)
print('✓ 阶段7 测试完成！')
print('=' * 70)
