"""
完整功能测试脚本
测试所有核心功能是否正常工作
"""

import os
import sys
import json

# 确保输出 UTF-8
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

print("=" * 70)
print("船舶静力学课程设计系统 - 完整功能测试")
print("=" * 70)

# 测试计数器
tests_passed = 0
tests_failed = 0

def test(name, func):
    """运行测试"""
    global tests_passed, tests_failed
    print(f"\n【{name}】")
    try:
        result = func()
        if result:
            print(f"✓ 通过")
            tests_passed += 1
            return True
        else:
            print(f"✗ 失败")
            tests_failed += 1
            return False
    except Exception as e:
        print(f"✗ 错误: {e}")
        tests_failed += 1
        return False

# ── 测试 1: 文件结构 ────────────────────────────────────────
def test_file_structure():
    """测试文件结构是否完整"""
    required_files = [
        "app.py",
        "main.py",
        "requirements.txt",
        "Procfile",
        "runtime.txt",
        "templates/index.html"
    ]
    
    missing = []
    for f in required_files:
        if not os.path.exists(f):
            missing.append(f)
    
    if missing:
        print(f"  缺失文件: {', '.join(missing)}")
        return False
    
    print(f"  所有必要文件存在")
    return True

test("文件结构检查", test_file_structure)

# ── 测试 2: 核心模块导入 ────────────────────────────────────
def test_imports():
    """测试核心模块导入"""
    try:
        print("  导入 Flask...")
        from flask import Flask
        print("  导入 NumPy...")
        import numpy
        print("  导入 Pandas...")
        import pandas
        print("  导入 Matplotlib...")
        import matplotlib
        print("  导入 openpyxl...")
        import openpyxl
        print("  导入 python-docx...")
        import docx
        print("  所有核心模块导入成功")
        return True
    except ImportError as e:
        print(f"  导入失败: {e}")
        return False

test("核心模块导入", test_imports)

# ── 测试 3: Flask 应用创建 ──────────────────────────────────
def test_flask_app():
    """测试 Flask 应用创建"""
    try:
        print("  创建 Flask 应用...")
        from app import app
        print(f"  应用名称: {app.name}")
        print("  Flask 应用创建成功")
        return True
    except Exception as e:
        print(f"  创建失败: {e}")
        return False

test("Flask 应用创建", test_flask_app)

# ── 测试 4: 路由配置 ────────────────────────────────────────
def test_routes():
    """测试路由配置"""
    try:
        from app import app
        
        # 获取所有路由
        routes = []
        for rule in app.url_map.iter_rules():
            routes.append(rule.rule)
        
        # 检查关键路由
        key_routes = [
            "/",
            "/upload_offsets",
            "/download_template",
            "/calc_table"
        ]
        
        missing = []
        for route in key_routes:
            if route not in routes:
                missing.append(route)
        
        if missing:
            print(f"  缺失路由: {', '.join(missing)}")
            return False
        
        print(f"  关键路由存在: {', '.join(key_routes)}")
        return True
    except Exception as e:
        print(f"  检查失败: {e}")
        return False

test("路由配置检查", test_routes)

# ── 测试 5: 目录权限 ────────────────────────────────────────
def test_directories():
    """测试目录权限"""
    try:
        # 测试 outputs 目录
        os.makedirs("outputs", exist_ok=True)
        test_file = "outputs/test_write.txt"
        with open(test_file, "w") as f:
            f.write("test")
        os.remove(test_file)
        
        # 测试 uploads 目录
        os.makedirs("uploads", exist_ok=True)
        test_file = "uploads/test_write.txt"
        with open(test_file, "w") as f:
            f.write("test")
        os.remove(test_file)
        
        print("  outputs/ 目录可写")
        print("  uploads/ 目录可写")
        return True
    except Exception as e:
        print(f"  权限错误: {e}")
        return False

test("目录权限检查", test_directories)

# ── 测试 6: 计算模块 ────────────────────────────────────────
def test_calculator():
    """测试计算模块"""
    try:
        print("  测试梯形法计算...")
        from core.calculator import calculate_hydrostatics
        
        # 测试数据
        stations = [-5, -4, -3, -2, -1, 0, 1, 2, 3, 4, 5]
        half_breadths = [0, 1.85, 4.62, 6.48, 7.35, 7.6, 7.35, 6.48, 4.62, 1.85, 0]
        L = 10.0
        draft = 5.0
        
        result = calculate_hydrostatics(stations, half_breadths, L, draft)
        
        print(f"  排水量: {result['delta']:.2f} t")
        print(f"  方形系数: {result['Cb']:.4f}")
        
        return True
    except Exception as e:
        print(f"  计算失败: {e}")
        import traceback
        traceback.print_exc()
        return False

test("计算模块测试", test_calculator)

# ── 测试 7: Excel 解析 ──────────────────────────────────────
def test_excel_parser():
    """测试 Excel 解析"""
    try:
        from core.template_generator import generate_template
        from core.excel_parser import parse_offsets_excel
        
        # 生成模板
        template_path = "outputs/test_template.xlsx"
        generate_template(template_path, example=True)
        
        if not os.path.exists(template_path):
            print("  模板生成失败")
            return False
        
        print("  模板生成成功")
        
        # 解析模板
        data = parse_offsets_excel(template_path)
        
        print(f"  站数: {data['n_stations']}")
        print(f"  水线数: {data['n_waterlines']}")
        
        # 清理
        os.remove(template_path)
        
        return True
    except Exception as e:
        print(f"  Excel 处理失败: {e}")
        import traceback
        traceback.print_exc()
        return False

test("Excel 解析测试", test_excel_parser)

# ── 测试 8: CORS 配置 ───────────────────────────────────────
def test_cors():
    """测试 CORS 配置"""
    try:
        from app import app
        
        # 检查是否有 CORS
        if hasattr(app, 'cors'):
            print("  CORS 已配置")
            return True
        
        # 检查 after_request 钩子
        has_cors = False
        for func in app.after_request_funcs.get(None, []):
            if 'cors' in func.__name__.lower():
                has_cors = True
                break
        
        if has_cors:
            print("  CORS 已配置（通过钩子）")
            return True
        
        # 检查是否导入了 CORS
        import sys
        if 'flask_cors' in sys.modules:
            print("  CORS 已配置（flask_cors）")
            return True
        
        print("  ⚠ CORS 未检测到，可能影响文件上传")
        return False
    except Exception as e:
        print(f"  CORS 检查失败: {e}")
        return False

test("CORS 配置检查", test_cors)

# ── 总结 ────────────────────────────────────────────────────
print("\n" + "=" * 70)
print("测试结果汇总")
print("=" * 70)

total = tests_passed + tests_failed
print(f"\n通过: {tests_passed}/{total}")
print(f"失败: {tests_failed}/{total}")

if tests_failed == 0:
    print("\n✓ 所有测试通过！系统可以正常部署。")
    print("\n部署步骤：")
    print("1. 推送代码到 GitHub")
    print("2. 在 Railway 连接仓库")
    print("3. 自动部署完成")
    print("4. 访问 Railway 提供的域名")
else:
    print("\n⚠ 部分测试失败，请先修复问题。")
    print("\n常见问题：")
    print("- 模块导入失败: pip install -r requirements.txt")
    print("- 目录权限问题: 检查 outputs/ 和 uploads/ 目录")
    print("- CORS 问题: 确认 app.py 中有 CORS(app)")
