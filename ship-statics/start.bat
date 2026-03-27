@echo off
chcp 65001 >nul
title 船舶静力学课程设计系统 - 本地测试与部署

echo.
echo ══════════════════════════════════════════════════════════
echo   船舶静力学课程设计系统 v3.0
echo   武汉理工大学 船舶与海洋工程学院
echo ══════════════════════════════════════════════════════════
echo.

:: 检查 Python
echo [步骤 1/5] 检查 Python 环境...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ✗ 未检测到 Python，请先安装 Python 3.8+
    pause
    exit /b 1
)
echo ✓ Python 已安装
echo.

:: 安装依赖
echo [步骤 2/5] 安装/更新依赖包...
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple --quiet
if %errorlevel% equ 0 (
    echo ✓ 依赖安装完成
) else (
    echo ⚠ 依赖安装可能有问题，继续尝试启动...
)
echo.

:: 创建目录
echo [步骤 3/5] 创建必要目录...
if not exist "outputs" mkdir outputs
if not exist "uploads" mkdir uploads
echo ✓ 目录创建完成
echo.

:: 测试启动
echo [步骤 4/5] 测试启动...
echo 启动本地服务器，访问 http://127.0.0.1:5000
echo 按 Ctrl+C 停止服务器
echo.

:: 延迟打开浏览器
start "" cmd /c "timeout /t 3 /nobreak >nul && start http://127.0.0.1:5000"

:: 启动应用
python main.py

pause
