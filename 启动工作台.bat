@echo off
chcp 65001 >nul
title HY的工作台
echo ============================================
echo    HY的工作台 启动中...
echo ============================================
echo.

cd /d "%~dp0"

:: 检查 Python 是否安装
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

:: 检查依赖
echo 正在检查依赖...
python -c "import openpyxl" 2>nul
if %errorlevel% neq 0 (
    echo 正在安装 openpyxl...
    pip install openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple
)
python -c "import pdfplumber" 2>nul
if %errorlevel% neq 0 (
    echo 正在安装 pdfplumber...
    pip install pdfplumber -i https://pypi.tuna.tsinghua.edu.cn/simple
)

echo.
echo 启动成功！浏览器将自动打开...
echo 如未自动打开，请手动访问: http://localhost:8080
echo 关闭此窗口即可停止工作台
echo.

:: 延迟1秒后打开浏览器
start "" http://localhost:8080

python app.py 8080

pause
