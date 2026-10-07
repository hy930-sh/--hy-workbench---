#!/bin/bash
# HY的工作台 Mac 启动脚本
cd "$(dirname "$0")"

echo "============================================"
echo "   HY的工作台 启动中..."
echo "============================================"
echo ""

# 检查 Python3
if ! command -v python3 &> /dev/null; then
    echo "[错误] 未检测到 Python3，请先安装 Python 3.8+"
    echo "可通过 Homebrew 安装: brew install python3"
    exit 1
fi

# 检查依赖
echo "正在检查依赖..."
python3 -c "import openpyxl" 2>/dev/null || { echo "安装 openpyxl..."; pip3 install openpyxl; }
python3 -c "import pdfplumber" 2>/dev/null || { echo "安装 pdfplumber..."; pip3 install pdfplumber; }

echo ""
echo "启动成功！浏览器将自动打开..."
echo "如未自动打开，请手动访问: http://localhost:8080"
echo ""

# 打开浏览器
sleep 1
open http://localhost:8080

python3 app.py 8080
