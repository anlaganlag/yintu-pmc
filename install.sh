#!/bin/bash

echo "========================================"
echo "银图PMC智能分析平台 - 依赖安装程序"
echo "Yintu PMC Platform - Dependency Installer"
echo "========================================"
echo

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "[错误] 未检测到Python，请先安装Python 3.8或更高版本"
    echo "[Error] Python not found. Please install Python 3.8 or higher"
    exit 1
fi

echo "[信息] 检测到Python版本:"
python3 --version
echo

# 升级pip
echo "[步骤1] 升级pip到最新版本..."
python3 -m pip install --upgrade pip

# 安装依赖
echo
echo "[步骤2] 安装项目依赖..."
pip install -r requirements.txt

# 检查安装结果
if [ $? -eq 0 ]; then
    echo
    echo "========================================"
    echo "[成功] 所有依赖已安装完成！"
    echo "[Success] All dependencies installed!"
    echo
    echo "运行以下命令启动系统:"
    echo "To start the system, run:"
    echo "  streamlit run streamlit_dashboard.py"
    echo "========================================"
else
    echo
    echo "[错误] 依赖安装失败，请检查错误信息"
    echo "[Error] Installation failed, please check error messages"
    exit 1
fi