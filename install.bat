@echo off
echo ========================================
echo 银图PMC智能分析平台 - 依赖安装程序
echo Yintu PMC Platform - Dependency Installer
echo ========================================
echo.

REM 检查Python是否安装
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到Python，请先安装Python 3.8或更高版本
    echo [Error] Python not found. Please install Python 3.8 or higher
    pause
    exit /b 1
)

echo [信息] 检测到Python版本:
python --version
echo.

REM 升级pip
echo [步骤1] 升级pip到最新版本...
python -m pip install --upgrade pip

REM 安装依赖
echo.
echo [步骤2] 安装项目依赖...
pip install -r requirements.txt

REM 检查安装结果
if %errorlevel% equ 0 (
    echo.
    echo ========================================
    echo [成功] 所有依赖已安装完成！
    echo [Success] All dependencies installed!
    echo.
    echo 运行以下命令启动系统:
    echo To start the system, run:
    echo   streamlit run streamlit_dashboard.py
    echo ========================================
) else (
    echo.
    echo [错误] 依赖安装失败，请检查错误信息
    echo [Error] Installation failed, please check error messages
)

pause