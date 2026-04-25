@echo off
echo ========================================
echo    课表生成器 - Streamlit版本
echo ========================================
echo.

echo 正在检查Python环境...
python --version
if %errorlevel% neq 0 (
    echo 错误: 未找到Python，请先安装Python 3.7或更高版本
    pause
    exit /b 1
)

echo 正在检查依赖包...
pip show streamlit >nul 2>nul
if %errorlevel% neq 0 (
    echo 正在安装依赖包，请稍候...
    pip install -r requirements.txt
)

echo.
echo ========================================
echo    启动课表生成器...
echo ========================================
echo.
echo 请在浏览器中访问: http://localhost:8501
echo 按 CTRL+C 停止程序
echo.

streamlit run app.py --server.port 8501

pause
