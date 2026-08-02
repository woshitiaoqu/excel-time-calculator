@echo off
chcp 65001 >nul
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.9 或更高版本（勾选 Add to PATH）。
    pause
    exit /b 1
)

if not exist ".venv" (
    echo [1/4] 创建虚拟环境...
    python -m venv .venv
)

call ".venv\Scripts\activate.bat"

echo [2/4] 安装依赖...
python -m pip install --upgrade pip
pip install -r requirements.txt

echo [3/4] 清理旧的打包文件...
if exist "dist\*.exe" del /q "dist\*.exe"
if exist "build" rmdir /s /q "build"

echo [4/4] 打包中（需 1-3 分钟）...
rem --collect-all selenium 打包 selenium-manager：运行时自动检测浏览器版本并匹配驱动
pyinstaller --noconfirm --onefile --noconsole --name ExcelTimeCalculator --collect-all tkinterdnd2 --collect-all selenium main.py

echo.
echo 打包完成！exe 位置：dist\ExcelTimeCalculator.exe
pause
