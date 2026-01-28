@echo off
REM Build script for packaging Amazon Product Parser as exe

echo ========================================
echo Amazon Product Parser - Build Script
echo ========================================
echo.

REM 检查虚拟环境
if exist ".venv\Scripts\activate.bat" (
    echo Activating virtual environment...
    call .venv\Scripts\activate.bat
) else (
    echo Warning: Virtual environment not found at .venv
    echo Please create it first: python -m venv .venv
    echo.
)

REM 检查 PyInstaller 是否已安装
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

REM 检查 Playwright 浏览器是否已安装
echo.
echo Checking Playwright browsers...
python -m playwright install chromium
if errorlevel 1 (
    echo Warning: Playwright browsers may not be installed.
    echo Users will need to run: playwright install chromium
)

REM 执行打包
echo.
echo Building executable...
pyinstaller amazon_parser.spec --clean

if errorlevel 1 (
    echo.
    echo Build failed! Please check the errors above.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build completed successfully!
echo ========================================
echo.
echo Executable location: dist\AmazonParser.exe
echo.
echo NOTE: To distribute to other users, they need:
echo   1. The exe file from dist\ folder
echo   2. Run: playwright install chromium (first time only)
echo.
pause
