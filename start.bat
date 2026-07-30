@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   Pi for Office
echo ============================================

call "%~dp0dev-server.bat"

echo Starting Excel...
start "" /MIN cmd /c "npx office-addin-debugging start manifest.xml desktop --app excel"

echo Done.
pause
