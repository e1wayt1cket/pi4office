@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: Auto-elevate to admin
net session >nul 2>&1
if errorlevel 1 (
    echo Requesting administrator permission...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo ============================================
echo   Uninstall Pi for Office
echo ============================================
echo.

:: ── Remove scheduled task ──
schtasks /Delete /TN "Pi4OfficeDevServer" /F >nul 2>&1
if errorlevel 1 (
    echo [INFO] No scheduled task found.
) else (
    echo [OK] Background auto-start removed.
)

:: ── Stop running server ──
echo Stopping dev server...
npx kill-port 3141 >nul 2>&1
echo [OK] Server stopped.

echo.
echo ============================================
echo   Uninstalled. Server will no longer auto-start.
echo ============================================
pause
