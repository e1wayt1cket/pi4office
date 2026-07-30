@echo off
chcp 65001 >nul

set "STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "LNK_PATH=%STARTUP_DIR%\pi4office-dev-server.lnk"

echo ============================================
echo   Uninstall Pi Dev Server Auto-Start
echo ============================================
echo.

if exist "%LNK_PATH%" (
    del "%LNK_PATH%"
    echo [OK] Startup shortcut removed. Dev server will no longer auto-start.
) else (
    echo [INFO] No startup shortcut found — nothing to uninstall.
)

echo.
pause
