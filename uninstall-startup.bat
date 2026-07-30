@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "LNK_PATH=%STARTUP_DIR%\pi4office-dev-server.lnk"

echo ============================================
echo   Uninstall Pi for Office
echo ============================================
echo.

:: ── Remove auto-start ──
if exist "%LNK_PATH%" (
    del "%LNK_PATH%"
    echo [OK] Auto-start removed.
) else (
    echo [INFO] No auto-start found.
)

:: ── Stop running server ──
echo Stopping dev server...
npx kill-port 3141 >nul 2>&1
echo [OK] Server stopped.

echo.
echo ============================================
echo   Uninstall complete.
echo.
echo   To remove the add-in from Office:
echo     Excel: Insert -^> Add-ins -^> My Add-ins -^> right-click Pi -^> Remove
echo     Word:  same steps
echo ============================================
pause
