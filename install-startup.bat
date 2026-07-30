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
echo   Pi for Office — One-Click Install
echo ============================================
echo.

:: ── Step 1: Create scheduled task for auto-start ──
echo [1/3] Registering background auto-start...
schtasks /Create /SC ONLOGON /TN "Pi4OfficeDevServer" /TR "\"%~dp0dev-server.bat\"" /F /RL HIGHEST /DELAY 0000:30 >nul 2>&1
if errorlevel 1 (
    echo [FAIL] Could not register scheduled task.
    echo        Try right-clicking this file and "Run as Administrator".
    pause
    exit /b 1
)
echo [OK] Dev server will auto-start on every login.

:: ── Step 2: Start server now ──
echo.
echo [2/3] Starting dev server now...
call "%~dp0dev-server.bat"

:: Wait for server to be ready
echo Waiting for server...
:wait_loop
timeout /t 2 /nobreak >nul
powershell -NoProfile -Command ^
  "try { Invoke-WebRequest -Uri 'https://localhost:3141/src/taskpane.html' -TimeoutSec 3 -UseBasicParsing -SkipCertificateCheck; exit 0 } catch { exit 1 }" >nul 2>&1
if errorlevel 1 goto wait_loop
echo [OK] Server ready at https://localhost:3141

:: ── Step 3: Sideload into Excel ──
echo.
echo [3/3] Sideloading add-in into Excel...
echo       (Excel will open — this registers the add-in permanently)
npx office-addin-debugging start manifest.xml desktop --app excel

echo.
echo ============================================
echo   Install complete!
echo.
echo   From now on:
echo     - Dev server auto-starts at login (invisible)
echo     - Open Excel/Word, click "Open Pi" to use
echo.
echo   To uninstall: right-click uninstall-startup.bat
echo                 and "Run as Administrator"
echo ============================================
pause
