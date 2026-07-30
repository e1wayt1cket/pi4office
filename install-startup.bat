@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "VBS_SRC=%~dp0dev-server.vbs"
set "LNK_DST=%STARTUP_DIR%\pi4office-dev-server.lnk"

echo ============================================
echo   Pi for Office — One-Click Install
echo ============================================
echo.

:: ── Step 1: Auto-start registration ──
echo [1/3] Registering dev server auto-start...
powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%LNK_DST%'); $s.TargetPath = '%VBS_SRC%'; $s.WorkingDirectory = '%~dp0'; $s.IconLocation = 'shell32.dll,13'; $s.Save()"

if errorlevel 1 (
    echo [FAIL] Could not create startup shortcut.
    pause
    exit /b 1
)
echo       Done. Dev server will auto-start on login.

:: ── Step 2: Start dev server now ──
echo.
echo [2/3] Starting dev server...
cscript //Nologo "%~dp0dev-server.vbs"

:: Wait for server to be ready
:wait_loop
timeout /t 2 /nobreak >nul
curl -s -o NUL https://localhost:3141/src/taskpane.html -k 2>nul
if errorlevel 1 goto wait_loop
echo       Server ready.

:: ── Step 3: Sideload into Excel and Word ──
echo.
echo [3/3] Sideloading add-in into Excel and Word...
echo       (Office will open — this registers the add-in permanently)
echo.

echo --- Excel ---
npx office-addin-debugging start manifest.xml desktop --app excel

echo.
echo --- Word ---
npx office-addin-debugging start manifest.xml desktop --app word

echo.
echo ============================================
echo   Install complete!
echo.
echo   What happens now:
echo     - Dev server starts automatically on every login
echo     - Open Excel or Word, click "Open Pi" to use
echo.
echo   To uninstall: double-click uninstall-startup.bat
echo ============================================
pause
