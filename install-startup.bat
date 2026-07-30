@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "STARTUP_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"
set "VBS_SRC=%~dp0dev-server.vbs"
set "LNK_DST=%STARTUP_DIR%\pi4office-dev-server.lnk"

echo ============================================
echo   Install Pi Dev Server Auto-Start
echo ============================================
echo.

:: Create a shortcut in Startup folder using PowerShell
powershell -NoProfile -Command ^
  "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%LNK_DST%'); $s.TargetPath = '%VBS_SRC%'; $s.WorkingDirectory = '%~dp0'; $s.IconLocation = 'shell32.dll,13'; $s.Save()"

if errorlevel 1 (
    echo [FAIL] Could not create startup shortcut.
    pause
    exit /b 1
)

echo [OK] Startup shortcut created.
echo.
echo The dev server will now start automatically when you log in.
echo You can close this window.
echo.
echo To uninstall: double-click uninstall-startup.bat
echo.
pause
