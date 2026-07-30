@echo off
cd /d "%~dp0"

:: Check if already running via PowerShell (more reliable than curl on Windows)
powershell -NoProfile -Command ^
  "try { $r = Invoke-WebRequest -Uri 'https://localhost:3141/src/taskpane.html' -TimeoutSec 3 -UseBasicParsing -SkipCertificateCheck; exit 0 } catch { exit 1 }" >nul 2>&1
if not errorlevel 1 exit /b 0

:: Start Vite dev server in a minimized window
start "" /MIN cmd /c "title Pi Dev Server && npm run dev"

exit /b 0
