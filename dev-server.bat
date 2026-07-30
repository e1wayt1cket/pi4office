@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: Check if server is already running
curl -s -o NUL https://localhost:3141/src/taskpane.html -k 2>nul
if not errorlevel 1 (
    echo Pi dev server is already running.
    exit /b 0
)

:: Start Vite dev server silently in background
start "" /MIN cmd /c "title Pi Dev Server && npm run dev"

echo Pi dev server starting...
exit /b 0
