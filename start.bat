@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   Pi for Office
echo ============================================

:: Start dev server minimized in background
echo Starting dev server...
start "" /MIN cmd /c "title Pi Dev Server && npm run dev"

:: Wait for server to be ready
echo Waiting for server...
:wait_loop
timeout /t 3 /nobreak >nul
curl -s -o NUL https://localhost:3141/src/taskpane.html -k 2>nul
if errorlevel 1 goto wait_loop

echo Server ready. Launching Excel...
start "" /MIN cmd /c "npx office-addin-debugging start manifest.xml desktop --app excel"

echo.
echo Done! Excel should open with Pi in the sidebar.
echo.
echo ============================================
echo   Quick commands:
echo     npm run dev          — start dev server
echo     npm run sideload     — reopen with add-in
echo     npx kill-port 3141   — stop dev server
echo ============================================
pause
