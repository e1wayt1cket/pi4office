@echo off
echo Stopping Pi dev server...

:: Kill vite/node processes on port 3141
npx kill-port 3141

:: Also kill office-addin-debugging if running
taskkill /FI "WINDOWTITLE eq Pi Dev Server" /F 2>nul
taskkill /FI "WINDOWTITLE eq office-addin*" /F 2>nul

echo Done.
pause
