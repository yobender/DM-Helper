@echo off
setlocal
cd /d "%~dp0.."
start "" /min cmd /c "npm run start"
exit /b 0
