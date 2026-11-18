@echo off
rem AttachLink one-click install launcher (all messages are inside scripts\setup.ps1)
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "scripts\setup.ps1"
echo.
pause
