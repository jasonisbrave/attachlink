@echo off
rem AttachLink one-click uninstall launcher (all messages are inside scripts\setup.ps1)
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "scripts\setup.ps1" -Uninstall
echo.
pause
