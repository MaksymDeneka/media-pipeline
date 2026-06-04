@echo off
rem === Restart the Media Pipeline watcher to apply config.ini changes ===
cd /d "%~dp0"
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0restart-watcher.ps1"

echo.
echo Press any key to close this window.
pause >nul
