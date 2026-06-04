@echo off
rem === Media Pipeline installer (asks for administrator, then runs install.ps1) ===

>nul 2>&1 net session
if %errorlevel% neq 0 (
    echo Requesting administrator permissions...
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

cd /d "%~dp0"
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"

echo.
echo Press any key to close this window.
pause >nul
