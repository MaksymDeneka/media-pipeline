@echo off
set "SCRIPT_DIR=%~dp0"
set "POWERSHELL_EXE=C:\Tools\pwsh\pwsh.exe"
start "" "%POWERSHELL_EXE%" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%SCRIPT_DIR%watch-media.ps1"
exit /b 0
