@echo off
setlocal
set "SCRIPT_DIR=%~dp0"

rem Prefer PowerShell 7 (enables parallel image processing); fall back to the
rem portable build, then to Windows PowerShell 5.1.
set "PWSH=C:\Program Files\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" set "PWSH=C:\Tools\pwsh\pwsh.exe"
if not exist "%PWSH%" set "PWSH=powershell.exe"

start "" "%PWSH%" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%SCRIPT_DIR%watch-media.ps1"
exit /b 0
