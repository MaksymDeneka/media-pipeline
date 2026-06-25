@echo off
"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "%~dp0sync-download.ps1" %*
set "SYNC_EXIT_CODE=%ERRORLEVEL%"
echo.
if "%SYNC_EXIT_CODE%"=="0" (
  echo Download finished successfully.
) else (
  echo Download failed with exit code %SYNC_EXIT_CODE%.
)
echo.
pause
exit /b %SYNC_EXIT_CODE%
