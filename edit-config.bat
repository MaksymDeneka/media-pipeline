@echo off
rem === Open the settings file (config.ini) in Notepad ===
cd /d "%~dp0"

if not exist "%~dp0config.ini" (
    echo config.ini was not found in this folder.
    echo The watcher is using its built-in default settings.
    echo Run Install.bat to set everything up, then try again.
    echo.
    echo Press any key to close.
    pause >nul
    exit /b 0
)

echo Opening config.ini in Notepad...
echo After you save your changes, run "Restart Watcher.bat" to apply them.
start "" notepad.exe "%~dp0config.ini"
