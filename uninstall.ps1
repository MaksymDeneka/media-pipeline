# =============================================================================
#  Media Pipeline - Uninstaller
# =============================================================================
#  Run via "Uninstall.bat" (which elevates to administrator). It stops the
#  watcher and removes the auto-start task. It does NOT delete your media files,
#  your settings (config.ini), or the installed tools (FFmpeg/ExifTool/pwsh).
# =============================================================================

$ErrorActionPreference = 'Stop'
$TaskName = 'Media Pipeline Watcher'

function Get-WatcherProcesses {
    try {
        return @(Get-CimInstance Win32_Process -Filter "CommandLine LIKE '%watch-media.ps1%'" -ErrorAction SilentlyContinue |
            Where-Object {
                $_.CommandLine -like '*-File*' -and
                $_.CommandLine -notlike '*-Command*' -and
                $_.CommandLine -notlike '*restart-watcher*' -and
                $_.CommandLine -notlike '*uninstall*'
            })
    }
    catch {
        return @()
    }
}

Write-Host ''
Write-Host '====================================================='
Write-Host '   Media Pipeline - Uninstaller'
Write-Host '====================================================='
Write-Host '   Your media files, config.ini, and installed tools are kept.'
Write-Host ''

# --- 1. Stop the watcher ---
$running = @(Get-WatcherProcesses)
if ($running.Count -gt 0) {
    foreach ($p in $running) {
        try {
            Stop-Process -Id $p.ProcessId -Force -ErrorAction Stop
            Write-Host ("  Stopped watcher (PID {0})." -f $p.ProcessId)
        }
        catch {
            Write-Host ("  Could not stop PID {0}: {1}" -f $p.ProcessId, $_.Exception.Message)
        }
    }
}
else {
    Write-Host '  No running watcher found.'
}

# --- 2. Remove the scheduled task ---
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host ("  Removed scheduled task '{0}'." -f $TaskName)
}
else {
    Write-Host '  No scheduled task to remove.'
}

Write-Host ''
Write-Host 'Done. The watcher will no longer start automatically.' -ForegroundColor Green
Write-Host 'You can delete this app folder to remove it completely.'
Write-Host 'FFmpeg, ExifTool and PowerShell 7 were left installed. Remove them'
Write-Host 'with winget (e.g. "winget uninstall Gyan.FFmpeg") if you no longer need them.'
Write-Host ''
