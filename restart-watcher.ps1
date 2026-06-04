# =============================================================================
#  Media Pipeline - Restart the watcher (so config.ini changes take effect)
# =============================================================================
#  Run via "Restart Watcher.bat". Stops the running watcher and starts a fresh
#  one. Settings are read when the watcher starts, so a restart is what applies
#  any edits you made in config.ini. No administrator rights are needed.
# =============================================================================

$ErrorActionPreference = 'Stop'

$AppDir = $PSScriptRoot
if (-not $AppDir) { $AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$TaskName = 'Media Pipeline Watcher'

# Find the real watcher process(es). The command line of a live watcher is the
# launcher form:  <pwsh> ... -File "...\watch-media.ps1". We deliberately exclude
# any -Command query process and this restart helper to avoid matching ourselves.
function Get-WatcherProcesses {
    try {
        $found = Get-CimInstance Win32_Process -Filter "CommandLine LIKE '%watch-media.ps1%'" -ErrorAction SilentlyContinue |
            Where-Object {
                $_.CommandLine -like '*-File*' -and
                $_.CommandLine -notlike '*-Command*' -and
                $_.CommandLine -notlike '*restart-watcher*'
            }
        # Callers must wrap this in @(); a single match would otherwise unroll to a
        # scalar CimInstance whose .Count is not the element count.
        return $found
    }
    catch {
        return @()
    }
}

Write-Host 'Restarting the Media Pipeline watcher...'

# --- 1. Stop the running watcher(s) ---
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
    Start-Sleep -Seconds 1
}
else {
    Write-Host '  No running watcher found (it may already be stopped).'
}

# --- 2. Start a fresh watcher ---
$started = $false
if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    try {
        Start-ScheduledTask -TaskName $TaskName
        $started = $true
        Write-Host ("  Started via scheduled task '{0}'." -f $TaskName)
    }
    catch {
        Write-Host ("  Scheduled task start failed: {0}" -f $_.Exception.Message)
    }
}
if (-not $started) {
    $vbs = Join-Path $AppDir 'start-watcher-hidden.vbs'
    if (Test-Path -LiteralPath $vbs) {
        Start-Process -FilePath 'wscript.exe' -ArgumentList ('"' + $vbs + '"')
        $started = $true
        Write-Host '  Started via start-watcher-hidden.vbs.'
    }
}
if (-not $started) {
    Write-Host ''
    Write-Host 'Could not start the watcher. Run Install.bat first.' -ForegroundColor Yellow
    exit 1
}

# --- 3. Confirm it came up (single instance is enforced by a mutex) ---
Start-Sleep -Seconds 3
$now = @(Get-WatcherProcesses)
Write-Host ''
if ($now.Count -ge 1) {
    Write-Host ("Done - the watcher is running ({0} instance)." -f $now.Count) -ForegroundColor Green
    Write-Host 'Your settings from config.ini are now in effect.'
}
else {
    Write-Host 'The watcher did not appear to start yet. Give it a few seconds,' -ForegroundColor Yellow
    Write-Host 'then check the newest log in your PipelineRoot\logs folder.' -ForegroundColor Yellow
}
