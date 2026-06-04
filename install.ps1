# =============================================================================
#  Media Pipeline - Installer
# =============================================================================
#  Run this via "Install.bat" (which elevates to administrator). It:
#    1. installs FFmpeg, ExifTool and PowerShell 7 with winget,
#    2. makes sure the settings file (config.ini) is usable,
#    3. moves any files left in the old folder layout into default\,
#    4. registers the watcher to start automatically when you sign in,
#    5. validates everything and starts the watcher.
#
#  Written for Windows PowerShell 5.1 (the version that ships with Windows), so
#  it runs before PowerShell 7 is installed.
# =============================================================================

$ErrorActionPreference = 'Stop'

$AppDir = $PSScriptRoot
if (-not $AppDir) { $AppDir = Split-Path -Parent $MyInvocation.MyCommand.Path }
$ScriptPath = Join-Path $AppDir 'watch-media.ps1'
$ConfigPath = Join-Path $AppDir 'config.ini'
$VbsPath = Join-Path $AppDir 'start-watcher-hidden.vbs'
$TaskName = 'Media Pipeline Watcher'

function Write-Step($m) { Write-Host ''; Write-Host ("==> " + $m) -ForegroundColor Cyan }
function Write-Ok($m)   { Write-Host ("    " + $m) -ForegroundColor Green }
function Write-Note($m) { Write-Host ("    " + $m) -ForegroundColor Yellow }

function Read-PipelineRoot {
    $root = 'D:\MediaPipeline'
    if (Test-Path -LiteralPath $ConfigPath) {
        foreach ($line in (Get-Content -LiteralPath $ConfigPath)) {
            if ($line -match '^\s*PipelineRoot\s*=\s*(.+)$') {
                $val = $matches[1].Trim()
                $val = ($val -replace '\s+[;#].*$', '').Trim()
                $val = $val.Trim('"').Trim("'")
                if ($val) { $root = $val }
                break
            }
        }
    }
    return $root
}

function Set-PipelineRoot($newRoot) {
    if (-not (Test-Path -LiteralPath $ConfigPath)) { return }
    $lines = Get-Content -LiteralPath $ConfigPath
    $replaced = $false
    $out = foreach ($line in $lines) {
        if (-not $replaced -and $line -match '^\s*PipelineRoot\s*=') {
            $replaced = $true
            'PipelineRoot = ' + $newRoot
        }
        else { $line }
    }
    if (-not $replaced) { $out = @($out) + ('PipelineRoot = ' + $newRoot) }
    Set-Content -LiteralPath $ConfigPath -Value $out -Encoding UTF8
}

function Install-Dependency($id, $friendly, $probeCmd) {
    Write-Step ("Installing {0}  ({1})" -f $friendly, $id)
    if ($probeCmd -and (Get-Command $probeCmd -ErrorAction SilentlyContinue)) {
        Write-Ok ("{0} is already available - skipping." -f $friendly)
        return
    }
    & winget install --exact --id $id --accept-source-agreements --accept-package-agreements --silent --disable-interactivity
    if ($LASTEXITCODE -eq 0) {
        Write-Ok ("{0} installed." -f $friendly)
    }
    else {
        Write-Note ("winget exit code {0} for {1}. If it is already installed this is fine; otherwise install {1} manually." -f $LASTEXITCODE, $friendly)
    }
}

try {
    Write-Host ''
    Write-Host '====================================================='
    Write-Host '   Media Pipeline - Installer'
    Write-Host '====================================================='
    Write-Host ("   App folder: {0}" -f $AppDir)

    # --- 1. winget present? ---
    Write-Step 'Checking for winget (Windows Package Manager)...'
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Note 'winget was not found on this PC.'
        Write-Note 'winget comes with "App Installer" from the Microsoft Store.'
        Write-Note 'Open the Microsoft Store, install or update "App Installer",'
        Write-Note 'then run this installer again.'
        throw 'winget is required for automatic dependency installation.'
    }
    Write-Ok 'winget found.'

    # --- 2. Dependencies ---
    Install-Dependency 'Gyan.FFmpeg'         'FFmpeg'        'ffmpeg'
    Install-Dependency 'OliverBetz.ExifTool' 'ExifTool'      'exiftool'
    Install-Dependency 'Microsoft.PowerShell' 'PowerShell 7' 'pwsh'

    # --- 3. Refresh PATH so freshly-installed tools are visible to this process ---
    Write-Step 'Refreshing PATH for this session...'
    $machinePath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = (@($machinePath, $userPath) | Where-Object { $_ }) -join ';'
    Write-Ok 'Done.'

    # --- 4. Settings file ---
    Write-Step 'Checking the settings file (config.ini)...'
    if (Test-Path -LiteralPath $ConfigPath) {
        Write-Ok 'config.ini found.'
    }
    else {
        Write-Note 'config.ini is missing - the watcher will use built-in defaults.'
        Write-Note 'You can create one later; "Edit Config.bat" will help.'
    }

    $pipelineRoot = Read-PipelineRoot
    Write-Step ("Pipeline folder root: {0}" -f $pipelineRoot)
    $qualifier = $null
    try { $qualifier = Split-Path -Qualifier $pipelineRoot } catch { $qualifier = $null }
    if ($qualifier -and -not (Test-Path -LiteralPath ($qualifier + '\'))) {
        $fallback = Join-Path $env:USERPROFILE 'MediaPipeline'
        Write-Note ("Drive {0} does not exist on this PC." -f $qualifier)
        Write-Note ("Changing PipelineRoot to {0}." -f $fallback)
        Set-PipelineRoot $fallback
        $pipelineRoot = $fallback
    }
    else {
        Write-Ok 'Location looks good.'
    }

    # --- 5. Migrate files from the old (root-level) layout into default\ ---
    Write-Step 'Checking for files from the old folder layout...'
    $defaultRoot = Join-Path $pipelineRoot 'default'
    $pairs = @(
        @{ Old = (Join-Path $pipelineRoot 'input');    New = (Join-Path $defaultRoot 'input') },
        @{ Old = (Join-Path $pipelineRoot 'output');   New = (Join-Path $defaultRoot 'output') },
        @{ Old = (Join-Path $pipelineRoot 'original'); New = (Join-Path $defaultRoot 'original') },
        @{ Old = (Join-Path $pipelineRoot 'failed');   New = (Join-Path $defaultRoot 'failed') }
    )
    $movedAny = $false
    foreach ($p in $pairs) {
        if (-not (Test-Path -LiteralPath $p.Old)) { continue }
        $items = @(Get-ChildItem -LiteralPath $p.Old -Force -ErrorAction SilentlyContinue)
        if ($items.Count -eq 0) { continue }
        New-Item -ItemType Directory -Path $p.New -Force | Out-Null
        foreach ($item in $items) {
            $dest = Join-Path $p.New $item.Name
            if (Test-Path -LiteralPath $dest) {
                Write-Note ("Kept in place (already exists in new folder): {0}" -f $item.Name)
                continue
            }
            Move-Item -LiteralPath $item.FullName -Destination $dest -ErrorAction SilentlyContinue
            $movedAny = $true
        }
        Write-Ok ("Moved files: {0}  ->  {1}" -f $p.Old, $p.New)
    }
    if (-not $movedAny) { Write-Ok 'Nothing to move.' }

    # --- 6. Scheduled task (start at logon) ---
    Write-Step ("Registering startup task '{0}'..." -f $TaskName)
    $action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ('"' + $VbsPath + '"')
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit (New-TimeSpan -Seconds 0)
    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Set-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings | Out-Null
        Write-Ok 'Updated the existing task.'
    }
    else {
        Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Description 'Runs the local media pipeline watcher silently at logon.' | Out-Null
        Write-Ok 'Task registered.'
    }

    # --- 7. Validate tools + create folders ---
    Write-Step 'Validating tools and creating folders...'
    $runnerPath = 'powershell.exe'
    $pwshCmd = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($pwshCmd) { $runnerPath = $pwshCmd.Source }
    elseif (Test-Path 'C:\Program Files\PowerShell\7\pwsh.exe') { $runnerPath = 'C:\Program Files\PowerShell\7\pwsh.exe' }
    & $runnerPath -NoProfile -ExecutionPolicy Bypass -File $ScriptPath -CheckOnly
    if ($LASTEXITCODE -eq 0) {
        Write-Ok 'Startup check passed (tools found, folders created).'
    }
    else {
        Write-Note 'Startup check did not finish cleanly - see the messages above.'
    }

    # --- 8. Start the watcher now ---
    Write-Step 'Starting the watcher...'
    Start-ScheduledTask -TaskName $TaskName
    Start-Sleep -Seconds 2
    Write-Ok 'Watcher start requested (it runs hidden in the background).'

    Write-Host ''
    Write-Host '====================================================='
    Write-Host '   Setup complete!'
    Write-Host '====================================================='
    Write-Host ''
    Write-Host 'Next steps:'
    Write-Host ('  1. Set your browser''s download folder to:')
    Write-Host ('       ' + (Join-Path (Join-Path $pipelineRoot 'default') 'input')) -ForegroundColor White
    Write-Host '  2. To change settings: run "Edit Config.bat", save, then'
    Write-Host '     run "Restart Watcher.bat" to apply them.'
    Write-Host '  3. The watcher will start by itself every time you sign in.'
    Write-Host ''
}
catch {
    Write-Host ''
    Write-Host ('SETUP FAILED: ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host 'Nothing was left half-running. Fix the issue above and run Install.bat again.' -ForegroundColor Red
    exit 1
}
