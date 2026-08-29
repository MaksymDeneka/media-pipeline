[CmdletBinding()]
param(
    [string]$BundleDirectory,
    [string]$InstallDirectory = (Join-Path $env:LOCALAPPDATA 'MediaPipelineNative'),
    [string]$ConfigPath,
    [int]$StopTimeoutSeconds = 120
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $BundleDirectory) {
    $BundleDirectory = Join-Path $repoRoot 'artifacts\native-worker\win-x64'
}
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $repoRoot 'config.ini'
}

$workerSource = Join-Path $BundleDirectory 'media-pipeline-worker.exe'
if (-not (Test-Path -LiteralPath $workerSource)) {
    throw "Native worker bundle not found. Run tools\Publish-NativeWorker.ps1 first."
}
if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "Configuration not found: $ConfigPath"
}

function Resolve-PipelineRoot([string]$Root, [string]$ConfigurationPath) {
    $expanded = $Root
    if ($expanded -eq '~') {
        $expanded = [Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)
    }
    elseif ($expanded.StartsWith('~/') -or $expanded.StartsWith('~\')) {
        $expanded = Join-Path `
            -Path ([Environment]::GetFolderPath([Environment+SpecialFolder]::UserProfile)) `
            -ChildPath ($expanded.Substring(2))
    }
    if ([IO.Path]::IsPathRooted($expanded)) {
        return [IO.Path]::GetFullPath($expanded)
    }
    return [IO.Path]::GetFullPath(
        $expanded,
        (Split-Path -Parent ([IO.Path]::GetFullPath($ConfigurationPath))))
}

function Read-PipelineRoot([string]$Path) {
    foreach ($line in Get-Content -LiteralPath $Path) {
        if ($line -match '^\s*PipelineRoot\s*=\s*(.+)$') {
            $value = $matches[1].Trim()
            if ($value -match '^"([^"]*)"\s*(?:[;#].*)?$' -or
                $value -match "^'([^']*)'\s*(?:[;#].*)?$") {
                $value = $matches[1]
            }
            else {
                $value = ($value -replace '\s+[;#].*$', '').Trim()
            }
            if ($value) { return Resolve-PipelineRoot $value $Path }
        }
    }
    return Resolve-PipelineRoot 'D:\MediaPipeline' $Path
}

function Get-LegacyMutexName([string]$Root) {
    $normalized = ([IO.Path]::GetFullPath($Root)).TrimEnd('\', '/').ToLowerInvariant()
    if ($normalized -eq 'd:\mediapipeline') { return 'Global\MediaPipelineWatcher' }
    $sha = [Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($normalized)
        $hash = [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '').Substring(0, 16)
        return "Global\MediaPipelineWatcher_$hash"
    }
    finally { $sha.Dispose() }
}

function Test-LegacyRunning([string]$MutexName) {
    try {
        $mutex = [Threading.Mutex]::OpenExisting($MutexName)
        $mutex.Dispose()
        return $true
    }
    catch [Threading.WaitHandleCannotBeOpenedException] { return $false }
    catch [UnauthorizedAccessException] { return $true }
}

function Wait-LegacyHealthy([string]$MutexName) {
    $deadline = (Get-Date).AddSeconds($StopTimeoutSeconds)
    $stableSince = $null
    while ((Get-Date) -lt $deadline) {
        if (Test-LegacyRunning $MutexName) {
            if ($null -eq $stableSince) { $stableSince = Get-Date }
            if ((Get-Date) - $stableSince -ge [TimeSpan]::FromSeconds(2)) { return $true }
        }
        else {
            $stableSince = $null
        }
        Start-Sleep -Milliseconds 250
    }
    return $false
}

$installedWorker = Join-Path $InstallDirectory 'media-pipeline-worker.exe'
$installedConfig = Join-Path $InstallDirectory 'config.ini'
$taskName = 'Media Pipeline Native Worker'

function Test-NativeRunning([string]$Worker, [string]$Configuration) {
    if (-not (Test-Path -LiteralPath $Worker) -or
        -not (Test-Path -LiteralPath $Configuration)) {
        return $false
    }
    try {
        $status = (& $Worker status --config $Configuration | ConvertFrom-Json)
        return [bool]$status.running
    }
    catch { return $false }
}

function Stop-NativeWorker([string]$Worker, [string]$Configuration) {
    if (-not (Test-NativeRunning $Worker $Configuration)) { return }
    & $Worker stop --config $Configuration
    if ($LASTEXITCODE -ne 0) { throw 'The installed native worker rejected its stop request.' }
    $deadline = (Get-Date).AddSeconds($StopTimeoutSeconds)
    while ((Get-Date) -lt $deadline -and (Test-NativeRunning $Worker $Configuration)) {
        Start-Sleep -Milliseconds 250
    }
    if (Test-NativeRunning $Worker $Configuration) {
        throw 'The installed native worker did not stop cleanly. Its files and config were not replaced.'
    }
}

$pipelineRoot = Read-PipelineRoot $ConfigPath
$legacyMutex = Get-LegacyMutexName $pipelineRoot

& $workerSource check --config $ConfigPath
if ($LASTEXITCODE -ne 0) { throw 'The requested native worker/config failed its startup check.' }

$nativeTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
$legacyTask = Get-ScheduledTask -TaskName 'Media Pipeline Watcher' -ErrorAction SilentlyContinue

try {
    if ($nativeTask) { Disable-ScheduledTask -InputObject $nativeTask | Out-Null }
    Stop-NativeWorker $installedWorker $installedConfig

    if ($legacyTask) {
        Disable-ScheduledTask -InputObject $legacyTask | Out-Null
    }
    if (Test-LegacyRunning $legacyMutex) {
        $control = Join-Path $pipelineRoot 'control'
        New-Item -ItemType Directory -Path $control -Force | Out-Null
        New-Item -ItemType File -Path (Join-Path $control 'stop') -Force | Out-Null
        $deadline = (Get-Date).AddSeconds($StopTimeoutSeconds)
        while ((Get-Date) -lt $deadline -and (Test-LegacyRunning $legacyMutex)) {
            Start-Sleep -Milliseconds 250
        }
        if (Test-LegacyRunning $legacyMutex) {
            throw 'The PowerShell watcher did not stop cleanly. The native worker was not started.'
        }
    }

    New-Item -ItemType Directory -Path $InstallDirectory -Force | Out-Null
    Copy-Item -Path (Join-Path $BundleDirectory '*') -Destination $InstallDirectory -Recurse -Force
    Copy-Item -LiteralPath $ConfigPath -Destination $installedConfig -Force

    & $installedWorker check --config $installedConfig
    if ($LASTEXITCODE -ne 0) { throw 'The installed native worker failed its startup check.' }

    $action = New-ScheduledTaskAction `
        -Execute $installedWorker `
        -Argument ('run --config "' + $installedConfig + '"')
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -RestartCount 3 `
        -RestartInterval (New-TimeSpan -Minutes 1) `
        -ExecutionTimeLimit (New-TimeSpan -Seconds 0)

    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
        Set-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings | Out-Null
    }
    else {
        Register-ScheduledTask `
            -TaskName $taskName `
            -Action $action `
            -Trigger $trigger `
            -Settings $settings `
            -Description 'Runs the cross-platform Media Pipeline worker at logon.' | Out-Null
    }

    Enable-ScheduledTask -TaskName $taskName | Out-Null
    Start-ScheduledTask -TaskName $taskName
    $deadline = (Get-Date).AddSeconds($StopTimeoutSeconds)
    while ((Get-Date) -lt $deadline -and -not (Test-NativeRunning $installedWorker $installedConfig)) {
        Start-Sleep -Milliseconds 250
    }
    if (-not (Test-NativeRunning $installedWorker $installedConfig)) {
        throw 'The native worker did not become healthy before the startup timeout.'
    }

    Write-Host 'Native worker installed and started.'
    Write-Host "Config: $installedConfig"
    Write-Host 'The PowerShell startup task is disabled, not deleted, so rollback remains available.'
}
catch {
    $cutoverError = $_.Exception.Message
    Stop-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
    Disable-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue | Out-Null
    $nativeDeadline = (Get-Date).AddSeconds($StopTimeoutSeconds)
    while ((Get-Date) -lt $nativeDeadline -and
        (Test-NativeRunning $installedWorker $installedConfig)) {
        Start-Sleep -Milliseconds 250
    }
    if (Test-NativeRunning $installedWorker $installedConfig) {
        throw "Native cutover failed: $cutoverError Rollback also failed because the native worker is still running."
    }
    if ($legacyTask) {
        Enable-ScheduledTask -TaskName 'Media Pipeline Watcher' | Out-Null
        if (-not (Test-LegacyRunning $legacyMutex)) {
            Start-ScheduledTask -TaskName 'Media Pipeline Watcher'
        }
        if (Wait-LegacyHealthy $legacyMutex) {
            throw "Native cutover failed: $cutoverError The legacy watcher was restored and verified."
        }
        throw "Native cutover failed: $cutoverError Legacy rollback also failed its health check."
    }
    throw "Native cutover failed: $cutoverError No legacy task exists for rollback."
}
