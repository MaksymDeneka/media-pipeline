[CmdletBinding()]
param(
    [string]$ConfigPath,
    [int]$StopTimeoutSeconds = 120
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $ConfigPath) {
    $nativeConfig = Join-Path $env:LOCALAPPDATA 'MediaPipelineNative\config.ini'
    $ConfigPath = if (Test-Path -LiteralPath $nativeConfig) {
        $nativeConfig
    }
    else {
        Join-Path $repoRoot 'config.ini'
    }
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
        else { $stableSince = $null }
        Start-Sleep -Milliseconds 250
    }
    return $false
}

function Test-NativeRunning([string]$LockPath) {
    if (-not (Test-Path -LiteralPath $LockPath)) { return $false }
    try {
        $stream = [IO.File]::Open($LockPath, 'Open', 'Read', 'ReadWrite')
        $stream.Dispose()
        return $false
    }
    catch [IO.IOException] { return $true }
}

$pipelineRoot = Read-PipelineRoot $ConfigPath
$legacyMutex = Get-LegacyMutexName $pipelineRoot
$control = Join-Path $pipelineRoot 'control'
$lock = Join-Path $pipelineRoot 'status\worker.lock'
if (Test-NativeRunning $lock) {
    New-Item -ItemType Directory -Path $control -Force | Out-Null
    New-Item -ItemType File -Path (Join-Path $control 'stop') -Force | Out-Null
    $deadline = (Get-Date).AddSeconds($StopTimeoutSeconds)
    while ((Get-Date) -lt $deadline -and (Test-NativeRunning $lock)) {
        Start-Sleep -Milliseconds 250
    }
    if (Test-NativeRunning $lock) {
        throw 'The native worker did not stop cleanly. The PowerShell watcher was not started.'
    }
}

$nativeTask = Get-ScheduledTask -TaskName 'Media Pipeline Native Worker' -ErrorAction SilentlyContinue
$legacyTask = Get-ScheduledTask -TaskName 'Media Pipeline Watcher' -ErrorAction SilentlyContinue
if (-not $legacyTask) {
    throw 'The PowerShell watcher task does not exist. Run the original installer to restore it.'
}

if ($nativeTask) { Disable-ScheduledTask -InputObject $nativeTask | Out-Null }
Enable-ScheduledTask -InputObject $legacyTask | Out-Null
Start-ScheduledTask -TaskName 'Media Pipeline Watcher'
if (Wait-LegacyHealthy $legacyMutex) {
    Write-Host 'PowerShell watcher restored and verified. The native startup task is disabled.'
    exit 0
}

Disable-ScheduledTask -InputObject $legacyTask | Out-Null
if ($nativeTask) {
    Enable-ScheduledTask -InputObject $nativeTask | Out-Null
    Start-ScheduledTask -TaskName 'Media Pipeline Native Worker'
    $deadline = (Get-Date).AddSeconds($StopTimeoutSeconds)
    while ((Get-Date) -lt $deadline -and -not (Test-NativeRunning $lock)) {
        Start-Sleep -Milliseconds 250
    }
    if (Test-NativeRunning $lock) {
        throw 'PowerShell restore failed its health check. The native worker was restarted and verified.'
    }
}
throw 'PowerShell restore failed its health check, and no healthy native worker could be restored.'
