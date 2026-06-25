[CmdletBinding()]
param(
    [ValidateSet('root', 'default', 'convert', 'images', 'imageclean', 'long', 'sets', 'setbatch', 'assetstore')]
    [string]$Lane = 'root',

    [string]$LocalPath,
    [string]$RemotePath,
    [string]$LocalPipelineRoot = 'D:\MediaPipeline',
    [string]$RemoteRoot = '\\100.124.72.13\MediaPipeline',

    [int]$Threads = 16,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

function Get-SyncRelativePath {
    param([string]$Name)

    if ($Name -eq 'root') {
        return 'sync'
    }

    return (Join-Path $Name 'sync')
}

function Invoke-RobocopyChecked {
    param([string[]]$Arguments)

    & robocopy @Arguments
    $exitCode = $LASTEXITCODE

    if ($exitCode -gt 7) {
        throw "robocopy failed with exit code $exitCode"
    }

    $script:RobocopyExitCode = $exitCode
}

$relativePath = Get-SyncRelativePath -Name $Lane

if ([string]::IsNullOrWhiteSpace($LocalPath)) {
    $LocalPath = Join-Path $LocalPipelineRoot $relativePath
}

if ([string]::IsNullOrWhiteSpace($RemotePath)) {
    $RemotePath = Join-Path $RemoteRoot $relativePath
}

New-Item -ItemType Directory -Force -Path $LocalPath | Out-Null
New-Item -ItemType Directory -Force -Path $RemotePath | Out-Null

$logRoot = Join-Path $LocalPipelineRoot 'logs'
New-Item -ItemType Directory -Force -Path $logRoot | Out-Null
$logFile = Join-Path $logRoot ("robocopy-download-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

$robocopyArgs = @(
    $RemotePath,
    $LocalPath,
    '/E',
    '/Z',
    "/MT:$Threads",
    '/R:3',
    '/W:5',
    '/XO',
    '/FFT',
    '/TEE',
    '/ETA',
    "/LOG+:$logFile"
)

if ($DryRun) {
    $robocopyArgs += '/L'
}

Write-Host "Fast download source:      $RemotePath"
Write-Host "Fast download destination: $LocalPath"
Write-Host "Threads:                   $Threads"
Write-Host "Restartable mode:          enabled"
Write-Host "Dry run:                   $($DryRun.IsPresent)"
Write-Host "Log file:                  $logFile"

Invoke-RobocopyChecked -Arguments $robocopyArgs
Write-Host "robocopy exit code:        $script:RobocopyExitCode"
