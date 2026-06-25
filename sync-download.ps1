[CmdletBinding()]
param(
    [ValidateSet('root', 'default', 'convert', 'images', 'imageclean', 'long', 'sets', 'setbatch', 'assetstore')]
    [string]$Lane = 'root',

    [string]$LocalPath,
    [string]$RemotePath,
    [string]$LocalPipelineRoot = 'D:\MediaPipeline',
    [string]$RemotePipelineRoot = '/D:/MediaPipeline',
    [string]$RemoteName = 'heatup-remote-sftp',

    [int]$Transfers = 3,
    [int]$Checkers = 8,
    [switch]$DryRun,
    [switch]$Mirror,
    [switch]$ConfirmDelete
)

$ErrorActionPreference = 'Stop'

function Get-SyncRelativePath {
    param([string]$Name)

    if ($Name -eq 'root') {
        return 'sync'
    }

    return (Join-Path $Name 'sync')
}

function Join-RcloneRemotePath {
    param(
        [string]$Root,
        [string]$RelativePath
    )

    $cleanRoot = $Root.Replace('\', '/').TrimEnd('/')
    $cleanRelative = $RelativePath.Replace('\', '/').TrimStart('/')
    return "$cleanRoot/$cleanRelative"
}

function Invoke-RcloneChecked {
    param([string[]]$Arguments)

    & rclone @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "rclone exited with code $LASTEXITCODE"
    }
}

$relativePath = Get-SyncRelativePath -Name $Lane

if ([string]::IsNullOrWhiteSpace($LocalPath)) {
    $LocalPath = Join-Path $LocalPipelineRoot $relativePath
}

if ([string]::IsNullOrWhiteSpace($RemotePath)) {
    $remoteSubPath = Join-RcloneRemotePath -Root $RemotePipelineRoot -RelativePath $relativePath
    $RemotePath = "${RemoteName}:$remoteSubPath"
}

if ($Mirror -and -not $DryRun -and -not $ConfirmDelete) {
    throw 'Mirror mode runs rclone sync and can delete destination files. Run a dry-run first, then rerun with -Mirror -ConfirmDelete if the output is correct.'
}

New-Item -ItemType Directory -Force -Path $LocalPath | Out-Null

$logRoot = Join-Path $LocalPipelineRoot 'logs'
New-Item -ItemType Directory -Force -Path $logRoot | Out-Null
$logFile = Join-Path $logRoot ("rclone-download-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))

$operation = if ($Mirror) { 'sync' } else { 'copy' }

$rcloneArgs = @(
    $operation,
    $RemotePath,
    $LocalPath,
    '--progress',
    '--stats', '10s',
    '--transfers', $Transfers.ToString(),
    '--checkers', $Checkers.ToString(),
    '--retries', '5',
    '--low-level-retries', '20',
    '--timeout', '10m',
    '--contimeout', '30s',
    '--modify-window', '2s',
    '--create-empty-src-dirs',
    '--partial-suffix', '.rclone-partial',
    '--sftp-disable-hashcheck',
    '--log-level', 'INFO',
    '--log-file', $logFile
)

if ($DryRun) {
    $rcloneArgs += '--dry-run'
}

Write-Host "Download source:      $RemotePath"
Write-Host "Download destination: $LocalPath"
Write-Host "Mode:                 $operation"
Write-Host "Dry run:              $($DryRun.IsPresent)"
Write-Host "Log file:             $logFile"

Invoke-RcloneChecked -Arguments $rcloneArgs
