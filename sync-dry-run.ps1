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
    [switch]$Mirror
)

$arguments = @{
    Lane = $Lane
    LocalPipelineRoot = $LocalPipelineRoot
    RemotePipelineRoot = $RemotePipelineRoot
    RemoteName = $RemoteName
    Transfers = $Transfers
    Checkers = $Checkers
    DryRun = $true
}

if (-not [string]::IsNullOrWhiteSpace($LocalPath)) {
    $arguments.LocalPath = $LocalPath
}

if (-not [string]::IsNullOrWhiteSpace($RemotePath)) {
    $arguments.RemotePath = $RemotePath
}

if ($Mirror) {
    $arguments.Mirror = $true
}

& (Join-Path $PSScriptRoot 'sync-upload.ps1') @arguments
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}
