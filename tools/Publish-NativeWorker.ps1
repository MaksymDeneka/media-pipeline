[CmdletBinding()]
param(
    [ValidateSet('win-x64', 'osx-arm64')]
    [string]$Runtime = 'win-x64',

    [string]$OutputRoot,

    [string]$DotnetPath
)

$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $OutputRoot) {
    $OutputRoot = Join-Path $repoRoot 'artifacts\native-worker'
}
if (-not $DotnetPath) {
    $command = Get-Command dotnet -ErrorAction SilentlyContinue
    $DotnetPath = if ($command) { $command.Source } else { 'C:\Program Files\dotnet\dotnet.exe' }
}
if (-not (Test-Path -LiteralPath $DotnetPath)) {
    throw "dotnet not found: $DotnetPath"
}

$project = Join-Path $repoRoot 'src\MediaPipeline.Worker\MediaPipeline.Worker.csproj'
$output = Join-Path $OutputRoot $Runtime
New-Item -ItemType Directory -Path $output -Force | Out-Null

& $DotnetPath publish $project `
    -c Release `
    -r $Runtime `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:PublishTrimmed=false `
    -p:DebugType=None `
    -p:DebugSymbols=false `
    -o $output `
    --nologo
if ($LASTEXITCODE -ne 0) {
    throw "Publishing the $Runtime worker failed."
}

$executable = if ($Runtime -eq 'win-x64') {
    Join-Path $output 'media-pipeline-worker.exe'
}
else {
    Join-Path $output 'media-pipeline-worker'
}
if (-not (Test-Path -LiteralPath $executable)) {
    throw "Published executable not found: $executable"
}

$hash = (Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash
[ordered]@{
    runtime = $Runtime
    executable = $executable
    bytes = (Get-Item -LiteralPath $executable).Length
    sha256 = $hash
    builtAtUtc = (Get-Date).ToUniversalTime().ToString('o')
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'release.json') -Encoding UTF8

Write-Host "Published $Runtime worker: $executable"
Write-Host "SHA-256: $hash"
