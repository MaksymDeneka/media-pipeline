[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,

    [ValidateRange(2, 31)]
    [int]$JpegQuality = 12,

    [ValidateRange(1, 32)]
    [int]$ThrottleLimit = 6,

    [string]$FFmpegPath = "ffmpeg"
)

$ErrorActionPreference = "Stop"

$resolvedOutputDirectory = (Resolve-Path -LiteralPath $OutputDirectory).Path
if (-not (Test-Path -LiteralPath $resolvedOutputDirectory -PathType Container)) {
    throw "Output directory does not exist: $OutputDirectory"
}

$pngFiles = @(
    Get-ChildItem -LiteralPath $resolvedOutputDirectory -File |
        Where-Object { $_.Extension -ieq ".png" }
)

if ($pngFiles.Count -eq 0) {
    Write-Host "No PNG outputs need normalization in: $resolvedOutputDirectory"
    exit 0
}

Write-Host "Normalizing $($pngFiles.Count) PNG output(s) in: $resolvedOutputDirectory"

$results = $pngFiles | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
    $sourcePath = $_.FullName
    $directory = $_.DirectoryName
    $baseName = $_.BaseName
    $targetPath = Join-Path $directory ($baseName + ".JPG")

    if (Test-Path -LiteralPath $targetPath) {
        $suffix = 1
        do {
            $targetPath = Join-Path $directory ("{0}-from-png-{1}.JPG" -f $baseName, $suffix)
            $suffix++
        } while (Test-Path -LiteralPath $targetPath)
    }

    $temporaryPath = Join-Path $directory (
        ".{0}.{1}.converting.jpg" -f $baseName, [guid]::NewGuid().ToString("N")
    )

    try {
        & $using:FFmpegPath `
            -y `
            -hide_banner `
            -loglevel error `
            -i $sourcePath `
            -frames:v 1 `
            -map_metadata -1 `
            -q:v $using:JpegQuality `
            $temporaryPath

        if ($LASTEXITCODE -ne 0) {
            throw "FFmpeg exited with code $LASTEXITCODE"
        }
        if (-not (Test-Path -LiteralPath $temporaryPath -PathType Leaf)) {
            throw "FFmpeg did not create the replacement"
        }
        if ((Get-Item -LiteralPath $temporaryPath).Length -le 0) {
            throw "FFmpeg created an empty replacement"
        }

        Move-Item -LiteralPath $temporaryPath -Destination $targetPath
        Remove-Item -LiteralPath $sourcePath

        [pscustomobject]@{
            Status = "Converted"
            SourceBytes = $_.Length
            TargetBytes = (Get-Item -LiteralPath $targetPath).Length
            Source = $sourcePath
            Target = $targetPath
            Error = $null
        }
    }
    catch {
        if (Test-Path -LiteralPath $temporaryPath) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        }

        [pscustomobject]@{
            Status = "Failed"
            SourceBytes = $_.Length
            TargetBytes = 0
            Source = $sourcePath
            Target = $targetPath
            Error = $_.Exception.Message
        }
    }
}

$converted = @($results | Where-Object { $_.Status -eq "Converted" })
$failed = @($results | Where-Object { $_.Status -eq "Failed" })
$sourceBytes = ($converted | Measure-Object SourceBytes -Sum).Sum
$targetBytes = ($converted | Measure-Object TargetBytes -Sum).Sum

Write-Host (
    "Converted: {0}; failed: {1}; size: {2:N1} MB -> {3:N1} MB" -f
    $converted.Count,
    $failed.Count,
    ($sourceBytes / 1MB),
    ($targetBytes / 1MB)
)

if ($failed.Count -gt 0) {
    $failed | Select-Object Source, Error | Format-Table -AutoSize
    exit 1
}
