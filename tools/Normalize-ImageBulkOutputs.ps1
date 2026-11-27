[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$OutputDirectory,

    [ValidateRange(2, 31)]
    [int]$JpegQuality = 12,

    [ValidateRange(1, 32)]
    [int]$ThrottleLimit = 6,

    [string]$FFmpegPath = "ffmpeg",

    [string]$ExifToolPath = "exiftool"
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

    $conversionId = [guid]::NewGuid().ToString("N")
    $stagingDirectory = Join-Path ([System.IO.Path]::GetTempPath()) "media-pipeline-jpeg-$conversionId"
    $encodedPath = Join-Path $stagingDirectory "image.jpg"
    $temporaryPath = Join-Path $directory ".$conversionId.converting"

    try {
        [void][System.IO.Directory]::CreateDirectory($stagingDirectory)
        & $using:FFmpegPath `
            -y `
            -hide_banner `
            -loglevel error `
            -i $sourcePath `
            -frames:v 1 `
            -map_metadata -1 `
            -q:v $using:JpegQuality `
            $encodedPath

        if ($LASTEXITCODE -ne 0) {
            throw "FFmpeg exited with code $LASTEXITCODE"
        }
        if (-not (Test-Path -LiteralPath $encodedPath -PathType Leaf)) {
            throw "FFmpeg did not create the replacement"
        }
        $exifOutput = & $using:ExifToolPath "-all=" "-o" $temporaryPath $encodedPath 2>&1
        if ($LASTEXITCODE -ne 0) {
            throw "ExifTool failed: $($exifOutput | Out-String)"
        }
        if ([System.IO.FileInfo]::new($temporaryPath).Length -le 0) {
            throw "The cleaned JPEG is empty"
        }

        [System.IO.File]::Move($temporaryPath, $targetPath)
        Remove-Item -LiteralPath $sourcePath -ErrorAction Stop

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
        [pscustomobject]@{
            Status = "Failed"
            SourceBytes = $_.Length
            TargetBytes = 0
            Source = $sourcePath
            Target = $targetPath
            Error = $_.Exception.Message
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path -LiteralPath $stagingDirectory) {
            Remove-Item -LiteralPath $stagingDirectory -Recurse -Force -ErrorAction SilentlyContinue
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
