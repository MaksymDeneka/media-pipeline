[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDirectory,

    [ValidateRange(1, 10000)]
    [int]$CopiesPerFile = 200,

    [ValidateRange(2, 31)]
    [int]$ConvertedJpegQuality = 12,

    [ValidateRange(2, 31)]
    [int]$NativeJpegQuality = 4,

    [ValidateRange(0, 100)]
    [int]$WebpQuality = 92,

    [string]$FFmpegPath = "ffmpeg",

    [string]$FFprobePath = "ffprobe"
)

$ErrorActionPreference = "Stop"

$resolvedSourceDirectory = (Resolve-Path -LiteralPath $SourceDirectory).Path
$sourceFiles = @(
    Get-ChildItem -LiteralPath $resolvedSourceDirectory -File -Recurse |
        Where-Object { $_.Extension -match "^\.(jpg|jpeg|png|webp|heic|heif)$" }
)

if ($sourceFiles.Count -eq 0) {
    throw "No supported images found in: $resolvedSourceDirectory"
}

$testDirectory = Join-Path ([System.IO.Path]::GetTempPath()) (
    "media-pipeline-estimate-" + [guid]::NewGuid().ToString("N")
)
New-Item -ItemType Directory -Path $testDirectory | Out-Null

$results = New-Object System.Collections.Generic.List[object]

try {
    $index = 0
    foreach ($sourceFile in $sourceFiles) {
        $index++
        $probeJson = & $FFprobePath `
            -v error `
            -select_streams v:0 `
            -show_entries stream=width,height `
            -of json `
            -- `
            $sourceFile.FullName
        if ($LASTEXITCODE -ne 0) {
            throw "FFprobe failed for: $($sourceFile.FullName)"
        }

        $probe = $probeJson | ConvertFrom-Json
        $width = [int]$probe.streams[0].width
        $height = [int]$probe.streams[0].height
        $sourceExtension = $sourceFile.Extension.ToLowerInvariant()

        if ($sourceExtension -in @(".png", ".heic", ".heif")) {
            $outputExtension = ".jpg"
            $quality = $ConvertedJpegQuality
        }
        elseif ($sourceExtension -in @(".jpg", ".jpeg")) {
            $outputExtension = $sourceExtension
            $quality = $NativeJpegQuality
        }
        else {
            $outputExtension = ".webp"
            $quality = $WebpQuality
        }

        $testOutput = Join-Path $testDirectory (
            ("{0:D5}" -f $index) + $outputExtension
        )
        $arguments = @(
            "-y",
            "-hide_banner",
            "-loglevel", "error",
            "-i", $sourceFile.FullName,
            "-frames:v", "1",
            "-map_metadata", "-1"
        )

        if ($width -ge 200 -and $height -ge 200) {
            # The actual pipeline chooses 5-20 permille randomly. Twelve is a
            # representative midpoint for estimating encoded size.
            $cropPixelsX = [Math]::Max(1, [int][Math]::Floor($width * 12 / 1000))
            $cropPixelsY = [Math]::Max(1, [int][Math]::Floor($height * 12 / 1000))
            $cropWidth = [Math]::Max(1, $width - ($cropPixelsX * 2))
            $cropHeight = [Math]::Max(1, $height - ($cropPixelsY * 2))
            $filter = "crop=${cropWidth}:${cropHeight}:${cropPixelsX}:${cropPixelsY},scale=${width}:${height}"
            $arguments += @("-vf", $filter)
        }

        if ($outputExtension -in @(".jpg", ".jpeg")) {
            $arguments += @("-q:v", [string]$quality)
        }
        else {
            $arguments += @("-quality", [string]$quality)
        }
        $arguments += @($testOutput)

        & $FFmpegPath @arguments
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $testOutput -PathType Leaf)) {
            throw "FFmpeg failed for: $($sourceFile.FullName)"
        }

        $testFile = Get-Item -LiteralPath $testOutput
        $results.Add([pscustomobject]@{
            Name = $sourceFile.Name
            SourceExtension = $sourceExtension
            Width = $width
            Height = $height
            SourceBytes = $sourceFile.Length
            EstimatedOutputBytes = $testFile.Length
            Copies = $CopiesPerFile
            ProjectedBytes = $testFile.Length * $CopiesPerFile
            GrowthRatio = [Math]::Round(
                $testFile.Length / [Math]::Max(1, $sourceFile.Length),
                2
            )
        })
    }
}
finally {
    foreach ($testFile in @(Get-ChildItem -LiteralPath $testDirectory -File -ErrorAction SilentlyContinue)) {
        Remove-Item -LiteralPath $testFile.FullName -Force -ErrorAction SilentlyContinue
    }
    Remove-Item -LiteralPath $testDirectory -Force -ErrorAction SilentlyContinue
}

Write-Output "PROJECTED SUMMARY BY SOURCE FORMAT"
$results |
    Group-Object SourceExtension |
    ForEach-Object {
        $sortedSizes = @($_.Group.EstimatedOutputBytes | Sort-Object)
        [pscustomobject]@{
            Format = $_.Name
            Sources = $_.Count
            Outputs = $_.Count * $CopiesPerFile
            SourceMB = [Math]::Round(
                (($_.Group | Measure-Object SourceBytes -Sum).Sum) / 1MB,
                1
            )
            OneSetMB = [Math]::Round(
                (($_.Group | Measure-Object EstimatedOutputBytes -Sum).Sum) / 1MB,
                1
            )
            ProjectedGB = [Math]::Round(
                (($_.Group | Measure-Object ProjectedBytes -Sum).Sum) / 1GB,
                2
            )
            MedianOutputKB = [Math]::Round(
                $sortedSizes[[Math]::Floor($sortedSizes.Count / 2)] / 1KB
            )
            MaxOutputMB = [Math]::Round(
                (($_.Group | Measure-Object EstimatedOutputBytes -Maximum).Maximum) / 1MB,
                2
            )
        }
    } |
    Format-Table -AutoSize

Write-Output "TOTAL"
$totalProjectedBytes = ($results | Measure-Object ProjectedBytes -Sum).Sum
[pscustomobject]@{
    Sources = $results.Count
    Outputs = $results.Count * $CopiesPerFile
    SourceMB = [Math]::Round(
        (($results | Measure-Object SourceBytes -Sum).Sum) / 1MB,
        1
    )
    OneSetMB = [Math]::Round(
        (($results | Measure-Object EstimatedOutputBytes -Sum).Sum) / 1MB,
        1
    )
    ProjectedGB = [Math]::Round($totalProjectedBytes / 1GB, 2)
    ProjectedWith20PercentSafetyGB = [Math]::Round(
        ($totalProjectedBytes * 1.2) / 1GB,
        2
    )
} | Format-List

Write-Output "LARGEST PROJECTED CONTRIBUTORS"
$results |
    Sort-Object ProjectedBytes -Descending |
    Select-Object -First 10 `
        Name,
        SourceExtension,
        @{ Name = "Dimensions"; Expression = { "$($_.Width)x$($_.Height)" } },
        @{ Name = "SourceMB"; Expression = { [Math]::Round($_.SourceBytes / 1MB, 2) } },
        @{ Name = "OutputEachMB"; Expression = { [Math]::Round($_.EstimatedOutputBytes / 1MB, 2) } },
        @{ Name = "GrowthX"; Expression = { $_.GrowthRatio } },
        @{ Name = "CopiesTotalGB"; Expression = { [Math]::Round($_.ProjectedBytes / 1GB, 2) } } |
    Format-Table -AutoSize -Wrap
