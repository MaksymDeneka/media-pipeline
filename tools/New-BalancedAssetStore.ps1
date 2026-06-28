param(
    [string]$SourceDirectory = "D:\Users\admin\Pictures\stores\heatup-phase2-store",
    [string]$OutputRoot = "D:\MediaPipeline\assetstore\LC\output",
    [int]$ProfileCount = 216,
    [int]$ImagesPerProfile = 7,
    [int]$Concurrency = 8,
    [string]$BatchName = ""
)

$ErrorActionPreference = "Stop"

$ImageExtensions = @(".jpg", ".jpeg", ".png", ".webp", ".heic")
$FirstWords = @(
    "amber", "autumn", "bright", "calm", "cedar", "clear", "coastal", "daily",
    "early", "fresh", "golden", "harbor", "light", "meadow", "modern", "natural",
    "quiet", "simple", "soft", "spring", "steady", "summer", "urban", "warm"
)
$SecondWords = @(
    "album", "capture", "clip", "collection", "frame", "gallery", "image", "media",
    "memory", "moment", "photo", "picture", "post", "project", "scene", "shot",
    "snapshot", "story", "take", "update", "upload", "video", "view", "work"
)

function Resolve-RequiredCommand {
    param([Parameter(Mandatory = $true)][string]$Name)

    $command = Get-Command $Name -ErrorAction SilentlyContinue
    if (-not $command) {
        throw "Required command is not available on PATH: $Name"
    }

    return $command.Source
}

function New-RandomBatchName {
    $first = $FirstWords[(Get-Random -Minimum 0 -Maximum $FirstWords.Count)]
    $second = $SecondWords[(Get-Random -Minimum 0 -Maximum $SecondWords.Count)]
    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $suffix = [Guid]::NewGuid().ToString("n").Substring(0, 8)
    return "heatup-phase2-{0}-{1}-{2}-{3}" -f $first, $second, $stamp, $suffix
}

function Get-FamilyKey {
    param(
        [Parameter(Mandatory = $true)][string]$FileName,
        [Parameter(Mandatory = $true)]$UsedKeys
    )

    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $key = ($base -replace '[^A-Za-z0-9._-]', '_').Trim("_")
    if ([string]::IsNullOrWhiteSpace($key)) {
        $key = "image"
    }

    $candidate = $key
    $suffix = 2
    while (-not $UsedKeys.Add($candidate)) {
        $candidate = "{0}_{1}" -f $key, $suffix
        $suffix++
    }

    return $candidate
}

function Get-ImageDimensions {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$FfprobePath
    )

    $arguments = @(
        "-v", "error",
        "-select_streams", "v:0",
        "-show_entries", "stream=width,height",
        "-of", "csv=s=x:p=0",
        $Path
    )

    $output = & $FfprobePath @arguments 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "ffprobe failed for '$Path': $($output | Out-String)"
    }

    $dimensionText = (($output | Out-String).Trim() -split "\s+")[0]
    if ($dimensionText -notmatch "^(\d+)x(\d+)") {
        throw "Unable to read image dimensions from ffprobe for: $Path"
    }

    return [pscustomobject]@{
        Width = [int]$Matches[1]
        Height = [int]$Matches[2]
    }
}

function Get-OutputExtension {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    $extension = [System.IO.Path]::GetExtension($SourcePath).ToLowerInvariant()
    if ($extension -eq ".heic") {
        return ".jpg"
    }

    return $extension
}

function ConvertTo-RelativePath {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $rootUri = [Uri](([System.IO.Path]::GetFullPath($Root).TrimEnd('\') + '\'))
    $pathUri = [Uri]([System.IO.Path]::GetFullPath($Path))
    return [Uri]::UnescapeDataString($rootUri.MakeRelativeUri($pathUri).ToString()).Replace('/', '\')
}

$sourceRoot = [System.IO.Path]::GetFullPath($SourceDirectory)
$outputRootFull = [System.IO.Path]::GetFullPath($OutputRoot)
if (-not (Test-Path -LiteralPath $sourceRoot -PathType Container)) {
    throw "Source directory does not exist: $sourceRoot"
}

New-Item -ItemType Directory -Path $outputRootFull -Force | Out-Null

$ffmpegPath = Resolve-RequiredCommand "ffmpeg"
$ffprobePath = Resolve-RequiredCommand "ffprobe"
$exifToolPath = Resolve-RequiredCommand "exiftool"

$files = @(Get-ChildItem -LiteralPath $sourceRoot -File -Recurse | Where-Object {
    $ImageExtensions -contains $_.Extension.ToLowerInvariant()
} | Sort-Object FullName)

if ($files.Count -eq 0) {
    throw "No supported image files found in: $sourceRoot"
}

$totalSlots = $ProfileCount * $ImagesPerProfile
if ($files.Count -gt $totalSlots) {
    throw "Source image count ($($files.Count)) is greater than available profile slots ($totalSlots)."
}

$baseUseCount = [int][Math]::Floor($totalSlots / $files.Count)
$extraUseCount = $totalSlots - ($baseUseCount * $files.Count)
$usedFamilyKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$extraSelection = @($files | Get-Random -Count $files.Count | Select-Object -First $extraUseCount)
$extraPathSet = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($file in $extraSelection) {
    [void]$extraPathSet.Add($file.FullName)
}

Write-Host "Reading dimensions for $($files.Count) source image(s)..."
$sources = New-Object System.Collections.Generic.List[object]
for ($i = 0; $i -lt $files.Count; $i++) {
    $file = $files[$i]
    $dimensions = Get-ImageDimensions -Path $file.FullName -FfprobePath $ffprobePath
    $targetUses = $baseUseCount
    if ($extraPathSet.Contains($file.FullName)) {
        $targetUses++
    }

    $sources.Add([pscustomobject]@{
        SourceId = $i + 1
        FullName = $file.FullName
        RelativePath = ConvertTo-RelativePath -Root $sourceRoot -Path $file.FullName
        Name = $file.Name
        FamilyKey = Get-FamilyKey -FileName $file.Name -UsedKeys $usedFamilyKeys
        Extension = $file.Extension.ToLowerInvariant()
        OutputExtension = Get-OutputExtension -SourcePath $file.FullName
        Width = $dimensions.Width
        Height = $dimensions.Height
        TargetUses = $targetUses
        RemainingUses = $targetUses
    })
}

Write-Host "Building balanced allocation: $ProfileCount profile(s), $ImagesPerProfile image(s) each, $totalSlots slot(s)."
$profiles = @()
for ($profileNumber = 1; $profileNumber -le $ProfileCount; $profileNumber++) {
    $profiles += [pscustomobject]@{
        ProfileNumber = $profileNumber
        ProfileName = "profile_{0:D3}" -f $profileNumber
        Items = New-Object System.Collections.Generic.List[object]
        UsedFamilies = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    }
}

foreach ($profile in $profiles) {
    for ($slot = 1; $slot -le $ImagesPerProfile; $slot++) {
        $available = @($sources | Where-Object {
            $_.RemainingUses -gt 0 -and -not $profile.UsedFamilies.Contains($_.FamilyKey)
        })
        if ($available.Count -eq 0) {
            throw "Unable to complete allocation for $($profile.ProfileName) slot $slot without duplicating a source."
        }

        $maxRemaining = ($available | Measure-Object -Property RemainingUses -Maximum).Maximum
        $candidate = @($available | Where-Object { $_.RemainingUses -eq $maxRemaining }) | Get-Random
        $useIndex = $candidate.TargetUses - $candidate.RemainingUses + 1
        $candidate.RemainingUses--
        [void]$profile.UsedFamilies.Add($candidate.FamilyKey)

        $profile.Items.Add([pscustomobject]@{
            ProfileNumber = $profile.ProfileNumber
            ProfileName = $profile.ProfileName
            SlotNumber = $slot
            SourceId = $candidate.SourceId
            SourcePath = $candidate.FullName
            SourceRelativePath = $candidate.RelativePath
            SourceOriginalName = $candidate.Name
            FamilyKey = $candidate.FamilyKey
            OutputExtension = $candidate.OutputExtension
            Width = $candidate.Width
            Height = $candidate.Height
            TargetUseCount = $candidate.TargetUses
            AssignmentUseIndex = $useIndex
        })
    }
}

$remainingTotal = ($sources | Measure-Object -Property RemainingUses -Sum).Sum
if ($remainingTotal -ne 0) {
    throw "Allocation ended with $remainingTotal remaining unassigned use(s)."
}

$duplicateProfile = $profiles | Where-Object {
    ($_.Items | Select-Object -ExpandProperty FamilyKey -Unique).Count -ne $_.Items.Count
} | Select-Object -First 1
if ($duplicateProfile) {
    throw "Allocation produced a duplicate source inside $($duplicateProfile.ProfileName)."
}

if ([string]::IsNullOrWhiteSpace($BatchName)) {
    $BatchName = New-RandomBatchName
}
$batchDirectory = Join-Path $outputRootFull $BatchName
if (Test-Path -LiteralPath $batchDirectory) {
    throw "Batch output directory already exists: $batchDirectory"
}
New-Item -ItemType Directory -Path $batchDirectory -Force | Out-Null

$tasks = New-Object System.Collections.Generic.List[object]
foreach ($profile in $profiles) {
    $profileDirectory = Join-Path $batchDirectory $profile.ProfileName
    New-Item -ItemType Directory -Path $profileDirectory -Force | Out-Null
    foreach ($item in $profile.Items) {
        $tasks.Add([pscustomobject]@{
            ProfileNumber = $item.ProfileNumber
            ProfileName = $item.ProfileName
            SlotNumber = $item.SlotNumber
            SourceId = $item.SourceId
            SourcePath = $item.SourcePath
            SourceRelativePath = $item.SourceRelativePath
            SourceOriginalName = $item.SourceOriginalName
            FamilyKey = $item.FamilyKey
            OutputExtension = $item.OutputExtension
            Width = $item.Width
            Height = $item.Height
            TargetUseCount = $item.TargetUseCount
            AssignmentUseIndex = $item.AssignmentUseIndex
            OutputDirectory = $profileDirectory
        })
    }
}

Write-Host "Generating $($tasks.Count) processed image file(s) in $batchDirectory ..."
$startedAt = Get-Date
$results = @($tasks | ForEach-Object -Parallel {
    $ErrorActionPreference = "Stop"
    $task = $_
    $outputName = "asset-{0}{1}" -f ([Guid]::NewGuid().ToString("n").Substring(0, 16)), $task.OutputExtension
    $outputPath = Join-Path $task.OutputDirectory $outputName
    $canCrop = ($task.Width -ge 40 -and $task.Height -ge 40)
    $ffmpegArgs = @(
        "-y",
        "-hide_banner",
        "-loglevel", "error",
        "-i", $task.SourcePath,
        "-frames:v", "1",
        "-map_metadata", "-1"
    )

    if ($canCrop) {
        $cropPermille = Get-Random -Minimum 5 -Maximum 21
        $cropPixelsX = [Math]::Max(1, [int][Math]::Floor($task.Width * $cropPermille / 1000))
        $cropPixelsY = [Math]::Max(1, [int][Math]::Floor($task.Height * $cropPermille / 1000))
        $cropWidth = [Math]::Max(1, $task.Width - ($cropPixelsX * 2))
        $cropHeight = [Math]::Max(1, $task.Height - ($cropPixelsY * 2))
        $offsetX = Get-Random -Minimum 0 -Maximum (($cropPixelsX * 2) + 1)
        $offsetY = Get-Random -Minimum 0 -Maximum (($cropPixelsY * 2) + 1)
        $filter = "crop=${cropWidth}:${cropHeight}:${offsetX}:${offsetY},scale=$($task.Width):$($task.Height)"
        $ffmpegArgs += @("-vf", $filter)
    }

    if ($task.OutputExtension -in @(".jpg", ".jpeg")) {
        $ffmpegArgs += @("-q:v", "2")
    }
    elseif ($task.OutputExtension -eq ".webp") {
        $ffmpegArgs += @("-quality", "92")
    }
    elseif ($task.OutputExtension -eq ".png") {
        $ffmpegArgs += @("-compression_level", "1")
    }

    $ffmpegArgs += @($outputPath)
    $ffmpegOutput = & $using:ffmpegPath @ffmpegArgs 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "ffmpeg failed for '$($task.SourcePath)': $($ffmpegOutput | Out-String)"
    }

    $exifOutput = & $using:exifToolPath "-all=" "-overwrite_original" $outputPath 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "exiftool failed for '$outputPath': $($exifOutput | Out-String)"
    }

    $fileInfo = Get-Item -LiteralPath $outputPath
    [pscustomobject]@{
        ProfileNumber = $task.ProfileNumber
        ProfileName = $task.ProfileName
        SlotNumber = $task.SlotNumber
        SourceId = $task.SourceId
        SourcePath = $task.SourcePath
        SourceRelativePath = $task.SourceRelativePath
        SourceOriginalName = $task.SourceOriginalName
        FamilyKey = $task.FamilyKey
        TargetUseCount = $task.TargetUseCount
        AssignmentUseIndex = $task.AssignmentUseIndex
        Width = $task.Width
        Height = $task.Height
        OutputPath = $outputPath
        OutputName = $outputName
        SizeBytes = [long]$fileInfo.Length
    }
} -ThrottleLimit $Concurrency)

if ($results.Count -ne $tasks.Count) {
    throw "Generated $($results.Count) output(s), expected $($tasks.Count)."
}

$generatedAt = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ", [System.Globalization.CultureInfo]::InvariantCulture)
$variants = @($results | Sort-Object ProfileNumber, SlotNumber | ForEach-Object {
    [ordered]@{
        familyKey = $_.FamilyKey
        variantKey = "{0}__{1}__slot_{2:D2}" -f $_.FamilyKey, $_.ProfileName, $_.SlotNumber
        path = "{0}/{1}" -f $_.ProfileName, $_.OutputName
        renditionSetKey = $_.ProfileName
        generationBatchKey = $BatchName
        sourceOriginalName = $_.SourceOriginalName
        sourceFamilyName = $_.FamilyKey
        sizeBytes = [long]$_.SizeBytes
        transformProfile = "asset_store_image_balanced_recrop"
        generatedAt = $generatedAt
        metadata = [ordered]@{
            sourceWidth = $_.Width
            sourceHeight = $_.Height
            sourceRelativePath = $_.SourceRelativePath
            profileNumber = $_.ProfileNumber
            slotNumber = $_.SlotNumber
            sourceTargetUseCount = $_.TargetUseCount
            sourceAssignmentUseIndex = $_.AssignmentUseIndex
        }
    }
})

$manifest = [ordered]@{
    schema = "heatup.assetStoreMediaManifest.v1"
    generatedAt = $generatedAt
    importRoot = "."
    variants = [object[]]$variants
}
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$manifestPath = Join-Path $batchDirectory "manifest.json"
[System.IO.File]::WriteAllText($manifestPath, ($manifest | ConvertTo-Json -Depth 12), $utf8NoBom)

$allocationCsvPath = Join-Path $batchDirectory "allocation.csv"
$results |
    Sort-Object ProfileNumber, SlotNumber |
    Select-Object ProfileName, ProfileNumber, SlotNumber, FamilyKey, SourceOriginalName, SourceRelativePath, TargetUseCount, AssignmentUseIndex, OutputName, SizeBytes |
    Export-Csv -LiteralPath $allocationCsvPath -NoTypeInformation -Encoding UTF8

$useCounts = @($results | Group-Object FamilyKey | ForEach-Object { $_.Count })
$summary = [ordered]@{
    generatedAt = $generatedAt
    sourceDirectory = $sourceRoot
    outputDirectory = $batchDirectory
    manifestPath = $manifestPath
    allocationCsvPath = $allocationCsvPath
    sourceImageCount = $files.Count
    profileCount = $ProfileCount
    imagesPerProfile = $ImagesPerProfile
    totalOutputImages = $results.Count
    baseUseCountPerSource = $baseUseCount
    extraUseSourceCount = $extraUseCount
    minUseCount = [int](($useCounts | Measure-Object -Minimum).Minimum)
    maxUseCount = [int](($useCounts | Measure-Object -Maximum).Maximum)
    sourceUseDistribution = [ordered]@{
        twoUses = [int](($useCounts | Where-Object { $_ -eq 2 } | Measure-Object).Count)
        threeUses = [int](($useCounts | Where-Object { $_ -eq 3 } | Measure-Object).Count)
    }
    elapsedSeconds = [Math]::Round(((Get-Date) - $startedAt).TotalSeconds, 3)
}
$summaryPath = Join-Path $batchDirectory "summary.json"
[System.IO.File]::WriteAllText($summaryPath, ($summary | ConvertTo-Json -Depth 8), $utf8NoBom)

Write-Host "Done."
Write-Host "Batch directory: $batchDirectory"
Write-Host "Manifest: $manifestPath"
Write-Host "Allocation CSV: $allocationCsvPath"
Write-Host "Summary: $summaryPath"
Write-Host "Use distribution: $($summary.sourceUseDistribution.twoUses) source(s) used twice, $($summary.sourceUseDistribution.threeUses) source(s) used three times."
