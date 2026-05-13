param(
    [switch]$CheckOnly
)

$ErrorActionPreference = "Stop"

# Editable settings
$PipelineRoot = "D:\MediaPipeline"
$InputDir = Join-Path $PipelineRoot "input"
$OutputDir = Join-Path $PipelineRoot "output"
$OriginalDir = Join-Path $PipelineRoot "original"
$FailedDir = Join-Path $PipelineRoot "failed"
$LogsDir = Join-Path $PipelineRoot "logs"

$CopiesPerFile = 3
$MinTrimMs = 50
$MaxTrimMs = 950
$Crf = 24
$Preset = "medium"
$AudioBitrate = "128k"
$MaxWidth = 1080
$StableSeconds = 3
$TimeoutSeconds = 600
$PollSeconds = 2

$VideoExtensions = @(".mp4", ".mov", ".mkv", ".webm", ".avi")
$ImageExtensions = @(".jpg", ".jpeg", ".png", ".webp", ".heic")
$TempExtensions = @(".crdownload", ".tmp", ".part", ".download")

$script:ProcessingPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:FFmpegPath = $null
$script:FFprobePath = $null
$script:ExifToolPath = $null
$script:InstanceMutex = $null

function Initialize-Folders {
    foreach ($directory in @($InputDir, $OutputDir, $OriginalDir, $FailedDir, $LogsDir)) {
        if (-not (Test-Path -LiteralPath $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    Write-Host $line

    try {
        $logFile = Join-Path $LogsDir ("media-pipeline-{0}.log" -f (Get-Date -Format "yyyyMMdd"))
        Add-Content -LiteralPath $logFile -Value $line -Encoding UTF8
    }
    catch {
        Write-Host "[$timestamp] [ERROR] Failed to write log file: $($_.Exception.Message)"
    }
}

function Resolve-RequiredTool {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $command = Get-Command $Name -ErrorAction SilentlyContinue
    if ($command) {
        return $command.Source
    }

    $fallbackPaths = @{
        ffmpeg = @(
            "C:\Tools\ffmpeg\bin\ffmpeg.exe"
        )
        ffprobe = @(
            "C:\Tools\ffmpeg\bin\ffprobe.exe"
        )
        exiftool = @(
            "C:\Tools\exiftool\exiftool.exe"
        )
    }

    foreach ($path in $fallbackPaths[$Name]) {
        if (Test-Path -LiteralPath $path) {
            return $path
        }
    }

    throw "Required tool '$Name' was not found in PATH or the default C:\Tools location. Install it and make sure '$Name' can be run from a new PowerShell window."
}

function Test-ExternalTools {
    $script:FFmpegPath = Resolve-RequiredTool "ffmpeg"
    $script:FFprobePath = Resolve-RequiredTool "ffprobe"
    $script:ExifToolPath = Resolve-RequiredTool "exiftool"

    Write-Log "Found ffmpeg: $script:FFmpegPath"
    Write-Log "Found ffprobe: $script:FFprobePath"
    Write-Log "Found exiftool: $script:ExifToolPath"
}

function Invoke-ExternalTool {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Command,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    $output = & $Command @Arguments 2>&1
    $exitCode = $LASTEXITCODE

    if ($exitCode -ne 0) {
        $outputText = ($output | Out-String).Trim()
        if ([string]::IsNullOrWhiteSpace($outputText)) {
            $outputText = "No command output."
        }

        throw "Command failed with exit code ${exitCode}: $Command $($Arguments -join ' ')`n$outputText"
    }

    return $output
}

function Test-IsTemporaryDownload {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    return $TempExtensions -contains $extension
}

function Test-IsSupportedMedia {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    return (($VideoExtensions -contains $extension) -or ($ImageExtensions -contains $extension))
}

function Test-IsVideo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    return ($VideoExtensions -contains $extension)
}

function Test-FileUnlocked {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::None)
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($stream) {
            $stream.Close()
            $stream.Dispose()
        }
    }
}

function Wait-FileReady {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Write-Log "Waiting for file ready: $Path"

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $lastSize = -1
    $stableSince = $null

    while ((Get-Date) -lt $deadline) {
        if (-not (Test-Path -LiteralPath $Path)) {
            throw "File disappeared before it was ready: $Path"
        }

        if (Test-IsTemporaryDownload $Path) {
            Start-Sleep -Seconds 1
            continue
        }

        $file = Get-Item -LiteralPath $Path
        $currentSize = $file.Length

        if ($currentSize -eq $lastSize -and $currentSize -gt 0) {
            if (-not $stableSince) {
                $stableSince = Get-Date
            }

            $stableFor = ((Get-Date) - $stableSince).TotalSeconds
            if ($stableFor -ge $StableSeconds -and (Test-FileUnlocked $Path)) {
                Write-Log "File is ready: $Path"
                return
            }
        }
        else {
            $lastSize = $currentSize
            $stableSince = $null
        }

        Start-Sleep -Seconds 1
    }

    throw "Timed out after $TimeoutSeconds seconds waiting for file to finish downloading: $Path"
}

function New-RandomToken {
    param(
        [int]$ByteCount = 8
    )

    $bytes = New-Object byte[] $ByteCount
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $rng.GetBytes($bytes)
    }
    finally {
        $rng.Dispose()
    }

    return (($bytes | ForEach-Object { $_.ToString("x2") }) -join "")
}

function New-RandomOutputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Extension
    )

    do {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $token = New-RandomToken
        $fileName = "media_{0}_{1}{2}" -f $timestamp, $token, $Extension.ToLowerInvariant()
        $path = Join-Path $OutputDir $fileName
    } while (Test-Path -LiteralPath $path)

    return $path
}

function Get-UniqueDestinationPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,

        [Parameter(Mandatory = $true)]
        [string]$OriginalFileName
    )

    $destination = Join-Path $Directory $OriginalFileName
    if (-not (Test-Path -LiteralPath $destination)) {
        return $destination
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OriginalFileName)
    $extension = [System.IO.Path]::GetExtension($OriginalFileName)

    do {
        $suffix = "{0}_{1}" -f (Get-Date -Format "yyyyMMdd_HHmmss"), (New-RandomToken 4)
        $fileName = "{0}_{1}{2}" -f $baseName, $suffix, $extension
        $destination = Join-Path $Directory $fileName
    } while (Test-Path -LiteralPath $destination)

    return $destination
}

function Move-InputFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$DestinationDirectory
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        Write-Log "Input file is no longer present, cannot move: $Path" "WARN"
        return
    }

    $destination = Get-UniqueDestinationPath -Directory $DestinationDirectory -OriginalFileName ([System.IO.Path]::GetFileName($Path))
    Move-Item -LiteralPath $Path -Destination $destination -Force
    Write-Log "Moved input file to: $destination"
}

function Remove-GeneratedOutputs {
    param(
        [AllowEmptyCollection()]
        [string[]]$Paths
    )

    if (-not $Paths -or $Paths.Count -eq 0) {
        return
    }

    foreach ($path in $Paths) {
        if ([string]::IsNullOrWhiteSpace($path)) {
            continue
        }

        try {
            if (Test-Path -LiteralPath $path) {
                Remove-Item -LiteralPath $path -Force
                Write-Log "Removed incomplete output after failure: $path" "WARN"
            }
        }
        catch {
            Write-Log "Could not remove incomplete output '$path': $($_.Exception.Message)" "WARN"
        }
    }
}

function Get-VideoDurationSeconds {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $arguments = @(
        "-v", "error",
        "-show_entries", "format=duration",
        "-of", "default=noprint_wrappers=1:nokey=1",
        $Path
    )

    $output = Invoke-ExternalTool -Command $script:FFprobePath -Arguments $arguments
    $durationText = (($output | Out-String).Trim() -split "\s+")[0]
    $duration = 0.0
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $parsed = [double]::TryParse($durationText, [System.Globalization.NumberStyles]::Float, $culture, [ref]$duration)

    if (-not $parsed -or $duration -le 0) {
        throw "Unable to read a valid duration from ffprobe for: $Path"
    }

    return $duration
}

function Get-TrimRange {
    param(
        [Parameter(Mandatory = $true)]
        [double]$DurationSeconds
    )

    $durationMs = [int][Math]::Floor($DurationSeconds * 1000)

    if ($durationMs -lt 500) {
        return [pscustomobject]@{
            CanTrim = $false
            MinMs = 0
            MaxMs = 0
            Reason = "video is shorter than 500 ms"
        }
    }

    if ($durationMs -lt 2000) {
        $safeMax = [int][Math]::Min(100, [Math]::Floor($durationMs * 0.10))
        if ($safeMax -lt 10) {
            return [pscustomobject]@{
                CanTrim = $false
                MinMs = 0
                MaxMs = 0
                Reason = "video is too short for safe trimming"
            }
        }

        return [pscustomobject]@{
            CanTrim = $true
            MinMs = 10
            MaxMs = $safeMax
            Reason = "short video safety range"
        }
    }

    $safeConfiguredMax = [int][Math]::Min($MaxTrimMs, $durationMs - 1000)
    if ($safeConfiguredMax -lt $MinTrimMs) {
        return [pscustomobject]@{
            CanTrim = $false
            MinMs = 0
            MaxMs = 0
            Reason = "configured trim range would make output too short"
        }
    }

    return [pscustomobject]@{
        CanTrim = $true
        MinMs = $MinTrimMs
        MaxMs = $safeConfiguredMax
        Reason = "configured trim range"
    }
}

function New-TrimMilliseconds {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Range,

        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[int]]$UsedValues
    )

    if (-not $Range.CanTrim) {
        return 0
    }

    $rangeSize = ($Range.MaxMs - $Range.MinMs) + 1
    $mustBeUnique = $rangeSize -ge $CopiesPerFile
    $attempts = 0

    do {
        $value = Get-Random -Minimum $Range.MinMs -Maximum ($Range.MaxMs + 1)
        $attempts++
    } while ($mustBeUnique -and $UsedValues.Contains($value) -and $attempts -lt 50)

    [void]$UsedValues.Add($value)
    return $value
}

function Clear-Metadata {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    Invoke-ExternalTool -Command $script:ExifToolPath -Arguments @("-all=", "-overwrite_original", $Path) | Out-Null
}

function Convert-VideoVariant {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [int]$VariantNumber,

        [Parameter(Mandatory = $true)]
        [double]$DurationSeconds,

        [Parameter(Mandatory = $true)]
        [int]$TrimMs
    )

    $outputPath = New-RandomOutputPath -Extension ".mp4"
    $trimSeconds = $TrimMs / 1000.0
    $targetDuration = [Math]::Max(0.1, $DurationSeconds - $trimSeconds)
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $targetDurationText = $targetDuration.ToString("0.###", $culture)
    $scaleFilter = "scale='trunc(min($MaxWidth,iw)/2)*2':-2"

    Write-Log "Video variant $VariantNumber trim: ${TrimMs}ms, target duration: ${targetDurationText}s"

    $arguments = @(
        "-y",
        "-hide_banner",
        "-loglevel", "error",
        "-i", $InputPath,
        "-t", $targetDurationText,
        "-map", "0:v:0",
        "-map", "0:a?",
        "-c:v", "libx264",
        "-crf", [string]$Crf,
        "-preset", $Preset,
        "-vf", $scaleFilter,
        "-c:a", "aac",
        "-b:a", $AudioBitrate,
        "-movflags", "+faststart",
        "-map_metadata", "-1",
        $outputPath
    )

    Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
    Clear-Metadata -Path $outputPath
    Write-Log "Created video output: $outputPath"

    return $outputPath
}

function Convert-ImageVariant {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [int]$VariantNumber
    )

    $extension = [System.IO.Path]::GetExtension($InputPath).ToLowerInvariant()
    $outputPath = New-RandomOutputPath -Extension $extension

    Copy-Item -LiteralPath $InputPath -Destination $outputPath -Force
    Clear-Metadata -Path $outputPath
    Write-Log "Created image output variant ${VariantNumber}: $outputPath"

    return $outputPath
}

function Process-VideoFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $createdOutputs = New-Object System.Collections.Generic.List[string]

    try {
        $duration = Get-VideoDurationSeconds -Path $Path
        $durationText = $duration.ToString("0.###", [System.Globalization.CultureInfo]::InvariantCulture)
        Write-Log "Video duration: ${durationText}s"

        $range = Get-TrimRange -DurationSeconds $duration
        if ($range.CanTrim) {
            Write-Log "Using trim range $($range.MinMs)-$($range.MaxMs) ms ($($range.Reason))"
        }
        else {
            Write-Log "Skipping duration trimming: $($range.Reason)" "WARN"
        }

        $usedTrimValues = [System.Collections.Generic.HashSet[int]]::new()

        for ($variant = 1; $variant -le $CopiesPerFile; $variant++) {
            $trimMs = New-TrimMilliseconds -Range $range -UsedValues $usedTrimValues
            $outputPath = Convert-VideoVariant -InputPath $Path -VariantNumber $variant -DurationSeconds $duration -TrimMs $trimMs
            $createdOutputs.Add($outputPath)
        }

        return $createdOutputs.ToArray()
    }
    catch {
        Remove-GeneratedOutputs -Paths $createdOutputs.ToArray()
        throw
    }
}

function Process-ImageFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $createdOutputs = New-Object System.Collections.Generic.List[string]

    try {
        for ($variant = 1; $variant -le $CopiesPerFile; $variant++) {
            $outputPath = Convert-ImageVariant -InputPath $Path -VariantNumber $variant
            $createdOutputs.Add($outputPath)
        }

        return $createdOutputs.ToArray()
    }
    catch {
        Remove-GeneratedOutputs -Paths $createdOutputs.ToArray()
        throw
    }
}

function Process-MediaFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }

    Write-Log "Detected file: $Path"
    Wait-FileReady -Path $Path

    if (-not (Test-IsSupportedMedia $Path)) {
        Write-Log "Skipping unsupported file: $Path" "WARN"
        return
    }

    Write-Log "Started processing: $Path"

    if (Test-IsVideo $Path) {
        [void](Process-VideoFile -Path $Path)
    }
    else {
        [void](Process-ImageFile -Path $Path)
    }

    Move-InputFile -Path $Path -DestinationDirectory $OriginalDir
    Write-Log "Successfully processed: $Path"
}

function Process-OneSafely {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)

    if (-not $script:ProcessingPaths.Add($fullPath)) {
        return
    }

    try {
        Process-MediaFile -Path $fullPath
    }
    catch {
        Write-Log "Failed processing '$fullPath': $($_.Exception.Message)" "ERROR"
        try {
            Move-InputFile -Path $fullPath -DestinationDirectory $FailedDir
        }
        catch {
            Write-Log "Could not move failed file '$fullPath': $($_.Exception.Message)" "ERROR"
        }
    }
    finally {
        [void]$script:ProcessingPaths.Remove($fullPath)
    }
}

function Get-CandidateInputFiles {
    if (-not (Test-Path -LiteralPath $InputDir)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $InputDir -File | Where-Object {
        (-not (Test-IsTemporaryDownload $_.FullName)) -and (Test-IsSupportedMedia $_.FullName)
    } | Sort-Object LastWriteTime, FullName)
}

function Start-PollingWatcher {
    Write-Log "Watcher started."
    Write-Log "Input: $InputDir"
    Write-Log "Output: $OutputDir"
    Write-Log "Original archive: $OriginalDir"
    Write-Log "Failed: $FailedDir"
    Write-Log "Polling every $PollSeconds seconds."

    while ($true) {
        try {
            $files = Get-CandidateInputFiles
            foreach ($file in $files) {
                Process-OneSafely -Path $file.FullName
            }
        }
        catch {
            Write-Log "Watcher loop error: $($_.Exception.Message)" "ERROR"
        }

        Start-Sleep -Seconds $PollSeconds
    }
}

try {
    $createdNew = $false
    $script:InstanceMutex = New-Object System.Threading.Mutex($true, "Global\MediaPipelineWatcher", [ref]$createdNew)
    if (-not $createdNew) {
        Initialize-Folders
        Write-Log "Another watcher instance is already running. Exiting this duplicate process." "WARN"
        exit 0
    }

    Initialize-Folders
    Test-ExternalTools

    if ($CheckOnly) {
        Write-Log "Startup check completed successfully."
        exit 0
    }

    Start-PollingWatcher
}
catch {
    try {
        Write-Log $_.Exception.Message "ERROR"
    }
    catch {
        Write-Host "[ERROR] $($_.Exception.Message)"
    }

    exit 1
}
finally {
    if ($script:InstanceMutex) {
        try {
            $script:InstanceMutex.ReleaseMutex()
            $script:InstanceMutex.Dispose()
        }
        catch {
        }
    }
}
