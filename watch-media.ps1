param(
    [switch]$CheckOnly,
    [switch]$RecompressLongOutputs
)

$ErrorActionPreference = "Stop"

# Editable settings
$PipelineRoot = "D:\MediaPipeline"
$InputDir = Join-Path $PipelineRoot "input"
$OutputDir = Join-Path $PipelineRoot "output"
$OriginalDir = Join-Path $PipelineRoot "original"
$FailedDir = Join-Path $PipelineRoot "failed"
$LogsDir = Join-Path $PipelineRoot "logs"
$RemuxRootDir = Join-Path $PipelineRoot "convert"
$RemuxInputDir = Join-Path $RemuxRootDir "input"
$RemuxOutputDir = Join-Path $RemuxRootDir "output"
$RemuxOriginalDir = Join-Path $RemuxRootDir "original"
$RemuxOriginalVideosDir = Join-Path $RemuxOriginalDir "videos"
$RemuxOriginalImagesDir = Join-Path $RemuxOriginalDir "images"
$RemuxFailedDir = Join-Path $RemuxRootDir "failed"
$LongRootDir = Join-Path $PipelineRoot "long"
$LongInputDir = Join-Path $LongRootDir "input"
$LongOutputDir = Join-Path $LongRootDir "output"
$LongOriginalDir = Join-Path $LongRootDir "original"
$LongFailedDir = Join-Path $LongRootDir "failed"
$LongWorkDir = Join-Path $LongRootDir "work"
$ImageBulkRootDir = Join-Path $PipelineRoot "images"
$ImageBulkInputDir = Join-Path $ImageBulkRootDir "input"
$ImageBulkOutputDir = Join-Path $ImageBulkRootDir "output"
$ImageBulkOriginalDir = Join-Path $ImageBulkRootDir "original"
$ImageBulkFailedDir = Join-Path $ImageBulkRootDir "failed"
$SetRootDir = Join-Path $PipelineRoot "sets"
$SetInputDir = Join-Path $SetRootDir "input"
$SetOutputDir = Join-Path $SetRootDir "output"
$SetOriginalDir = Join-Path $SetRootDir "original"
$SetFailedDir = Join-Path $SetRootDir "failed"

$DefaultPipelineMinCopiesPerFile = 7
$DefaultPipelineAlternatingCopiesPerFile = 8
$LongCopiesPerSegment = 3
$ImageBulkCopiesPerFile = 20
$SetCopiesPerFile = 10
$ImageBulkCropMinPermille = 5
$ImageBulkCropMaxPermille = 20
$MinTrimMs = 15
$MaxTrimMs = 95
$PreferNvenc = $true
$Crf = 24
$Preset = "medium"
$NvencPreset = "p4"
$NvencCq = 26
$LongNvencCq = 28
$LongNvencPrimaryMaxrateScale = 0.92
$AudioBitrate = "128k"
$MaxWidth = 1080
$StableSeconds = 3
$TimeoutSeconds = 600
$PollSeconds = 2
$LongSegmentTargetSeconds = 15
$LongSegmentMinSeconds = 11
$LongMaxOutputSizeMB = 8
$LongSizeCapFallbackMaxWidth = 720
$ArchiveEnabled = $true
$ArchiveAgeHours = 15
$ArchiveCheckIntervalMinutes = 30
$ArchiveRootDir = Join-Path $PipelineRoot "archive"
$ArchiveDefaultOutputDir = Join-Path $ArchiveRootDir "output"
$ArchiveImageBulkOutputDir = Join-Path $ArchiveRootDir "images"
$ArchiveRemuxOutputDir = Join-Path $ArchiveRootDir "convert"
$ArchiveLongOutputDir = Join-Path $ArchiveRootDir "long"
$ArchiveSetOutputDir = Join-Path $ArchiveRootDir "sets"

$VideoExtensions = @(".mp4", ".mov", ".mkv", ".webm", ".avi")
$ImageExtensions = @(".jpg", ".jpeg", ".png", ".webp", ".heic")
$TempExtensions = @(".crdownload", ".tmp", ".part", ".download")

# Convert pipeline: source formats that get rewritten into widely supported ones.
$RemuxVideoSourceExtensions = @(".mov")
$RemuxImageSourceExtensions = @(".heic")
$RemuxImageOutputExtension = ".jpg"

$script:ProcessingPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$script:FFmpegPath = $null
$script:FFprobePath = $null
$script:ExifToolPath = $null
$script:UseNvenc = $false
$script:InstanceMutex = $null
$script:LastArchiveCheck = $null
$script:DefaultPipelineEntryCount = 0

function Get-DefaultPipelineCopyCount {
    $script:DefaultPipelineEntryCount++
    if (($script:DefaultPipelineEntryCount % 2) -eq 1) {
        return $DefaultPipelineAlternatingCopiesPerFile
    }

    return $DefaultPipelineMinCopiesPerFile
}

function Initialize-Folders {
    foreach ($directory in @($InputDir, $OutputDir, $OriginalDir, $FailedDir, $LogsDir, $RemuxInputDir, $RemuxOutputDir, $RemuxOriginalDir, $RemuxOriginalVideosDir, $RemuxOriginalImagesDir, $RemuxFailedDir, $LongInputDir, $LongOutputDir, $LongOriginalDir, $LongFailedDir, $LongWorkDir, $ImageBulkInputDir, $ImageBulkOutputDir, $ImageBulkOriginalDir, $ImageBulkFailedDir, $SetInputDir, $SetOutputDir, $SetOriginalDir, $SetFailedDir, $ArchiveDefaultOutputDir, $ArchiveImageBulkOutputDir, $ArchiveRemuxOutputDir, $ArchiveLongOutputDir, $ArchiveSetOutputDir)) {
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

function Test-NvencEncoderAvailable {
    try {
        $output = & $script:FFmpegPath -hide_banner -encoders 2>&1 | Out-String
        return $output -match '\bh264_nvenc\b'
    }
    catch {
        return $false
    }
}

function Initialize-VideoEncoder {
    $script:UseNvenc = $false

    if ($PreferNvenc -and (Test-NvencEncoderAvailable)) {
        $script:UseNvenc = $true
        Write-Log "Video encoder: h264_nvenc (GPU, preset $NvencPreset, CQ $NvencCq, long CQ $LongNvencCq)"
        return
    }

    if ($PreferNvenc) {
        Write-Log "NVENC encoder not available in FFmpeg; falling back to libx264." "WARN"
    }

    Write-Log "Video encoder: libx264 (CPU, preset $Preset, CRF $Crf)"
}

function Get-VideoEncoderName {
    if ($script:UseNvenc) {
        return "h264_nvenc"
    }

    return "libx264"
}

function Get-VideoScaleFilter {
    param(
        [Parameter(Mandatory = $true)]
        [int]$MaxWidthValue
    )

    return "scale='trunc(min($MaxWidthValue,iw)/2)*2':-2"
}

function New-VideoEncoderArguments {
    param(
        [Parameter(Mandatory = $true)]
        [int]$QualityValue,

        [Parameter(Mandatory = $true)]
        [int]$MaxWidthValue,

        [int]$MaxVideoBitrateKbps = 0
    )

    if ($script:UseNvenc) {
        $arguments = @(
            "-c:v", "h264_nvenc",
            "-preset", $NvencPreset,
            "-tune", "hq",
            "-rc", "vbr",
            "-cq", [string]$QualityValue,
            "-b:v", "0",
            "-spatial_aq", "1",
            "-temporal_aq", "1",
            "-vf", (Get-VideoScaleFilter -MaxWidthValue $MaxWidthValue),
            "-pix_fmt", "yuv420p"
        )
    }
    else {
        $arguments = @(
            "-c:v", "libx264",
            "-crf", [string]$QualityValue,
            "-preset", $Preset,
            "-vf", (Get-VideoScaleFilter -MaxWidthValue $MaxWidthValue),
            "-pix_fmt", "yuv420p"
        )
    }

    if ($MaxVideoBitrateKbps -gt 0) {
        $arguments += @(
            "-maxrate", ("{0}k" -f $MaxVideoBitrateKbps),
            "-bufsize", ("{0}k" -f ($MaxVideoBitrateKbps * 2))
        )
    }

    return $arguments
}

function Get-LongPrimaryMaxVideoBitrateKbps {
    param(
        [Parameter(Mandatory = $true)]
        [double]$DurationSeconds
    )

    if (-not $script:UseNvenc -or $LongMaxOutputSizeMB -le 0 -or $DurationSeconds -le 0) {
        return 0
    }

    $targetKbps = Get-LongTargetVideoBitrateKbps -DurationSeconds $DurationSeconds -MaxSizeMegabytes $LongMaxOutputSizeMB
    return [int][Math]::Max(200, [Math]::Floor($targetKbps * $LongNvencPrimaryMaxrateScale))
}

function Get-LongSizeCapQualityProfiles {
    if ($script:UseNvenc) {
        return @(
            @{ Quality = 30; MaxWidth = $MaxWidth; Bitrate = 0 },
            @{ Quality = 32; MaxWidth = $MaxWidth; Bitrate = 0 },
            @{ Quality = 34; MaxWidth = $LongSizeCapFallbackMaxWidth; Bitrate = 0 },
            @{ Quality = 36; MaxWidth = $LongSizeCapFallbackMaxWidth; Bitrate = 0 }
        )
    }

    return @(
        @{ Quality = 28; MaxWidth = $MaxWidth; Bitrate = 0 },
        @{ Quality = 30; MaxWidth = $MaxWidth; Bitrate = 0 },
        @{ Quality = 32; MaxWidth = $LongSizeCapFallbackMaxWidth; Bitrate = 0 },
        @{ Quality = 32; MaxWidth = $LongSizeCapFallbackMaxWidth; Bitrate = 0 }
    )
}

function Test-ExternalTools {
    $script:FFmpegPath = Resolve-RequiredTool "ffmpeg"
    $script:FFprobePath = Resolve-RequiredTool "ffprobe"
    $script:ExifToolPath = Resolve-RequiredTool "exiftool"

    Write-Log "Found ffmpeg: $script:FFmpegPath"
    Write-Log "Found ffprobe: $script:FFprobePath"
    Write-Log "Found exiftool: $script:ExifToolPath"

    Initialize-VideoEncoder
}

function Invoke-ExternalTool {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Command,

        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    # Native tools such as exiftool write warnings to stderr. With
    # $ErrorActionPreference = "Stop", PowerShell treats those as terminating errors
    # even when the tool exits successfully.
    $previousErrorAction = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $output = & $Command @Arguments 2>&1
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorAction
    }

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

function New-RandomFilePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,

        [Parameter(Mandatory = $true)]
        [string]$Prefix,

        [Parameter(Mandatory = $true)]
        [string]$Extension
    )

    do {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $token = New-RandomToken
        $fileName = "{0}_{1}_{2}{3}" -f $Prefix, $timestamp, $token, $Extension.ToLowerInvariant()
        $path = Join-Path $Directory $fileName
    } while (Test-Path -LiteralPath $path)

    return $path
}

function New-ImageBulkBatchId {
    return "{0}_{1}" -f (Get-Date -Format "yyyyMMdd_HHmmss"), (New-RandomToken)
}

function New-ImageBulkOutputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BatchId,

        [Parameter(Mandatory = $true)]
        [int]$VariantNumber,

        [Parameter(Mandatory = $true)]
        [string]$Extension
    )

    do {
        $token = New-RandomToken
        $fileName = "image_{0}_v{1:00}_{2}{3}" -f $BatchId, $VariantNumber, $token, $Extension.ToLowerInvariant()
        $path = Join-Path $ImageBulkOutputDir $fileName
    } while (Test-Path -LiteralPath $path)

    return $path
}

function New-RandomOutputDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,

        [Parameter(Mandatory = $true)]
        [string]$Prefix
    )

    do {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $token = New-RandomToken
        $directoryName = "{0}_{1}_{2}" -f $Prefix, $timestamp, $token
        $path = Join-Path $Directory $directoryName
    } while (Test-Path -LiteralPath $path)

    New-Item -ItemType Directory -Path $path -Force | Out-Null
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

function Get-OutputArchiveCutoffTime {
    return (Get-Date).AddHours(-1 * $ArchiveAgeHours)
}

function Move-OldOutputFile {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory = $true)]
        [string]$ArchiveDirectory
    )

    if (-not (Test-Path -LiteralPath $ArchiveDirectory)) {
        New-Item -ItemType Directory -Path $ArchiveDirectory -Force | Out-Null
    }

    $destination = Get-UniqueDestinationPath -Directory $ArchiveDirectory -OriginalFileName $File.Name
    Move-Item -LiteralPath $File.FullName -Destination $destination -Force
    return $destination
}

function Move-OldOutputDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$Directory,

        [Parameter(Mandatory = $true)]
        [string]$ArchiveDirectory
    )

    if (-not (Test-Path -LiteralPath $ArchiveDirectory)) {
        New-Item -ItemType Directory -Path $ArchiveDirectory -Force | Out-Null
    }

    $destination = Get-UniqueDestinationPath -Directory $ArchiveDirectory -OriginalFileName $Directory.Name
    Move-Item -LiteralPath $Directory.FullName -Destination $destination -Force
    return $destination
}

function Invoke-FlatOutputArchive {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceDirectory,

        [Parameter(Mandatory = $true)]
        [string]$ArchiveDirectory,

        [Parameter(Mandatory = $true)]
        [string]$Label,

        [Parameter(Mandatory = $true)]
        [datetime]$CutoffTime
    )

    if (-not (Test-Path -LiteralPath $SourceDirectory)) {
        return 0
    }

    $count = 0
    $files = @(Get-ChildItem -LiteralPath $SourceDirectory -File -ErrorAction SilentlyContinue)
    foreach ($file in $files) {
        if ($file.LastWriteTime -gt $CutoffTime) {
            continue
        }

        try {
            [void](Move-OldOutputFile -File $file -ArchiveDirectory $ArchiveDirectory)
            $count++
        }
        catch {
            Write-Log "Could not archive output file '$($file.FullName)': $($_.Exception.Message)" "WARN"
        }
    }

    if ($count -gt 0) {
        Write-Log "Archived $count file(s) from $Label output."
    }

    return $count
}

function Invoke-SetOutputArchive {
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$CutoffTime
    )

    if (-not (Test-Path -LiteralPath $SetOutputDir)) {
        return 0
    }

    $count = 0
    $directories = @(Get-ChildItem -LiteralPath $SetOutputDir -Directory -ErrorAction SilentlyContinue)
    foreach ($directory in $directories) {
        if ($directory.LastWriteTime -gt $CutoffTime) {
            continue
        }

        try {
            [void](Move-OldOutputDirectory -Directory $directory -ArchiveDirectory $ArchiveSetOutputDir)
            $count++
        }
        catch {
            Write-Log "Could not archive set output directory '$($directory.FullName)': $($_.Exception.Message)" "WARN"
        }
    }

    if ($count -gt 0) {
        Write-Log "Archived $count set folder(s) from sets output."
    }

    return $count
}

function Get-OutputArchiveTargets {
    return @(
        [pscustomobject]@{
            SourceDirectory = $OutputDir
            ArchiveDirectory = $ArchiveDefaultOutputDir
            Label = "default"
        },
        [pscustomobject]@{
            SourceDirectory = $ImageBulkOutputDir
            ArchiveDirectory = $ArchiveImageBulkOutputDir
            Label = "images"
        },
        [pscustomobject]@{
            SourceDirectory = $LongOutputDir
            ArchiveDirectory = $ArchiveLongOutputDir
            Label = "long"
        },
        [pscustomobject]@{
            SourceDirectory = $RemuxOutputDir
            ArchiveDirectory = $ArchiveRemuxOutputDir
            Label = "convert"
        }
    )
}

function Invoke-OutputArchiveIfDue {
    if (-not $ArchiveEnabled) {
        return
    }

    $now = Get-Date
    if ($script:LastArchiveCheck -and (($now - $script:LastArchiveCheck).TotalMinutes -lt $ArchiveCheckIntervalMinutes)) {
        return
    }

    $script:LastArchiveCheck = $now
    $cutoffTime = Get-OutputArchiveCutoffTime

    Write-Log "Running scheduled output archive check (older than $ArchiveAgeHours hours)."

    foreach ($target in Get-OutputArchiveTargets) {
        [void](Invoke-FlatOutputArchive -SourceDirectory $target.SourceDirectory -ArchiveDirectory $target.ArchiveDirectory -Label $target.Label -CutoffTime $cutoffTime)
    }

    [void](Invoke-SetOutputArchive -CutoffTime $cutoffTime)
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

function Remove-GeneratedOutputDirectory {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    try {
        if (Test-Path -LiteralPath $Path) {
            Remove-Item -LiteralPath $Path -Recurse -Force
            Write-Log "Removed incomplete output directory after failure: $Path" "WARN"
        }
    }
    catch {
        Write-Log "Could not remove incomplete output directory '$Path': $($_.Exception.Message)" "WARN"
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

function Get-MediaDimensions {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $arguments = @(
        "-v", "error",
        "-select_streams", "v:0",
        "-show_entries", "stream=width,height",
        "-of", "csv=s=x:p=0",
        $Path
    )

    $output = Invoke-ExternalTool -Command $script:FFprobePath -Arguments $arguments
    $dimensionText = (($output | Out-String).Trim() -split "\s+")[0]
    if ($dimensionText -notmatch "^(\d+)x(\d+)") {
        throw "Unable to read image dimensions from ffprobe for: $Path"
    }

    return [pscustomobject]@{
        Width = [int]$Matches[1]
        Height = [int]$Matches[2]
    }
}

function Resolve-ImageProcessingSource {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $source = [pscustomobject]@{
        SourcePath = $Path
        ProcessingPath = $Path
        TempPath = $null
    }

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($extension -ne ".heic") {
        return $source
    }

    $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("media-pipeline-heic-{0}.png" -f [Guid]::NewGuid().ToString("n"))
    $arguments = @(
        "-y",
        "-hide_banner",
        "-loglevel", "error",
        "-i", $Path,
        "-frames:v", "1",
        "-map_metadata", "-1",
        $tempPath
    )

    Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
    Write-Log "Decoded HEIC working copy for processing: $tempPath"

    $source.ProcessingPath = $tempPath
    $source.TempPath = $tempPath
    return $source
}

function Remove-HeicWorkingCopy {
    param(
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    try {
        if (Test-Path -LiteralPath $Path) {
            Remove-Item -LiteralPath $Path -Force
        }
    }
    catch {
        Write-Log "Could not remove HEIC working copy '$Path': $($_.Exception.Message)" "WARN"
    }
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
        [System.Collections.Generic.HashSet[int]]$UsedValues,

        [Parameter(Mandatory = $true)]
        [int]$CopyCount
    )

    if (-not $Range.CanTrim) {
        return 0
    }

    $rangeSize = ($Range.MaxMs - $Range.MinMs) + 1
    $mustBeUnique = $rangeSize -ge $CopyCount
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

    Write-Log "Video variant $VariantNumber trim: ${TrimMs}ms, target duration: ${targetDurationText}s"

    $qualityValue = if ($script:UseNvenc) { $NvencCq } else { $Crf }
    $arguments = @(
        "-y",
        "-hide_banner",
        "-loglevel", "error",
        "-i", $InputPath,
        "-t", $targetDurationText,
        "-map", "0:v:0",
        "-map", "0:a?"
    )
    $arguments += New-VideoEncoderArguments -QualityValue $qualityValue -MaxWidthValue $MaxWidth
    $arguments += @(
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
    $outputExtension = $extension
    if ($extension -eq ".heic") {
        $outputExtension = ".png"
    }

    $outputPath = New-RandomOutputPath -Extension $outputExtension

    if ($extension -eq ".heic") {
        $arguments = @(
            "-y",
            "-hide_banner",
            "-loglevel", "error",
            "-i", $InputPath,
            "-frames:v", "1",
            "-map_metadata", "-1",
            $outputPath
        )

        Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
    }
    else {
        Copy-Item -LiteralPath $InputPath -Destination $outputPath -Force
    }

    Clear-Metadata -Path $outputPath
    Write-Log "Created image output variant ${VariantNumber}: $outputPath"

    return $outputPath
}

function Process-VideoFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [int]$CopyCount
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

        for ($variant = 1; $variant -le $CopyCount; $variant++) {
            $trimMs = New-TrimMilliseconds -Range $range -UsedValues $usedTrimValues -CopyCount $CopyCount
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
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [int]$CopyCount
    )

    $createdOutputs = New-Object System.Collections.Generic.List[string]

    try {
        for ($variant = 1; $variant -le $CopyCount; $variant++) {
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

function Get-ImageBulkOutputExtension {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath
    )

    $extension = [System.IO.Path]::GetExtension($InputPath).ToLowerInvariant()
    if ($extension -eq ".heic") {
        return ".png"
    }

    return $extension
}

function Convert-ImageBulkVariant {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [int]$VariantNumber,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Dimensions,

        [Parameter(Mandatory = $true)]
        [string]$BatchId,

        [string]$SourcePath = $InputPath
    )

    $outputExtension = Get-ImageBulkOutputExtension -InputPath $SourcePath
    $outputPath = New-ImageBulkOutputPath -BatchId $BatchId -VariantNumber $VariantNumber -Extension $outputExtension
    $width = $Dimensions.Width
    $height = $Dimensions.Height
    $canCrop = ($width -ge 200 -and $height -ge 200)

    $arguments = @(
        "-y",
        "-hide_banner",
        "-loglevel", "error",
        "-i", $InputPath,
        "-frames:v", "1",
        "-map_metadata", "-1"
    )

    if ($canCrop) {
        $cropPermille = Get-Random -Minimum $ImageBulkCropMinPermille -Maximum ($ImageBulkCropMaxPermille + 1)
        $cropPixelsX = [Math]::Max(1, [int][Math]::Floor($width * $cropPermille / 1000))
        $cropPixelsY = [Math]::Max(1, [int][Math]::Floor($height * $cropPermille / 1000))
        $cropWidth = [Math]::Max(1, $width - ($cropPixelsX * 2))
        $cropHeight = [Math]::Max(1, $height - ($cropPixelsY * 2))
        $offsetX = Get-Random -Minimum 0 -Maximum (($cropPixelsX * 2) + 1)
        $offsetY = Get-Random -Minimum 0 -Maximum (($cropPixelsY * 2) + 1)
        $filter = "crop=${cropWidth}:${cropHeight}:${offsetX}:${offsetY},scale=${width}:${height}"
        $arguments += @("-vf", $filter)
        Write-Log "Image bulk variant $VariantNumber crop: ${cropWidth}x${cropHeight}+${offsetX}+${offsetY}, restored to ${width}x${height}"
    }
    else {
        Write-Log "Image bulk variant $VariantNumber skipping crop because image is small: ${width}x${height}" "WARN"
    }

    if ($outputExtension -in @(".jpg", ".jpeg")) {
        $arguments += @("-q:v", "2")
    }
    elseif ($outputExtension -eq ".webp") {
        $arguments += @("-quality", "92")
    }
    elseif ($outputExtension -eq ".png") {
        $arguments += @("-compression_level", "6")
    }

    $arguments += @($outputPath)

    Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
    Clear-Metadata -Path $outputPath
    Write-Log "Created image bulk output variant ${VariantNumber}: $outputPath"

    return $outputPath
}

function Process-ImageBulkFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $createdOutputs = New-Object System.Collections.Generic.List[string]
    $processingSource = Resolve-ImageProcessingSource -Path $Path

    try {
        $dimensions = Get-MediaDimensions -Path $processingSource.ProcessingPath
        Write-Log "Image bulk dimensions: $($dimensions.Width)x$($dimensions.Height)"

        $batchId = New-ImageBulkBatchId
        Write-Log "Image bulk batch id: $batchId"

        for ($variant = 1; $variant -le $ImageBulkCopiesPerFile; $variant++) {
            $outputPath = Convert-ImageBulkVariant -InputPath $processingSource.ProcessingPath -SourcePath $Path -VariantNumber $variant -Dimensions $dimensions -BatchId $batchId
            $createdOutputs.Add($outputPath)
        }

        Move-InputFile -Path $Path -DestinationDirectory $ImageBulkOriginalDir
        Write-Log "Successfully processed image bulk file: $Path"
    }
    catch {
        Remove-GeneratedOutputs -Paths $createdOutputs.ToArray()
        throw
    }
    finally {
        Remove-HeicWorkingCopy -Path $processingSource.TempPath
    }
}

function Process-ImageBulkFileSafely {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)

    if (-not $script:ProcessingPaths.Add($fullPath)) {
        return
    }

    try {
        Write-Log "Detected image bulk file: $fullPath"
        Wait-FileReady -Path $fullPath
        Process-ImageBulkFile -Path $fullPath
    }
    catch {
        Write-Log "Failed image bulk processing '$fullPath': $($_.Exception.Message)" "ERROR"
        try {
            Move-InputFile -Path $fullPath -DestinationDirectory $ImageBulkFailedDir
        }
        catch {
            Write-Log "Could not move failed image bulk file '$fullPath': $($_.Exception.Message)" "ERROR"
        }
    }
    finally {
        [void]$script:ProcessingPaths.Remove($fullPath)
    }
}

function Convert-SetVideoVariant {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [int]$VariantNumber,

        [Parameter(Mandatory = $true)]
        [double]$DurationSeconds,

        [Parameter(Mandatory = $true)]
        [int]$TrimMs
    )

    $outputPath = New-RandomFilePath -Directory $OutputDirectory -Prefix ("media_v{0:00}" -f $VariantNumber) -Extension ".mp4"
    $trimSeconds = $TrimMs / 1000.0
    $targetDuration = [Math]::Max(0.1, $DurationSeconds - $trimSeconds)
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $targetDurationText = $targetDuration.ToString("0.###", $culture)

    Write-Log "Set video variant $VariantNumber trim: ${TrimMs}ms, target duration: ${targetDurationText}s"

    $qualityValue = if ($script:UseNvenc) { $NvencCq } else { $Crf }
    $arguments = @(
        "-y",
        "-hide_banner",
        "-loglevel", "error",
        "-i", $InputPath,
        "-t", $targetDurationText,
        "-map", "0:v:0",
        "-map", "0:a?"
    )
    $arguments += New-VideoEncoderArguments -QualityValue $qualityValue -MaxWidthValue $MaxWidth
    $arguments += @(
        "-c:a", "aac",
        "-b:a", $AudioBitrate,
        "-movflags", "+faststart",
        "-map_metadata", "-1",
        $outputPath
    )

    Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
    Clear-Metadata -Path $outputPath
    Write-Log "Created set video output: $outputPath"

    return $outputPath
}

function Convert-SetImageVariant {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [int]$VariantNumber,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Dimensions,

        [string]$SourcePath = $InputPath
    )

    $outputExtension = Get-ImageBulkOutputExtension -InputPath $SourcePath
    $outputPath = New-RandomFilePath -Directory $OutputDirectory -Prefix ("media_v{0:00}" -f $VariantNumber) -Extension $outputExtension
    $width = $Dimensions.Width
    $height = $Dimensions.Height
    $canCrop = ($width -ge 200 -and $height -ge 200)

    $arguments = @(
        "-y",
        "-hide_banner",
        "-loglevel", "error",
        "-i", $InputPath,
        "-frames:v", "1",
        "-map_metadata", "-1"
    )

    if ($canCrop) {
        $cropPermille = Get-Random -Minimum $ImageBulkCropMinPermille -Maximum ($ImageBulkCropMaxPermille + 1)
        $cropPixelsX = [Math]::Max(1, [int][Math]::Floor($width * $cropPermille / 1000))
        $cropPixelsY = [Math]::Max(1, [int][Math]::Floor($height * $cropPermille / 1000))
        $cropWidth = [Math]::Max(1, $width - ($cropPixelsX * 2))
        $cropHeight = [Math]::Max(1, $height - ($cropPixelsY * 2))
        $offsetX = Get-Random -Minimum 0 -Maximum (($cropPixelsX * 2) + 1)
        $offsetY = Get-Random -Minimum 0 -Maximum (($cropPixelsY * 2) + 1)
        $filter = "crop=${cropWidth}:${cropHeight}:${offsetX}:${offsetY},scale=${width}:${height}"
        $arguments += @("-vf", $filter)
        Write-Log "Set image variant $VariantNumber crop: ${cropWidth}x${cropHeight}+${offsetX}+${offsetY}, restored to ${width}x${height}"
    }
    else {
        Write-Log "Set image variant $VariantNumber skipping crop because image is small: ${width}x${height}" "WARN"
    }

    if ($outputExtension -in @(".jpg", ".jpeg")) {
        $arguments += @("-q:v", "2")
    }
    elseif ($outputExtension -eq ".webp") {
        $arguments += @("-quality", "92")
    }
    elseif ($outputExtension -eq ".png") {
        $arguments += @("-compression_level", "6")
    }

    $arguments += @($outputPath)

    Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
    Clear-Metadata -Path $outputPath
    Write-Log "Created set image output: $outputPath"

    return $outputPath
}

function Process-SetMediaFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $outputDirectory = $null

    try {
        $outputDirectory = New-RandomOutputDirectory -Directory $SetOutputDir -Prefix "set"
        Write-Log "Set output directory: $outputDirectory"

        if (Test-IsVideo $Path) {
            $duration = Get-VideoDurationSeconds -Path $Path
            $durationText = $duration.ToString("0.###", [System.Globalization.CultureInfo]::InvariantCulture)
            Write-Log "Set video duration: ${durationText}s"

            $range = Get-TrimRange -DurationSeconds $duration
            if ($range.CanTrim) {
                Write-Log "Set video trim range $($range.MinMs)-$($range.MaxMs) ms"
            }
            else {
                Write-Log "Set video skipping trim: $($range.Reason)" "WARN"
            }

            $usedTrimValues = [System.Collections.Generic.HashSet[int]]::new()
            for ($variant = 1; $variant -le $SetCopiesPerFile; $variant++) {
                $trimMs = New-TrimMilliseconds -Range $range -UsedValues $usedTrimValues -CopyCount $SetCopiesPerFile
                [void](Convert-SetVideoVariant -InputPath $Path -OutputDirectory $outputDirectory -VariantNumber $variant -DurationSeconds $duration -TrimMs $trimMs)
            }
        }
        else {
            $processingSource = Resolve-ImageProcessingSource -Path $Path

            try {
                $dimensions = Get-MediaDimensions -Path $processingSource.ProcessingPath
                Write-Log "Set image dimensions: $($dimensions.Width)x$($dimensions.Height)"

                for ($variant = 1; $variant -le $SetCopiesPerFile; $variant++) {
                    [void](Convert-SetImageVariant -InputPath $processingSource.ProcessingPath -SourcePath $Path -OutputDirectory $outputDirectory -VariantNumber $variant -Dimensions $dimensions)
                }
            }
            finally {
                Remove-HeicWorkingCopy -Path $processingSource.TempPath
            }
        }

        Move-InputFile -Path $Path -DestinationDirectory $SetOriginalDir
        Write-Log "Successfully processed set media file: $Path"
    }
    catch {
        Remove-GeneratedOutputDirectory -Path $outputDirectory
        throw
    }
}

function Process-SetMediaFileSafely {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)

    if (-not $script:ProcessingPaths.Add($fullPath)) {
        return
    }

    try {
        Write-Log "Detected set media file: $fullPath"
        Wait-FileReady -Path $fullPath

        if (-not (Test-IsSupportedMedia $fullPath)) {
            Write-Log "Skipping unsupported set media file: $fullPath" "WARN"
            return
        }

        Process-SetMediaFile -Path $fullPath
    }
    catch {
        Write-Log "Failed set media processing '$fullPath': $($_.Exception.Message)" "ERROR"
        try {
            Move-InputFile -Path $fullPath -DestinationDirectory $SetFailedDir
        }
        catch {
            Write-Log "Could not move failed set media file '$fullPath': $($_.Exception.Message)" "ERROR"
        }
    }
    finally {
        [void]$script:ProcessingPaths.Remove($fullPath)
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

    $copyCount = Get-DefaultPipelineCopyCount
    Write-Log "Default pipeline copy count for entry $($script:DefaultPipelineEntryCount): $copyCount"

    if (Test-IsVideo $Path) {
        [void](Process-VideoFile -Path $Path -CopyCount $copyCount)
    }
    else {
        [void](Process-ImageFile -Path $Path -CopyCount $copyCount)
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

function Invoke-MovToMp4Remux {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    $arguments = @(
        "-y",
        "-hide_banner",
        "-loglevel", "error",
        "-i", $InputPath,
        "-map", "0:v:0",
        "-map", "0:a?",
        "-dn",
        "-c", "copy",
        "-map_metadata", "-1",
        "-movflags", "+faststart",
        $OutputPath
    )

    Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
}

function Invoke-RemuxImageConvert {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    $arguments = @(
        "-y",
        "-hide_banner",
        "-loglevel", "error",
        "-i", $InputPath,
        "-frames:v", "1",
        "-map_metadata", "-1"
    )

    $outputExtension = [System.IO.Path]::GetExtension($OutputPath).ToLowerInvariant()
    if ($outputExtension -in @(".jpg", ".jpeg")) {
        $arguments += @("-q:v", "2")
    }
    elseif ($outputExtension -eq ".webp") {
        $arguments += @("-quality", "92")
    }
    elseif ($outputExtension -eq ".png") {
        $arguments += @("-compression_level", "6")
    }

    $arguments += @($OutputPath)

    Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
}

function Convert-RemuxMediaFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()

    if ($RemuxVideoSourceExtensions -contains $extension) {
        $outputPath = New-RandomFilePath -Directory $RemuxOutputDir -Prefix "remux" -Extension ".mp4"

        Write-Log "Started convert (video) for: $Path"
        Write-Log "Convert output path: $outputPath"

        try {
            Invoke-MovToMp4Remux -InputPath $Path -OutputPath $outputPath
            Clear-Metadata -Path $outputPath
            Move-InputFile -Path $Path -DestinationDirectory $RemuxOriginalVideosDir
            Write-Log "Successfully converted video to MP4: $outputPath"
        }
        catch {
            Remove-GeneratedOutputs -Paths @($outputPath)
            throw
        }

        return
    }

    if ($RemuxImageSourceExtensions -contains $extension) {
        $outputPath = New-RandomFilePath -Directory $RemuxOutputDir -Prefix "convert" -Extension $RemuxImageOutputExtension

        Write-Log "Started convert (image) for: $Path"
        Write-Log "Convert output path: $outputPath"

        try {
            Invoke-RemuxImageConvert -InputPath $Path -OutputPath $outputPath
            Clear-Metadata -Path $outputPath
            Move-InputFile -Path $Path -DestinationDirectory $RemuxOriginalImagesDir
            Write-Log "Successfully converted image to $($RemuxImageOutputExtension): $outputPath"
        }
        catch {
            Remove-GeneratedOutputs -Paths @($outputPath)
            throw
        }

        return
    }

    if (Test-IsSupportedMedia $Path) {
        Write-Log "Convert pass-through (already a supported format): $Path"
        Move-InputFile -Path $Path -DestinationDirectory $RemuxOutputDir
        Write-Log "Passed through to convert output unchanged: $Path"
        return
    }

    throw "Unsupported convert source format '$extension' for: $Path"
}

function Process-RemuxFileSafely {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)

    if (-not $script:ProcessingPaths.Add($fullPath)) {
        return
    }

    try {
        Write-Log "Detected convert file: $fullPath"
        Wait-FileReady -Path $fullPath
        Convert-RemuxMediaFile -Path $fullPath
    }
    catch {
        Write-Log "Failed converting '$fullPath': $($_.Exception.Message)" "ERROR"
        try {
            Move-InputFile -Path $fullPath -DestinationDirectory $RemuxFailedDir
        }
        catch {
            Write-Log "Could not move failed remux file '$fullPath': $($_.Exception.Message)" "ERROR"
        }
    }
    finally {
        [void]$script:ProcessingPaths.Remove($fullPath)
    }
}

function Get-LongSegmentPlan {
    param(
        [Parameter(Mandatory = $true)]
        [double]$DurationSeconds
    )

    $durationMs = [int][Math]::Floor($DurationSeconds * 1000)
    $targetMs = [int]($LongSegmentTargetSeconds * 1000)
    $minMs = [int]($LongSegmentMinSeconds * 1000)
    $durations = New-Object System.Collections.Generic.List[int]

    if ($durationMs -le 0) {
        throw "Cannot segment a video with invalid duration: $DurationSeconds"
    }

    if ($durationMs -le $targetMs) {
        $durations.Add($durationMs)
    }
    else {
        $fullCount = [int][Math]::Floor($durationMs / $targetMs)
        $remainderMs = $durationMs - ($fullCount * $targetMs)

        for ($i = 0; $i -lt $fullCount; $i++) {
            $durations.Add($targetMs)
        }

        if ($remainderMs -gt 0) {
            if ($remainderMs -ge $minMs) {
                $durations.Add($remainderMs)
            }
            else {
                $neededMs = $minMs - $remainderMs
                $borrowedMs = 0

                for ($i = $durations.Count - 1; $i -ge 0 -and $borrowedMs -lt $neededMs; $i--) {
                    $availableMs = $durations[$i] - $minMs
                    if ($availableMs -le 0) {
                        continue
                    }

                    $takeMs = [Math]::Min($availableMs, $neededMs - $borrowedMs)
                    $durations[$i] = $durations[$i] - $takeMs
                    $borrowedMs += $takeMs
                }

                if ($borrowedMs -eq $neededMs) {
                    $durations.Add($remainderMs + $borrowedMs)
                }
                else {
                    $lastIndex = $durations.Count - 1
                    $durations[$lastIndex] = $durations[$lastIndex] + $remainderMs + $borrowedMs
                }
            }
        }
    }

    $segments = New-Object System.Collections.Generic.List[object]
    $startMs = 0
    for ($i = 0; $i -lt $durations.Count; $i++) {
        $duration = $durations[$i] / 1000.0
        $start = $startMs / 1000.0
        $segments.Add([pscustomobject]@{
            Index = $i + 1
            StartSeconds = $start
            DurationSeconds = $duration
        })
        $startMs += $durations[$i]
    }

    return $segments.ToArray()
}

function Get-LongVideoScaleFilter {
    param(
        [Parameter(Mandatory = $true)]
        [int]$MaxWidthValue
    )

    return Get-VideoScaleFilter -MaxWidthValue $MaxWidthValue
}

function Get-LongTargetVideoBitrateKbps {
    param(
        [Parameter(Mandatory = $true)]
        [double]$DurationSeconds,

        [Parameter(Mandatory = $true)]
        [double]$MaxSizeMegabytes
    )

    if ($DurationSeconds -le 0) {
        return 0
    }

    $audioBitrateText = ($AudioBitrate -replace "[^0-9.]", "")
    $audioBitrateKbps = 128.0
    $parsedAudioBitrate = 0.0
    if (-not [string]::IsNullOrWhiteSpace($audioBitrateText)) {
        if ([double]::TryParse($audioBitrateText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsedAudioBitrate)) {
            $audioBitrateKbps = $parsedAudioBitrate
        }
    }

    $totalBitrateKbps = ($MaxSizeMegabytes * 8192.0) / $DurationSeconds
    $videoBitrateKbps = [Math]::Max(200, $totalBitrateKbps - $audioBitrateKbps)

    return [int][Math]::Floor($videoBitrateKbps * 0.90)
}

function Invoke-LongSegmentExtract {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $true)]
        [double]$StartSeconds,

        [Parameter(Mandatory = $true)]
        [double]$DurationSeconds
    )

    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $startText = $StartSeconds.ToString("0.###", $culture)
    $durationText = $DurationSeconds.ToString("0.###", $culture)

    $arguments = @(
        "-y",
        "-hide_banner",
        "-loglevel", "error",
        "-ss", $startText,
        "-i", $InputPath,
        "-t", $durationText,
        "-map", "0:v:0",
        "-map", "0:a?",
        "-dn",
        "-c", "copy",
        "-map_metadata", "-1",
        "-movflags", "+faststart",
        $OutputPath
    )

    Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
}

function Invoke-LongVideoEncode {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [int]$QualityValue,

        [int]$MaxWidthValue,

        [double]$StartSeconds = -1,

        [double]$DurationSeconds = -1,

        [int]$MaxVideoBitrateKbps = 0
    )

    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $arguments = @(
        "-y",
        "-hide_banner",
        "-loglevel", "error"
    )

    if ($StartSeconds -ge 0 -and $DurationSeconds -gt 0) {
        $startText = $StartSeconds.ToString("0.###", $culture)
        $durationText = $DurationSeconds.ToString("0.###", $culture)
        $arguments += @("-ss", $startText, "-i", $InputPath, "-t", $durationText)
    }
    else {
        $arguments += @("-i", $InputPath)

        if ($DurationSeconds -gt 0) {
            $durationText = $DurationSeconds.ToString("0.###", $culture)
            $arguments += @("-t", $durationText)
        }
    }

    $arguments += @(
        "-map", "0:v:0",
        "-map", "0:a?"
    )
    $arguments += New-VideoEncoderArguments -QualityValue $QualityValue -MaxWidthValue $MaxWidthValue -MaxVideoBitrateKbps $MaxVideoBitrateKbps
    $arguments += @(
        "-c:a", "aac",
        "-b:a", $AudioBitrate,
        "-movflags", "+faststart",
        "-map_metadata", "-1",
        $OutputPath
    )

    Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
}

function Invoke-LongOutputSizeCap {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [string]$SourceInputPath = "",

        [double]$StartSeconds = -1,

        [double]$SegmentDurationSeconds = -1,

        [int]$TrimMs = 0
    )

    if ($LongMaxOutputSizeMB -le 0) {
        return
    }

    $maxBytes = [long]($LongMaxOutputSizeMB * 1024 * 1024)
    $initialSize = (Get-Item -LiteralPath $OutputPath).Length

    if ($initialSize -le $maxBytes) {
        return
    }

    Write-Log "Long output exceeds size cap ($([math]::Round($initialSize / 1MB, 2)) MB > $LongMaxOutputSizeMB MB): $OutputPath" "WARN"

    $reencodeFromSource = -not [string]::IsNullOrWhiteSpace($SourceInputPath)
    $encodeInputPath = if ($reencodeFromSource) { $SourceInputPath } else { $OutputPath }
    $encodeStartSeconds = -1
    $encodeDurationSeconds = -1
    $durationForBitrate = Get-VideoDurationSeconds -Path $OutputPath

    if ($reencodeFromSource) {
        $trimSeconds = $TrimMs / 1000.0
        $encodeDurationSeconds = [Math]::Max(0.1, $SegmentDurationSeconds - $trimSeconds)
        $encodeStartSeconds = $StartSeconds
        $durationForBitrate = $encodeDurationSeconds
    }

    $bitrateKbps = Get-LongTargetVideoBitrateKbps -DurationSeconds $durationForBitrate -MaxSizeMegabytes $LongMaxOutputSizeMB
    $profiles = Get-LongSizeCapQualityProfiles
    $profiles[$profiles.Count - 1].Bitrate = $bitrateKbps
    $qualityLabel = if ($script:UseNvenc) { "CQ" } else { "CRF" }

    $outputDirectory = [System.IO.Path]::GetDirectoryName($OutputPath)
    $chosenTempPath = $null
    $chosenSize = [long]::MaxValue

    foreach ($profile in $profiles) {
        $tempPath = Join-Path $outputDirectory ("sizecap_{0}.mp4" -f (New-RandomToken 8))

        try {
            Invoke-LongVideoEncode -InputPath $encodeInputPath -OutputPath $tempPath -StartSeconds $encodeStartSeconds -DurationSeconds $encodeDurationSeconds -QualityValue $profile.Quality -MaxWidthValue $profile.MaxWidth -MaxVideoBitrateKbps $profile.Bitrate
            $newSize = (Get-Item -LiteralPath $tempPath).Length
            $bitrateLabel = if ($profile.Bitrate -gt 0) { "$($profile.Bitrate)k maxrate" } else { "no maxrate" }
            Write-Log "Long size-cap attempt $qualityLabel $($profile.Quality), max width $($profile.MaxWidth), $bitrateLabel -> $([math]::Round($newSize / 1MB, 2)) MB"

            if ($newSize -lt $chosenSize) {
                if ($chosenTempPath -and (Test-Path -LiteralPath $chosenTempPath)) {
                    Remove-Item -LiteralPath $chosenTempPath -Force
                }

                $chosenTempPath = $tempPath
                $chosenSize = $newSize
                $tempPath = $null
            }

            if ($newSize -le $maxBytes) {
                break
            }
        }
        finally {
            if ($tempPath -and (Test-Path -LiteralPath $tempPath)) {
                Remove-Item -LiteralPath $tempPath -Force
            }
        }
    }

    if (-not $chosenTempPath -or -not (Test-Path -LiteralPath $chosenTempPath)) {
        Write-Log "Long output size-cap re-encode did not produce a candidate: $OutputPath" "WARN"
        return
    }

    Move-Item -LiteralPath $chosenTempPath -Destination $OutputPath -Force
    Clear-Metadata -Path $OutputPath

    if ($chosenSize -gt $maxBytes) {
        Write-Log "Long output still above size cap after all attempts ($([math]::Round($chosenSize / 1MB, 2)) MB): $OutputPath" "WARN"
    }
    else {
        Write-Log "Long output compressed to size cap ($([math]::Round($chosenSize / 1MB, 2)) MB): $OutputPath"
    }
}

function Get-LongOutputRecompressTargets {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Directories
    )

    $maxBytes = [long]($LongMaxOutputSizeMB * 1024 * 1024)
    $targets = New-Object System.Collections.Generic.List[string]

    foreach ($directory in $Directories) {
        if (-not (Test-Path -LiteralPath $directory)) {
            continue
        }

        Get-ChildItem -LiteralPath $directory -File -Filter "*.mp4" | ForEach-Object {
            if ($_.Length -gt $maxBytes) {
                $targets.Add($_.FullName)
            }
        }
    }

    return $targets.ToArray()
}

function Start-LongOutputRecompressBatch {
    $directories = @($LongOutputDir)
    $targets = Get-LongOutputRecompressTargets -Directories $directories

    Write-Log "Long output recompress: found $($targets.Count) file(s) over $LongMaxOutputSizeMB MB in $($directories -join ', ')"

    $processed = 0
    $failed = 0

    foreach ($path in $targets) {
        try {
            $before = (Get-Item -LiteralPath $path).Length
            Invoke-LongOutputSizeCap -OutputPath $path
            $after = (Get-Item -LiteralPath $path).Length
            Write-Log "Recompressed long output: $path ($([math]::Round($before / 1MB, 2)) MB -> $([math]::Round($after / 1MB, 2)) MB)"
            $processed++
        }
        catch {
            Write-Log "Failed to recompress long output '$path': $($_.Exception.Message)" "ERROR"
            $failed++
        }
    }

    Write-Log "Long output recompress finished: $processed succeeded, $failed failed"
}

function Convert-LongVideoVariant {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [int]$SegmentNumber,

        [Parameter(Mandatory = $true)]
        [int]$VariantNumber,

        [Parameter(Mandatory = $true)]
        [double]$SegmentDurationSeconds,

        [Parameter(Mandatory = $true)]
        [int]$TrimMs
    )

    $outputPath = New-RandomFilePath -Directory $LongOutputDir -Prefix ("long_s{0:00}_v{1:00}" -f $SegmentNumber, $VariantNumber) -Extension ".mp4"
    $trimSeconds = $TrimMs / 1000.0
    $targetDuration = [Math]::Max(0.1, $SegmentDurationSeconds - $trimSeconds)
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $targetDurationText = $targetDuration.ToString("0.###", $culture)

    Write-Log "Long segment $SegmentNumber variant $VariantNumber trim: ${TrimMs}ms, target duration: ${targetDurationText}s"

    $qualityValue = if ($script:UseNvenc) { $LongNvencCq } else { $Crf }
    $maxVideoBitrateKbps = Get-LongPrimaryMaxVideoBitrateKbps -DurationSeconds $targetDuration
    Invoke-LongVideoEncode -InputPath $InputPath -OutputPath $outputPath -DurationSeconds $targetDuration -QualityValue $qualityValue -MaxWidthValue $MaxWidth -MaxVideoBitrateKbps $maxVideoBitrateKbps
    Clear-Metadata -Path $outputPath
    Invoke-LongOutputSizeCap -OutputPath $outputPath -SourceInputPath $InputPath -SegmentDurationSeconds $SegmentDurationSeconds -TrimMs $TrimMs
    Write-Log "Created long output: $outputPath"

    return $outputPath
}

function Process-LongVideoFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $createdOutputs = New-Object System.Collections.Generic.List[string]
    $workDir = Join-Path $LongWorkDir ("job_{0}_{1}" -f (Get-Date -Format "yyyyMMdd_HHmmss"), (New-RandomToken 4))

    try {
        New-Item -ItemType Directory -Path $workDir -Force | Out-Null

        $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
        $sourcePath = $Path
        if ($extension -eq ".mov") {
            $sourcePath = Join-Path $workDir "source.mp4"
            Write-Log "Long pipeline remuxing MOV source before segmentation: $Path"
            Invoke-MovToMp4Remux -InputPath $Path -OutputPath $sourcePath
        }

        $duration = Get-VideoDurationSeconds -Path $sourcePath
        $durationText = $duration.ToString("0.###", [System.Globalization.CultureInfo]::InvariantCulture)
        $segments = Get-LongSegmentPlan -DurationSeconds $duration
        $segmentSummary = (($segments | ForEach-Object { $_.DurationSeconds.ToString("0.###", [System.Globalization.CultureInfo]::InvariantCulture) + "s" }) -join ", ")

        Write-Log "Long video duration: ${durationText}s"
        Write-Log "Long segment plan: $($segments.Count) segment(s): $segmentSummary"

        $segmentPaths = @{}
        foreach ($segment in $segments) {
            $segmentPath = Join-Path $workDir ("segment_{0:00}.mp4" -f $segment.Index)
            $startText = $segment.StartSeconds.ToString("0.###", [System.Globalization.CultureInfo]::InvariantCulture)
            $segmentDurationText = $segment.DurationSeconds.ToString("0.###", [System.Globalization.CultureInfo]::InvariantCulture)
            Write-Log "Long pipeline extracting segment $($segment.Index) (${startText}s, ${segmentDurationText}s)"
            Invoke-LongSegmentExtract -InputPath $sourcePath -OutputPath $segmentPath -StartSeconds $segment.StartSeconds -DurationSeconds $segment.DurationSeconds
            $segmentPaths[$segment.Index] = $segmentPath
        }

        foreach ($segment in $segments) {
            $segmentInputPath = $segmentPaths[$segment.Index]
            $range = Get-TrimRange -DurationSeconds $segment.DurationSeconds
            if ($range.CanTrim) {
                Write-Log "Long segment $($segment.Index) trim range $($range.MinMs)-$($range.MaxMs) ms"
            }
            else {
                Write-Log "Long segment $($segment.Index) skipping trim: $($range.Reason)" "WARN"
            }

            $usedTrimValues = [System.Collections.Generic.HashSet[int]]::new()
            for ($variant = 1; $variant -le $LongCopiesPerSegment; $variant++) {
                $trimMs = New-TrimMilliseconds -Range $range -UsedValues $usedTrimValues -CopyCount $LongCopiesPerSegment
                $outputPath = Convert-LongVideoVariant -InputPath $segmentInputPath -SegmentNumber $segment.Index -VariantNumber $variant -SegmentDurationSeconds $segment.DurationSeconds -TrimMs $trimMs
                $createdOutputs.Add($outputPath)
            }
        }

        Move-InputFile -Path $Path -DestinationDirectory $LongOriginalDir
        Write-Log "Successfully processed long video: $Path"
    }
    catch {
        Remove-GeneratedOutputs -Paths $createdOutputs.ToArray()
        throw
    }
    finally {
        try {
            if (Test-Path -LiteralPath $workDir) {
                Remove-Item -LiteralPath $workDir -Recurse -Force
            }
        }
        catch {
            Write-Log "Could not remove long pipeline work directory '$workDir': $($_.Exception.Message)" "WARN"
        }
    }
}

function Process-LongFileSafely {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)

    if (-not $script:ProcessingPaths.Add($fullPath)) {
        return
    }

    try {
        Write-Log "Detected long pipeline file: $fullPath"
        Wait-FileReady -Path $fullPath
        Process-LongVideoFile -Path $fullPath
    }
    catch {
        Write-Log "Failed long pipeline processing '$fullPath': $($_.Exception.Message)" "ERROR"
        try {
            Move-InputFile -Path $fullPath -DestinationDirectory $LongFailedDir
        }
        catch {
            Write-Log "Could not move failed long pipeline file '$fullPath': $($_.Exception.Message)" "ERROR"
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

function Get-CandidateRemuxFiles {
    if (-not (Test-Path -LiteralPath $RemuxInputDir)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $RemuxInputDir -File | Where-Object {
        (-not (Test-IsTemporaryDownload $_.FullName)) -and (Test-IsSupportedMedia $_.FullName)
    } | Sort-Object LastWriteTime, FullName)
}

function Get-CandidateLongFiles {
    if (-not (Test-Path -LiteralPath $LongInputDir)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $LongInputDir -File | Where-Object {
        (-not (Test-IsTemporaryDownload $_.FullName)) -and (Test-IsVideo $_.FullName)
    } | Sort-Object LastWriteTime, FullName)
}

function Get-CandidateImageBulkFiles {
    if (-not (Test-Path -LiteralPath $ImageBulkInputDir)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $ImageBulkInputDir -File | Where-Object {
        (-not (Test-IsTemporaryDownload $_.FullName)) -and ($ImageExtensions -contains $_.Extension.ToLowerInvariant())
    } | Sort-Object LastWriteTime, FullName)
}

function Get-CandidateSetMediaFiles {
    if (-not (Test-Path -LiteralPath $SetInputDir)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $SetInputDir -File | Where-Object {
        (-not (Test-IsTemporaryDownload $_.FullName)) -and (Test-IsSupportedMedia $_.FullName)
    } | Sort-Object LastWriteTime, FullName)
}

function Start-PollingWatcher {
    Write-Log "Watcher started."
    Write-Log "Input: $InputDir"
    Write-Log "Output: $OutputDir"
    Write-Log "Original archive: $OriginalDir"
    Write-Log "Failed: $FailedDir"
    Write-Log "Convert input: $RemuxInputDir"
    Write-Log "Convert output: $RemuxOutputDir"
    Write-Log "Long pipeline input: $LongInputDir"
    Write-Log "Long pipeline output: $LongOutputDir"
    if ($LongMaxOutputSizeMB -gt 0) {
        Write-Log "Long pipeline size cap: $LongMaxOutputSizeMB MB (fallback max width: $LongSizeCapFallbackMaxWidth px)"
    }
    else {
        Write-Log "Long pipeline size cap: disabled"
    }
    Write-Log "Image bulk input: $ImageBulkInputDir"
    Write-Log "Image bulk output: $ImageBulkOutputDir"
    Write-Log "Set pipeline input: $SetInputDir"
    Write-Log "Set pipeline output: $SetOutputDir"
    if ($ArchiveEnabled) {
        Write-Log "Output archive enabled: files older than $ArchiveAgeHours hours move under $ArchiveRootDir (checked every $ArchiveCheckIntervalMinutes minutes)."
        foreach ($target in Get-OutputArchiveTargets) {
            Write-Log "Output archive target: $($target.Label) -> $($target.ArchiveDirectory)"
        }
        Write-Log "Output archive target: sets -> $ArchiveSetOutputDir"
    }
    Write-Log "Polling every $PollSeconds seconds."

    while ($true) {
        try {
            Invoke-OutputArchiveIfDue

            $setMediaFiles = Get-CandidateSetMediaFiles
            foreach ($file in $setMediaFiles) {
                Process-SetMediaFileSafely -Path $file.FullName
            }

            $imageBulkFiles = Get-CandidateImageBulkFiles
            foreach ($file in $imageBulkFiles) {
                Process-ImageBulkFileSafely -Path $file.FullName
            }

            $longFiles = Get-CandidateLongFiles
            foreach ($file in $longFiles) {
                Process-LongFileSafely -Path $file.FullName
            }

            $remuxFiles = Get-CandidateRemuxFiles
            foreach ($file in $remuxFiles) {
                Process-RemuxFileSafely -Path $file.FullName
            }

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

    if ($RecompressLongOutputs) {
        Start-LongOutputRecompressBatch
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
