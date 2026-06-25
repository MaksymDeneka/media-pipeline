param(
    [switch]$CheckOnly,
    [switch]$RecompressLongOutputs,
    [switch]$AsLibrary
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
# All user-tunable settings live in config.ini next to this script. Every value
# below has a built-in default, so the watcher still runs if config.ini is
# missing or a key is absent/garbled. To change settings, run "Edit Config.bat"
# (opens config.ini in Notepad), then run "Restart Watcher.bat".

# Reads a simple key=value INI file (lines starting with # or ; are comments,
# [section] headers are ignored). Returns a case-insensitive hashtable of raw
# string values. Returns an empty table if the file is missing or unreadable.
function Read-IniSettings {
    param([string]$Path)

    $settings = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) {
        return $settings
    }

    try {
        foreach ($rawLine in (Get-Content -LiteralPath $Path -ErrorAction Stop)) {
            $line = $rawLine.Trim()
            if ($line.Length -eq 0) { continue }
            if ($line.StartsWith('#') -or $line.StartsWith(';') -or $line.StartsWith('[')) { continue }

            $eq = $line.IndexOf('=')
            if ($eq -lt 1) { continue }

            $key = $line.Substring(0, $eq).Trim()
            $value = $line.Substring($eq + 1).Trim()
            if ($value -match '^"(.*)"$' -or $value -match "^'(.*)'$") {
                # Quoted value: take it verbatim (quotes can protect ; # and spaces).
                $value = $matches[1]
            }
            else {
                # Unquoted value: an inline comment starts at the first whitespace
                # followed by ; or #, e.g.  Crf = 20   ; default: 24
                $value = ($value -replace '\s+[;#].*$', '').Trim()
            }
            if ($key.Length -gt 0) { $settings[$key] = $value }
        }
    }
    catch {
        # A malformed config file must never stop the watcher; fall back to defaults.
    }

    return $settings
}

# Returns the config.ini value for $Key coerced to the type of $Default, or
# $Default when the key is missing, blank, or cannot be parsed.
function Get-Setting {
    param(
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $true)]$Default
    )

    if (-not $script:ConfigSettings.ContainsKey($Key)) { return $Default }
    $raw = [string]$script:ConfigSettings[$Key]
    if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
    $raw = $raw.Trim()

    try {
        if ($Default -is [bool]) {
            if ($raw -match '^(true|1|yes|on)$') { return $true }
            if ($raw -match '^(false|0|no|off)$') { return $false }
            return $Default
        }
        elseif ($Default -is [int]) {
            return [int]::Parse($raw, [System.Globalization.CultureInfo]::InvariantCulture)
        }
        elseif ($Default -is [double]) {
            return [double]::Parse($raw, [System.Globalization.CultureInfo]::InvariantCulture)
        }
        else {
            return $raw
        }
    }
    catch {
        return $Default
    }
}

# Locate config.ini next to this script (works both when run with -File and when
# dot-sourced by parallel worker runspaces with -AsLibrary).
$script:ConfigPath = $null
if ($PSScriptRoot) {
    $script:ConfigPath = Join-Path $PSScriptRoot "config.ini"
}
elseif ($PSCommandPath) {
    $script:ConfigPath = Join-Path (Split-Path -Parent $PSCommandPath) "config.ini"
}
$script:ConfigSettings = Read-IniSettings -Path $script:ConfigPath

# --- Tunable scalar settings (loaded from config.ini, with built-in defaults) ---
$PipelineRoot = Get-Setting 'PipelineRoot' 'D:\MediaPipeline'

$DefaultPipelineMinCopiesPerFile = Get-Setting 'DefaultPipelineMinCopiesPerFile' 7
$DefaultPipelineAlternatingCopiesPerFile = Get-Setting 'DefaultPipelineAlternatingCopiesPerFile' 8
$LongCopiesPerSegment = Get-Setting 'LongCopiesPerSegment' 3
$ImageBulkCopiesPerFile = Get-Setting 'ImageBulkCopiesPerFile' 20
$SetCopiesPerFile = Get-Setting 'SetCopiesPerFile' 10
$SetBatchCount = Get-Setting 'SetBatchCount' 10

# How many image conversions run at once (convert pipeline files; bulk pipeline
# variants). Requires PowerShell 7. "auto" (or blank) = min(6, CPU count).
$ImageProcessingConcurrencyRaw = Get-Setting 'ImageProcessingConcurrency' 'auto'
$ImageProcessingConcurrencyParsed = 0
if ([int]::TryParse([string]$ImageProcessingConcurrencyRaw, [ref]$ImageProcessingConcurrencyParsed) -and $ImageProcessingConcurrencyParsed -ge 1) {
    $ImageProcessingConcurrency = $ImageProcessingConcurrencyParsed
}
else {
    $ImageProcessingConcurrency = [Math]::Max(1, [Math]::Min(6, [Environment]::ProcessorCount))
}

$ImageBulkCropMinPermille = Get-Setting 'ImageBulkCropMinPermille' 5
$ImageBulkCropMaxPermille = Get-Setting 'ImageBulkCropMaxPermille' 20
$ImageBulkPngCompressionLevel = Get-Setting 'ImageBulkPngCompressionLevel' 1
if ($ImageBulkPngCompressionLevel -lt 0) { $ImageBulkPngCompressionLevel = 0 }
elseif ($ImageBulkPngCompressionLevel -gt 9) { $ImageBulkPngCompressionLevel = 9 }
$ImageCleanPngCompressionLevel = Get-Setting 'ImageCleanPngCompressionLevel' 1
if ($ImageCleanPngCompressionLevel -lt 0) { $ImageCleanPngCompressionLevel = 0 }
elseif ($ImageCleanPngCompressionLevel -gt 9) { $ImageCleanPngCompressionLevel = 9 }
$MinTrimMs = Get-Setting 'MinTrimMs' 15
$MaxTrimMs = Get-Setting 'MaxTrimMs' 95
$PreferNvenc = Get-Setting 'PreferNvenc' $true
$PreferAmf = Get-Setting 'PreferAmf' $true
$Crf = Get-Setting 'Crf' 24
$Preset = Get-Setting 'Preset' 'medium'
$NvencPreset = Get-Setting 'NvencPreset' 'p4'
$NvencCq = Get-Setting 'NvencCq' 26
$LongNvencCq = Get-Setting 'LongNvencCq' 28
$AmfQuality = Get-Setting 'AmfQuality' 'balanced'
$AmfQp = Get-Setting 'AmfQp' 24
$LongAmfQp = Get-Setting 'LongAmfQp' 26
$LongNvencPrimaryMaxrateScale = Get-Setting 'LongNvencPrimaryMaxrateScale' 0.92
$AudioBitrate = Get-Setting 'AudioBitrate' '128k'
$MaxWidth = Get-Setting 'MaxWidth' 1080
$DefaultMaxOutputSizeMB = Get-Setting 'DefaultMaxOutputSizeMB' 8
$DefaultSizeCapFallbackMaxWidth = Get-Setting 'DefaultSizeCapFallbackMaxWidth' 720
$DefaultNvencPrimaryMaxrateScale = Get-Setting 'DefaultNvencPrimaryMaxrateScale' 0.92
$StableSeconds = Get-Setting 'StableSeconds' 3
$TimeoutSeconds = Get-Setting 'TimeoutSeconds' 600
$PollSeconds = Get-Setting 'PollSeconds' 2
$LongSegmentTargetSeconds = Get-Setting 'LongSegmentTargetSeconds' 15
$LongSegmentMinSeconds = Get-Setting 'LongSegmentMinSeconds' 11
$LongMaxOutputSizeMB = Get-Setting 'LongMaxOutputSizeMB' 8
$LongSizeCapFallbackMaxWidth = Get-Setting 'LongSizeCapFallbackMaxWidth' 720
$ArchiveEnabled = Get-Setting 'ArchiveEnabled' $true
$ArchiveAgeHours = Get-Setting 'ArchiveAgeHours' 15
$ArchiveCheckIntervalMinutes = Get-Setting 'ArchiveCheckIntervalMinutes' 30

# Asset store pipeline: like set-batch (one processed copy of every source file
# per set), but it also writes a heatup.assetStoreMediaManifest.v1 JSON next to
# the generated sets and uses a deliberately tiny end-trim (tens of ms at most)
# so each rendition differs without noticeably changing its length.
$AssetStoreSetCount = Get-Setting 'AssetStoreSetCount' 15
$AssetStoreMinTrimMs = Get-Setting 'AssetStoreMinTrimMs' 10
$AssetStoreMaxTrimMs = Get-Setting 'AssetStoreMaxTrimMs' 40
$AssetStoreManifestSchema = Get-Setting 'AssetStoreManifestSchema' 'heatup.assetStoreMediaManifest.v1'

# --- Derived directory paths (computed from $PipelineRoot above) ---
# Default pipeline folders live under "default" so they no longer sit loose in
# the pipeline root next to the other pipeline folders.
$DefaultRootDir = Join-Path $PipelineRoot "default"
$InputDir = Join-Path $DefaultRootDir "input"
$OutputDir = Join-Path $DefaultRootDir "output"
$OriginalDir = Join-Path $DefaultRootDir "original"
$FailedDir = Join-Path $DefaultRootDir "failed"
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
$ImageCleanRootDir = Join-Path $PipelineRoot "imageclean"
$ImageCleanInputDir = Join-Path $ImageCleanRootDir "input"
$ImageCleanOutputDir = Join-Path $ImageCleanRootDir "output"
$ImageCleanOriginalDir = Join-Path $ImageCleanRootDir "original"
$ImageCleanFailedDir = Join-Path $ImageCleanRootDir "failed"
$SetRootDir = Join-Path $PipelineRoot "sets"
$SetInputDir = Join-Path $SetRootDir "input"
$SetOutputDir = Join-Path $SetRootDir "output"
$SetOriginalDir = Join-Path $SetRootDir "original"
$SetFailedDir = Join-Path $SetRootDir "failed"
$SetBatchRootDir = Join-Path $PipelineRoot "setbatch"
$SetBatchInputDir = Join-Path $SetBatchRootDir "input"
$SetBatchOutputDir = Join-Path $SetBatchRootDir "output"
$SetBatchOriginalDir = Join-Path $SetBatchRootDir "original"
$SetBatchFailedDir = Join-Path $SetBatchRootDir "failed"
$AssetStoreRootDir = Join-Path $PipelineRoot "assetstore"
$AssetStoreInputDir = Join-Path $AssetStoreRootDir "input"
$AssetStoreOutputDir = Join-Path $AssetStoreRootDir "output"
$AssetStoreOriginalDir = Join-Path $AssetStoreRootDir "original"
$AssetStoreFailedDir = Join-Path $AssetStoreRootDir "failed"
$ArchiveRootDir = Join-Path $PipelineRoot "archive"
$ArchiveDefaultOutputDir = Join-Path $ArchiveRootDir "output"
$ArchiveImageBulkOutputDir = Join-Path $ArchiveRootDir "images"
$ArchiveImageCleanOutputDir = Join-Path $ArchiveRootDir "imageclean"
$ArchiveRemuxOutputDir = Join-Path $ArchiveRootDir "convert"
$ArchiveLongOutputDir = Join-Path $ArchiveRootDir "long"
$ArchiveSetOutputDir = Join-Path $ArchiveRootDir "sets"
$ArchiveSetBatchOutputDir = Join-Path $ArchiveRootDir "setbatch"
$ArchiveAssetStoreOutputDir = Join-Path $ArchiveRootDir "assetstore"

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
$script:UseAmf = $false
$script:InstanceMutex = $null
$script:LastArchiveCheck = $null
$script:LogMutex = $null
$script:ScriptPath = $PSCommandPath
$script:SupportsParallel = ($PSVersionTable.PSVersion.Major -ge 7)
$script:DefaultPipelineEntryCount = 0
$script:LastSetBatchSignature = $null
$script:LastAssetStoreSignature = $null

function Get-DefaultPipelineCopyCount {
    $script:DefaultPipelineEntryCount++
    if (($script:DefaultPipelineEntryCount % 2) -eq 1) {
        return $DefaultPipelineAlternatingCopiesPerFile
    }

    return $DefaultPipelineMinCopiesPerFile
}

function Initialize-Folders {
    $directories = @(
        $InputDir, $OutputDir, $OriginalDir, $FailedDir, $LogsDir,
        $RemuxInputDir, $RemuxOutputDir, $RemuxOriginalDir, $RemuxOriginalVideosDir, $RemuxOriginalImagesDir, $RemuxFailedDir,
        $LongInputDir, $LongOutputDir, $LongOriginalDir, $LongFailedDir, $LongWorkDir,
        $ImageBulkInputDir, $ImageBulkOutputDir, $ImageBulkOriginalDir, $ImageBulkFailedDir,
        $ImageCleanInputDir, $ImageCleanOutputDir, $ImageCleanOriginalDir, $ImageCleanFailedDir,
        $SetInputDir, $SetOutputDir, $SetOriginalDir, $SetFailedDir,
        $SetBatchInputDir, $SetBatchOutputDir, $SetBatchOriginalDir, $SetBatchFailedDir,
        $AssetStoreInputDir, $AssetStoreOutputDir, $AssetStoreOriginalDir, $AssetStoreFailedDir,
        $ArchiveDefaultOutputDir, $ArchiveImageBulkOutputDir, $ArchiveImageCleanOutputDir, $ArchiveRemuxOutputDir, $ArchiveLongOutputDir,
        $ArchiveSetOutputDir, $ArchiveSetBatchOutputDir, $ArchiveAssetStoreOutputDir
    )

    foreach ($directory in $directories) {
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
        # Parallel workers (PS7 ForEach-Object -Parallel) run in separate runspaces, so serialize
        # appends through a named system mutex shared by name across all runspaces/processes.
        if (-not $script:LogMutex) {
            $script:LogMutex = [System.Threading.Mutex]::new($false, "Local\MediaPipelineLogMutex")
        }
        $acquired = $false
        try {
            try { $acquired = $script:LogMutex.WaitOne(5000) } catch [System.Threading.AbandonedMutexException] { $acquired = $true }
            Add-Content -LiteralPath $logFile -Value $line -Encoding UTF8
        }
        finally {
            if ($acquired) { $script:LogMutex.ReleaseMutex() }
        }
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

    # winget installs these and adds them to PATH, but a freshly-installed PATH may
    # not be visible yet to an already-running process. The WinGet\Links shims and
    # the C:\Tools portable layout are checked as fallbacks.
    $wingetLinks = Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Links"
    $fallbackPaths = @{
        ffmpeg = @(
            (Join-Path $wingetLinks "ffmpeg.exe"),
            "C:\Tools\ffmpeg\bin\ffmpeg.exe"
        )
        ffprobe = @(
            (Join-Path $wingetLinks "ffprobe.exe"),
            "C:\Tools\ffmpeg\bin\ffprobe.exe"
        )
        exiftool = @(
            (Join-Path $wingetLinks "exiftool.exe"),
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

function Test-FfmpegEncoderUsable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$EncoderName
    )

    try {
        # FFmpeg's -encoders list shows a hardware encoder (h264_nvenc, h264_amf)
        # whenever FFmpeg was *compiled* with it, even on machines without the
        # matching GPU. Trusting the list makes the watcher pick a GPU encoder that
        # then fails at runtime ("Cannot load nvcuda.dll" / "No NVIDIA capable
        # devices found" / AMF "DLL not found"), sending every output to the failed
        # folder. So confirm with a tiny throwaway encode and only trust a clean
        # exit code.
        $listed = & $script:FFmpegPath -hide_banner -encoders 2>&1 | Out-String
        if ($listed -notmatch ("\b{0}\b" -f [regex]::Escape($EncoderName))) {
            return $false
        }

        # 256x256 stays above the hardware encoders' minimum frame size (a smaller
        # probe fails with "Frame Dimension less than the minimum supported value"
        # even on a working GPU), and yuv420p is the format the real encodes use.
        $probeArguments = @(
            "-hide_banner",
            "-loglevel", "error",
            "-f", "lavfi",
            "-i", "color=c=black:s=256x256:r=1:d=1",
            "-frames:v", "1",
            "-c:v", $EncoderName,
            "-pix_fmt", "yuv420p",
            "-f", "null",
            "-"
        )

        $previousErrorAction = $ErrorActionPreference
        try {
            # A failing probe writes to stderr; keep that from becoming a terminating
            # error so we can fall back to another encoder on the exit code instead.
            $ErrorActionPreference = "Continue"
            & $script:FFmpegPath @probeArguments 2>&1 | Out-Null
            $probeExitCode = $LASTEXITCODE
        }
        finally {
            $ErrorActionPreference = $previousErrorAction
        }

        return ($probeExitCode -eq 0)
    }
    catch {
        return $false
    }
}

function Test-NvencEncoderAvailable {
    return Test-FfmpegEncoderUsable -EncoderName "h264_nvenc"
}

function Test-AmfEncoderAvailable {
    return Test-FfmpegEncoderUsable -EncoderName "h264_amf"
}

function Initialize-VideoEncoder {
    $script:UseNvenc = $false
    $script:UseAmf = $false

    # Preference order: NVIDIA GPU (NVENC) -> AMD GPU (AMF) -> CPU (libx264). Each
    # GPU option is confirmed with a real probe encode, so a machine that lists the
    # encoder but cannot run it cleanly falls through to the next option.
    if ($PreferNvenc -and (Test-NvencEncoderAvailable)) {
        $script:UseNvenc = $true
        Write-Log "Video encoder: h264_nvenc (NVIDIA GPU, preset $NvencPreset, CQ $NvencCq, long CQ $LongNvencCq)"
        return
    }

    if ($PreferAmf -and (Test-AmfEncoderAvailable)) {
        $script:UseAmf = $true
        Write-Log "Video encoder: h264_amf (AMD GPU, quality $AmfQuality, QP $AmfQp, long QP $LongAmfQp)"
        return
    }

    if ($PreferNvenc -or $PreferAmf) {
        Write-Log "No usable GPU encoder (NVENC/AMF) found in FFmpeg; falling back to libx264 (CPU)." "WARN"
    }

    Write-Log "Video encoder: libx264 (CPU, preset $Preset, CRF $Crf)"
}

function Get-VideoEncoderName {
    if ($script:UseNvenc) {
        return "h264_nvenc"
    }

    if ($script:UseAmf) {
        return "h264_amf"
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
    elseif ($script:UseAmf) {
        if ($MaxVideoBitrateKbps -gt 0) {
            # Size-targeted: AMF constant-QP ignores a bitrate ceiling, so use
            # peak-constrained VBR aimed at the ceiling (the shared block below
            # adds -maxrate/-bufsize). This lands under the size cap in one pass.
            $arguments = @(
                "-c:v", "h264_amf",
                "-usage", "transcoding",
                "-quality", $AmfQuality,
                "-rc", "vbr_peak",
                "-b:v", ("{0}k" -f $MaxVideoBitrateKbps),
                "-vf", (Get-VideoScaleFilter -MaxWidthValue $MaxWidthValue),
                "-pix_fmt", "yuv420p"
            )
        }
        else {
            # Quality-targeted: constant QP, analogous to NVENC's CQ / libx264's CRF.
            $arguments = @(
                "-c:v", "h264_amf",
                "-usage", "transcoding",
                "-quality", $AmfQuality,
                "-rc", "cqp",
                "-qp_i", [string]$QualityValue,
                "-qp_p", [string]$QualityValue,
                "-qp_b", [string]$QualityValue,
                "-vf", (Get-VideoScaleFilter -MaxWidthValue $MaxWidthValue),
                "-pix_fmt", "yuv420p"
            )
        }
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

function Get-PrimaryMaxVideoBitrateKbps {
    param(
        [Parameter(Mandatory = $true)]
        [double]$DurationSeconds,

        [Parameter(Mandatory = $true)]
        [double]$MaxSizeMegabytes,

        [Parameter(Mandatory = $true)]
        [double]$MaxrateScale
    )

    if ((-not $script:UseNvenc -and -not $script:UseAmf) -or $MaxSizeMegabytes -le 0 -or $DurationSeconds -le 0) {
        return 0
    }

    $targetKbps = Get-TargetVideoBitrateKbps -DurationSeconds $DurationSeconds -MaxSizeMegabytes $MaxSizeMegabytes
    return [int][Math]::Max(200, [Math]::Floor($targetKbps * $MaxrateScale))
}

function Get-OutputSizeCapQualityProfiles {
    param(
        [Parameter(Mandatory = $true)]
        [int]$FallbackMaxWidth
    )

    if ($script:UseNvenc) {
        return @(
            @{ Quality = 30; MaxWidth = $MaxWidth; Bitrate = 0 },
            @{ Quality = 32; MaxWidth = $MaxWidth; Bitrate = 0 },
            @{ Quality = 34; MaxWidth = $FallbackMaxWidth; Bitrate = 0 },
            @{ Quality = 36; MaxWidth = $FallbackMaxWidth; Bitrate = 0 }
        )
    }

    if ($script:UseAmf) {
        return @(
            @{ Quality = 28; MaxWidth = $MaxWidth; Bitrate = 0 },
            @{ Quality = 30; MaxWidth = $MaxWidth; Bitrate = 0 },
            @{ Quality = 32; MaxWidth = $FallbackMaxWidth; Bitrate = 0 },
            @{ Quality = 34; MaxWidth = $FallbackMaxWidth; Bitrate = 0 }
        )
    }

    return @(
        @{ Quality = 28; MaxWidth = $MaxWidth; Bitrate = 0 },
        @{ Quality = 30; MaxWidth = $MaxWidth; Bitrate = 0 },
        @{ Quality = 32; MaxWidth = $FallbackMaxWidth; Bitrate = 0 },
        @{ Quality = 32; MaxWidth = $FallbackMaxWidth; Bitrate = 0 }
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

function Get-OutputDateStamp {
    return (Get-Date).ToString("dd-MM-yyyy")
}

function New-RandomOutputPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Extension
    )

    do {
        $token = New-RandomToken
        $fileName = "dt_{0}_{1}{2}" -f (Get-OutputDateStamp), $token, $Extension.ToLowerInvariant()
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
        [string]$Extension,

        [string]$NameSuffix
    )

    do {
        $token = New-RandomToken
        if ($NameSuffix) {
            $fileName = "{0}_{1}_{2}_{3}{4}" -f $Prefix, (Get-OutputDateStamp), $NameSuffix, $token, $Extension.ToLowerInvariant()
        }
        else {
            $fileName = "{0}_{1}_{2}{3}" -f $Prefix, (Get-OutputDateStamp), $token, $Extension.ToLowerInvariant()
        }
        $path = Join-Path $Directory $fileName
    } while (Test-Path -LiteralPath $path)

    return $path
}

function New-PureRandomFilePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,

        [Parameter(Mandatory = $true)]
        [string]$Extension,

        [int]$TokenBytes = 12
    )

    do {
        $token = New-RandomToken -ByteCount $TokenBytes
        $fileName = "{0}{1}" -f $token, $Extension.ToLowerInvariant()
        $path = Join-Path $Directory $fileName
    } while (Test-Path -LiteralPath $path)

    return $path
}

function New-ImageBulkBatchId {
    return "{0}_{1}" -f (Get-OutputDateStamp), (New-RandomToken)
}

function New-ImageBulkOutputPath {
    param(
        [Parameter(Mandatory = $true)]
        [int]$VariantNumber,

        [Parameter(Mandatory = $true)]
        [string]$Extension
    )

    do {
        $token = New-RandomToken
        $fileName = "img_v{0:00}_{1}_{2}{3}" -f $VariantNumber, (Get-OutputDateStamp), $token, $Extension.ToLowerInvariant()
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
        $token = New-RandomToken
        $directoryName = "{0}_{1}_{2}" -f $Prefix, (Get-OutputDateStamp), $token
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
        $suffix = "{0}_{1}" -f (Get-OutputDateStamp), (New-RandomToken 4)
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

function Invoke-DirectoryOutputArchive {
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
    $directories = @(Get-ChildItem -LiteralPath $SourceDirectory -Directory -ErrorAction SilentlyContinue)
    foreach ($directory in $directories) {
        if ($directory.LastWriteTime -gt $CutoffTime) {
            continue
        }

        try {
            [void](Move-OldOutputDirectory -Directory $directory -ArchiveDirectory $ArchiveDirectory)
            $count++
        }
        catch {
            Write-Log "Could not archive $Label output directory '$($directory.FullName)': $($_.Exception.Message)" "WARN"
        }
    }

    if ($count -gt 0) {
        Write-Log "Archived $count folder(s) from $Label output."
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
            SourceDirectory = $ImageCleanOutputDir
            ArchiveDirectory = $ArchiveImageCleanOutputDir
            Label = "imageclean"
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

    [void](Invoke-DirectoryOutputArchive -SourceDirectory $SetOutputDir -ArchiveDirectory $ArchiveSetOutputDir -Label "sets" -CutoffTime $cutoffTime)
    [void](Invoke-DirectoryOutputArchive -SourceDirectory $SetBatchOutputDir -ArchiveDirectory $ArchiveSetBatchOutputDir -Label "setbatch" -CutoffTime $cutoffTime)
    [void](Invoke-DirectoryOutputArchive -SourceDirectory $AssetStoreOutputDir -ArchiveDirectory $ArchiveAssetStoreOutputDir -Label "assetstore" -CutoffTime $cutoffTime)
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

    $qualityValue = if ($script:UseNvenc) { $NvencCq } elseif ($script:UseAmf) { $AmfQp } else { $Crf }
    $maxVideoBitrateKbps = Get-PrimaryMaxVideoBitrateKbps -DurationSeconds $targetDuration -MaxSizeMegabytes $DefaultMaxOutputSizeMB -MaxrateScale $DefaultNvencPrimaryMaxrateScale
    $arguments = @(
        "-y",
        "-hide_banner",
        "-loglevel", "error",
        "-i", $InputPath,
        "-t", $targetDurationText,
        "-map", "0:v:0",
        "-map", "0:a?"
    )
    $arguments += New-VideoEncoderArguments -QualityValue $qualityValue -MaxWidthValue $MaxWidth -MaxVideoBitrateKbps $maxVideoBitrateKbps
    $arguments += @(
        "-c:a", "aac",
        "-b:a", $AudioBitrate,
        "-movflags", "+faststart",
        "-map_metadata", "-1",
        $outputPath
    )

    Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
    Clear-Metadata -Path $outputPath
    Invoke-OutputSizeCap -OutputPath $outputPath -MaxSizeMegabytes $DefaultMaxOutputSizeMB -FallbackMaxWidth $DefaultSizeCapFallbackMaxWidth -SourceInputPath $InputPath -SegmentDurationSeconds $DurationSeconds -TrimMs $TrimMs
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
    $outputPath = New-ImageBulkOutputPath -VariantNumber $VariantNumber -Extension $outputExtension
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
        $arguments += @("-compression_level", ([string]$ImageBulkPngCompressionLevel))
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
        [string]$Path,

        [int]$VariantConcurrency = $ImageProcessingConcurrency
    )

    $createdOutputs = New-Object System.Collections.Generic.List[string]
    $processingSource = Resolve-ImageProcessingSource -Path $Path

    try {
        $effectiveVariantConcurrency = [Math]::Max(1, $VariantConcurrency)
        $dimensions = Get-MediaDimensions -Path $processingSource.ProcessingPath
        Write-Log "Image bulk dimensions: $($dimensions.Width)x$($dimensions.Height)"

        $batchId = New-ImageBulkBatchId
        Write-Log "Image bulk batch id: $batchId"

        if ($script:SupportsParallel -and $ImageBulkCopiesPerFile -gt 1 -and $effectiveVariantConcurrency -gt 1) {
            $libPath = $script:ScriptPath
            $ffPath = $script:FFmpegPath
            $fpPath = $script:FFprobePath
            $exPath = $script:ExifToolPath
            $procPath = $processingSource.ProcessingPath
            $srcPath = $Path
            $dims = $dimensions
            $bId = $batchId
            $variantResults = 1..$ImageBulkCopiesPerFile | ForEach-Object -ThrottleLimit $effectiveVariantConcurrency -Parallel {
                . $using:libPath -AsLibrary
                $script:FFmpegPath = $using:ffPath
                $script:FFprobePath = $using:fpPath
                $script:ExifToolPath = $using:exPath
                try {
                    $out = Convert-ImageBulkVariant -InputPath $using:procPath -SourcePath $using:srcPath -VariantNumber $_ -Dimensions $using:dims -BatchId $using:bId
                    [pscustomobject]@{ Output = $out; Error = $null }
                }
                catch {
                    [pscustomobject]@{ Output = $null; Error = $_.Exception.Message }
                }
            }
            foreach ($vr in $variantResults) {
                if ($vr.Output) { $createdOutputs.Add($vr.Output) }
            }
            $variantErrors = @($variantResults | Where-Object { $_.Error })
            if ($variantErrors.Count -gt 0) {
                throw "Failed $($variantErrors.Count)/$ImageBulkCopiesPerFile image bulk variants. First error: $($variantErrors[0].Error)"
            }
        }
        else {
            for ($variant = 1; $variant -le $ImageBulkCopiesPerFile; $variant++) {
                $outputPath = Convert-ImageBulkVariant -InputPath $processingSource.ProcessingPath -SourcePath $Path -VariantNumber $variant -Dimensions $dimensions -BatchId $batchId
                $createdOutputs.Add($outputPath)
            }
        }

        Move-InputFile -Path $Path -DestinationDirectory $ImageBulkOriginalDir
        Write-Log "Successfully processed image bulk file: $Path"
    }
    catch {
        if ($createdOutputs.Count -gt 0) {
            Write-Log "Preserving $($createdOutputs.Count) completed image bulk output(s) after failure: $Path" "WARN"
        }
        throw
    }
    finally {
        Remove-HeicWorkingCopy -Path $processingSource.TempPath
    }
}

function Process-ImageBulkFileSafely {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [int]$VariantConcurrency = $ImageProcessingConcurrency
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)

    if (-not $script:ProcessingPaths.Add($fullPath)) {
        return
    }

    try {
        Write-Log "Detected image bulk file: $fullPath"
        Wait-FileReady -Path $fullPath
        Process-ImageBulkFile -Path $fullPath -VariantConcurrency $VariantConcurrency
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

function Convert-ImageCleanFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Dimensions,

        [string]$SourcePath = $InputPath
    )

    $outputExtension = Get-ImageBulkOutputExtension -InputPath $SourcePath
    $outputPath = New-PureRandomFilePath -Directory $ImageCleanOutputDir -Extension $outputExtension
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
        Write-Log "Image clean crop: ${cropWidth}x${cropHeight}+${offsetX}+${offsetY}, restored to ${width}x${height}"
    }
    else {
        Write-Log "Image clean skipping crop because image is small: ${width}x${height}" "WARN"
    }

    if ($outputExtension -in @(".jpg", ".jpeg")) {
        $arguments += @("-q:v", "2")
    }
    elseif ($outputExtension -eq ".webp") {
        $arguments += @("-quality", "92")
    }
    elseif ($outputExtension -eq ".png") {
        $arguments += @("-compression_level", ([string]$ImageCleanPngCompressionLevel))
    }

    $arguments += @($outputPath)

    try {
        Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
        Clear-Metadata -Path $outputPath
        Write-Log "Created image clean output: $outputPath"
        return $outputPath
    }
    catch {
        Remove-GeneratedOutputs -Paths @($outputPath)
        throw
    }
}

function Process-ImageCleanFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $createdOutputs = New-Object System.Collections.Generic.List[string]
    $processingSource = Resolve-ImageProcessingSource -Path $Path

    try {
        $dimensions = Get-MediaDimensions -Path $processingSource.ProcessingPath
        Write-Log "Image clean dimensions: $($dimensions.Width)x$($dimensions.Height)"

        $outputPath = Convert-ImageCleanFile -InputPath $processingSource.ProcessingPath -SourcePath $Path -Dimensions $dimensions
        $createdOutputs.Add($outputPath)

        Move-InputFile -Path $Path -DestinationDirectory $ImageCleanOriginalDir
        Write-Log "Successfully processed image clean file: $Path"
    }
    catch {
        Remove-GeneratedOutputs -Paths $createdOutputs.ToArray()
        throw
    }
    finally {
        Remove-HeicWorkingCopy -Path $processingSource.TempPath
    }
}

function Process-ImageCleanFileSafely {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)

    if (-not $script:ProcessingPaths.Add($fullPath)) {
        return
    }

    try {
        Write-Log "Detected image clean file: $fullPath"
        Wait-FileReady -Path $fullPath
        Process-ImageCleanFile -Path $fullPath
    }
    catch {
        Write-Log "Failed image clean processing '$fullPath': $($_.Exception.Message)" "ERROR"
        try {
            Move-InputFile -Path $fullPath -DestinationDirectory $ImageCleanFailedDir
        }
        catch {
            Write-Log "Could not move failed image clean file '$fullPath': $($_.Exception.Message)" "ERROR"
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

    $outputPath = New-RandomFilePath -Directory $OutputDirectory -Prefix ("dt_v{0:00}" -f $VariantNumber) -Extension ".mp4"
    $trimSeconds = $TrimMs / 1000.0
    $targetDuration = [Math]::Max(0.1, $DurationSeconds - $trimSeconds)
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $targetDurationText = $targetDuration.ToString("0.###", $culture)

    Write-Log "Set video variant $VariantNumber trim: ${TrimMs}ms, target duration: ${targetDurationText}s"

    $qualityValue = if ($script:UseNvenc) { $NvencCq } elseif ($script:UseAmf) { $AmfQp } else { $Crf }
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
    $outputPath = New-RandomFilePath -Directory $OutputDirectory -Prefix ("dt_v{0:00}" -f $VariantNumber) -Extension $outputExtension
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
        $outputDirectory = New-RandomOutputDirectory -Directory $SetOutputDir -Prefix "st"
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

function Get-SetBatchOutputExtension {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath
    )

    $extension = [System.IO.Path]::GetExtension($InputPath).ToLowerInvariant()
    if ($extension -eq ".heic") {
        return ".jpg"
    }

    return $extension
}

function Convert-SetBatchImageVariant {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [int]$SetNumber,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Dimensions,

        [string]$SourcePath = $InputPath
    )

    $outputExtension = Get-SetBatchOutputExtension -InputPath $SourcePath
    $outputPath = New-RandomFilePath -Directory $OutputDirectory -Prefix "dt" -Extension $outputExtension
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
        Write-Log "Set batch image set $SetNumber crop: ${cropWidth}x${cropHeight}+${offsetX}+${offsetY}, restored to ${width}x${height}"
    }
    else {
        Write-Log "Set batch image set $SetNumber skipping crop because image is small: ${width}x${height}" "WARN"
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
    Write-Log "Created set batch output (set $SetNumber): $outputPath"

    return $outputPath
}

function Process-SetBatchSourceFile {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory = $true)]
        [string[]]$SetDirectories
    )

    $path = $File.FullName

    if (Test-IsVideo $path) {
        $duration = Get-VideoDurationSeconds -Path $path
        $durationText = $duration.ToString("0.###", [System.Globalization.CultureInfo]::InvariantCulture)
        Write-Log "Set batch video duration: ${durationText}s ($($File.Name))"

        $range = Get-TrimRange -DurationSeconds $duration
        if ($range.CanTrim) {
            Write-Log "Set batch video trim range $($range.MinMs)-$($range.MaxMs) ms ($($File.Name))"
        }
        else {
            Write-Log "Set batch video skipping trim: $($range.Reason) ($($File.Name))" "WARN"
        }

        $usedTrimValues = [System.Collections.Generic.HashSet[int]]::new()
        for ($setNumber = 1; $setNumber -le $SetBatchCount; $setNumber++) {
            $trimMs = New-TrimMilliseconds -Range $range -UsedValues $usedTrimValues -CopyCount $SetBatchCount
            [void](Convert-SetVideoVariant -InputPath $path -OutputDirectory $SetDirectories[$setNumber - 1] -VariantNumber $setNumber -DurationSeconds $duration -TrimMs $trimMs)
        }

        return
    }

    $processingSource = Resolve-ImageProcessingSource -Path $path

    try {
        $dimensions = Get-MediaDimensions -Path $processingSource.ProcessingPath
        Write-Log "Set batch image dimensions: $($dimensions.Width)x$($dimensions.Height) ($($File.Name))"

        if ($script:SupportsParallel -and $SetBatchCount -gt 1) {
            $libPath = $script:ScriptPath
            $ffPath = $script:FFmpegPath
            $fpPath = $script:FFprobePath
            $exPath = $script:ExifToolPath
            $procPath = $processingSource.ProcessingPath
            $srcPath = $path
            $dims = $dimensions
            $dirs = $SetDirectories
            $variantResults = 1..$SetBatchCount | ForEach-Object -ThrottleLimit $ImageProcessingConcurrency -Parallel {
                . $using:libPath -AsLibrary
                $script:FFmpegPath = $using:ffPath
                $script:FFprobePath = $using:fpPath
                $script:ExifToolPath = $using:exPath
                $targetDirs = $using:dirs
                try {
                    $out = Convert-SetBatchImageVariant -InputPath $using:procPath -SourcePath $using:srcPath -OutputDirectory $targetDirs[$_ - 1] -SetNumber $_ -Dimensions $using:dims
                    [pscustomobject]@{ Output = $out; Error = $null }
                }
                catch {
                    [pscustomobject]@{ Output = $null; Error = $_.Exception.Message }
                }
            }
            $variantErrors = @($variantResults | Where-Object { $_.Error })
            if ($variantErrors.Count -gt 0) {
                throw "Failed $($variantErrors.Count)/$SetBatchCount set-batch copies for '$($File.Name)'. First error: $($variantErrors[0].Error)"
            }
        }
        else {
            for ($setNumber = 1; $setNumber -le $SetBatchCount; $setNumber++) {
                [void](Convert-SetBatchImageVariant -InputPath $processingSource.ProcessingPath -SourcePath $path -OutputDirectory $SetDirectories[$setNumber - 1] -SetNumber $setNumber -Dimensions $dimensions)
            }
        }
    }
    finally {
        Remove-HeicWorkingCopy -Path $processingSource.TempPath
    }
}

function Process-SetBatch {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$Files
    )

    $batchDirectory = New-RandomOutputDirectory -Directory $SetBatchOutputDir -Prefix "bt"
    Write-Log "Set batch output directory: $batchDirectory ($SetBatchCount sets, $($Files.Count) source file(s))"

    try {
        $setDirectories = @()
        for ($setNumber = 1; $setNumber -le $SetBatchCount; $setNumber++) {
            $setDirectory = Join-Path $batchDirectory ("set_{0:00}" -f $setNumber)
            New-Item -ItemType Directory -Path $setDirectory -Force | Out-Null
            $setDirectories += $setDirectory
        }

        foreach ($file in $Files) {
            Write-Log "Set batch processing source file: $($file.Name)"
            Process-SetBatchSourceFile -File $file -SetDirectories $setDirectories
        }
    }
    catch {
        Remove-GeneratedOutputDirectory -Path $batchDirectory
        throw
    }

    # Outputs are complete. Archiving the source files is best-effort and must not
    # discard the finished sets, so it runs after the transactional block above.
    foreach ($file in $Files) {
        try {
            Move-InputFile -Path $file.FullName -DestinationDirectory $SetBatchOriginalDir
        }
        catch {
            Write-Log "Set batch sets are complete but could not archive source '$($file.FullName)': $($_.Exception.Message)" "WARN"
        }
    }

    Write-Log "Successfully processed set batch of $($Files.Count) file(s) into $SetBatchCount sets: $batchDirectory"
}

function Process-SetBatchSafely {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$Files
    )

    try {
        Write-Log "Detected set batch: $($Files.Count) file(s)."
        Process-SetBatch -Files $Files
    }
    catch {
        Write-Log "Failed set batch processing: $($_.Exception.Message)" "ERROR"
        foreach ($file in $Files) {
            try {
                Move-InputFile -Path $file.FullName -DestinationDirectory $SetBatchFailedDir
            }
            catch {
                Write-Log "Could not move failed set batch file '$($file.FullName)': $($_.Exception.Message)" "ERROR"
            }
        }
    }
}

# ---------------------------------------------------------------------------
# Asset store manifest pipeline
# ---------------------------------------------------------------------------
# Treats everything dropped in assetstore\input as one batch. Produces
# $AssetStoreSetCount sets (set_01 .. set_NN), each holding one processed,
# metadata-stripped copy of every source file, then writes a
# heatup.assetStoreMediaManifest.v1 manifest describing every generated variant.
# Each video copy gets a tiny end-trim (tens of ms at most, see
# $AssetStoreMinTrimMs/$AssetStoreMaxTrimMs) so the renditions differ.

function Get-UtcIsoTimestamp {
    # e.g. 2026-06-04T12:00:00.000Z — matches the manifest example format.
    return (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-AssetStoreFamilyKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $sanitized = ($base -replace '[^A-Za-z0-9._-]', '_').Trim('_')
    if ([string]::IsNullOrWhiteSpace($sanitized)) {
        $sanitized = "media"
    }

    return $sanitized
}

function Get-AssetStoreTrimRange {
    param(
        [Parameter(Mandatory = $true)]
        [double]$DurationSeconds
    )

    $durationMs = [int][Math]::Floor($DurationSeconds * 1000)

    # The whole point of this lane is a near-invisible trim, so only trim when
    # the clip has comfortable headroom over the largest possible micro-trim.
    if ($durationMs -lt ($AssetStoreMaxTrimMs + 300)) {
        return [pscustomobject]@{
            CanTrim = $false
            MinMs = 0
            MaxMs = 0
            Reason = "video is too short for an asset-store micro-trim"
        }
    }

    $minMs = [Math]::Max(1, $AssetStoreMinTrimMs)
    $maxMs = [Math]::Max($minMs, $AssetStoreMaxTrimMs)

    return [pscustomobject]@{
        CanTrim = $true
        MinMs = $minMs
        MaxMs = $maxMs
        Reason = "asset-store micro-trim range"
    }
}

function New-AssetStoreVideoVariant {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [string]$SetName,

        [Parameter(Mandatory = $true)]
        [int]$SetNumber,

        [Parameter(Mandatory = $true)]
        [string]$FamilyKey,

        [Parameter(Mandatory = $true)]
        [string]$SourceOriginalName,

        [Parameter(Mandatory = $true)]
        [string]$BatchKey,

        [Parameter(Mandatory = $true)]
        [double]$DurationSeconds,

        [Parameter(Mandatory = $true)]
        [int]$TrimMs
    )

    # Reuse the set lane's encoder (H.264 MP4, AAC, width cap, FFmpeg + ExifTool
    # metadata stripping) so asset-store videos match the other lanes exactly.
    $outputPath = Convert-SetVideoVariant -InputPath $InputPath -OutputDirectory $OutputDirectory -VariantNumber $SetNumber -DurationSeconds $DurationSeconds -TrimMs $TrimMs
    $fileName = [System.IO.Path]::GetFileName($outputPath)
    $sizeBytes = (Get-Item -LiteralPath $outputPath).Length
    $variantDuration = [Math]::Round([Math]::Max(0.1, $DurationSeconds - ($TrimMs / 1000.0)), 3)

    return [ordered]@{
        familyKey          = $FamilyKey
        variantKey         = "{0}__{1}" -f $FamilyKey, $SetName
        path               = "{0}/{1}" -f $SetName, $fileName
        renditionSetKey    = $SetName
        generationBatchKey = $BatchKey
        sourceOriginalName = $SourceOriginalName
        sourceFamilyName   = $FamilyKey
        durationSeconds    = [double]$variantDuration
        sizeBytes          = [long]$sizeBytes
        transformProfile   = "asset_store_video_micro_trim"
        generatedAt        = (Get-UtcIsoTimestamp)
        metadata           = [ordered]@{
            encoder  = (Get-VideoEncoderName)
            trimMs   = $TrimMs
            maxWidth = $MaxWidth
        }
    }
}

function New-AssetStoreImageVariant {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [string]$SetName,

        [Parameter(Mandatory = $true)]
        [int]$SetNumber,

        [Parameter(Mandatory = $true)]
        [string]$FamilyKey,

        [Parameter(Mandatory = $true)]
        [string]$SourceOriginalName,

        [Parameter(Mandatory = $true)]
        [string]$BatchKey,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Dimensions
    )

    # Reuse the set-batch image variant (FFmpeg re-encode, tiny randomized crop
    # back to original dimensions, metadata stripped) for per-set differentiation.
    $outputPath = Convert-SetBatchImageVariant -InputPath $InputPath -SourcePath $SourcePath -OutputDirectory $OutputDirectory -SetNumber $SetNumber -Dimensions $Dimensions
    $fileName = [System.IO.Path]::GetFileName($outputPath)
    $sizeBytes = (Get-Item -LiteralPath $outputPath).Length

    return [ordered]@{
        familyKey          = $FamilyKey
        variantKey         = "{0}__{1}" -f $FamilyKey, $SetName
        path               = "{0}/{1}" -f $SetName, $fileName
        renditionSetKey    = $SetName
        generationBatchKey = $BatchKey
        sourceOriginalName = $SourceOriginalName
        sourceFamilyName   = $FamilyKey
        sizeBytes          = [long]$sizeBytes
        transformProfile   = "asset_store_image_recrop"
        generatedAt        = (Get-UtcIsoTimestamp)
        metadata           = [ordered]@{
            sourceWidth  = $Dimensions.Width
            sourceHeight = $Dimensions.Height
        }
    }
}

function Process-AssetStoreSourceFile {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo]$File,

        [Parameter(Mandatory = $true)]
        [string[]]$SetDirectories,

        [Parameter(Mandatory = $true)]
        [string[]]$SetNames,

        [Parameter(Mandatory = $true)]
        [string]$BatchKey,

        [Parameter(Mandatory = $true)]
        [string]$FamilyKey
    )

    $path = $File.FullName
    $records = New-Object System.Collections.Generic.List[object]

    if (Test-IsVideo $path) {
        $duration = Get-VideoDurationSeconds -Path $path
        $durationText = $duration.ToString("0.###", [System.Globalization.CultureInfo]::InvariantCulture)
        Write-Log "Asset store video duration: ${durationText}s ($($File.Name))"

        $range = Get-AssetStoreTrimRange -DurationSeconds $duration
        if ($range.CanTrim) {
            Write-Log "Asset store micro-trim range $($range.MinMs)-$($range.MaxMs) ms ($($File.Name))"
        }
        else {
            Write-Log "Asset store skipping micro-trim: $($range.Reason) ($($File.Name))" "WARN"
        }

        $usedTrimValues = [System.Collections.Generic.HashSet[int]]::new()
        for ($setNumber = 1; $setNumber -le $AssetStoreSetCount; $setNumber++) {
            $trimMs = New-TrimMilliseconds -Range $range -UsedValues $usedTrimValues -CopyCount $AssetStoreSetCount
            $record = New-AssetStoreVideoVariant -InputPath $path -OutputDirectory $SetDirectories[$setNumber - 1] -SetName $SetNames[$setNumber - 1] -SetNumber $setNumber -FamilyKey $FamilyKey -SourceOriginalName $File.Name -BatchKey $BatchKey -DurationSeconds $duration -TrimMs $trimMs
            $records.Add($record)
        }

        return $records.ToArray()
    }

    $processingSource = Resolve-ImageProcessingSource -Path $path

    try {
        $dimensions = Get-MediaDimensions -Path $processingSource.ProcessingPath
        Write-Log "Asset store image dimensions: $($dimensions.Width)x$($dimensions.Height) ($($File.Name))"

        for ($setNumber = 1; $setNumber -le $AssetStoreSetCount; $setNumber++) {
            $record = New-AssetStoreImageVariant -InputPath $processingSource.ProcessingPath -SourcePath $path -OutputDirectory $SetDirectories[$setNumber - 1] -SetName $SetNames[$setNumber - 1] -SetNumber $setNumber -FamilyKey $FamilyKey -SourceOriginalName $File.Name -BatchKey $BatchKey -Dimensions $dimensions
            $records.Add($record)
        }
    }
    finally {
        Remove-HeicWorkingCopy -Path $processingSource.TempPath
    }

    return $records.ToArray()
}

function Write-AssetStoreManifest {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BatchDirectory,

        [Parameter(Mandatory = $true)]
        [string]$GeneratedAt,

        [AllowEmptyCollection()]
        [object[]]$Variants
    )

    $manifest = [ordered]@{
        schema      = $AssetStoreManifestSchema
        generatedAt = $GeneratedAt
        importRoot  = "."
        variants    = [object[]]$Variants
    }

    $json = $manifest | ConvertTo-Json -Depth 12
    $manifestPath = Join-Path $BatchDirectory "manifest.json"
    # Write UTF-8 without a BOM so strict JSON parsers accept the file.
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($manifestPath, $json, $utf8NoBom)
    Write-Log "Wrote asset store manifest: $manifestPath ($($Variants.Count) variant(s))"

    return $manifestPath
}

function Process-AssetStoreBatch {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$Files
    )

    $batchDirectory = New-RandomOutputDirectory -Directory $AssetStoreOutputDir -Prefix "as"
    $batchKey = [System.IO.Path]::GetFileName($batchDirectory)
    $generatedAt = Get-UtcIsoTimestamp
    Write-Log "Asset store batch output directory: $batchDirectory ($AssetStoreSetCount set(s), $($Files.Count) source file(s))"

    $variants = New-Object System.Collections.Generic.List[object]

    try {
        $setDirectories = @()
        $setNames = @()
        for ($setNumber = 1; $setNumber -le $AssetStoreSetCount; $setNumber++) {
            $setName = "set_{0:00}" -f $setNumber
            $setDirectory = Join-Path $batchDirectory $setName
            New-Item -ItemType Directory -Path $setDirectory -Force | Out-Null
            $setDirectories += $setDirectory
            $setNames += $setName
        }

        # Each source file is one family; keep family keys unique within a batch
        # even when two sources share a base name (e.g. clip.mov and clip.mp4).
        $usedFamilyKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($file in $Files) {
            $familyKey = Get-AssetStoreFamilyKey -FileName $file.Name
            $candidate = $familyKey
            $suffix = 2
            while (-not $usedFamilyKeys.Add($candidate)) {
                $candidate = "{0}_{1}" -f $familyKey, $suffix
                $suffix++
            }
            $familyKey = $candidate

            Write-Log "Asset store processing source file: $($file.Name) (family $familyKey)"
            $records = Process-AssetStoreSourceFile -File $file -SetDirectories $setDirectories -SetNames $setNames -BatchKey $batchKey -FamilyKey $familyKey
            foreach ($record in $records) {
                $variants.Add($record)
            }
        }

        [void](Write-AssetStoreManifest -BatchDirectory $batchDirectory -GeneratedAt $generatedAt -Variants $variants.ToArray())
    }
    catch {
        Remove-GeneratedOutputDirectory -Path $batchDirectory
        throw
    }

    # Outputs and manifest are complete; archiving the sources is best-effort and
    # must not discard the finished sets, so it runs after the transactional block.
    foreach ($file in $Files) {
        try {
            Move-InputFile -Path $file.FullName -DestinationDirectory $AssetStoreOriginalDir
        }
        catch {
            Write-Log "Asset store sets are complete but could not archive source '$($file.FullName)': $($_.Exception.Message)" "WARN"
        }
    }

    Write-Log "Successfully processed asset store batch of $($Files.Count) file(s) into $AssetStoreSetCount set(s): $batchDirectory"
}

function Process-AssetStoreBatchSafely {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileInfo[]]$Files
    )

    try {
        Write-Log "Detected asset store batch: $($Files.Count) file(s)."
        Process-AssetStoreBatch -Files $Files
    }
    catch {
        Write-Log "Failed asset store batch processing: $($_.Exception.Message)" "ERROR"
        foreach ($file in $Files) {
            try {
                Move-InputFile -Path $file.FullName -DestinationDirectory $AssetStoreFailedDir
            }
            catch {
                Write-Log "Could not move failed asset store file '$($file.FullName)': $($_.Exception.Message)" "ERROR"
            }
        }
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
        $outputPath = New-RandomFilePath -Directory $RemuxOutputDir -Prefix "rx" -Extension ".mp4"

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
        $outputPath = New-RandomFilePath -Directory $RemuxOutputDir -Prefix "cv" -Extension $RemuxImageOutputExtension

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

function Get-TargetVideoBitrateKbps {
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

function Invoke-OutputSizeCap {
    param(
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $true)]
        [double]$MaxSizeMegabytes,

        [Parameter(Mandatory = $true)]
        [int]$FallbackMaxWidth,

        [string]$SourceInputPath = "",

        [double]$StartSeconds = -1,

        [double]$SegmentDurationSeconds = -1,

        [int]$TrimMs = 0
    )

    if ($MaxSizeMegabytes -le 0) {
        return
    }

    $maxBytes = [long]($MaxSizeMegabytes * 1024 * 1024)
    $initialSize = (Get-Item -LiteralPath $OutputPath).Length

    if ($initialSize -le $maxBytes) {
        return
    }

    Write-Log "Output exceeds size cap ($([math]::Round($initialSize / 1MB, 2)) MB > $MaxSizeMegabytes MB): $OutputPath" "WARN"

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

    $bitrateKbps = Get-TargetVideoBitrateKbps -DurationSeconds $durationForBitrate -MaxSizeMegabytes $MaxSizeMegabytes
    $profiles = Get-OutputSizeCapQualityProfiles -FallbackMaxWidth $FallbackMaxWidth
    $profiles[$profiles.Count - 1].Bitrate = $bitrateKbps
    $qualityLabel = if ($script:UseNvenc) { "CQ" } elseif ($script:UseAmf) { "QP" } else { "CRF" }

    $outputDirectory = [System.IO.Path]::GetDirectoryName($OutputPath)
    $chosenTempPath = $null
    $chosenSize = [long]::MaxValue

    foreach ($profile in $profiles) {
        $tempPath = Join-Path $outputDirectory ("sizecap_{0}.mp4" -f (New-RandomToken 8))

        try {
            Invoke-LongVideoEncode -InputPath $encodeInputPath -OutputPath $tempPath -StartSeconds $encodeStartSeconds -DurationSeconds $encodeDurationSeconds -QualityValue $profile.Quality -MaxWidthValue $profile.MaxWidth -MaxVideoBitrateKbps $profile.Bitrate
            $newSize = (Get-Item -LiteralPath $tempPath).Length
            $bitrateLabel = if ($profile.Bitrate -gt 0) { "$($profile.Bitrate)k maxrate" } else { "no maxrate" }
            Write-Log "Size-cap attempt $qualityLabel $($profile.Quality), max width $($profile.MaxWidth), $bitrateLabel -> $([math]::Round($newSize / 1MB, 2)) MB"

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
        Write-Log "Output size-cap re-encode did not produce a candidate: $OutputPath" "WARN"
        return
    }

    Move-Item -LiteralPath $chosenTempPath -Destination $OutputPath -Force
    Clear-Metadata -Path $OutputPath

    if ($chosenSize -gt $maxBytes) {
        Write-Log "Output still above size cap after all attempts ($([math]::Round($chosenSize / 1MB, 2)) MB): $OutputPath" "WARN"
    }
    else {
        Write-Log "Output compressed to size cap ($([math]::Round($chosenSize / 1MB, 2)) MB): $OutputPath"
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
            Invoke-OutputSizeCap -OutputPath $path -MaxSizeMegabytes $LongMaxOutputSizeMB -FallbackMaxWidth $LongSizeCapFallbackMaxWidth
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

    $outputPath = New-RandomFilePath -Directory $LongOutputDir -Prefix "lg" -NameSuffix ("s{0:00}_v{1:00}" -f $SegmentNumber, $VariantNumber) -Extension ".mp4"
    $trimSeconds = $TrimMs / 1000.0
    $targetDuration = [Math]::Max(0.1, $SegmentDurationSeconds - $trimSeconds)
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $targetDurationText = $targetDuration.ToString("0.###", $culture)

    Write-Log "Long segment $SegmentNumber variant $VariantNumber trim: ${TrimMs}ms, target duration: ${targetDurationText}s"

    $qualityValue = if ($script:UseNvenc) { $LongNvencCq } elseif ($script:UseAmf) { $LongAmfQp } else { $Crf }
    $maxVideoBitrateKbps = Get-PrimaryMaxVideoBitrateKbps -DurationSeconds $targetDuration -MaxSizeMegabytes $LongMaxOutputSizeMB -MaxrateScale $LongNvencPrimaryMaxrateScale
    Invoke-LongVideoEncode -InputPath $InputPath -OutputPath $outputPath -DurationSeconds $targetDuration -QualityValue $qualityValue -MaxWidthValue $MaxWidth -MaxVideoBitrateKbps $maxVideoBitrateKbps
    Clear-Metadata -Path $outputPath
    Invoke-OutputSizeCap -OutputPath $outputPath -MaxSizeMegabytes $LongMaxOutputSizeMB -FallbackMaxWidth $LongSizeCapFallbackMaxWidth -SourceInputPath $InputPath -SegmentDurationSeconds $SegmentDurationSeconds -TrimMs $TrimMs
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

function Get-CandidateImageCleanFiles {
    if (-not (Test-Path -LiteralPath $ImageCleanInputDir)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $ImageCleanInputDir -File | Where-Object {
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

function Get-CandidateSetBatchFiles {
    if (-not (Test-Path -LiteralPath $SetBatchInputDir)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $SetBatchInputDir -File | Where-Object {
        (-not (Test-IsTemporaryDownload $_.FullName)) -and (Test-IsSupportedMedia $_.FullName)
    } | Sort-Object LastWriteTime, FullName)
}

function Get-CandidateAssetStoreFiles {
    if (-not (Test-Path -LiteralPath $AssetStoreInputDir)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $AssetStoreInputDir -File | Where-Object {
        (-not (Test-IsTemporaryDownload $_.FullName)) -and (Test-IsSupportedMedia $_.FullName)
    } | Sort-Object LastWriteTime, FullName)
}

function Start-PollingWatcher {
    Write-Log "Watcher started."
    Write-Log "Input: $InputDir"
    Write-Log "Output: $OutputDir"
    if ($DefaultMaxOutputSizeMB -gt 0) {
        Write-Log "Default pipeline size cap: $DefaultMaxOutputSizeMB MB (fallback max width: $DefaultSizeCapFallbackMaxWidth px)"
    }
    else {
        Write-Log "Default pipeline size cap: disabled"
    }
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
    Write-Log "Image clean input: $ImageCleanInputDir"
    Write-Log "Image clean output: $ImageCleanOutputDir"
    Write-Log "Set pipeline input: $SetInputDir"
    Write-Log "Set pipeline output: $SetOutputDir"
    Write-Log "Set batch input: $SetBatchInputDir"
    Write-Log "Set batch output: $SetBatchOutputDir ($SetBatchCount sets per batch)"
    Write-Log "Asset store input: $AssetStoreInputDir"
    Write-Log "Asset store output: $AssetStoreOutputDir ($AssetStoreSetCount sets per batch, micro-trim $AssetStoreMinTrimMs-$AssetStoreMaxTrimMs ms, manifest schema $AssetStoreManifestSchema)"
    if ($ArchiveEnabled) {
        Write-Log "Output archive enabled: files older than $ArchiveAgeHours hours move under $ArchiveRootDir (checked every $ArchiveCheckIntervalMinutes minutes)."
        foreach ($target in Get-OutputArchiveTargets) {
            Write-Log "Output archive target: $($target.Label) -> $($target.ArchiveDirectory)"
        }
        Write-Log "Output archive target: sets -> $ArchiveSetOutputDir"
        Write-Log "Output archive target: setbatch -> $ArchiveSetBatchOutputDir"
        Write-Log "Output archive target: assetstore -> $ArchiveAssetStoreOutputDir"
    }
    if ($script:SupportsParallel) {
        Write-Log "Image processing concurrency: $ImageProcessingConcurrency (parallel enabled on PowerShell $($PSVersionTable.PSVersion))."
    }
    else {
        Write-Log "Image processing runs sequentially (PowerShell $($PSVersionTable.PSVersion) has no -Parallel; needs 7+)." "WARN"
    }
    Write-Log "Polling every $PollSeconds seconds."

    while ($true) {
        try {
            Invoke-OutputArchiveIfDue

            $setMediaFiles = Get-CandidateSetMediaFiles
            foreach ($file in $setMediaFiles) {
                Process-SetMediaFileSafely -Path $file.FullName
            }

            $setBatchFiles = Get-CandidateSetBatchFiles
            if ($setBatchFiles.Count -gt 0) {
                $batchReady = $true
                foreach ($batchFile in $setBatchFiles) {
                    if (-not (Test-FileUnlocked $batchFile.FullName)) {
                        $batchReady = $false
                        break
                    }
                }

                $batchSignature = (($setBatchFiles | ForEach-Object { '{0}|{1}' -f $_.FullName, $_.Length }) -join ';')
                $newestWrite = ($setBatchFiles | Measure-Object -Property LastWriteTime -Maximum).Maximum
                $batchSettled = ((Get-Date) - $newestWrite).TotalSeconds -ge $StableSeconds

                if ($batchReady -and $batchSettled -and $batchSignature -eq $script:LastSetBatchSignature) {
                    Process-SetBatchSafely -Files $setBatchFiles
                    $script:LastSetBatchSignature = $null
                }
                else {
                    if ($batchSignature -ne $script:LastSetBatchSignature) {
                        Write-Log "Set batch: $($setBatchFiles.Count) file(s) detected; waiting for the batch to settle before processing."
                    }
                    $script:LastSetBatchSignature = $batchSignature
                }
            }
            else {
                $script:LastSetBatchSignature = $null
            }

            $assetStoreFiles = Get-CandidateAssetStoreFiles
            if ($assetStoreFiles.Count -gt 0) {
                $assetReady = $true
                foreach ($assetFile in $assetStoreFiles) {
                    if (-not (Test-FileUnlocked $assetFile.FullName)) {
                        $assetReady = $false
                        break
                    }
                }

                $assetSignature = (($assetStoreFiles | ForEach-Object { '{0}|{1}' -f $_.FullName, $_.Length }) -join ';')
                $assetNewestWrite = ($assetStoreFiles | Measure-Object -Property LastWriteTime -Maximum).Maximum
                $assetSettled = ((Get-Date) - $assetNewestWrite).TotalSeconds -ge $StableSeconds

                if ($assetReady -and $assetSettled -and $assetSignature -eq $script:LastAssetStoreSignature) {
                    Process-AssetStoreBatchSafely -Files $assetStoreFiles
                    $script:LastAssetStoreSignature = $null
                }
                else {
                    if ($assetSignature -ne $script:LastAssetStoreSignature) {
                        Write-Log "Asset store batch: $($assetStoreFiles.Count) file(s) detected; waiting for the batch to settle before processing."
                    }
                    $script:LastAssetStoreSignature = $assetSignature
                }
            }
            else {
                $script:LastAssetStoreSignature = $null
            }

            $imageCleanFiles = Get-CandidateImageCleanFiles
            if ($script:SupportsParallel -and $imageCleanFiles.Count -gt 1) {
                $libPath = $script:ScriptPath
                $ffPath = $script:FFmpegPath
                $fpPath = $script:FFprobePath
                $exPath = $script:ExifToolPath
                $imageCleanFiles | ForEach-Object -ThrottleLimit $ImageProcessingConcurrency -Parallel {
                    . $using:libPath -AsLibrary
                    $script:FFmpegPath = $using:ffPath
                    $script:FFprobePath = $using:fpPath
                    $script:ExifToolPath = $using:exPath
                    Process-ImageCleanFileSafely -Path $_.FullName
                }
            }
            else {
                foreach ($file in $imageCleanFiles) {
                    Process-ImageCleanFileSafely -Path $file.FullName
                }
            }

            $imageBulkFiles = Get-CandidateImageBulkFiles
            if ($script:SupportsParallel -and $imageBulkFiles.Count -gt 1) {
                $libPath = $script:ScriptPath
                $ffPath = $script:FFmpegPath
                $fpPath = $script:FFprobePath
                $exPath = $script:ExifToolPath
                Write-Log "Image bulk processing $($imageBulkFiles.Count) file(s) with file concurrency $ImageProcessingConcurrency and per-file variant concurrency 1."
                $imageBulkFiles | ForEach-Object -ThrottleLimit $ImageProcessingConcurrency -Parallel {
                    . $using:libPath -AsLibrary
                    $script:FFmpegPath = $using:ffPath
                    $script:FFprobePath = $using:fpPath
                    $script:ExifToolPath = $using:exPath
                    Process-ImageBulkFileSafely -Path $_.FullName -VariantConcurrency 1
                }
            }
            else {
                foreach ($file in $imageBulkFiles) {
                    Process-ImageBulkFileSafely -Path $file.FullName
                }
            }

            $longFiles = Get-CandidateLongFiles
            foreach ($file in $longFiles) {
                Process-LongFileSafely -Path $file.FullName
            }

            $remuxFiles = Get-CandidateRemuxFiles
            if ($script:SupportsParallel -and $remuxFiles.Count -gt 1) {
                $libPath = $script:ScriptPath
                $ffPath = $script:FFmpegPath
                $fpPath = $script:FFprobePath
                $exPath = $script:ExifToolPath
                $remuxFiles | ForEach-Object -ThrottleLimit $ImageProcessingConcurrency -Parallel {
                    . $using:libPath -AsLibrary
                    $script:FFmpegPath = $using:ffPath
                    $script:FFprobePath = $using:fpPath
                    $script:ExifToolPath = $using:exPath
                    Process-RemuxFileSafely -Path $_.FullName
                }
            }
            else {
                foreach ($file in $remuxFiles) {
                    Process-RemuxFileSafely -Path $file.FullName
                }
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

# When dot-sourced by a parallel worker runspace, only load functions/config and return —
# do not take the single-instance mutex or start the polling loop.
if ($AsLibrary) { return }

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
