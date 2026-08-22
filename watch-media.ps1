param(
    [switch]$CheckOnly,
    [switch]$RecompressOutputs,
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
# Parses config.ini into global settings plus any [preset <name>] sections.
#
# Ordinary section headers stay decorative: their keys land in the flat global table, which
# is how every setting behaved before presets existed. A [preset <name>] header instead
# collects the keys that follow it under that preset, so a preset can override any global.
function Read-IniDocument {
    param([string]$Path)

    $globals = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
    $presets = New-Object 'System.Collections.Specialized.OrderedDictionary' ([System.StringComparer]::OrdinalIgnoreCase)

    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) {
        return [pscustomobject]@{ Globals = $globals; Presets = $presets }
    }

    try {
        $currentPreset = $null

        foreach ($rawLine in (Get-Content -LiteralPath $Path -ErrorAction Stop)) {
            $line = $rawLine.Trim()
            if ($line.Length -eq 0) { continue }
            if ($line.StartsWith('#') -or $line.StartsWith(';')) { continue }

            if ($line.StartsWith('[')) {
                $header = $line.TrimStart('[').TrimEnd(']').Trim()
                if ($header -match '^preset\s+(.+)$') {
                    $currentPreset = $matches[1].Trim()
                    if (-not $presets.Contains($currentPreset)) {
                        $presets[$currentPreset] = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
                    }
                }
                else {
                    $currentPreset = $null
                }
                continue
            }

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

            if ($key.Length -eq 0) { continue }

            if ($currentPreset) {
                $presets[$currentPreset][$key] = $value
            }
            else {
                $globals[$key] = $value
            }
        }
    }
    catch {
        # A malformed config file must never stop the watcher; fall back to defaults.
    }

    return [pscustomobject]@{ Globals = $globals; Presets = $presets }
}

# Coerces a raw config string to the type of $Default, returning $Default when the value is
# missing, blank, or cannot be parsed.
function ConvertTo-SettingValue {
    param(
        [string]$Raw,
        [Parameter(Mandatory = $true)]$Default
    )

    if ([string]::IsNullOrWhiteSpace($Raw)) { return $Default }
    $trimmed = $Raw.Trim()

    try {
        if ($Default -is [bool]) {
            if ($trimmed -match '^(true|1|yes|on)$') { return $true }
            if ($trimmed -match '^(false|0|no|off)$') { return $false }
            return $Default
        }
        elseif ($Default -is [int]) {
            return [int]::Parse($trimmed, [System.Globalization.CultureInfo]::InvariantCulture)
        }
        elseif ($Default -is [double]) {
            return [double]::Parse($trimmed, [System.Globalization.CultureInfo]::InvariantCulture)
        }
        else {
            return $trimmed
        }
    }
    catch {
        return $Default
    }
}

# Returns the config.ini global value for $Key coerced to the type of $Default, or
# $Default when the key is missing, blank, or cannot be parsed.
function Get-Setting {
    param(
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $true)]$Default
    )

    if (-not $script:ConfigSettings.ContainsKey($Key)) { return $Default }
    return ConvertTo-SettingValue -Raw ([string]$script:ConfigSettings[$Key]) -Default $Default
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
$script:ConfigDocument = Read-IniDocument -Path $script:ConfigPath
$script:ConfigSettings = $script:ConfigDocument.Globals
$script:ConfigPresetSections = $script:ConfigDocument.Presets

# --- Global settings (loaded from config.ini, with built-in defaults) ---
#
# Every key here that has a matching preset option is simply the default for that option:
# a preset inherits it unless its own [preset <name>] section overrides it. The script
# variables are prefixed Default so a preset object passed around as $Preset can never
# shadow them, which PowerShell's scope resolution would otherwise allow.

$PipelineRoot = Get-Setting 'PipelineRoot' 'D:\MediaPipeline'

# How many files are processed at once. Requires PowerShell 7.
# "auto" (or blank) = min(6, CPU count).
$ImageProcessingConcurrencyRaw = Get-Setting 'ImageProcessingConcurrency' 'auto'
$ImageProcessingConcurrencyParsed = 0
if ([int]::TryParse([string]$ImageProcessingConcurrencyRaw, [ref]$ImageProcessingConcurrencyParsed) -and $ImageProcessingConcurrencyParsed -ge 1) {
    $ImageProcessingConcurrency = $ImageProcessingConcurrencyParsed
}
else {
    $ImageProcessingConcurrency = [Math]::Max(1, [Math]::Min(6, [Environment]::ProcessorCount))
}

# Image defaults
$DefaultCropMinPermille = Get-Setting 'CropMinPermille' 5
$DefaultCropMaxPermille = Get-Setting 'CropMaxPermille' 20

$DefaultJpegQuality = Get-Setting 'JpegQuality' 4
if ($DefaultJpegQuality -lt 2) { $DefaultJpegQuality = 2 }
elseif ($DefaultJpegQuality -gt 31) { $DefaultJpegQuality = 31 }

# Sources that had to be decoded first (HEIC) are already a re-encode, so they get a little
# more headroom than an untouched JPEG.
$DefaultConvertedJpegQuality = Get-Setting 'ConvertedJpegQuality' 12
if ($DefaultConvertedJpegQuality -lt 2) { $DefaultConvertedJpegQuality = 2 }
elseif ($DefaultConvertedJpegQuality -gt 31) { $DefaultConvertedJpegQuality = 31 }

$DefaultPngCompressionLevel = Get-Setting 'PngCompressionLevel' 6
if ($DefaultPngCompressionLevel -lt 0) { $DefaultPngCompressionLevel = 0 }
elseif ($DefaultPngCompressionLevel -gt 9) { $DefaultPngCompressionLevel = 9 }

# Video defaults
$DefaultMinTrimMs = Get-Setting 'MinTrimMs' 15
$DefaultMaxTrimMs = Get-Setting 'MaxTrimMs' 95
$PreferNvenc = Get-Setting 'PreferNvenc' $true
$PreferAmf = Get-Setting 'PreferAmf' $true
$DefaultCrf = Get-Setting 'Crf' 24

# Named X264Preset rather than Preset so it cannot be shadowed by the pipeline preset object
# the processing functions pass around.
$X264Preset = Get-Setting 'X264Preset' 'medium'
$NvencPreset = Get-Setting 'NvencPreset' 'p4'
$AmfQuality = Get-Setting 'AmfQuality' 'balanced'
$DefaultNvencCq = Get-Setting 'NvencCq' 26
$DefaultAmfQp = Get-Setting 'AmfQp' 24
$DefaultAudioBitrate = Get-Setting 'AudioBitrate' '128k'
$DefaultMaxWidth = Get-Setting 'MaxWidth' 1080

# Size cap defaults. Zero disables both the first-encode bitrate ceiling and the retry pass.
$DefaultSizeCapMB = Get-Setting 'SizeCapMB' 8
$DefaultSizeCapFallbackMaxWidth = Get-Setting 'SizeCapFallbackMaxWidth' 720
$DefaultMaxrateScale = Get-Setting 'MaxrateScale' 0.92

# Segmenting defaults, used by presets with Segment = true.
$DefaultSegmentTargetSeconds = Get-Setting 'SegmentTargetSeconds' 15
$DefaultSegmentMinSeconds = Get-Setting 'SegmentMinSeconds' 11

# Manifest schema, used by presets with Manifest = true.
$DefaultManifestSchema = Get-Setting 'ManifestSchema' 'heatup.assetStoreMediaManifest.v1'

# Timing
$StableSeconds = Get-Setting 'StableSeconds' 3
$TimeoutSeconds = Get-Setting 'TimeoutSeconds' 600
$PollSeconds = Get-Setting 'PollSeconds' 2

# Archive and retention
$ArchiveEnabled = Get-Setting 'ArchiveEnabled' $true
$ArchiveAgeHours = Get-Setting 'ArchiveAgeHours' 15
$ArchiveCheckIntervalMinutes = Get-Setting 'ArchiveCheckIntervalMinutes' 30
$AssetRetentionDays = Get-Setting 'AssetRetentionDays' 5

$WorkspaceNames = @("LC", "MD", "YL", "PL", "general")
$DefaultWorkspaceName = "LC"

# ---------------------------------------------------------------------------
# Presets
# ---------------------------------------------------------------------------
#
# A preset is a named bundle of processing options, and it owns one folder tree per
# workspace: <PipelineRoot>\<preset>\<workspace>\{input,output,original,failed}.
#
# The nine hardcoded pipelines this replaces were one pipeline with different orchestration.
# What actually varied between them is exactly the option set below: how many copies per
# media type, how output is grouped, whether a whole folder is treated as one batch, whether
# long videos are split, and whether a manifest is written.
#
# Every option falls back to the matching global setting, so a preset section usually needs
# only the two or three values that differ.

$script:PresetGroupingValues = @('Flat', 'PerSource', 'PerSet')
$script:PresetBatchValues = @('PerFile', 'PerGroup')
$script:PresetFailureValues = @('PreservePartial', 'DeleteFiles', 'DeleteContainer')
$script:PresetParallelValues = @('OverFiles', 'OverVariants', 'Sequential')

# Returns a preset's override for $Key coerced to the type of $GlobalValue, or $GlobalValue
# when the preset does not override it. Passing the already-resolved global as the fallback
# keeps preset defaults from drifting away from the global defaults.
function Get-PresetValue {
    param(
        [hashtable]$Overrides,
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $true)]$GlobalValue
    )

    if ($Overrides -and $Overrides.ContainsKey($Key)) {
        return ConvertTo-SettingValue -Raw ([string]$Overrides[$Key]) -Default $GlobalValue
    }

    return $GlobalValue
}

# Same as Get-PresetValue for options that accept only a fixed set of words. An unrecognized
# value logs a warning and falls back rather than failing the whole watcher.
function Get-PresetChoice {
    param(
        [hashtable]$Overrides,
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $true)][string]$Default,
        [Parameter(Mandatory = $true)][string[]]$Allowed,
        [Parameter(Mandatory = $true)][string]$PresetName
    )

    if (-not $Overrides -or -not $Overrides.ContainsKey($Key)) { return $Default }

    $raw = ([string]$Overrides[$Key]).Trim()
    $match = @($Allowed | Where-Object { $_ -eq $raw })[0]
    if ($match) { return $match }

    Write-Log "Preset '$PresetName': '$Key = $raw' is not one of $($Allowed -join ', '). Using $Default." "WARN"
    return $Default
}

function New-PipelinePreset {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [hashtable]$Overrides
    )

    $grouping = Get-PresetChoice -Overrides $Overrides -Key 'Grouping' -Default 'Flat' -Allowed $script:PresetGroupingValues -PresetName $Name
    $batch = Get-PresetChoice -Overrides $Overrides -Key 'Batch' -Default 'PerFile' -Allowed $script:PresetBatchValues -PresetName $Name

    # A grouped preset owns its output folder and can safely delete it after a failure.
    # A flat preset shares one folder with every other file, so it keeps what it produced.
    $defaultFailure = if ($grouping -eq 'Flat') { 'PreservePartial' } else { 'DeleteContainer' }

    return [pscustomobject]@{
        Name                    = $Name
        Enabled                 = Get-PresetValue -Overrides $Overrides -Key 'Enabled' -GlobalValue $true

        # Copy counts are per media type, which is what lets one inbox take a mixed folder
        # of videos and photos. Zero disables that media type for this preset.
        VideoCopies             = Get-PresetValue -Overrides $Overrides -Key 'VideoCopies' -GlobalValue 1
        ImageCopies             = Get-PresetValue -Overrides $Overrides -Key 'ImageCopies' -GlobalValue 1
        # When set, consecutive files alternate between the copy count above and this one,
        # so a run of files does not all produce the same number of outputs.
        CopiesAlternate         = Get-PresetValue -Overrides $Overrides -Key 'CopiesAlternate' -GlobalValue 0

        Grouping                = $grouping
        SetCount                = Get-PresetValue -Overrides $Overrides -Key 'SetCount' -GlobalValue 1
        Batch                   = $batch

        Segment                 = Get-PresetValue -Overrides $Overrides -Key 'Segment' -GlobalValue $false
        SegmentTargetSeconds    = Get-PresetValue -Overrides $Overrides -Key 'SegmentTargetSeconds' -GlobalValue $DefaultSegmentTargetSeconds
        SegmentMinSeconds       = Get-PresetValue -Overrides $Overrides -Key 'SegmentMinSeconds' -GlobalValue $DefaultSegmentMinSeconds

        Manifest                = Get-PresetValue -Overrides $Overrides -Key 'Manifest' -GlobalValue $false
        ManifestSchema          = Get-PresetValue -Overrides $Overrides -Key 'ManifestSchema' -GlobalValue $DefaultManifestSchema

        # Converts .mov and .heic sources to workable formats before processing.
        Normalize               = Get-PresetValue -Overrides $Overrides -Key 'Normalize' -GlobalValue $true

        OnFailure               = Get-PresetChoice -Overrides $Overrides -Key 'OnFailure' -Default $defaultFailure -Allowed $script:PresetFailureValues -PresetName $Name
        Parallel                = Get-PresetChoice -Overrides $Overrides -Key 'Parallel' -Default 'OverFiles' -Allowed $script:PresetParallelValues -PresetName $Name

        MaxWidth                = Get-PresetValue -Overrides $Overrides -Key 'MaxWidth' -GlobalValue $DefaultMaxWidth
        AudioBitrate            = Get-PresetValue -Overrides $Overrides -Key 'AudioBitrate' -GlobalValue $DefaultAudioBitrate
        # Zero means no bitrate ceiling on the first encode and no size-cap retry pass.
        SizeCapMB               = Get-PresetValue -Overrides $Overrides -Key 'SizeCapMB' -GlobalValue $DefaultSizeCapMB
        SizeCapFallbackMaxWidth = Get-PresetValue -Overrides $Overrides -Key 'SizeCapFallbackMaxWidth' -GlobalValue $DefaultSizeCapFallbackMaxWidth
        MaxrateScale            = Get-PresetValue -Overrides $Overrides -Key 'MaxrateScale' -GlobalValue $DefaultMaxrateScale
        NvencCq                 = Get-PresetValue -Overrides $Overrides -Key 'NvencCq' -GlobalValue $DefaultNvencCq
        AmfQp                   = Get-PresetValue -Overrides $Overrides -Key 'AmfQp' -GlobalValue $DefaultAmfQp
        Crf                     = Get-PresetValue -Overrides $Overrides -Key 'Crf' -GlobalValue $DefaultCrf

        MinTrimMs               = Get-PresetValue -Overrides $Overrides -Key 'MinTrimMs' -GlobalValue $DefaultMinTrimMs
        MaxTrimMs               = Get-PresetValue -Overrides $Overrides -Key 'MaxTrimMs' -GlobalValue $DefaultMaxTrimMs

        CropMinPermille         = Get-PresetValue -Overrides $Overrides -Key 'CropMinPermille' -GlobalValue $DefaultCropMinPermille
        CropMaxPermille         = Get-PresetValue -Overrides $Overrides -Key 'CropMaxPermille' -GlobalValue $DefaultCropMaxPermille
        JpegQuality             = Get-PresetValue -Overrides $Overrides -Key 'JpegQuality' -GlobalValue $DefaultJpegQuality
        ConvertedJpegQuality    = Get-PresetValue -Overrides $Overrides -Key 'ConvertedJpegQuality' -GlobalValue $DefaultConvertedJpegQuality
        PngCompressionLevel     = Get-PresetValue -Overrides $Overrides -Key 'PngCompressionLevel' -GlobalValue $DefaultPngCompressionLevel
    }
}

# The lane layout that shipped before presets existed, expressed as preset overrides. When
# config.ini declares no [preset ...] sections these are synthesized, so an existing install
# keeps working and its folders keep their meaning.
#
# The old "convert" lane is deliberately absent: format conversion is now the Normalize stage
# that every preset runs, rather than a destination of its own.
function Get-BuiltInPresetOverrides {
    return [ordered]@{
        'default'    = @{ VideoCopies = '8';   ImageCopies = '8'; CopiesAlternate = '7' }
        'videoclean' = @{ VideoCopies = '1';   ImageCopies = '0' }
        'imageclean' = @{ VideoCopies = '0';   ImageCopies = '1' }
        'images'     = @{ VideoCopies = '0';   ImageCopies = '20' }
        'sets'       = @{ VideoCopies = '10';  ImageCopies = '10'; Grouping = 'PerSource'; SizeCapMB = '0' }
        'setbatch'   = @{ VideoCopies = '1';   ImageCopies = '1';  Grouping = 'PerSet'; SetCount = '10'; Batch = 'PerGroup'; SizeCapMB = '0' }
        'assetstore' = @{ VideoCopies = '1';   ImageCopies = '1';  Grouping = 'PerSet'; SetCount = '15'; Batch = 'PerGroup'; SizeCapMB = '0'; Manifest = 'true'; MinTrimMs = '10'; MaxTrimMs = '40' }
        'long'       = @{ VideoCopies = '3';   ImageCopies = '0';  Segment = 'true'; NvencCq = '28'; AmfQp = '26' }
    }
}

# All enabled presets, built once per process (and once per parallel worker runspace, which
# re-loads this script).
function Get-PipelinePresets {
    if ($script:PipelinePresets) { return $script:PipelinePresets }

    $sections = $script:ConfigPresetSections
    if ($sections -and $sections.Count -gt 0) {
        $overrides = $sections
    }
    else {
        $overrides = Get-BuiltInPresetOverrides
    }

    $presets = New-Object System.Collections.Generic.List[object]
    foreach ($name in @($overrides.Keys)) {
        $preset = New-PipelinePreset -Name $name -Overrides $overrides[$name]
        if ($preset.Enabled) {
            $presets.Add($preset) | Out-Null
        }
    }

    $script:PipelinePresets = $presets.ToArray()
    return $script:PipelinePresets
}

function Get-PipelinePreset {
    param([Parameter(Mandatory = $true)][string]$Name)

    return @(Get-PipelinePresets | Where-Object { $_.Name -eq $Name })[0]
}

# Every preset owns the same folder shape under its own name. This replaces the fifty
# hardcoded per-lane path properties the nine pipelines needed.
function Get-PresetWorkspacePaths {
    param(
        [Parameter(Mandatory = $true)][string]$PresetName,
        [Parameter(Mandatory = $true)][string]$WorkspaceName
    )

    $presetRoot = Join-Path (Join-Path $PipelineRoot $PresetName) $WorkspaceName
    $archiveRoot = Join-Path (Join-Path $ArchiveRootDir $PresetName) $WorkspaceName

    return [pscustomobject]@{
        PresetName    = $PresetName
        WorkspaceName = $WorkspaceName
        InputDir      = Join-Path $presetRoot "input"
        OutputDir     = Join-Path $presetRoot "output"
        OriginalDir   = Join-Path $presetRoot "original"
        FailedDir     = Join-Path $presetRoot "failed"
        WorkDir       = Join-Path $presetRoot "work"
        ArchiveDir    = Join-Path $archiveRoot "output"
    }
}

$DefaultRootDir = Join-Path $PipelineRoot "default"
$VideoCleanRootDir = Join-Path $PipelineRoot "videoclean"
$LogsDir = Join-Path $PipelineRoot "logs"
$RemuxRootDir = Join-Path $PipelineRoot "convert"
$LongRootDir = Join-Path $PipelineRoot "long"
$ImageBulkRootDir = Join-Path $PipelineRoot "images"
$ImageCleanRootDir = Join-Path $PipelineRoot "imageclean"
$SetRootDir = Join-Path $PipelineRoot "sets"
$SetBatchRootDir = Join-Path $PipelineRoot "setbatch"
$AssetStoreRootDir = Join-Path $PipelineRoot "assetstore"
$ArchiveRootDir = Join-Path $PipelineRoot "archive"






$VideoExtensions = @(".mp4", ".mov", ".mkv", ".webm", ".avi")
$ImageExtensions = @(".jpg", ".jpeg", ".png", ".webp", ".heic", ".heif")
$TempExtensions = @(".crdownload", ".tmp", ".part", ".download")

# Convert pipeline: source formats that get rewritten into widely supported ones.
$RemuxVideoSourceExtensions = @(".mov")
$RemuxImageSourceExtensions = @(".heic", ".heif")
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
$script:PresetEntryCount = 0
$script:LastSetBatchSignature = $null
$script:LastAssetStoreSignature = $null
$script:WorkspaceRuntimeState = @{}

function Use-PipelineWorkspace {
    param(
        [Parameter(Mandatory = $true)]
        [string]$WorkspaceName
    )

    $script:CurrentWorkspaceName = $WorkspaceName

    if (-not $script:WorkspaceRuntimeState.ContainsKey($WorkspaceName)) {
        $script:WorkspaceRuntimeState[$WorkspaceName] = @{
            LastArchiveCheck = $null
            BatchSignatures = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
        }
    }

    $state = $script:WorkspaceRuntimeState[$WorkspaceName]
    $script:LastArchiveCheck = $state.LastArchiveCheck

    # One settle signature per batch preset, so two batch presets in the same workspace
    # cannot clobber each other's debounce state the way the old two fixed slots could.
    if (-not $state.BatchSignatures) {
        $state.BatchSignatures = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
    }
    $script:BatchSignatures = $state.BatchSignatures
}

function Save-PipelineWorkspaceState {
    if ([string]::IsNullOrWhiteSpace($script:CurrentWorkspaceName)) {
        return
    }

    if (-not $script:WorkspaceRuntimeState.ContainsKey($script:CurrentWorkspaceName)) {
        $script:WorkspaceRuntimeState[$script:CurrentWorkspaceName] = @{}
    }

    $state = $script:WorkspaceRuntimeState[$script:CurrentWorkspaceName]
    $state.LastArchiveCheck = $script:LastArchiveCheck
    $state.BatchSignatures = $script:BatchSignatures
}


function Initialize-Folders {
    $directories = New-Object System.Collections.Generic.List[string]
    $directories.Add($LogsDir) | Out-Null

    if (-not (Test-Path -LiteralPath $LogsDir)) {
        New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null
    }

    Move-LegacyPipelineAssetsToDefaultWorkspace

    # Each preset owns one folder tree per workspace. A preset that takes no video or no
    # images still gets the full tree, because its copy counts can change at any time.
    foreach ($preset in Get-PipelinePresets) {
        foreach ($workspaceName in $WorkspaceNames) {
            $paths = Get-PresetWorkspacePaths -PresetName $preset.Name -WorkspaceName $workspaceName
            $directories.Add($paths.InputDir) | Out-Null
            $directories.Add($paths.OutputDir) | Out-Null
            $directories.Add($paths.OriginalDir) | Out-Null
            $directories.Add($paths.FailedDir) | Out-Null
        }
    }

    foreach ($directory in $directories) {
        if (-not (Test-Path -LiteralPath $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
    }

    Use-PipelineWorkspace -WorkspaceName $DefaultWorkspaceName
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
        Write-Log "Video encoder: h264_nvenc (NVIDIA GPU, preset $NvencPreset, CQ $DefaultNvencCq)"
        return
    }

    if ($PreferAmf -and (Test-AmfEncoderAvailable)) {
        $script:UseAmf = $true
        Write-Log "Video encoder: h264_amf (AMD GPU, quality $AmfQuality, QP $DefaultAmfQp)"
        return
    }

    if ($PreferNvenc -or $PreferAmf) {
        Write-Log "No usable GPU encoder (NVENC/AMF) found in FFmpeg; falling back to libx264 (CPU)." "WARN"
    }

    Write-Log "Video encoder: libx264 (CPU, preset $X264Preset, CRF $DefaultCrf)"
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
            "-preset", $script:X264Preset,
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
            @{ Quality = 30; MaxWidth = $DefaultMaxWidth; Bitrate = 0 },
            @{ Quality = 32; MaxWidth = $DefaultMaxWidth; Bitrate = 0 },
            @{ Quality = 34; MaxWidth = $FallbackMaxWidth; Bitrate = 0 },
            @{ Quality = 36; MaxWidth = $FallbackMaxWidth; Bitrate = 0 }
        )
    }

    if ($script:UseAmf) {
        return @(
            @{ Quality = 28; MaxWidth = $DefaultMaxWidth; Bitrate = 0 },
            @{ Quality = 30; MaxWidth = $DefaultMaxWidth; Bitrate = 0 },
            @{ Quality = 32; MaxWidth = $FallbackMaxWidth; Bitrate = 0 },
            @{ Quality = 34; MaxWidth = $FallbackMaxWidth; Bitrate = 0 }
        )
    }

    return @(
        @{ Quality = 28; MaxWidth = $DefaultMaxWidth; Bitrate = 0 },
        @{ Quality = 30; MaxWidth = $DefaultMaxWidth; Bitrate = 0 },
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

$script:OutputNameDescriptors = @(
    "autumn", "bright", "calm", "cedar", "clear", "coastal", "daily", "evening",
    "fresh", "garden", "golden", "harbor", "local", "maple", "meadow", "modern",
    "morning", "natural", "open", "quiet", "river", "silver", "simple", "spring",
    "studio", "summer", "sunny", "travel", "urban", "warm", "weekend", "winter"
)

$script:OutputNameSubjects = @(
    "album", "capture", "clip", "collection", "frame", "gallery", "image", "media",
    "memory", "moment", "photo", "picture", "post", "project", "scene", "shot",
    "snapshot", "story", "take", "update", "upload", "video", "view", "work"
)

$script:OutputNameContexts = @(
    "archive", "backup", "camera", "desktop", "draft", "edit", "export", "folder",
    "home", "inbox", "library", "mobile", "notes", "phone", "review", "share",
    "social", "temp", "today", "trip", "week", "workshop"
)

function Get-RandomInt {
    param(
        [int]$Minimum = 0,

        [Parameter(Mandatory = $true)]
        [int]$Maximum
    )

    if ($Maximum -le $Minimum) {
        throw "Maximum must be greater than minimum."
    }

    $range = $Maximum - $Minimum
    $limit = [int]::MaxValue - ([int]::MaxValue % $range)
    $bytes = New-Object byte[] 4
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()

    try {
        do {
            $rng.GetBytes($bytes)
            $value = [System.BitConverter]::ToInt32($bytes, 0) -band 0x7fffffff
        } while ($value -ge $limit)
    }
    finally {
        $rng.Dispose()
    }

    return $Minimum + ($value % $range)
}

function Get-RandomChoice {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Values
    )

    return $Values[(Get-RandomInt -Maximum $Values.Count)]
}

function Convert-OutputNamePart {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [int]$Style
    )

    if ($Style -eq 1) {
        return ($Value.Substring(0, 1).ToUpperInvariant() + $Value.Substring(1))
    }

    return $Value
}

function New-RegularRandomNumberText {
    $digits = Get-RandomInt -Minimum 2 -Maximum 7
    $minimum = [int][Math]::Pow(10, $digits - 1)
    $maximum = [int][Math]::Pow(10, $digits)

    return [string](Get-RandomInt -Minimum $minimum -Maximum $maximum)
}

function Join-RegularRandomNameParts {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Parts
    )

    $separator = Get-RandomChoice -Values @("-", "_", " ")
    return ($Parts -join $separator)
}

function New-RegularRandomName {
    $style = Get-RandomInt -Maximum 2
    $descriptor = Convert-OutputNamePart -Value (Get-RandomChoice -Values $script:OutputNameDescriptors) -Style $style
    $subject = Convert-OutputNamePart -Value (Get-RandomChoice -Values $script:OutputNameSubjects) -Style $style
    $context = Convert-OutputNamePart -Value (Get-RandomChoice -Values $script:OutputNameContexts) -Style $style
    $number = New-RegularRandomNumberText

    switch (Get-RandomInt -Maximum 12) {
        0 { return (Join-RegularRandomNameParts -Parts @($descriptor, $subject)) }
        1 { return (Join-RegularRandomNameParts -Parts @($subject, $number)) }
        2 { return (Join-RegularRandomNameParts -Parts @($descriptor, $subject, $number)) }
        3 { return (Join-RegularRandomNameParts -Parts @($context, $subject)) }
        4 { return (Join-RegularRandomNameParts -Parts @($subject, $context, $number)) }
        5 { return (Join-RegularRandomNameParts -Parts @($descriptor, $context, $subject)) }
        6 { return (Join-RegularRandomNameParts -Parts @($context, $number)) }
        7 { return (Join-RegularRandomNameParts -Parts @($subject, $descriptor)) }
        8 { return (Join-RegularRandomNameParts -Parts @($context, $descriptor, $number)) }
        9 { return (Join-RegularRandomNameParts -Parts @($descriptor, $number)) }
        10 { return (Join-RegularRandomNameParts -Parts @($subject, $context)) }
        default { return (Join-RegularRandomNameParts -Parts @($descriptor, $context, $subject, $number)) }
    }
}

function New-IPhoneRandomFilePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,

        [Parameter(Mandatory = $true)]
        [string]$Extension
    )

    $normalizedExtension = $Extension.ToUpperInvariant()
    if (-not $normalizedExtension.StartsWith(".")) {
        $normalizedExtension = ".{0}" -f $normalizedExtension
    }

    do {
        # Match the four-digit naming convention used by the iPhone Camera app
        # while choosing the number independently for every generated file.
        $fileName = "IMG_{0:D4}{1}" -f (Get-RandomInt -Minimum 1 -Maximum 10000), $normalizedExtension
        $path = Join-Path $Directory $fileName
    } while (Test-Path -LiteralPath $path)

    return $path
}

function New-RegularRandomDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory
    )

    do {
        $directoryName = New-RegularRandomName
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
        $fileName = "{0}-{1}{2}" -f $baseName, (New-RegularRandomName), $extension
        $destination = Join-Path $Directory $fileName
    } while (Test-Path -LiteralPath $destination)

    return $destination
}

function Move-LegacyDirectoryContents {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceDirectory,

        [Parameter(Mandatory = $true)]
        [string]$DestinationDirectory,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    if (-not (Test-Path -LiteralPath $SourceDirectory)) {
        return 0
    }

    $sourceFullPath = [System.IO.Path]::GetFullPath($SourceDirectory).TrimEnd('\')
    $destinationFullPath = [System.IO.Path]::GetFullPath($DestinationDirectory).TrimEnd('\')
    if ($sourceFullPath.Equals($destinationFullPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        return 0
    }

    $items = @(Get-ChildItem -LiteralPath $SourceDirectory -Force -ErrorAction SilentlyContinue)
    if ($items.Count -eq 0) {
        return 0
    }

    if (-not (Test-Path -LiteralPath $DestinationDirectory)) {
        New-Item -ItemType Directory -Path $DestinationDirectory -Force | Out-Null
    }

    $moved = 0
    foreach ($item in $items) {
        if ($item.PSIsContainer -and ($WorkspaceNames -contains $item.Name)) {
            continue
        }

        $itemFullPath = [System.IO.Path]::GetFullPath($item.FullName).TrimEnd('\')
        if ($itemFullPath.Equals($destinationFullPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }
        if ($destinationFullPath.StartsWith($itemFullPath + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        try {
            $destination = Get-UniqueDestinationPath -Directory $DestinationDirectory -OriginalFileName $item.Name
            Move-Item -LiteralPath $item.FullName -Destination $destination -Force
            $moved++
        }
        catch {
            Write-Log "Could not migrate legacy $Label item '$($item.FullName)': $($_.Exception.Message)" "WARN"
        }
    }

    if ($moved -gt 0) {
        Write-Log "Migrated $moved legacy $Label item(s) to $DestinationDirectory."
    }

    return $moved
}

# Upgrade path from installs that predate workspaces, where a lane's folders sat directly
# under <root>\<lane>\ instead of <root>\<lane>\<workspace>\. Everything found is moved into
# the default workspace. Also handles the oldest layout, where the default lane's folders sat
# at the pipeline root.
function Move-LegacyPipelineAssetsToDefaultWorkspace {
    $pairs = New-Object System.Collections.Generic.List[object]

    $rootPaths = Get-PresetWorkspacePaths -PresetName "default" -WorkspaceName $DefaultWorkspaceName
    $pairs.Add(@{ Label = "root default input";    Old = (Join-Path $PipelineRoot "input");    New = $rootPaths.InputDir }) | Out-Null
    $pairs.Add(@{ Label = "root default output";   Old = (Join-Path $PipelineRoot "output");   New = $rootPaths.OutputDir }) | Out-Null
    $pairs.Add(@{ Label = "root default original"; Old = (Join-Path $PipelineRoot "original"); New = $rootPaths.OriginalDir }) | Out-Null
    $pairs.Add(@{ Label = "root default failed";   Old = (Join-Path $PipelineRoot "failed");   New = $rootPaths.FailedDir }) | Out-Null

    foreach ($preset in Get-PipelinePresets) {
        $paths = Get-PresetWorkspacePaths -PresetName $preset.Name -WorkspaceName $DefaultWorkspaceName
        $presetRoot = Join-Path $PipelineRoot $preset.Name

        $pairs.Add(@{ Label = "$($preset.Name) input";    Old = (Join-Path $presetRoot "input");    New = $paths.InputDir }) | Out-Null
        $pairs.Add(@{ Label = "$($preset.Name) output";   Old = (Join-Path $presetRoot "output");   New = $paths.OutputDir }) | Out-Null
        $pairs.Add(@{ Label = "$($preset.Name) original"; Old = (Join-Path $presetRoot "original"); New = $paths.OriginalDir }) | Out-Null
        $pairs.Add(@{ Label = "$($preset.Name) failed";   Old = (Join-Path $presetRoot "failed");   New = $paths.FailedDir }) | Out-Null

        $pairs.Add(@{
            Label = "archive $($preset.Name)"
            Old   = (Join-Path $ArchiveRootDir $preset.Name)
            New   = $paths.ArchiveDir
        }) | Out-Null
    }

    $pairs.Add(@{
        Label = "legacy archive output"
        Old   = (Join-Path $ArchiveRootDir "output")
        New   = $rootPaths.ArchiveDir
    }) | Out-Null

    $totalMoved = 0
    foreach ($pair in $pairs) {
        $totalMoved += Move-LegacyDirectoryContents -SourceDirectory $pair.Old -DestinationDirectory $pair.New -Label $pair.Label
    }

    if ($totalMoved -gt 0) {
        Write-Log "Legacy workspace migration complete: $totalMoved item(s) moved into $DefaultWorkspaceName."
    }
}

function Get-OutputArchiveCutoffTime {
    return (Get-Date).AddHours(-1 * $ArchiveAgeHours)
}

function Get-AssetRetentionCutoffTime {
    return (Get-Date).AddDays(-1 * $AssetRetentionDays)
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
    $creationTime = $File.CreationTime
    Move-Item -LiteralPath $File.FullName -Destination $destination -Force
    try {
        (Get-Item -LiteralPath $destination).CreationTime = $creationTime
    }
    catch {
        Write-Log "Could not preserve archive creation time for '$destination': $($_.Exception.Message)" "WARN"
    }
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
    $creationTime = $Directory.CreationTime
    Move-Item -LiteralPath $Directory.FullName -Destination $destination -Force
    try {
        (Get-Item -LiteralPath $destination).CreationTime = $creationTime
    }
    catch {
        Write-Log "Could not preserve archive creation time for '$destination': $($_.Exception.Message)" "WARN"
    }
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


# Retention sweeps every preset's archive folder for the current workspace. The images
# preset is excluded, matching the long-standing rule that archived images are kept.
function Get-ArchiveRetentionTargets {
    $targets = New-Object System.Collections.Generic.List[object]

    foreach ($preset in Get-PipelinePresets) {
        if ($preset.Name -eq "images") { continue }

        $paths = Get-PresetWorkspacePaths -PresetName $preset.Name -WorkspaceName $script:CurrentWorkspaceName
        $targets.Add([pscustomobject]@{
            TargetDirectory = $paths.ArchiveDir
            Label = "archive $($preset.Name)"
        }) | Out-Null
    }

    return $targets.ToArray()
}

function Get-LegacyArchiveRetentionTargets {
    return @(
        [pscustomobject]@{
            TargetDirectory = Join-Path $ArchiveRootDir "output"
            Label = "legacy archive output"
            ExcludedNames = @()
        },
        [pscustomobject]@{
            TargetDirectory = Join-Path $ArchiveRootDir "default"
            Label = "legacy archive default"
            ExcludedNames = $WorkspaceNames
        },
        [pscustomobject]@{
            TargetDirectory = Join-Path $ArchiveRootDir "videoclean"
            Label = "legacy archive videoclean"
            ExcludedNames = $WorkspaceNames
        },
        [pscustomobject]@{
            TargetDirectory = Join-Path $ArchiveRootDir "convert"
            Label = "legacy archive convert"
            ExcludedNames = $WorkspaceNames
        },
        [pscustomobject]@{
            TargetDirectory = Join-Path $ArchiveRootDir "long"
            Label = "legacy archive long"
            ExcludedNames = $WorkspaceNames
        },
        [pscustomobject]@{
            TargetDirectory = Join-Path $ArchiveRootDir "sets"
            Label = "legacy archive sets"
            ExcludedNames = $WorkspaceNames
        },
        [pscustomobject]@{
            TargetDirectory = Join-Path $ArchiveRootDir "setbatch"
            Label = "legacy archive setbatch"
            ExcludedNames = $WorkspaceNames
        },
        [pscustomobject]@{
            TargetDirectory = Join-Path $ArchiveRootDir "assetstore"
            Label = "legacy archive assetstore"
            ExcludedNames = $WorkspaceNames
        }
    )
}

# Retention sweeps the original, failed and work folders of every preset in the current
# workspace. The images preset keeps its assets, matching the archive rule above.
function Get-PipelineAssetRetentionTargets {
    $targets = New-Object System.Collections.Generic.List[object]

    foreach ($preset in Get-PipelinePresets) {
        if ($preset.Name -eq "images") { continue }

        $paths = Get-PresetWorkspacePaths -PresetName $preset.Name -WorkspaceName $script:CurrentWorkspaceName

        $targets.Add([pscustomobject]@{ TargetDirectory = $paths.OriginalDir; Label = "$($preset.Name) original" }) | Out-Null
        $targets.Add([pscustomobject]@{ TargetDirectory = $paths.FailedDir;   Label = "$($preset.Name) failed" }) | Out-Null

        if ($preset.Segment) {
            $targets.Add([pscustomobject]@{ TargetDirectory = $paths.WorkDir; Label = "$($preset.Name) work" }) | Out-Null
        }
    }

    return $targets.ToArray()
}

# Sync folders sit above the workspace level: <root>\sync and <root>\<preset>\sync. They are
# written by the upload scripts rather than the watcher, which only ages them out.
function Get-SyncRetentionTargets {
    $targets = New-Object System.Collections.Generic.List[object]

    $targets.Add([pscustomobject]@{
        TargetDirectory = Join-Path $PipelineRoot "sync"
        Label = "root sync"
    }) | Out-Null

    foreach ($preset in Get-PipelinePresets) {
        $targets.Add([pscustomobject]@{
            TargetDirectory = Join-Path (Join-Path $PipelineRoot $preset.Name) "sync"
            Label = "$($preset.Name) sync"
        }) | Out-Null
    }

    $targets.Add([pscustomobject]@{
        TargetDirectory = Join-Path $PipelineRoot ".sync-parts"
        Label = "sync parts"
    }) | Out-Null

    return $targets.ToArray()
}

function Invoke-AssetRetentionCleanup {
    if ($AssetRetentionDays -le 0) {
        return
    }

    $cutoffTime = Get-AssetRetentionCutoffTime
    Write-Log "Running asset retention cleanup (entries created before $($cutoffTime.ToString('yyyy-MM-dd HH:mm:ss')))."

    foreach ($target in Get-ArchiveRetentionTargets) {
        [void](Invoke-RetentionCleanup -TargetDirectory $target.TargetDirectory -Label $target.Label -CutoffTime $cutoffTime)
    }

    foreach ($target in Get-PipelineAssetRetentionTargets) {
        [void](Invoke-RetentionCleanup -TargetDirectory $target.TargetDirectory -Label $target.Label -CutoffTime $cutoffTime)
    }

    if ($script:CurrentWorkspaceName -eq $DefaultWorkspaceName) {
        foreach ($target in Get-LegacyArchiveRetentionTargets) {
            [void](Invoke-RetentionCleanup -TargetDirectory $target.TargetDirectory -Label $target.Label -CutoffTime $cutoffTime -ExcludedNames $target.ExcludedNames)
        }

        foreach ($target in Get-SyncRetentionTargets) {
            [void](Invoke-RetentionCleanup -TargetDirectory $target.TargetDirectory -Label $target.Label -CutoffTime $cutoffTime)
        }
    }
}

function Get-RetentionEntryTime {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.FileSystemInfo]$Entry
    )

    if ($Entry.CreationTime -le $Entry.LastWriteTime) {
        return $Entry.CreationTime
    }

    return $Entry.LastWriteTime
}

function Invoke-RetentionCleanup {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TargetDirectory,

        [Parameter(Mandatory = $true)]
        [string]$Label,

        [Parameter(Mandatory = $true)]
        [datetime]$CutoffTime,

        [string[]]$ExcludedNames = @()
    )

    if (-not (Test-Path -LiteralPath $TargetDirectory)) {
        return 0
    }

    $excludedNameSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $ExcludedNames) {
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            [void]$excludedNameSet.Add($name)
        }
    }

    $count = 0
    $entries = @(Get-ChildItem -LiteralPath $TargetDirectory -Force -ErrorAction SilentlyContinue)
    foreach ($entry in $entries) {
        if ($excludedNameSet.Contains($entry.Name)) {
            continue
        }

        if ((Get-RetentionEntryTime -Entry $entry) -gt $CutoffTime) {
            continue
        }

        try {
            Remove-Item -LiteralPath $entry.FullName -Recurse -Force
            $count++
        }
        catch {
            Write-Log "Could not delete expired $Label entry '$($entry.FullName)': $($_.Exception.Message)" "WARN"
        }
    }

    if ($count -gt 0) {
        Write-Log "Deleted $count expired entr$(if ($count -eq 1) { 'y' } else { 'ies' }) from $Label."
    }

    return $count
}

function Invoke-OutputArchiveIfDue {
    if ((-not $ArchiveEnabled) -and ($AssetRetentionDays -le 0)) {
        return
    }

    $now = Get-Date
    if ($script:LastArchiveCheck -and (($now - $script:LastArchiveCheck).TotalMinutes -lt $ArchiveCheckIntervalMinutes)) {
        return
    }

    $script:LastArchiveCheck = $now
    if ($ArchiveEnabled) {
        $cutoffTime = Get-OutputArchiveCutoffTime

        Write-Log "Running scheduled output archive check (older than $ArchiveAgeHours hours)."

        # A flat preset archives loose files; a grouped preset moves whole output folders,
        # which is exactly the distinction its Grouping option already encodes.
        foreach ($preset in Get-PipelinePresets) {
            $paths = Get-PresetWorkspacePaths -PresetName $preset.Name -WorkspaceName $script:CurrentWorkspaceName

            if ($preset.Grouping -eq "Flat") {
                [void](Invoke-FlatOutputArchive -SourceDirectory $paths.OutputDir -ArchiveDirectory $paths.ArchiveDir -Label $preset.Name -CutoffTime $cutoffTime)
            }
            else {
                [void](Invoke-DirectoryOutputArchive -SourceDirectory $paths.OutputDir -ArchiveDirectory $paths.ArchiveDir -Label $preset.Name -CutoffTime $cutoffTime)
            }
        }
    }

    Invoke-AssetRetentionCleanup
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
        "-read_intervals", "%+#1",
        "-show_frames",
        "-show_entries", "frame=width,height:frame_side_data=rotation",
        "-of", "json",
        $Path
    )

    $output = Invoke-ExternalTool -Command $script:FFprobePath -Arguments $arguments
    $probe = (($output | Out-String).Trim() | ConvertFrom-Json)
    $frame = @($probe.frames)[0]
    if ($null -eq $frame -or [int]$frame.width -le 0 -or [int]$frame.height -le 0) {
        throw "Unable to read image dimensions from ffprobe for: $Path"
    }

    $width = [int]$frame.width
    $height = [int]$frame.height
    $rotation = 0.0
    foreach ($sideData in @($frame.side_data_list)) {
        if ($null -ne $sideData -and $sideData.PSObject.Properties.Name -contains "rotation") {
            $rotation = [double]$sideData.rotation
            break
        }
    }

    # FFmpeg auto-rotates image inputs before filters run. A quarter-turn stored
    # in EXIF/display-matrix metadata therefore swaps the filter's input axes.
    $normalizedRotation = (($rotation % 360.0) + 360.0) % 360.0
    if ([Math]::Abs($normalizedRotation - 90.0) -lt 0.5 -or [Math]::Abs($normalizedRotation - 270.0) -lt 0.5) {
        $temp = $width
        $width = $height
        $height = $temp
    }

    return [pscustomobject]@{
        Width = $width
        Height = $height
    }
}

function Test-IsHeicContainer {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $arguments = @(
        "-v", "error",
        "-show_entries", "format_tags=major_brand,compatible_brands",
        "-of", "default=noprint_wrappers=1:nokey=1",
        $Path
    )

    try {
        $brandText = ((Invoke-ExternalTool -Command $script:FFprobePath -Arguments $arguments | Out-String).Trim()).ToLowerInvariant()
        return ($brandText -match "(^|\s)(heic|heix|hevc|hevx|mif1|msf1)(\s|$)")
    }
    catch {
        return $false
    }
}

function Test-IsHeicImage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    return (($extension -in @(".heic", ".heif")) -or (Test-IsHeicContainer -Path $Path))
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

    if (-not (Test-IsHeicImage -Path $Path)) {
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

# Decides how much may be trimmed off the end of a video variant. ConfiguredMinMs and
# ConfiguredMaxMs default to the global trim range; a preset passes its own to override it.
function Get-TrimRange {
    param(
        [Parameter(Mandatory = $true)]
        [double]$DurationSeconds,

        [int]$ConfiguredMinMs = -1,

        [int]$ConfiguredMaxMs = -1
    )

    if ($ConfiguredMinMs -lt 0) { $ConfiguredMinMs = $DefaultMinTrimMs }
    if ($ConfiguredMaxMs -lt 0) { $ConfiguredMaxMs = $DefaultMaxTrimMs }

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

    $safeConfiguredMax = [int][Math]::Min($ConfiguredMaxMs, $durationMs - 1000)
    if ($safeConfiguredMax -lt $ConfiguredMinMs) {
        return [pscustomobject]@{
            CanTrim = $false
            MinMs = 0
            MaxMs = 0
            Reason = "configured trim range would make output too short"
        }
    }

    return [pscustomobject]@{
        CanTrim = $true
        MinMs = $ConfiguredMinMs
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

# Produces one randomly named .mp4 variant: trims the tail, encodes, strips metadata, and
# optionally enforces a size cap.
#
# Every video lane goes through here. The lane-specific policy is expressed by two arguments:
#   MaxVideoBitrateKbps  0 means no bitrate ceiling on the first encode.
#   MaxSizeMegabytes     0 skips the size-cap retry pass entirely.
# The set, set-batch and asset-store lanes pass 0 for both; the default lane passes neither.
#
# LogLabel identifies the lane in the log, e.g. "Video variant 3" or "Set video variant 2".
function New-VideoVariant {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [double]$DurationSeconds,

        [Parameter(Mandatory = $true)]
        [int]$TrimMs,

        [Parameter(Mandatory = $true)]
        [string]$LogLabel,

        # Encoder settings come from the caller's preset. Below zero, or empty, falls back to
        # the global default, which is what the size-cap retry pass relies on.
        [int]$QualityValue = -1,

        [int]$MaxWidthValue = -1,

        [string]$AudioBitrateValue = "",

        [int]$MaxVideoBitrateKbps = 0,

        [double]$MaxSizeMegabytes = 0,

        [int]$SizeCapFallbackMaxWidth = 0
    )

    if ($QualityValue -lt 0) {
        $QualityValue = if ($script:UseNvenc) { $DefaultNvencCq } elseif ($script:UseAmf) { $DefaultAmfQp } else { $DefaultCrf }
    }
    if ($MaxWidthValue -lt 0) { $MaxWidthValue = $DefaultMaxWidth }
    if ([string]::IsNullOrWhiteSpace($AudioBitrateValue)) { $AudioBitrateValue = $DefaultAudioBitrate }

    $outputPath = New-IPhoneRandomFilePath -Directory $OutputDirectory -Extension ".mp4"
    $trimSeconds = $TrimMs / 1000.0
    $targetDuration = [Math]::Max(0.1, $DurationSeconds - $trimSeconds)
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $targetDurationText = $targetDuration.ToString("0.###", $culture)

    Write-Log "$LogLabel trim: ${TrimMs}ms, target duration: ${targetDurationText}s"

    $encodeArgs = @{
        InputPath           = $InputPath
        OutputPath          = $outputPath
        QualityValue        = $QualityValue
        MaxWidthValue       = $MaxWidthValue
        AudioBitrateValue   = $AudioBitrateValue
        DurationSeconds     = $targetDuration
        MaxVideoBitrateKbps = $MaxVideoBitrateKbps
    }

    Invoke-VideoEncode @encodeArgs

    Clear-Metadata -Path $outputPath

    if ($MaxSizeMegabytes -gt 0) {
        Invoke-OutputSizeCap -OutputPath $outputPath -MaxSizeMegabytes $MaxSizeMegabytes -FallbackMaxWidth $SizeCapFallbackMaxWidth -SourceInputPath $InputPath -SegmentDurationSeconds $DurationSeconds -TrimMs $TrimMs
    }

    Write-Log "Created $LogLabel output: $outputPath"

    return $outputPath
}







# Produces one randomly named image variant: a tiny randomized crop scaled back to the
# original dimensions, so each copy differs while looking identical, with metadata stripped.
#
# Quality is passed in rather than read from config because the lanes disagree today. The bulk
# lane uses the configured quality, the clean lane uses 4 for HEIC sources and 2 otherwise, and
# the set family hardcodes 2. Unifying those values is a separate change.
#
# RemoveOutputOnFailure deletes a partial output before rethrowing. Only the image-clean lane
# does that today; the others leave partial outputs for their caller to reap.
function New-ImageVariant {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputDirectory,

        [Parameter(Mandatory = $true)]
        [string]$OutputExtension,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Dimensions,

        [Parameter(Mandatory = $true)]
        [string]$LogLabel,

        [string]$JpegQuality = "2",

        [string]$WebpQuality = "92",

        [int]$PngCompressionLevel = 6,

        # Below zero means "use the global crop range". A preset passes its own.
        [int]$CropMinPermille = -1,

        [int]$CropMaxPermille = -1,

        [switch]$RemoveOutputOnFailure
    )

    if ($CropMinPermille -lt 0) { $CropMinPermille = $DefaultCropMinPermille }
    if ($CropMaxPermille -lt 0) { $CropMaxPermille = $DefaultCropMaxPermille }

    $outputPath = New-IPhoneRandomFilePath -Directory $OutputDirectory -Extension $OutputExtension
    $width = $Dimensions.Width
    $height = $Dimensions.Height

    $arguments = @(
        "-y",
        "-hide_banner",
        "-loglevel", "error",
        "-i", $InputPath,
        "-frames:v", "1",
        "-map_metadata", "-1"
    )

    # Cropping a very small image would visibly degrade it, so leave those untouched.
    if ($width -ge 200 -and $height -ge 200) {
        $cropPermille = Get-Random -Minimum $CropMinPermille -Maximum ($CropMaxPermille + 1)
        $cropPixelsX = [Math]::Max(1, [int][Math]::Floor($width * $cropPermille / 1000))
        $cropPixelsY = [Math]::Max(1, [int][Math]::Floor($height * $cropPermille / 1000))
        $cropWidth = [Math]::Max(1, $width - ($cropPixelsX * 2))
        $cropHeight = [Math]::Max(1, $height - ($cropPixelsY * 2))
        $offsetX = Get-Random -Minimum 0 -Maximum (($cropPixelsX * 2) + 1)
        $offsetY = Get-Random -Minimum 0 -Maximum (($cropPixelsY * 2) + 1)
        $filter = "crop=${cropWidth}:${cropHeight}:${offsetX}:${offsetY},scale=${width}:${height}"
        $arguments += @("-filter_complex", "[0:v:0]$filter[v]", "-map", "[v]")
        Write-Log "$LogLabel crop: ${cropWidth}x${cropHeight}+${offsetX}+${offsetY}, restored to ${width}x${height}"
    }
    else {
        Write-Log "$LogLabel skipping crop because image is small: ${width}x${height}" "WARN"
    }

    if ($OutputExtension -in @(".jpg", ".jpeg")) {
        $arguments += @("-q:v", $JpegQuality)
    }
    elseif ($OutputExtension -eq ".webp") {
        $arguments += @("-quality", $WebpQuality)
    }
    elseif ($OutputExtension -eq ".png") {
        $arguments += @("-compression_level", ([string]$PngCompressionLevel))
    }

    $arguments += @($outputPath)

    try {
        Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
        Clear-Metadata -Path $outputPath
    }
    catch {
        if ($RemoveOutputOnFailure) {
            Remove-GeneratedOutputs -Paths @($outputPath)
        }
        throw
    }

    Write-Log "Created $LogLabel output: $outputPath"

    return $outputPath
}
















# ---------------------------------------------------------------------------
# Asset store manifest pipeline
# ---------------------------------------------------------------------------
# Treats everything dropped in assetstore\<workspace>\input as one batch. Produces
# SetCount randomly named sets, each holding one processed,
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
        "-map", "0:a:0?",
        "-dn",
        "-c", "copy",
        "-map_metadata", "-1",
        "-movflags", "+faststart",
        $OutputPath
    )

    Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
}




function Get-LongSegmentPlan {
    param(
        [Parameter(Mandatory = $true)]
        [double]$DurationSeconds
    )

    $durationMs = [int][Math]::Floor($DurationSeconds * 1000)
    $targetMs = [int]($DefaultSegmentTargetSeconds * 1000)
    $minMs = [int]($DefaultSegmentMinSeconds * 1000)
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

    $audioBitrateText = ($DefaultAudioBitrate -replace "[^0-9.]", "")
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
        "-map", "0:a:0?",
        "-dn",
        "-c", "copy",
        "-map_metadata", "-1",
        "-movflags", "+faststart",
        $OutputPath
    )

    Invoke-ExternalTool -Command $script:FFmpegPath -Arguments $arguments | Out-Null
}

# The single H.264/AAC ffmpeg invocation used by every video lane, and by the size-cap retry
# pass. StartSeconds below zero encodes from the beginning, DurationSeconds at or below zero
# encodes to the end, and MaxVideoBitrateKbps of zero applies no bitrate ceiling.
function Invoke-VideoEncode {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [int]$QualityValue,

        [int]$MaxWidthValue,

        [double]$StartSeconds = -1,

        [double]$DurationSeconds = -1,

        [int]$MaxVideoBitrateKbps = 0,

        [string]$AudioBitrateValue = ""
    )

    if ([string]::IsNullOrWhiteSpace($AudioBitrateValue)) { $AudioBitrateValue = $DefaultAudioBitrate }

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
        "-map", "0:a:0?"
    )
    $arguments += New-VideoEncoderArguments -QualityValue $QualityValue -MaxWidthValue $MaxWidthValue -MaxVideoBitrateKbps $MaxVideoBitrateKbps
    $arguments += @(
        "-c:a", "aac",
        "-b:a", $AudioBitrateValue,
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
            Invoke-VideoEncode -InputPath $encodeInputPath -OutputPath $tempPath -StartSeconds $encodeStartSeconds -DurationSeconds $encodeDurationSeconds -QualityValue $profile.Quality -MaxWidthValue $profile.MaxWidth -MaxVideoBitrateKbps $profile.Bitrate
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

# Maintenance mode (-RecompressOutputs): walks every preset that has a size cap and re-runs
# the size-cap pass on any output that is over it. Useful after lowering a cap, since the
# watcher only applies caps at creation time.
function Get-OversizedOutputs {
    param(
        [Parameter(Mandatory = $true)][string]$Directory,
        [Parameter(Mandatory = $true)][double]$MaxSizeMegabytes
    )

    if (-not (Test-Path -LiteralPath $Directory)) { return @() }

    $maxBytes = [long]($MaxSizeMegabytes * 1024 * 1024)

    return @(
        Get-ChildItem -LiteralPath $Directory -File -Filter "*.mp4" -Recurse |
            Where-Object { $_.Length -gt $maxBytes } |
            ForEach-Object { $_.FullName }
    )
}

function Start-OutputRecompressBatch {
    $processed = 0
    $failed = 0

    foreach ($preset in Get-PipelinePresets) {
        if ($preset.SizeCapMB -le 0) { continue }

        foreach ($workspaceName in $WorkspaceNames) {
            $paths = Get-PresetWorkspacePaths -PresetName $preset.Name -WorkspaceName $workspaceName
            $targets = @(Get-OversizedOutputs -Directory $paths.OutputDir -MaxSizeMegabytes $preset.SizeCapMB)

            if ($targets.Count -eq 0) { continue }

            Write-Log "Recompress: $($targets.Count) file(s) over $($preset.SizeCapMB) MB in $($paths.OutputDir)"

            foreach ($path in $targets) {
                try {
                    $before = (Get-Item -LiteralPath $path).Length
                    Invoke-OutputSizeCap -OutputPath $path -MaxSizeMegabytes $preset.SizeCapMB -FallbackMaxWidth $preset.SizeCapFallbackMaxWidth
                    $after = (Get-Item -LiteralPath $path).Length
                    Write-Log "Recompressed: $path ($([math]::Round($before / 1MB, 2)) MB -> $([math]::Round($after / 1MB, 2)) MB)"
                    $processed++
                }
                catch {
                    Write-Log "Failed to recompress '$path': $($_.Exception.Message)" "ERROR"
                    $failed++
                }
            }
        }
    }

    Write-Log "Recompress finished: $processed succeeded, $failed failed"
}













# ---------------------------------------------------------------------------
# Event stream, control surface, and status
# ---------------------------------------------------------------------------
#
# The daily text log is written for people and cannot be parsed reliably: the preset and
# workspace appear only inside English prose, and with parallel runspaces the per-variant
# lines of different files interleave with nothing tying them back to a source file.
#
# These three surfaces exist so a monitoring UI does not have to guess:
#
#   logs\events-YYYYMMDD.jsonl   append-only event stream, one JSON object per line
#   status\watcher.json          current state, rewritten each sweep
#   control\                     flag files the poll loop checks each tick

$script:WatcherStartedUtc = (Get-Date).ToUniversalTime().ToString('o')
$script:LaneSnapshot = New-Object 'System.Collections.Specialized.OrderedDictionary'

# The name of the single-instance mutex for a given pipeline root.
#
# The default root keeps the original unqualified name, so an existing installation and any
# tooling that probes for it are unaffected. Any other root gets a name derived from its
# path, which lets a second root run alongside without the two fighting over one lock.
#
# A UI can use this to answer "is the watcher running" for free, with Mutex.OpenExisting,
# instead of scanning the process list.
function Get-WatcherMutexName {
    param([Parameter(Mandatory = $true)][string]$Root)

    $normalized = $Root.TrimEnd('\', '/').ToLowerInvariant()
    if ($normalized -eq 'd:\mediapipeline') {
        return "Global\MediaPipelineWatcher"
    }

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($normalized)
    $sha = [System.Security.Cryptography.SHA256]::Create()

    try {
        $hash = [System.BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '').Substring(0, 16)
    }
    finally {
        $sha.Dispose()
    }

    return "Global\MediaPipelineWatcher_$hash"
}

function Get-ControlDirectory {
    return (Join-Path $PipelineRoot "control")
}

function Get-StatusDirectory {
    return (Join-Path $PipelineRoot "status")
}

# Appends one event to the daily JSONL stream. Serialized through its own named mutex so
# parallel worker runspaces cannot interleave partial lines, the same way Write-Log is.
#
# JobId is what makes parallel progress attributable: every event from one group carries the
# same id, so a reader can reassemble "file X is on variant 12 of 20" from an interleaved
# stream.
function Write-PipelineEvent {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [hashtable]$Data
    )

    try {
        if (-not (Test-Path -LiteralPath $LogsDir)) {
            New-Item -ItemType Directory -Path $LogsDir -Force | Out-Null
        }

        $record = [ordered]@{
            ts = (Get-Date).ToUniversalTime().ToString('o')
            ev = $Name
        }

        if ($Data) {
            foreach ($key in $Data.Keys) {
                $record[$key] = $Data[$key]
            }
        }

        $line = $record | ConvertTo-Json -Depth 6 -Compress
        $eventFile = Join-Path $LogsDir ("events-{0}.jsonl" -f (Get-Date -Format "yyyyMMdd"))

        $mutex = New-Object System.Threading.Mutex($false, "Local\MediaPipelineEventMutex")
        $acquired = $false

        try {
            try { $acquired = $mutex.WaitOne(5000) } catch [System.Threading.AbandonedMutexException] { $acquired = $true }
            Add-Content -LiteralPath $eventFile -Value $line -Encoding UTF8
        }
        finally {
            if ($acquired) { $mutex.ReleaseMutex() }
            $mutex.Dispose()
        }
    }
    catch {
        # Losing an event must never stop media processing.
    }
}

function Test-ControlFlag {
    param([Parameter(Mandatory = $true)][string]$Name)

    return (Test-Path -LiteralPath (Join-Path (Get-ControlDirectory) $Name))
}

# Pausing is checked from the most specific flag outwards, so a single lane can be paused
# without touching the rest.
function Test-PresetPaused {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][string]$WorkspaceName
    )

    if (Test-ControlFlag -Name "pause") { return $true }
    if (Test-ControlFlag -Name "pause.$($Preset.Name)") { return $true }
    if (Test-ControlFlag -Name "pause.$($Preset.Name).$WorkspaceName") { return $true }

    return $false
}

function Test-StopRequested {
    return (Test-ControlFlag -Name "stop")
}

function Clear-StopRequest {
    $stopFile = Join-Path (Get-ControlDirectory) "stop"
    if (Test-Path -LiteralPath $stopFile) {
        Remove-Item -LiteralPath $stopFile -Force -ErrorAction SilentlyContinue
    }
}

function Set-LaneSnapshot {
    param(
        [Parameter(Mandatory = $true)][string]$PresetName,
        [Parameter(Mandatory = $true)][string]$WorkspaceName,
        [int]$Queued,
        [bool]$Paused
    )

    $key = "$PresetName/$WorkspaceName"
    $script:LaneSnapshot[$key] = [ordered]@{
        preset    = $PresetName
        workspace = $WorkspaceName
        queued    = $Queued
        paused    = $Paused
    }
}

# Rewrites status\watcher.json after each full sweep. A UI reads this instead of guessing
# at process state, and it is cheap because the queue counts were already gathered by the
# poll that just ran.
function Write-WatcherStatus {
    try {
        $statusDir = Get-StatusDirectory
        if (-not (Test-Path -LiteralPath $statusDir)) {
            New-Item -ItemType Directory -Path $statusDir -Force | Out-Null
        }

        $presetSummaries = New-Object System.Collections.Generic.List[object]
        foreach ($preset in Get-PipelinePresets) {
            $presetSummaries.Add([ordered]@{
                name        = $preset.Name
                videoCopies = $preset.VideoCopies
                imageCopies = $preset.ImageCopies
                grouping    = $preset.Grouping
                setCount    = $preset.SetCount
                batch       = $preset.Batch
                segment     = $preset.Segment
                manifest    = $preset.Manifest
                sizeCapMB   = $preset.SizeCapMB
            }) | Out-Null
        }

        $status = [ordered]@{
            schema       = "mediaPipeline.status.v1"
            pid          = $PID
            startedUtc   = $script:WatcherStartedUtc
            updatedUtc   = (Get-Date).ToUniversalTime().ToString('o')
            pipelineRoot = $PipelineRoot
            encoder      = Get-VideoEncoderName
            pollSeconds  = $PollSeconds
            pausedAll    = (Test-ControlFlag -Name "pause")
            workspaces   = $WorkspaceNames
            presets      = $presetSummaries.ToArray()
            lanes        = @($script:LaneSnapshot.Values)
        }

        $json = $status | ConvertTo-Json -Depth 8
        $statusPath = Join-Path $statusDir "watcher.json"
        $tempPath = "$statusPath.tmp"

        # Write then move, so a reader never sees a half-written file.
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText($tempPath, $json, $utf8NoBom)
        Move-Item -LiteralPath $tempPath -Destination $statusPath -Force
    }
    catch {
        # Status is a convenience for the UI, never a reason to stop processing.
    }
}

# ---------------------------------------------------------------------------
# Unified processing core
# ---------------------------------------------------------------------------
#
# One function processes every preset. What used to be nine hand-written pipelines is now
# a group of source files, a preset that says what to make of them, and three axes:
#
#   Grouping   where outputs land: one shared folder, one folder per source, or N set folders
#   Batch      whether a whole input folder is one transaction or each file is its own
#   Segment    whether long videos are split before variants are made
#
# Media type is detected per file, and the preset carries a separate copy count for video and
# for images, so a single inbox can take a mixed folder without choosing a lane.

# Emits one progress event per finished variant. The n-of-total pair is what a UI needs for
# a real progress bar, and the old text log could not provide it: the total was only ever
# logged for one lane, and parallel per-variant lines were unattributable.
function Write-PresetVariantEvent {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$File,
        [Parameter(Mandatory = $true)][int]$Index,
        [Parameter(Mandatory = $true)][int]$Total,
        [Parameter(Mandatory = $true)][string]$OutputPath
    )

    Write-PipelineEvent -Name "job.variant" -Data @{
        jobId     = $script:CurrentJobId
        preset    = $Preset.Name
        workspace = $script:CurrentWorkspaceName
        file      = $File.Name
        n         = $Index
        total     = $Total
        output    = [System.IO.Path]::GetFileName($OutputPath)
    }
}

function Get-MediaKind {
    param([Parameter(Mandatory = $true)][string]$Path)

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($VideoExtensions -contains $extension) { return "Video" }
    if ($ImageExtensions -contains $extension) { return "Image" }
    return "Unsupported"
}

# A preset ignores a media type whose copy count is zero, which is how a video-only or
# image-only preset is expressed.
function Test-PresetAcceptsFile {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][string]$Path
    )

    switch (Get-MediaKind -Path $Path) {
        "Video" { return ($Preset.VideoCopies -gt 0) }
        "Image" { return ($Preset.ImageCopies -gt 0) }
        default { return $false }
    }
}

function Get-PresetCandidateFiles {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][pscustomobject]$Paths
    )

    if (-not (Test-Path -LiteralPath $Paths.InputDir)) { return @() }

    return @(
        Get-ChildItem -LiteralPath $Paths.InputDir -File |
            Where-Object {
                (-not (Test-IsTemporaryDownload $_.FullName)) -and
                (Test-PresetAcceptsFile -Preset $Preset -Path $_.FullName)
            } |
            Sort-Object LastWriteTime, FullName
    )
}

# One rule for every preset: images come out as .jpg, except WebP which is already small and
# web-ready. This replaces three mappers that disagreed with each other, so the same .heic no
# longer becomes .png in one lane and .jpg in another.
function Get-PresetOutputExtension {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    $extension = [System.IO.Path]::GetExtension($SourcePath).ToLowerInvariant()
    if ($extension -eq ".webp") { return ".webp" }
    return ".jpg"
}

# Prepares a source for processing.
#
# HEIC decodes to a temporary working copy because ffmpeg cannot filter it directly. A .mov
# bound for segmenting is remuxed to .mp4 first, because lossless segment extraction is
# unreliable on that container. Everything else is fed to ffmpeg as-is.
#
# Returns ProcessingPath (what ffmpeg reads), SourcePath (what decides the output format),
# and TempPath (a working copy to delete afterwards, or $null).
function Resolve-PresetSource {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Kind,
        [string]$WorkDir
    )

    if ($Kind -eq "Image") {
        return Resolve-ImageProcessingSource -Path $Path
    }

    $source = [pscustomobject]@{
        SourcePath     = $Path
        ProcessingPath = $Path
        TempPath       = $null
    }

    $needsRemux = (
        $Preset.Normalize -and
        $Preset.Segment -and
        ([System.IO.Path]::GetExtension($Path).ToLowerInvariant() -eq ".mov")
    )

    if ($needsRemux) {
        if (-not (Test-Path -LiteralPath $WorkDir)) {
            New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
        }

        $remuxPath = Join-Path $WorkDir ([System.IO.Path]::GetFileNameWithoutExtension($Path) + ".mp4")
        Invoke-MovToMp4Remux -InputPath $Path -OutputPath $remuxPath
        Write-Log "Normalized .mov source to .mp4 before segmenting: $remuxPath"
        $source.ProcessingPath = $remuxPath
        $source.TempPath = $remuxPath
    }

    return $source
}

# How many variants this file gets. When CopiesAlternate is set, consecutive files alternate
# between the two counts so a run of files does not all produce the same number of outputs.
function Get-PresetCopyCount {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][string]$Kind
    )

    $count = if ($Kind -eq "Video") { $Preset.VideoCopies } else { $Preset.ImageCopies }

    if ($Preset.CopiesAlternate -gt 0) {
        $script:PresetEntryCount++
        if (($script:PresetEntryCount % 2) -eq 1) {
            $count = $Preset.CopiesAlternate
        }
    }

    return $count
}

# The destination folder for each variant, one entry per variant to produce.
#
#   Flat        every variant into the preset's shared output folder
#   PerSource   every variant into one new folder belonging to this source file
#   PerSet      the file contributes CopyCount variants to each of the SetCount set folders
function New-PresetVariantTargets {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][pscustomobject]$Paths,
        [Parameter(Mandatory = $true)][int]$CopyCount,
        [AllowEmptyCollection()][string[]]$SetDirectories,
        [System.Collections.Generic.List[string]]$CreatedDirectories
    )

    switch ($Preset.Grouping) {
        "PerSource" {
            $directory = New-RegularRandomDirectory -Directory $Paths.OutputDir
            if ($CreatedDirectories) { $CreatedDirectories.Add($directory) | Out-Null }
            Write-Log "Preset '$($Preset.Name)' output directory: $directory"
            return @(1..$CopyCount | ForEach-Object { $directory })
        }
        "PerSet" {
            $targets = New-Object System.Collections.Generic.List[string]
            foreach ($setDirectory in $SetDirectories) {
                for ($copy = 1; $copy -le $CopyCount; $copy++) {
                    $targets.Add($setDirectory) | Out-Null
                }
            }
            return $targets.ToArray()
        }
        default {
            return @(1..$CopyCount | ForEach-Object { $Paths.OutputDir })
        }
    }
}

# Produces every variant of one already-resolved source file, returning the output paths.
# Manifest records are appended to $Records when the preset writes a manifest.
function Invoke-PresetFileVariants {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][pscustomobject]$Paths,
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$File,
        [Parameter(Mandatory = $true)][pscustomobject]$Source,
        [Parameter(Mandatory = $true)][string]$Kind,
        [AllowEmptyCollection()][string[]]$SetDirectories,
        [AllowEmptyCollection()][string[]]$SetNames,
        [System.Collections.Generic.List[string]]$CreatedDirectories,
        [System.Collections.Generic.List[object]]$Records,
        [string]$BatchKey,
        [string]$FamilyKey
    )

    $created = New-Object System.Collections.Generic.List[string]
    $copyCount = Get-PresetCopyCount -Preset $Preset -Kind $Kind
    $targets = @(New-PresetVariantTargets -Preset $Preset -Paths $Paths -CopyCount $copyCount -SetDirectories $SetDirectories -CreatedDirectories $CreatedDirectories)

    if ($Kind -eq "Image") {
        $dimensions = Get-MediaDimensions -Path $Source.ProcessingPath
        $outputExtension = Get-PresetOutputExtension -SourcePath $Source.SourcePath

        # A source that had to be decoded first is already a re-encode, so give it a little
        # more quality headroom than an untouched JPEG.
        $jpegQuality = if ($Source.TempPath) { [string]$Preset.ConvertedJpegQuality } else { [string]$Preset.JpegQuality }

        for ($index = 0; $index -lt $targets.Count; $index++) {
            $variantArgs = @{
                InputPath           = $Source.ProcessingPath
                OutputDirectory     = $targets[$index]
                OutputExtension     = $outputExtension
                Dimensions          = $dimensions
                LogLabel            = "$($Preset.Name) image variant $($index + 1)"
                JpegQuality         = $jpegQuality
                PngCompressionLevel = $Preset.PngCompressionLevel
                CropMinPermille     = $Preset.CropMinPermille
                CropMaxPermille     = $Preset.CropMaxPermille
            }

            $outputPath = New-ImageVariant @variantArgs
            $created.Add($outputPath) | Out-Null
            Write-PresetVariantEvent -Preset $Preset -File $File -Index ($index + 1) -Total $targets.Count -OutputPath $outputPath

            if ($Records -ne $null) {
                $Records.Add((New-PresetManifestRecord -Preset $Preset -File $File -OutputPath $outputPath -SetName $SetNames[[int]($index / $copyCount)] -BatchKey $BatchKey -FamilyKey $FamilyKey -Kind $Kind -Dimensions $dimensions)) | Out-Null
            }
        }

        return $created.ToArray()
    }

    # Video. Segmenting splits the source first and then makes CopyCount variants of every
    # segment; otherwise the whole file is the single segment.
    $duration = Get-VideoDurationSeconds -Path $Source.ProcessingPath
    $segments = @([pscustomobject]@{ Index = 1; StartSeconds = -1; DurationSeconds = $duration; Path = $Source.ProcessingPath })
    $segmentTemps = New-Object System.Collections.Generic.List[string]

    if ($Preset.Segment) {
        $segments = @(Invoke-PresetSegmentExtract -Preset $Preset -Paths $Paths -Source $Source -DurationSeconds $duration -TempPaths $segmentTemps)
    }

    # Progress is reported across every segment, so a segmented job reads as one bar rather
    # than restarting at zero for each segment.
    $totalVariants = $targets.Count * $segments.Count
    $variantsDone = 0

    # Which quality knob applies depends on the encoder that was probed at startup.
    $presetQuality = if ($script:UseNvenc) { $Preset.NvencCq } elseif ($script:UseAmf) { $Preset.AmfQp } else { $Preset.Crf }

    try {
        foreach ($segment in $segments) {
            $trimRange = Get-TrimRange -DurationSeconds $segment.DurationSeconds -ConfiguredMinMs $Preset.MinTrimMs -ConfiguredMaxMs $Preset.MaxTrimMs
            if (-not $trimRange.CanTrim) {
                Write-Log "Preset '$($Preset.Name)': not trimming ($($trimRange.Reason))." "WARN"
            }

            $usedTrims = New-Object 'System.Collections.Generic.HashSet[int]'

            for ($index = 0; $index -lt $targets.Count; $index++) {
                $trimMs = if ($trimRange.CanTrim) {
                    New-TrimMilliseconds -Range $trimRange -UsedValues $usedTrims -CopyCount $targets.Count
                }
                else {
                    0
                }

                $segmentLabel = if ($Preset.Segment) { " segment $($segment.Index)" } else { "" }

                $variantArgs = @{
                    InputPath               = $segment.Path
                    OutputDirectory         = $targets[$index]
                    DurationSeconds         = $segment.DurationSeconds
                    TrimMs                  = $trimMs
                    LogLabel                = "$($Preset.Name)$segmentLabel video variant $($index + 1)"
                    QualityValue            = $presetQuality
                    MaxWidthValue           = $Preset.MaxWidth
                    AudioBitrateValue       = $Preset.AudioBitrate
                    MaxSizeMegabytes        = $Preset.SizeCapMB
                    SizeCapFallbackMaxWidth = $Preset.SizeCapFallbackMaxWidth
                }

                if ($Preset.SizeCapMB -gt 0) {
                    $targetDuration = [Math]::Max(0.1, $segment.DurationSeconds - ($trimMs / 1000.0))
                    $variantArgs.MaxVideoBitrateKbps = Get-PrimaryMaxVideoBitrateKbps -DurationSeconds $targetDuration -MaxSizeMegabytes $Preset.SizeCapMB -MaxrateScale $Preset.MaxrateScale
                }

                $outputPath = New-VideoVariant @variantArgs
                $created.Add($outputPath) | Out-Null
                $variantsDone++
                Write-PresetVariantEvent -Preset $Preset -File $File -Index $variantsDone -Total $totalVariants -OutputPath $outputPath

                if ($Records -ne $null) {
                    $Records.Add((New-PresetManifestRecord -Preset $Preset -File $File -OutputPath $outputPath -SetName $SetNames[[int]($index / $copyCount)] -BatchKey $BatchKey -FamilyKey $FamilyKey -Kind $Kind -TrimMs $trimMs -DurationSeconds $segment.DurationSeconds)) | Out-Null
                }
            }
        }
    }
    finally {
        foreach ($tempPath in $segmentTemps) {
            if ($tempPath -and (Test-Path -LiteralPath $tempPath)) {
                Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
            }
        }
    }

    return $created.ToArray()
}

# Splits a long video into losslessly extracted segments in the preset's work folder.
function Invoke-PresetSegmentExtract {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][pscustomobject]$Paths,
        [Parameter(Mandatory = $true)][pscustomobject]$Source,
        [Parameter(Mandatory = $true)][double]$DurationSeconds,
        [AllowEmptyCollection()][System.Collections.Generic.List[string]]$TempPaths
    )

    $plan = @(Get-LongSegmentPlan -DurationSeconds $DurationSeconds)
    $summary = ($plan | ForEach-Object { "{0:0.##}s" -f $_.DurationSeconds }) -join ", "
    Write-Log "Preset '$($Preset.Name)' segment plan: $($plan.Count) segment(s): $summary"

    if (-not (Test-Path -LiteralPath $Paths.WorkDir)) {
        New-Item -ItemType Directory -Path $Paths.WorkDir -Force | Out-Null
    }

    $segments = New-Object System.Collections.Generic.List[object]
    $jobToken = [Guid]::NewGuid().ToString("n").Substring(0, 8)

    foreach ($entry in $plan) {
        $segmentPath = Join-Path $Paths.WorkDir ("segment_{0}_{1:D3}.mp4" -f $jobToken, $entry.Index)
        Invoke-LongSegmentExtract -InputPath $Source.ProcessingPath -OutputPath $segmentPath -StartSeconds $entry.StartSeconds -DurationSeconds $entry.DurationSeconds
        $TempPaths.Add($segmentPath) | Out-Null

        $segments.Add([pscustomobject]@{
            Index           = $entry.Index
            StartSeconds    = -1
            DurationSeconds = $entry.DurationSeconds
            Path            = $segmentPath
        }) | Out-Null
    }

    return $segments.ToArray()
}

function New-PresetManifestRecord {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][System.IO.FileInfo]$File,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [string]$SetName,
        [string]$BatchKey,
        [string]$FamilyKey,
        [Parameter(Mandatory = $true)][string]$Kind,
        [int]$TrimMs = 0,
        [double]$DurationSeconds = 0,
        [pscustomobject]$Dimensions
    )

    $record = [ordered]@{
        familyKey          = $FamilyKey
        variantKey         = "{0}__{1}" -f $FamilyKey, $SetName
        path               = "{0}/{1}" -f $SetName, [System.IO.Path]::GetFileName($OutputPath)
        renditionSetKey    = $SetName
        generationBatchKey = $BatchKey
        sourceOriginalName = $File.Name
        sourceFamilyName   = $FamilyKey
        sizeBytes          = (Get-Item -LiteralPath $OutputPath).Length
        generatedAt        = Get-UtcIsoTimestamp
        metadata           = [ordered]@{
            encoder  = Get-VideoEncoderName
            trimMs   = $TrimMs
            maxWidth = $Preset.MaxWidth
        }
    }

    if ($Kind -eq "Video") {
        $record.durationSeconds = [Math]::Max(0.1, $DurationSeconds - ($TrimMs / 1000.0))
        $record.transformProfile = "preset_video_micro_trim"
    }
    else {
        $record.durationSeconds = 0
        $record.transformProfile = "preset_image_recrop"
        if ($Dimensions) {
            $record.metadata.sourceWidth = $Dimensions.Width
            $record.metadata.sourceHeight = $Dimensions.Height
        }
    }

    return [pscustomobject]$record
}

# Processes one group of files as a single transaction.
#
# A per-file preset calls this with a group of one, which is what removes the loop inversion
# the old code had between the per-file lanes and the batch lanes: creating the destination
# before the source loop is correct in both cases, and a group of one degenerates to the
# old per-file behavior.
function Invoke-PresetGroup {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][pscustomobject]$Paths,
        [Parameter(Mandatory = $true)][object[]]$Files
    )

    $createdOutputs = New-Object System.Collections.Generic.List[string]
    $createdDirectories = New-Object System.Collections.Generic.List[string]
    $records = if ($Preset.Manifest) { New-Object System.Collections.Generic.List[object] } else { $null }
    $processedFiles = New-Object System.Collections.Generic.List[object]

    $setDirectories = @()
    $setNames = @()
    $batchDirectory = $null
    $batchKey = $null

    try {
        # PerSet builds its whole container tree up front so every source contributes to the
        # same set folders.
        if ($Preset.Grouping -eq "PerSet") {
            $batchDirectory = New-RegularRandomDirectory -Directory $Paths.OutputDir
            $createdDirectories.Add($batchDirectory) | Out-Null
            $batchKey = [System.IO.Path]::GetFileName($batchDirectory)
            Write-Log "Preset '$($Preset.Name)' batch directory: $batchDirectory"

            for ($setNumber = 1; $setNumber -le $Preset.SetCount; $setNumber++) {
                $setDirectory = New-RegularRandomDirectory -Directory $batchDirectory
                $setDirectories += $setDirectory
                $setNames += [System.IO.Path]::GetFileName($setDirectory)
            }
        }
        else {
            # A flat or per-source preset has no set folders, but manifest records still want
            # a stable name, so the preset's own output folder stands in.
            $setNames = @(".")
        }

        $usedFamilyKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)

        foreach ($file in $Files) {
            $kind = Get-MediaKind -Path $file.FullName
            $source = $null

            $familyKey = $null
            if ($Preset.Manifest) {
                $familyKey = Get-AssetStoreFamilyKey -FileName $file.Name
                $candidate = $familyKey
                $suffix = 2
                while (-not $usedFamilyKeys.Add($candidate)) {
                    $candidate = "{0}_{1}" -f $familyKey, $suffix
                    $suffix++
                }
                $familyKey = $candidate
            }

            try {
                # A batch preset already proved the whole folder settled before it got here.
                # A per-file preset waits on this one file, and throws if it never settles.
                if ($Preset.Batch -eq "PerFile") {
                    Wait-FileReady -Path $file.FullName
                }

                $source = Resolve-PresetSource -Preset $Preset -Path $file.FullName -Kind $kind -WorkDir $Paths.WorkDir

                $variantArgs = @{
                    Preset             = $Preset
                    Paths              = $Paths
                    File               = $file
                    Source             = $source
                    Kind               = $kind
                    SetDirectories     = $setDirectories
                    SetNames           = $setNames
                    CreatedDirectories = $createdDirectories
                    Records            = $records
                    BatchKey           = $batchKey
                    FamilyKey          = $familyKey
                }

                foreach ($outputPath in (Invoke-PresetFileVariants @variantArgs)) {
                    $createdOutputs.Add($outputPath) | Out-Null
                }

                $processedFiles.Add($file) | Out-Null
            }
            finally {
                if ($source -and $source.TempPath) {
                    Remove-HeicWorkingCopy -Path $source.TempPath
                }
            }
        }

        if ($Preset.Manifest -and $batchDirectory) {
            Write-PresetManifest -Preset $Preset -BatchDirectory $batchDirectory -Variants $records.ToArray()
        }
    }
    catch {
        Invoke-PresetRollback -Preset $Preset -CreatedOutputs $createdOutputs -CreatedDirectories $createdDirectories

        foreach ($file in $Files) {
            if (Test-Path -LiteralPath $file.FullName) {
                Move-InputFile -Path $file.FullName -DestinationDirectory $Paths.FailedDir
            }
        }

        throw
    }

    # Outputs are complete. Archiving the sources is best-effort and must never discard
    # finished work, so it happens after the transactional block above.
    foreach ($file in $processedFiles) {
        if (Test-Path -LiteralPath $file.FullName) {
            Move-InputFile -Path $file.FullName -DestinationDirectory $Paths.OriginalDir
        }
    }

    return $createdOutputs.ToArray()
}

# Applies the preset's rollback policy after a failed group.
function Invoke-PresetRollback {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [AllowEmptyCollection()][System.Collections.Generic.List[string]]$CreatedOutputs,
        [AllowEmptyCollection()][System.Collections.Generic.List[string]]$CreatedDirectories
    )

    switch ($Preset.OnFailure) {
        "DeleteContainer" {
            foreach ($directory in $CreatedDirectories) {
                Remove-GeneratedOutputDirectory -Path $directory
            }
            if ($CreatedDirectories.Count -eq 0) {
                Remove-GeneratedOutputs -Paths $CreatedOutputs.ToArray()
            }
        }
        "DeleteFiles" {
            Remove-GeneratedOutputs -Paths $CreatedOutputs.ToArray()
        }
        default {
            if ($CreatedOutputs.Count -gt 0) {
                Write-Log "Preset '$($Preset.Name)': preserving $($CreatedOutputs.Count) completed output(s) after failure." "WARN"
            }
        }
    }
}

function Write-PresetManifest {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][string]$BatchDirectory,
        [AllowEmptyCollection()][object[]]$Variants
    )

    $manifest = [ordered]@{
        schema      = $Preset.ManifestSchema
        generatedAt = Get-UtcIsoTimestamp
        importRoot  = "."
        variants    = [object[]]$Variants
    }

    $json = $manifest | ConvertTo-Json -Depth 12
    $manifestPath = Join-Path $BatchDirectory "manifest.json"
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($manifestPath, $json, $utf8NoBom)
    Write-Log "Wrote manifest: $manifestPath ($($Variants.Count) variant(s))"
}

function Get-PresetBatchSignature {
    param([Parameter(Mandatory = $true)][string]$PresetName)

    if ($script:BatchSignatures -and $script:BatchSignatures.ContainsKey($PresetName)) {
        return $script:BatchSignatures[$PresetName]
    }

    return $null
}

function Set-PresetBatchSignature {
    param(
        [Parameter(Mandatory = $true)][string]$PresetName,
        [string]$Signature
    )

    if (-not $script:BatchSignatures) { return }

    if ($null -eq $Signature) {
        [void]$script:BatchSignatures.Remove($PresetName)
    }
    else {
        $script:BatchSignatures[$PresetName] = $Signature
    }
}

# Runs one group and turns any failure into a log line rather than stopping the watcher.
# The group's own transaction has already applied the preset's rollback policy and moved the
# sources to failed\ by the time an exception reaches here.
function Invoke-PresetGroupSafely {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][pscustomobject]$Paths,
        [Parameter(Mandatory = $true)][object[]]$Files
    )

    $names = ($Files | ForEach-Object { $_.Name }) -join ", "
    $jobId = [Guid]::NewGuid().ToString("n").Substring(0, 8)

    # Every variant produced under this group carries the same job id, which is how a reader
    # attributes interleaved progress from parallel runspaces back to one source.
    $script:CurrentJobId = $jobId
    $script:CurrentJobTotal = 0
    $script:CurrentJobDone = 0

    Write-Log "Preset '$($Preset.Name)' [$($Paths.WorkspaceName)] processing $($Files.Count) file(s): $names"
    Write-PipelineEvent -Name "job.start" -Data @{
        jobId     = $jobId
        preset    = $Preset.Name
        workspace = $Paths.WorkspaceName
        files     = @($Files | ForEach-Object { $_.Name })
        bytes     = (($Files | Measure-Object -Property Length -Sum).Sum)
    }

    try {
        $outputs = @(Invoke-PresetGroup -Preset $Preset -Paths $Paths -Files $Files)
        Write-Log "Preset '$($Preset.Name)' [$($Paths.WorkspaceName)] created $($outputs.Count) output(s)."
        Write-PipelineEvent -Name "job.done" -Data @{
            jobId     = $jobId
            preset    = $Preset.Name
            workspace = $Paths.WorkspaceName
            outputs   = $outputs.Count
        }
    }
    catch {
        $origin = if ($_.InvocationInfo) { " at $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())" } else { "" }
        Write-Log "Preset '$($Preset.Name)' [$($Paths.WorkspaceName)] failed: $($_.Exception.Message)$origin" "ERROR"
        Write-PipelineEvent -Name "job.failed" -Data @{
            jobId     = $jobId
            preset    = $Preset.Name
            workspace = $Paths.WorkspaceName
            error     = $_.Exception.Message
        }
    }
    finally {
        $script:CurrentJobId = $null
    }
}

# Decides whether a PerGroup preset's input folder has settled. All three conditions must
# hold: every file unlocked, the newest write older than StableSeconds, and the folder
# contents unchanged since the previous poll.
function Test-PresetBatchSettled {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][object[]]$Files
    )

    foreach ($file in $Files) {
        if (-not (Test-FileUnlocked $file.FullName)) { return $false }
    }

    $newestWrite = ($Files | Measure-Object -Property LastWriteTime -Maximum).Maximum
    if (((Get-Date) - $newestWrite).TotalSeconds -lt $StableSeconds) { return $false }

    $signature = (($Files | ForEach-Object { '{0}|{1}' -f $_.FullName, $_.Length }) -join ';')
    $previous = Get-PresetBatchSignature -PresetName $Preset.Name

    if ($signature -eq $previous) { return $true }

    if ($signature -ne $previous) {
        Write-Log "Preset '$($Preset.Name)': $($Files.Count) file(s) detected; waiting for the batch to settle."
    }

    Set-PresetBatchSignature -PresetName $Preset.Name -Signature $signature
    return $false
}

# One poll of one preset in one workspace.
function Invoke-PresetPoll {
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Preset,
        [Parameter(Mandatory = $true)][string]$WorkspaceName
    )

    $paths = Get-PresetWorkspacePaths -PresetName $Preset.Name -WorkspaceName $WorkspaceName
    $files = @(Get-PresetCandidateFiles -Preset $Preset -Paths $paths)
    $paused = Test-PresetPaused -Preset $Preset -WorkspaceName $WorkspaceName

    Set-LaneSnapshot -PresetName $Preset.Name -WorkspaceName $WorkspaceName -Queued $files.Count -Paused $paused

    # A paused lane keeps its queue and stops picking work up. Nothing is discarded.
    if ($paused) { return }

    if ($files.Count -eq 0) {
        Set-PresetBatchSignature -PresetName $Preset.Name -Signature $null
        return
    }

    if ($Preset.Batch -eq "PerGroup") {
        if (Test-PresetBatchSettled -Preset $Preset -Files $files) {
            Set-PresetBatchSignature -PresetName $Preset.Name -Signature $null
            Invoke-PresetGroupSafely -Preset $Preset -Paths $paths -Files $files
        }

        return
    }

    # A per-file preset is the same transaction with a group of one.
    #
    # Files are processed concurrently when the preset allows it. Alternating copy counts are
    # excluded because the alternation counter lives in script scope, and each worker runspace
    # loads its own copy, which would make the alternation arbitrary.
    $canParallelize = (
        $script:SupportsParallel -and
        $Preset.Parallel -eq "OverFiles" -and
        $Preset.CopiesAlternate -le 0 -and
        $files.Count -gt 1 -and
        $ImageProcessingConcurrency -gt 1
    )

    if ($canParallelize) {
        Write-Log "Preset '$($Preset.Name)' [$WorkspaceName] processing $($files.Count) file(s) with concurrency $ImageProcessingConcurrency."

        $libPath = $script:ScriptPath
        $ffPath = $script:FFmpegPath
        $fpPath = $script:FFprobePath
        $exPath = $script:ExifToolPath
        $presetName = $Preset.Name

        $files | ForEach-Object -ThrottleLimit $ImageProcessingConcurrency -Parallel {
            . $using:libPath -AsLibrary

            $script:FFmpegPath = $using:ffPath
            $script:FFprobePath = $using:fpPath
            $script:ExifToolPath = $using:exPath
            $script:CurrentWorkspaceName = $using:WorkspaceName

            $workerPreset = Get-PipelinePreset -Name $using:presetName
            $workerPaths = Get-PresetWorkspacePaths -PresetName $using:presetName -WorkspaceName $using:WorkspaceName

            Invoke-PresetGroupSafely -Preset $workerPreset -Paths $workerPaths -Files @($_)
        }

        return
    }

    foreach ($file in $files) {
        Invoke-PresetGroupSafely -Preset $Preset -Paths $paths -Files @($file)
    }
}

function Write-WatcherStartupBanner {
    Write-Log "Watcher started."
    Write-Log "Pipeline root: $PipelineRoot"
    Write-Log "Workspaces: $($WorkspaceNames -join ', ')"
    Write-Log "Polling every $PollSeconds seconds."

    foreach ($preset in Get-PipelinePresets) {
        $details = New-Object System.Collections.Generic.List[string]
        $details.Add("video x$($preset.VideoCopies)") | Out-Null
        $details.Add("image x$($preset.ImageCopies)") | Out-Null
        $details.Add($preset.Grouping.ToLowerInvariant()) | Out-Null

        if ($preset.Grouping -eq "PerSet") { $details.Add("$($preset.SetCount) set(s)") | Out-Null }
        if ($preset.Batch -eq "PerGroup") { $details.Add("batched") | Out-Null }
        if ($preset.Segment) { $details.Add("segmented") | Out-Null }
        if ($preset.Manifest) { $details.Add("manifest") | Out-Null }
        $details.Add($(if ($preset.SizeCapMB -gt 0) { "cap $($preset.SizeCapMB) MB" } else { "no size cap" })) | Out-Null

        Write-Log "Preset '$($preset.Name)': $($details -join ', ')"
    }

    if ($ArchiveEnabled) {
        Write-Log "Archiving outputs older than $ArchiveAgeHours hours every $ArchiveCheckIntervalMinutes minutes."
    }
    else {
        Write-Log "Archiving disabled."
    }

    if ($AssetRetentionDays -gt 0) {
        Write-Log "Asset retention: deleting retained entries after $AssetRetentionDays day(s)."
    }
}

function Start-PollingWatcher {
    Write-WatcherStartupBanner

    if (-not (Test-Path -LiteralPath (Get-ControlDirectory))) {
        New-Item -ItemType Directory -Path (Get-ControlDirectory) -Force | Out-Null
    }

    # A stop flag left over from a previous run would exit immediately.
    Clear-StopRequest

    Write-PipelineEvent -Name "watcher.start" -Data @{
        pid          = $PID
        pipelineRoot = $PipelineRoot
        encoder      = Get-VideoEncoderName
        presets      = @(Get-PipelinePresets | ForEach-Object { $_.Name })
        workspaces   = $WorkspaceNames
    }

    $stopping = $false

    while (-not $stopping) {
        foreach ($workspaceName in $WorkspaceNames) {
            if (Test-StopRequested) { $stopping = $true; break }

            Use-PipelineWorkspace -WorkspaceName $workspaceName

            try {
                Invoke-OutputArchiveIfDue

                foreach ($preset in Get-PipelinePresets) {
                    if (Test-StopRequested) { $stopping = $true; break }
                    Invoke-PresetPoll -Preset $preset -WorkspaceName $workspaceName
                }
            }
            catch {
                Write-Log "Watcher loop error [$($script:CurrentWorkspaceName)]: $($_.Exception.Message)" "ERROR"
            }
            finally {
                Save-PipelineWorkspaceState
            }
        }

        Write-WatcherStatus

        if ($stopping) { break }

        Start-Sleep -Seconds $PollSeconds
    }

    # Reaching here means a stop was requested rather than the process being killed, so the
    # current file finished cleanly and the mutex is about to be released properly.
    Write-Log "Stop requested. Watcher shutting down cleanly."
    Write-PipelineEvent -Name "watcher.stop" -Data @{ pid = $PID; reason = "control" }
    Clear-StopRequest
    Write-WatcherStatus
}

# When dot-sourced by a parallel worker runspace, only load functions/config and return —
# do not take the single-instance mutex or start the polling loop.
if ($AsLibrary) { return }

try {
    $createdNew = $false
    # The single-instance lock is scoped to the pipeline root, so one watcher per root rather
    # than one per machine. Two roots can then run side by side, which is what makes it
    # possible to exercise a sandbox watcher without stopping the real one.
    $script:InstanceMutexName = Get-WatcherMutexName -Root $PipelineRoot
    $script:InstanceMutex = New-Object System.Threading.Mutex($true, $script:InstanceMutexName, [ref]$createdNew)
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

    if ($RecompressOutputs) {
        foreach ($workspaceName in $WorkspaceNames) {
            Use-PipelineWorkspace -WorkspaceName $workspaceName
            try {
                Start-OutputRecompressBatch
            }
            finally {
                Save-PipelineWorkspaceState
            }
        }
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
