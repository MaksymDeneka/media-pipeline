<#
.SYNOPSIS
    Characterization harness for watch-media.ps1.

.DESCRIPTION
    Drives every pipeline handler directly against a throwaway sandbox root and records a
    structural fingerprint of what each one produced. Capture a baseline before refactoring,
    compare after, and any unintended behavior change shows up as a diff.

    Output filenames are generated from a crypto RNG and can never be reproduced, so the
    fingerprint deliberately ignores names. It compares what is actually contractual:
    how many outputs, of which extensions, at which folder depth, with which pixel
    dimensions and (bucketed) durations, and where the source files ended up.

    The script never touches the real pipeline root and never starts the watcher. It
    dot-sources watch-media.ps1 with -AsLibrary, which loads the functions without taking
    the single-instance mutex or entering the poll loop.

.EXAMPLE
    # Before refactoring
    pwsh -File tools\Test-PipelineParity.ps1 -Mode Capture -BuildCorpus

.EXAMPLE
    # After refactoring
    pwsh -File tools\Test-PipelineParity.ps1 -Mode Compare
#>
[CmdletBinding()]
param(
    [ValidateSet('Capture', 'Compare')]
    [string]$Mode = 'Capture',

    # Where the baseline fingerprint is written / read. Text only, safe to commit.
    [string]$BaselinePath,

    # Scratch area for the sandbox pipeline root and the media corpus. Never the real root.
    [string]$SandboxRoot = (Join-Path $env:TEMP 'mp-parity'),

    # The watcher under test.
    [string]$ScriptPath,

    # Legacy runs watch-media.ps1. Native runs the C# worker against the same sandbox.
    [ValidateSet('Legacy', 'Native')]
    [string]$Engine = 'Legacy',

    [string]$WorkerProject,

    [string]$DotnetPath,

    # Where -BuildCorpus harvests sample media from.
    [string]$CorpusSource = 'D:\MediaPipeline',

    # Harvest a fresh corpus before running. Needed on first use.
    [switch]$BuildCorpus,

    # Generates a deterministic local corpus without reading the real pipeline tree.
    [switch]$BuildSyntheticCorpus,

    # Leave the sandbox tree on disk for inspection.
    [switch]$KeepSandbox,

    # Only run these scenarios (by name). Default: all.
    [string[]]$Only
)

$ErrorActionPreference = 'Stop'

# Deliberately no Set-StrictMode: this script dot-sources the watcher into its own scope,
# and strict mode would change the watcher's behavior rather than observe it.

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $ScriptPath)   { $ScriptPath   = Join-Path $repoRoot 'watch-media.ps1' }
if (-not $BaselinePath) { $BaselinePath = Join-Path $PSScriptRoot 'parity-baseline.json' }
if (-not $WorkerProject) { $WorkerProject = Join-Path $repoRoot 'src\MediaPipeline.Worker\MediaPipeline.Worker.csproj' }
if (-not $DotnetPath) {
    $dotnetCommand = Get-Command dotnet -ErrorAction SilentlyContinue
    $DotnetPath = if ($dotnetCommand) { $dotnetCommand.Source } else { 'C:\Program Files\dotnet\dotnet.exe' }
}

$corpusDir  = Join-Path $SandboxRoot 'corpus'
$appDir     = Join-Path $SandboxRoot 'app'
$pipelineRoot = Join-Path $SandboxRoot 'root'

# ---------------------------------------------------------------------------
# Corpus
# ---------------------------------------------------------------------------

# Deterministic selection: sort by full path, take the first N of each kind, so repeated
# runs against an unchanged source pick exactly the same files.
$CorpusSpec = @(
    @{ Kind = 'mp4';  Pattern = '*.mp4';  Count = 3 }
    @{ Kind = 'mov';  Pattern = '*.mov';  Count = 2 }
    @{ Kind = 'jpg';  Pattern = '*.jpg';  Count = 3 }
    @{ Kind = 'png';  Pattern = '*.png';  Count = 2 }
    @{ Kind = 'heic'; Pattern = '*.heic'; Count = 2 }
)

function Build-Corpus {
    param([string]$Source, [string]$Destination)

    if (-not (Test-Path -LiteralPath $Source)) {
        throw "Corpus source not found: $Source"
    }

    if (Test-Path -LiteralPath $Destination) {
        Remove-Item -LiteralPath $Destination -Recurse -Force
    }
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null

    $manifest = [System.Collections.Generic.List[object]]::new()

    foreach ($spec in $CorpusSpec) {
        $found = @(
            Get-ChildItem -LiteralPath $Source -Recurse -File -Filter $spec.Pattern -ErrorAction SilentlyContinue |
                Where-Object { $_.Length -gt 0 } |
                Sort-Object FullName |
                Select-Object -First $spec.Count
        )

        if ($found.Count -lt $spec.Count) {
            Write-Warning "Corpus: wanted $($spec.Count) $($spec.Kind) file(s), found $($found.Count)."
        }

        $index = 0
        foreach ($file in $found) {
            $index++
            # Stable, kind-tagged names so scenarios can request media by type.
            $name = '{0}-{1:D2}{2}' -f $spec.Kind, $index, $file.Extension.ToLowerInvariant()
            $target = Join-Path $Destination $name
            Copy-Item -LiteralPath $file.FullName -Destination $target -Force

            $manifest.Add([ordered]@{
                name       = $name
                kind       = $spec.Kind
                sourcePath = $file.FullName
                bytes      = $file.Length
                sha256     = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
            })
        }
    }

    $manifestPath = Join-Path $Destination 'corpus.json'
    $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
    Write-Host "Corpus built: $($manifest.Count) file(s) in $Destination"
}

function Build-SyntheticCorpus {
    param([string]$Destination)

    $ffmpegCandidates = @(
        'C:\Tools\ffmpeg\bin\ffmpeg.exe',
        (Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Links\ffmpeg.exe')
    )
    $ffmpegCommand = Get-Command ffmpeg -ErrorAction SilentlyContinue
    if ($ffmpegCommand) { $ffmpegCandidates += $ffmpegCommand.Source }
    $ffmpeg = $ffmpegCandidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1
    if (-not $ffmpeg) { throw 'FFmpeg is required to build the synthetic parity corpus.' }

    if (Test-Path -LiteralPath $Destination) {
        Remove-Item -LiteralPath $Destination -Recurse -Force
    }
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null

    function Invoke-CorpusFfmpeg {
        param([string[]]$Arguments)
        & $ffmpeg -y -hide_banner -loglevel error @Arguments
        if ($LASTEXITCODE -ne 0) { throw "FFmpeg failed while building the parity corpus." }
    }

    $colors = @('red', 'green', 'blue')
    for ($index = 1; $index -le 3; $index++) {
        $target = Join-Path $Destination ('mp4-{0:D2}.mp4' -f $index)
        Invoke-CorpusFfmpeg -Arguments @(
            '-f', 'lavfi', '-i', "testsrc2=size=320x240:rate=24",
            '-f', 'lavfi', '-i', "sine=frequency=$((400 + $index * 100)):sample_rate=44100",
            '-t', [string](6 + $index), '-c:v', 'libx264', '-preset', 'ultrafast',
            '-pix_fmt', 'yuv420p', '-c:a', 'aac', '-shortest', $target
        )
    }

    for ($index = 1; $index -le 2; $index++) {
        $target = Join-Path $Destination ('mov-{0:D2}.mov' -f $index)
        Invoke-CorpusFfmpeg -Arguments @(
            '-f', 'lavfi', '-i', 'testsrc2=size=240x320:rate=24',
            '-f', 'lavfi', '-i', "sine=frequency=$((700 + $index * 100)):sample_rate=44100",
            '-t', [string](6 + $index), '-c:v', 'libx264', '-preset', 'ultrafast',
            '-pix_fmt', 'yuv420p', '-c:a', 'aac', '-shortest', $target
        )
    }

    for ($index = 1; $index -le 3; $index++) {
        Invoke-CorpusFfmpeg -Arguments @(
            '-f', 'lavfi', '-i', "color=c=$($colors[$index - 1]):s=320x240",
            '-frames:v', '1', (Join-Path $Destination ('jpg-{0:D2}.jpg' -f $index))
        )
    }

    for ($index = 1; $index -le 2; $index++) {
        Invoke-CorpusFfmpeg -Arguments @(
            '-f', 'lavfi', '-i', "testsrc2=size=300x220",
            '-frames:v', '1', (Join-Path $Destination ('png-{0:D2}.png' -f $index))
        )
    }

    for ($index = 1; $index -le 2; $index++) {
        $temporary = Join-Path $Destination ('heic-{0:D2}.avif' -f $index)
        Invoke-CorpusFfmpeg -Arguments @(
            '-f', 'lavfi', '-i', "color=c=$($colors[$index]):s=280x360",
            '-frames:v', '1', '-c:v', 'libaom-av1', '-still-picture', '1',
            '-f', 'avif', $temporary
        )
        Move-Item -LiteralPath $temporary -Destination (Join-Path $Destination ('heic-{0:D2}.heic' -f $index))
    }

    Get-ChildItem -LiteralPath $Destination -File | Sort-Object Name | ForEach-Object {
        [ordered]@{
            name = $_.Name
            kind = $_.Extension.TrimStart('.').ToLowerInvariant()
            sourcePath = 'synthetic'
            bytes = $_.Length
            sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
        }
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $Destination 'corpus.json') -Encoding UTF8

    Write-Host "Synthetic corpus built in $Destination"
}

function Get-CorpusFiles {
    param([string]$Kind, [int]$Count = 1)

    $matched = @(
        Get-ChildItem -LiteralPath $corpusDir -File -Filter "$Kind-*" -ErrorAction SilentlyContinue |
            Sort-Object Name |
            Select-Object -First $Count
    )

    if ($matched.Count -lt $Count) {
        throw "Corpus is missing $Count '$Kind' file(s) (found $($matched.Count)). Re-run with -BuildCorpus."
    }

    return $matched
}

# ---------------------------------------------------------------------------
# Sandbox
# ---------------------------------------------------------------------------

# Small copy counts keep a full run to a few minutes while still exercising every loop.
# The alternating default counts stay different from each other so the alternation stays
# observable in the fingerprint.
$SandboxConfig = @'
[General]
PipelineRoot = {0}

[Video]
Crf = 28
X264Preset = ultrafast
AudioBitrate = 96k
MaxWidth = 640
PreferNvenc = false
PreferAmf = false
SizeCapMB = 8
SizeCapFallbackMaxWidth = 480
MinTrimMs = 15
MaxTrimMs = 95
SegmentTargetSeconds = 4
SegmentMinSeconds = 3

[Images]
; One at a time keeps the run deterministic and the fingerprint stable.
ImageProcessingConcurrency = 1
CropMinPermille = 5
CropMaxPermille = 20
JpegQuality = 4
ConvertedJpegQuality = 12
PngCompressionLevel = 6

[Timing]
StableSeconds = 0
TimeoutSeconds = 120
PollSeconds = 2

[Archive]
ArchiveEnabled = false
ArchiveAgeHours = 999
ArchiveCheckIntervalMinutes = 999
AssetRetentionDays = 0

; Small copy counts keep a full run to a couple of minutes while still exercising every
; loop. The default preset keeps two different counts so its alternation stays observable.
[preset bulk]
VideoCopies = 2
ImageCopies = 2
CopiesAlternate = 3

[preset video-clean]
VideoCopies = 1
ImageCopies = 0

[preset image-clean]
VideoCopies = 0
ImageCopies = 1

[preset image-bulk]
VideoCopies = 0
ImageCopies = 3

[preset sets]
VideoCopies = 2
ImageCopies = 2
Grouping = PerSource
SizeCapMB = 0

[preset sets-batch]
VideoCopies = 1
ImageCopies = 1
Grouping = PerSet
SetCount = 2
Batch = PerGroup
SizeCapMB = 0

[preset asset-store]
VideoCopies = 1
ImageCopies = 1
Grouping = PerSet
SetCount = 2
Batch = PerGroup
SizeCapMB = 0
Manifest = true
MinTrimMs = 10
MaxTrimMs = 40

[preset video-long]
VideoCopies = 2
ImageCopies = 0
Segment = true
'@

function New-Sandbox {
    if (Test-Path -LiteralPath $appDir)       { Remove-Item -LiteralPath $appDir -Recurse -Force }
    if (Test-Path -LiteralPath $pipelineRoot) { Remove-Item -LiteralPath $pipelineRoot -Recurse -Force }

    New-Item -ItemType Directory -Path $appDir -Force | Out-Null
    New-Item -ItemType Directory -Path $pipelineRoot -Force | Out-Null

    # The watcher resolves config.ini relative to its own location, so the copy in the
    # sandbox picks up the sandbox config and never the real one.
    Copy-Item -LiteralPath $ScriptPath -Destination (Join-Path $appDir 'watch-media.ps1') -Force
    ($SandboxConfig -f $pipelineRoot) | Set-Content -LiteralPath (Join-Path $appDir 'config.ini') -Encoding UTF8
}

# ---------------------------------------------------------------------------
# Fingerprinting
# ---------------------------------------------------------------------------

$script:ProbePath = $null

function Get-MediaDescriptor {
    param([System.IO.FileInfo]$File)

    $ext = $File.Extension.ToLowerInvariant()
    $descriptor = [ordered]@{ ext = $ext }

    if (-not $script:ProbePath) {
        return $descriptor
    }

    try {
        $raw = & $script:ProbePath -v error `
            -select_streams v:0 `
            -show_entries 'stream=width,height:format=duration' `
            -of json $File.FullName 2>$null

        if ($LASTEXITCODE -eq 0 -and $raw) {
            $probe = ($raw -join '') | ConvertFrom-Json

            if ($probe.PSObject.Properties.Name -contains 'streams' -and $probe.streams.Count -gt 0) {
                $descriptor.width  = [int]$probe.streams[0].width
                $descriptor.height = [int]$probe.streams[0].height
            }

            if ($probe.PSObject.Properties.Name -contains 'format' -and $probe.format.duration) {
                # Random per-variant trim is tens of milliseconds, so bucket to half a
                # second: real duration changes show up, trim jitter does not.
                $seconds = [double]$probe.format.duration
                $descriptor.durationBucket = [Math]::Round($seconds * 2, 0) / 2
            }
        }
    }
    catch {
        $descriptor.probeError = $_.Exception.Message
    }

    return $descriptor
}

function Get-TreeFingerprint {
    param([string]$Root, [switch]$ProbeMedia)

    if (-not (Test-Path -LiteralPath $Root)) {
        return [ordered]@{ exists = $false }
    }

    $files = @(Get-ChildItem -LiteralPath $Root -Recurse -File -ErrorAction SilentlyContinue)
    $dirs  = @(Get-ChildItem -LiteralPath $Root -Recurse -Directory -ErrorAction SilentlyContinue)

    $rootPrefix = $Root.TrimEnd('\') + '\'
    $depths = @{}
    $extensions = @{}
    $shapes = @{}

    foreach ($file in $files) {
        $relative = $file.FullName.Substring($rootPrefix.Length)
        $depth = ($relative -split '\\').Count - 1
        $depths["$depth"] = 1 + ($(if ($depths.ContainsKey("$depth")) { $depths["$depth"] } else { 0 }))

        $ext = $file.Extension.ToLowerInvariant()
        $extensions[$ext] = 1 + ($(if ($extensions.ContainsKey($ext)) { $extensions[$ext] } else { 0 }))

        if ($ProbeMedia -and $ext -ne '.json') {
            $descriptor = Get-MediaDescriptor -File $file
            $shapeKey = '{0}|{1}x{2}|{3}' -f `
                $descriptor.ext,
                $(if ($descriptor.Contains('width'))  { $descriptor.width }  else { '?' }),
                $(if ($descriptor.Contains('height')) { $descriptor.height } else { '?' }),
                $(if ($descriptor.Contains('durationBucket')) { $descriptor.durationBucket } else { 'na' })
            $shapes[$shapeKey] = 1 + ($(if ($shapes.ContainsKey($shapeKey)) { $shapes[$shapeKey] } else { 0 }))
        }
    }

    # Directory fan-out matters: flat vs folder-per-source vs folder-per-set is exactly the
    # axis the refactor is collapsing, so record it explicitly.
    $dirDepths = @{}
    foreach ($dir in $dirs) {
        $relative = $dir.FullName.Substring($rootPrefix.Length)
        $depth = ($relative -split '\\').Count
        $dirDepths["$depth"] = 1 + ($(if ($dirDepths.ContainsKey("$depth")) { $dirDepths["$depth"] } else { 0 }))
    }

    $result = [ordered]@{
        exists      = $true
        fileCount   = $files.Count
        dirCount    = $dirs.Count
        extensions  = [ordered]@{}
        fileDepths  = [ordered]@{}
        dirDepths   = [ordered]@{}
    }

    foreach ($key in ($extensions.Keys | Sort-Object)) { $result.extensions[$key] = $extensions[$key] }
    foreach ($key in ($depths.Keys     | Sort-Object)) { $result.fileDepths[$key] = $depths[$key] }
    foreach ($key in ($dirDepths.Keys  | Sort-Object)) { $result.dirDepths[$key]  = $dirDepths[$key] }

    if ($ProbeMedia) {
        $result.shapes = [ordered]@{}
        foreach ($key in ($shapes.Keys | Sort-Object)) { $result.shapes[$key] = $shapes[$key] }
    }

    # Manifests are contractual output; record their structure, not their random keys.
    $manifests = @($files | Where-Object { $_.Name -eq 'manifest.json' })
    if ($manifests.Count -gt 0) {
        $manifestSummaries = [System.Collections.Generic.List[object]]::new()
        foreach ($manifest in $manifests) {
            try {
                $parsed = Get-Content -LiteralPath $manifest.FullName -Raw | ConvertFrom-Json
                $variantFields = @()
                if ($parsed.variants -and $parsed.variants.Count -gt 0) {
                    $variantFields = @($parsed.variants[0].PSObject.Properties.Name | Sort-Object)
                }
                $manifestSummaries.Add([ordered]@{
                    schema        = $parsed.schema
                    importRoot    = $parsed.importRoot
                    variantCount  = @($parsed.variants).Count
                    variantFields = $variantFields
                })
            }
            catch {
                $manifestSummaries.Add([ordered]@{ parseError = $_.Exception.Message })
            }
        }
        $result.manifests = $manifestSummaries
    }

    return $result
}

# ---------------------------------------------------------------------------
# Scenarios
# ---------------------------------------------------------------------------

# Each scenario stages files into one lane's input folder and invokes that lane's handler
# exactly the way the poll loop does, then fingerprints output, original and failed.
# Each scenario stages files into one preset's input folder and polls that preset exactly
# the way the watcher loop does, then fingerprints output, original and failed.
function Get-Scenarios {
    return @(
        @{ Name = 'default';    Preset = 'bulk';    Workspace = 'LC'; Stage = { @(Get-CorpusFiles -Kind 'mp4' -Count 1) + @(Get-CorpusFiles -Kind 'jpg' -Count 1) + @(Get-CorpusFiles -Kind 'heic' -Count 1) } }
        @{ Name = 'videoclean'; Preset = 'video-clean'; Workspace = 'LC'; Stage = { Get-CorpusFiles -Kind 'mp4' -Count 1 } }
        @{ Name = 'images';     Preset = 'image-bulk';     Workspace = 'LC'; Stage = { @(Get-CorpusFiles -Kind 'jpg' -Count 1) + @(Get-CorpusFiles -Kind 'png' -Count 1) + @(Get-CorpusFiles -Kind 'heic' -Count 1) } }
        @{ Name = 'imageclean'; Preset = 'image-clean'; Workspace = 'LC'; Stage = { @(Get-CorpusFiles -Kind 'jpg' -Count 1) + @(Get-CorpusFiles -Kind 'png' -Count 1) + @(Get-CorpusFiles -Kind 'heic' -Count 1) } }
        @{ Name = 'sets';       Preset = 'sets';       Workspace = 'LC'; Stage = { @(Get-CorpusFiles -Kind 'mp4' -Count 1) + @(Get-CorpusFiles -Kind 'jpg' -Count 1) } }
        @{ Name = 'setbatch';   Preset = 'sets-batch';   Workspace = 'LC'; Stage = { @(Get-CorpusFiles -Kind 'mp4' -Count 2) + @(Get-CorpusFiles -Kind 'jpg' -Count 2) } }
        @{ Name = 'assetstore'; Preset = 'asset-store'; Workspace = 'LC'; Stage = { @(Get-CorpusFiles -Kind 'mp4' -Count 2) + @(Get-CorpusFiles -Kind 'jpg' -Count 2) } }
        @{ Name = 'long';       Preset = 'video-long';       Workspace = 'LC'; Stage = { Get-CorpusFiles -Kind 'mp4' -Count 1 } }
        # A .mov no longer needs a lane of its own: any preset that takes video normalizes it.
        @{ Name = 'mov-input';  Preset = 'video-clean'; Workspace = 'MD'; Stage = { Get-CorpusFiles -Kind 'mov' -Count 1 } }
    )
}

# ---------------------------------------------------------------------------
# Run
# ---------------------------------------------------------------------------

if ($BuildCorpus) {
    New-Item -ItemType Directory -Path $SandboxRoot -Force | Out-Null
    Build-Corpus -Source $CorpusSource -Destination $corpusDir
}
elseif ($BuildSyntheticCorpus) {
    New-Item -ItemType Directory -Path $SandboxRoot -Force | Out-Null
    Build-SyntheticCorpus -Destination $corpusDir
}

if (-not (Test-Path -LiteralPath $corpusDir)) {
    throw "No corpus at $corpusDir. Run once with -BuildCorpus."
}

Write-Host "Sandbox:      $pipelineRoot"
Write-Host "Under test:   $ScriptPath"
Write-Host "Engine:       $Engine"
Write-Host "Mode:         $Mode"
Write-Host ''

New-Sandbox

# Loading with -AsLibrary gives us every function with no mutex and no poll loop.
. (Join-Path $appDir 'watch-media.ps1') -AsLibrary

Initialize-Folders
Test-ExternalTools
$script:ProbePath = $script:FFprobePath

$nativeWorker = $null
if ($Engine -eq 'Native') {
    if (-not (Test-Path -LiteralPath $DotnetPath)) { throw "dotnet not found: $DotnetPath" }
    if (-not (Test-Path -LiteralPath $WorkerProject)) { throw "Worker project not found: $WorkerProject" }

    $nativeOutput = Join-Path $appDir 'native-worker'
    & $DotnetPath publish $WorkerProject -c Release --no-self-contained -o $nativeOutput --nologo | Out-Host
    if ($LASTEXITCODE -ne 0) { throw 'Could not publish the native worker for parity testing.' }
    $nativeWorker = Join-Path $nativeOutput 'media-pipeline-worker.exe'
    if (-not (Test-Path -LiteralPath $nativeWorker)) {
        throw "Published worker not found: $nativeWorker"
    }
}

$scenarios = Get-Scenarios
if ($Only) {
    $scenarios = @($scenarios | Where-Object { $Only -contains $_.Name })
    if ($scenarios.Count -eq 0) { throw "No scenarios matched: $($Only -join ', ')" }
}

$results = [ordered]@{}

foreach ($scenario in $scenarios) {
    Write-Host ("Running {0} ..." -f $scenario.Name) -NoNewline

    Use-PipelineWorkspace -WorkspaceName $scenario.Workspace

    $preset = Get-PipelinePreset -Name $scenario.Preset
    $paths = Get-PresetWorkspacePaths -PresetName $scenario.Preset -WorkspaceName $scenario.Workspace

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $scenarioError = $null

    try {
        foreach ($file in (& $scenario.Stage)) {
            Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $paths.InputDir $file.Name) -Force
        }

        if ($Engine -eq 'Native') {
            & $nativeWorker once --config (Join-Path $appDir 'config.ini') --assume-stable | Out-Host
            if ($LASTEXITCODE -ne 0) { throw "Native worker exited with $LASTEXITCODE." }
        }
        else {
            # A batch preset needs two polls to settle: the first records the signature, the
            # second sees it unchanged and processes.
            Invoke-PresetPoll -Preset $preset -WorkspaceName $scenario.Workspace
            if ($preset.Batch -eq 'PerGroup') {
                Invoke-PresetPoll -Preset $preset -WorkspaceName $scenario.Workspace
            }
        }
    }
    catch {
        $scenarioError = $_.Exception.Message
    }

    $stopwatch.Stop()

    $fingerprint = [ordered]@{
        output         = Get-TreeFingerprint -Root $paths.OutputDir -ProbeMedia
        original       = Get-TreeFingerprint -Root $paths.OriginalDir
        failed         = Get-TreeFingerprint -Root $paths.FailedDir
        inputRemaining = Get-TreeFingerprint -Root $paths.InputDir
    }

    $results[$scenario.Name] = [ordered]@{
        error = $scenarioError
        tree  = $fingerprint
    }

    if ($scenarioError) {
        Write-Host (" error after {0:N1}s" -f $stopwatch.Elapsed.TotalSeconds) -ForegroundColor Red
        Write-Host "    $scenarioError" -ForegroundColor Red
    }
    else {
        Write-Host (" ok ({0:N1}s)" -f $stopwatch.Elapsed.TotalSeconds) -ForegroundColor Green
    }
}

$report = [ordered]@{
    capturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
    scriptPath    = $ScriptPath
    scriptSha256  = (Get-FileHash -LiteralPath $ScriptPath -Algorithm SHA256).Hash
    scenarios     = $results
}

$json = $report | ConvertTo-Json -Depth 12

if ($Mode -eq 'Capture') {
    $baselineDir = Split-Path -Parent $BaselinePath
    if ($baselineDir -and -not (Test-Path -LiteralPath $baselineDir)) {
        New-Item -ItemType Directory -Path $baselineDir -Force | Out-Null
    }
    $json | Set-Content -LiteralPath $BaselinePath -Encoding UTF8
    Write-Host ''
    Write-Host "Baseline written: $BaselinePath"
    $exitCode = 0
}
else {
    if (-not (Test-Path -LiteralPath $BaselinePath)) {
        throw "No baseline at $BaselinePath. Run with -Mode Capture first."
    }

    $baseline = Get-Content -LiteralPath $BaselinePath -Raw | ConvertFrom-Json
    $comparisonBaseline = $baseline
    $acceptedManifestFix = $false

    # The legacy asset-store path currently drops every record before serializing the
    # manifest. The native worker deliberately fixes that defect. Accept only that precise
    # difference: the manifest must account for every generated media file and expose the
    # contractual identity and path fields.
    if ($Engine -eq 'Native') {
        $beforeManifest = @($baseline.scenarios.assetstore.tree.output.manifests)[0]
        $afterManifest = @($report.scenarios.assetstore.tree.output.manifests)[0]
        $expectedVariants = $report.scenarios.assetstore.tree.output.fileCount - 1
        $requiredFields = @('familyKey', 'variantKey', 'path', 'renditionSetKey', 'sourceOriginalName')
        $hasRequiredFields = $requiredFields.Count -eq @(
            $requiredFields | Where-Object { $afterManifest.variantFields -contains $_ }
        ).Count

        if ($beforeManifest.variantCount -eq 0 -and
            $afterManifest.variantCount -eq $expectedVariants -and
            $afterManifest.schema -eq $beforeManifest.schema -and
            $hasRequiredFields) {
            $comparisonBaseline = ($baseline | ConvertTo-Json -Depth 12 | ConvertFrom-Json)
            $comparisonBaseline.scenarios.assetstore.tree.output.manifests =
                $report.scenarios.assetstore.tree.output.manifests
            $acceptedManifestFix = $true
        }
    }

    # Compare the scenario trees only. Timestamps and the script hash are expected to differ.
    $baselineTrees = ($comparisonBaseline.scenarios | ConvertTo-Json -Depth 12)
    $currentTrees  = ($report.scenarios   | ConvertTo-Json -Depth 12)

    if ($baselineTrees -eq $currentTrees) {
        Write-Host ''
        Write-Host 'PARITY OK: every scenario matches the baseline.' -ForegroundColor Green
        if ($acceptedManifestFix) {
            Write-Host 'Accepted fix: asset-store manifests now contain one record per generated variant.' -ForegroundColor Green
        }
        $exitCode = 0
    }
    else {
        $actualPath = [System.IO.Path]::ChangeExtension($BaselinePath, '.actual.json')
        $json | Set-Content -LiteralPath $actualPath -Encoding UTF8

        Write-Host ''
        Write-Host 'PARITY DIFF: behavior changed.' -ForegroundColor Yellow
        Write-Host "  baseline: $BaselinePath"
        Write-Host "  actual:   $actualPath"
        Write-Host ''

        # Per-scenario summary so the offending lane is obvious without reading the JSON.
        foreach ($name in $results.Keys) {
            $before = $comparisonBaseline.scenarios.PSObject.Properties[$name]
            $beforeJson = if ($before) { $before.Value | ConvertTo-Json -Depth 12 } else { '<missing>' }
            $afterJson  = $results[$name] | ConvertTo-Json -Depth 12

            if ($beforeJson -eq $afterJson) {
                Write-Host ("  {0,-12} same" -f $name) -ForegroundColor DarkGray
            }
            else {
                Write-Host ("  {0,-12} CHANGED" -f $name) -ForegroundColor Yellow
            }
        }

        $exitCode = 1
    }
}

if (-not $KeepSandbox) {
    try {
        Remove-Item -LiteralPath $appDir -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $pipelineRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
    catch {
    }
}
else {
    Write-Host ''
    Write-Host "Sandbox kept: $pipelineRoot"
}

exit $exitCode
