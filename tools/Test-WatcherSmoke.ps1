<#
.SYNOPSIS
    End-to-end smoke test: starts a real watcher process against a sandbox root.

.DESCRIPTION
    The parity harness calls the handlers directly, which never exercises the poll loop, the
    single-instance mutex, the control folder, or the status file. This starts an actual
    watcher process, drops files in, waits for output, and stops it through the control
    folder rather than killing it.

    It uses its own pipeline root, so the single-instance lock does not collide with a
    watcher already running against the real root.
#>
[CmdletBinding()]
param(
    [string]$SandboxRoot = (Join-Path $env:TEMP 'mp-smoke'),
    [string]$ScriptPath,
    [string]$CorpusSource = (Join-Path $env:TEMP 'mp-parity\corpus'),
    [int]$TimeoutSeconds = 120,
    [switch]$KeepSandbox
)

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
if (-not $ScriptPath) { $ScriptPath = Join-Path $repoRoot 'watch-media.ps1' }

$appDir = Join-Path $SandboxRoot 'app'
$pipelineRoot = Join-Path $SandboxRoot 'root'

$config = @'
[General]
PipelineRoot = {0}

[Video]
Crf = 30
X264Preset = ultrafast
MaxWidth = 480
PreferNvenc = false
PreferAmf = false
SizeCapMB = 0

[Images]
ImageProcessingConcurrency = 99
JpegQuality = 8

[Timing]
StableSeconds = 0
TimeoutSeconds = 60
PollSeconds = 1

[Archive]
ArchiveEnabled = false
AssetRetentionDays = 0

[preset quick]
VideoCopies = 1
ImageCopies = 2
'@

function Write-Step { param([string]$Text) Write-Host "`n== $Text" -ForegroundColor Cyan }

if (Test-Path -LiteralPath $SandboxRoot) { Remove-Item -LiteralPath $SandboxRoot -Recurse -Force }
New-Item -ItemType Directory -Path $appDir -Force | Out-Null
New-Item -ItemType Directory -Path $pipelineRoot -Force | Out-Null

Copy-Item -LiteralPath $ScriptPath -Destination (Join-Path $appDir 'watch-media.ps1') -Force
($config -f $pipelineRoot) | Set-Content -LiteralPath (Join-Path $appDir 'config.ini') -Encoding UTF8

if (-not (Test-Path -LiteralPath $CorpusSource)) {
    throw "No corpus at $CorpusSource. Run Test-PipelineParity.ps1 -BuildCorpus first."
}

$failures = New-Object System.Collections.Generic.List[string]
function Assert-That {
    param([string]$Label, [bool]$Condition, [string]$Detail)
    if ($Condition) {
        Write-Host ("   PASS  {0}" -f $Label) -ForegroundColor Green
    }
    else {
        Write-Host ("   FAIL  {0} :: {1}" -f $Label, $Detail) -ForegroundColor Red
        $failures.Add($Label) | Out-Null
    }
}

Write-Step "Starting a real watcher process"
$watcher = Start-Process -FilePath 'C:\Tools\pwsh\pwsh.exe' `
    -ArgumentList @('-NoProfile', '-File', (Join-Path $appDir 'watch-media.ps1')) `
    -PassThru -WindowStyle Hidden

Write-Host "   pid $($watcher.Id)"

try {
    $inputDir = Join-Path $pipelineRoot 'LC\quick\input'
    $outputDir = Join-Path $pipelineRoot 'LC\quick\output'
    $statusFile = Join-Path $pipelineRoot 'status\watcher.json'

    $deadline = (Get-Date).AddSeconds(30)
    while (-not (Test-Path -LiteralPath $inputDir) -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 250
    }
    Assert-That "watcher creates its folders" (Test-Path -LiteralPath $inputDir) "no $inputDir"

    $deadline = (Get-Date).AddSeconds(30)
    while (-not (Test-Path -LiteralPath $statusFile) -and (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 250
    }
    Assert-That "watcher writes status\watcher.json" (Test-Path -LiteralPath $statusFile) "no status file"

    if (Test-Path -LiteralPath $statusFile) {
        $status = Get-Content -LiteralPath $statusFile -Raw | ConvertFrom-Json
        Assert-That "status reports the running pid" ($status.pid -eq $watcher.Id) "status pid $($status.pid), process $($watcher.Id)"
        Assert-That "status lists the quick preset" (@($status.presets | Where-Object { $_.name -eq 'quick' }).Count -eq 1) "presets: $($status.presets.name -join ',')"
    }

    Write-Step "Second instance must refuse to start"
    $duplicate = Start-Process -FilePath 'C:\Tools\pwsh\pwsh.exe' `
        -ArgumentList @('-NoProfile', '-File', (Join-Path $appDir 'watch-media.ps1')) `
        -PassThru -WindowStyle Hidden -Wait
    Assert-That "duplicate exits without taking over" ($duplicate.ExitCode -eq 0) "exit code $($duplicate.ExitCode)"
    Assert-That "original still running" (-not $watcher.HasExited) "original exited"

    Write-Step "Dropping media"
    $video = @(Get-ChildItem -LiteralPath $CorpusSource -File -Filter 'mp4-*')[0]
    $image = @(Get-ChildItem -LiteralPath $CorpusSource -File -Filter 'jpg-*')[0]
    Copy-Item -LiteralPath $video.FullName -Destination (Join-Path $inputDir $video.Name)
    Copy-Item -LiteralPath $image.FullName -Destination (Join-Path $inputDir $image.Name)

    # quick makes 1 video copy and 2 image copies
    $expected = 3

    # Wait on the input folder draining, not on a file count in output. FFmpeg creates its
    # output file before it finishes writing it, so counting output files would treat an
    # encode still in progress as finished.
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $remaining = @(Get-ChildItem -LiteralPath $inputDir -File -ErrorAction SilentlyContinue).Count
        if ($remaining -eq 0) {
            # The source moves after the last variant lands, so give the move a moment to
            # settle before inspecting the folders.
            Start-Sleep -Milliseconds 750
            break
        }

        Start-Sleep -Milliseconds 500
    }

    $produced = @(Get-ChildItem -LiteralPath $outputDir -File -ErrorAction SilentlyContinue)
    Assert-That "produces the expected outputs" ($produced.Count -eq $expected) "got $($produced.Count), wanted $expected"
    Assert-That "image outputs are .jpg" (@($produced | Where-Object { $_.Extension -eq '.JPG' }).Count -eq 2) "extensions: $(($produced.Extension | Sort-Object -Unique) -join ',')"
    Assert-That "sources moved to original" (@(Get-ChildItem -LiteralPath (Join-Path $pipelineRoot 'LC\quick\original') -File -ErrorAction SilentlyContinue).Count -eq 2) "original is not 2"
    Assert-That "input drained" (@(Get-ChildItem -LiteralPath $inputDir -File).Count -eq 0) "input not empty"

    $watcherLog = Get-Content -LiteralPath (@(Get-ChildItem -LiteralPath (Join-Path $pipelineRoot 'logs') -Filter 'media-pipeline-*.log')[0].FullName) -Raw
    Assert-That "caps unsafe image concurrency" ($watcherLog -match 'requested 99 workers; capped at \d+') "no concurrency cap warning"

    Write-Step "Event stream"
    $eventFile = @(Get-ChildItem -LiteralPath (Join-Path $pipelineRoot 'logs') -Filter 'events-*.jsonl' -ErrorAction SilentlyContinue)[0]
    Assert-That "event stream written" ($null -ne $eventFile) "no events file"

    if ($eventFile) {
        $events = @(Get-Content -LiteralPath $eventFile.FullName | Where-Object { $_.Trim() } | ForEach-Object { $_ | ConvertFrom-Json })
        $variants = @($events | Where-Object { $_.ev -eq 'job.variant' })
        Assert-That "watcher.start recorded" (@($events | Where-Object { $_.ev -eq 'watcher.start' }).Count -eq 1) "not exactly one"
        Assert-That "one variant event per output" ($variants.Count -eq $expected) "got $($variants.Count)"
        Assert-That "variants carry a job id" (@($variants | Where-Object { $_.jobId }).Count -eq $variants.Count) "some missing jobId"
        Assert-That "variants report n of total" (@($variants | Where-Object { $_.total -gt 0 }).Count -eq $variants.Count) "some missing total"
    }

    Write-Step "Pause blocks new work"
    $control = Join-Path $pipelineRoot 'control'
    New-Item -ItemType File -Path (Join-Path $control 'pause.quick') -Force | Out-Null
    Start-Sleep -Seconds 2
    Copy-Item -LiteralPath $image.FullName -Destination (Join-Path $inputDir "paused-$($image.Name)")
    Start-Sleep -Seconds 4
    Assert-That "paused lane leaves its queue alone" (@(Get-ChildItem -LiteralPath $inputDir -File).Count -eq 1) "queue drained while paused"

    Remove-Item -LiteralPath (Join-Path $control 'pause.quick') -Force
    $deadline = (Get-Date).AddSeconds(60)
    while ((Get-Date) -lt $deadline -and @(Get-ChildItem -LiteralPath $inputDir -File).Count -gt 0) {
        Start-Sleep -Milliseconds 500
    }
    Assert-That "resumes after pause is cleared" (@(Get-ChildItem -LiteralPath $inputDir -File).Count -eq 0) "still queued"

    Write-Step "Graceful stop"
    New-Item -ItemType File -Path (Join-Path $control 'stop') -Force | Out-Null
    $exited = $watcher.WaitForExit(30000)
    Assert-That "stops on the control flag" $exited "still running after 30s"
    Assert-That "clears the stop flag on the way out" (-not (Test-Path -LiteralPath (Join-Path $control 'stop'))) "stop flag left behind"

    if ($exited) {
        $log = Get-Content -LiteralPath (@(Get-ChildItem -LiteralPath (Join-Path $pipelineRoot 'logs') -Filter 'media-pipeline-*.log')[0].FullName) -Raw
        Assert-That "logs a clean shutdown" ($log -match 'shutting down cleanly') "no shutdown line"
    }
}
finally {
    if (-not $watcher.HasExited) {
        Write-Host "`n   forcing watcher exit" -ForegroundColor Yellow
        $watcher.Kill()
    }
}

Write-Host ""
if ($failures.Count -eq 0) {
    Write-Host "SMOKE OK: watcher starts, processes, pauses, and stops cleanly." -ForegroundColor Green
    $exitCode = 0
}
else {
    Write-Host ("SMOKE FAILED: {0}" -f ($failures -join '; ')) -ForegroundColor Red
    $exitCode = 1
}

if (-not $KeepSandbox) {
    Remove-Item -LiteralPath $SandboxRoot -Recurse -Force -ErrorAction SilentlyContinue
}
else {
    Write-Host "Sandbox kept: $pipelineRoot"
}

exit $exitCode
