# Native worker migration

The cross-platform C# worker now reproduces the PowerShell pipeline and is ready for an explicit,
reversible Windows cutover. The two implementations must never process the same pipeline root at
the same time.

## Chosen application structure

```text
MediaPipeline.Core       shared configuration, paths, processing, and runtime contracts
MediaPipeline.Worker     headless Windows and macOS executable
MediaPipelineTray        existing Windows WPF interface
MediaPipelineMac         SwiftUI interface with AppKit lifecycle integration
```

The macOS app bundles an `osx-arm64` worker under `Contents/Helpers`. FFmpeg, FFprobe,
ExifTool, rclone, and SSH can be resolved from an application tools directory or Homebrew.

The worker owns media and queue state. Platform interfaces handle windows, the Windows
notification area or macOS menu bar, notifications, opening folders, and start-at-login.

## Implemented

- A `net8.0` core with no Windows desktop dependency
- Typed resolution of the existing `config.ini`, including preset inheritance and safe fallbacks
- Workspace-first native paths on Windows and macOS
- Cross-platform control flags and an exclusive worker file lock
- Atomic versioned status output for native interfaces
- The existing JSONL activity-event contract with serialized writes
- Process execution with cancellation and child-process cleanup
- Tool discovery from an application bundle, `PATH`, Homebrew, WinGet, or the current portable tools
- Runtime H.264 probing for VideoToolbox, NVENC, AMF, and libx264
- Media classification, stable-file tracking, rotation-aware probing, and HEIC working copies
- Image recrop variants, video micro-trim variants, size caps, MOV remuxing, and segmentation
- VideoToolbox on Apple silicon, NVENC and AMF on Windows, with libx264 fallback
- Flat, per-source, and per-set output grouping with transactional rollback
- Per-file and settled-folder batch modes, manifests, archiving, and retention
- Resumable chunked upload with per-part SHA-256 verification and streaming remote assembly
- Job zip creation, queue re-entry, pause scopes, clean stop, status, and JSONL events
- Commands for continuous operation, one pass, status, control, uploads, archives, and recompression
- Self-contained `win-x64` and `osx-arm64` single-file releases
- Structural parity across every shipped preset, HEIC, MOV, grouped output, and long video
- A SwiftUI sidebar app plus menu bar panel, native notifications, login-item support, settings,
  preset editing, activity controls, archiving, and upload progress

The parity harness reports one intentional correction. The PowerShell asset-store path creates the
right media files but serializes an empty `variants` array. The native worker records every
generated variant and verifies that the manifest count matches the media file count.

## Worker commands

```text
check        validate configuration, tools, and encoder
run          watch continuously
once         process one polling pass
status       return worker liveness and the last status document as JSON
pause        pause all work, one preset, or one preset/workspace lane
resume       clear the matching pause
stop         request a clean stop after the current operation
requeue      move failed files back into a lane input folder
archive      collect selected job outputs into a workspace zip
upload       upload one staged file
upload-all   upload staged files sequentially
recompress   apply current size caps to existing oversized outputs
```

## Release and Windows cutover

Publish both worker targets:

```powershell
pwsh -File tools\Publish-NativeWorker.ps1 -Runtime win-x64
pwsh -File tools\Publish-NativeWorker.ps1 -Runtime osx-arm64
```

`Switch-ToNativeWorker.ps1` validates the native executable before touching startup state. It asks
the PowerShell watcher to stop cleanly, waits for its mutex to disappear, disables its scheduled
task, registers the native task, and starts it. It aborts instead of killing a watcher that does
not stop within the timeout.

```powershell
pwsh -File tools\Switch-ToNativeWorker.ps1
```

The original scheduled task is disabled rather than deleted. Roll back with:

```powershell
pwsh -File tools\Restore-PowerShellWorker.ps1
```

The INI reader is transitional. A later configuration version can use typed JSON, but the worker
should first reproduce the current processing behavior. Changing format and behavior together
would make parity failures harder to diagnose.

## Commands

The SDK on the current Windows development machine is outside `PATH`, so these examples use its
absolute path.

```powershell
& 'C:\Program Files\dotnet\dotnet.exe' build src\MediaPipeline.Worker\MediaPipeline.Worker.csproj -c Release
& 'C:\Program Files\dotnet\dotnet.exe' run --project tests\MediaPipeline.Core.SelfTest\MediaPipeline.Core.SelfTest.csproj -c Release
& 'C:\Program Files\dotnet\dotnet.exe' run --project tests\MediaPipeline.Integration.SelfTest\MediaPipeline.Integration.SelfTest.csproj -c Release
pwsh -File tools\Test-PipelineParity.ps1 -Mode Capture -Engine Legacy -BuildSyntheticCorpus -BaselinePath "$env:TEMP\native-parity.json"
pwsh -File tools\Test-PipelineParity.ps1 -Mode Compare -Engine Native -BaselinePath "$env:TEMP\native-parity.json"
```
