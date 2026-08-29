# media-pipeline

A Windows watcher that turns media you drop into a folder into processed, differentiated copies.

Drop files into an input folder, get processed copies out. Video is re-encoded to H.264 MP4 with
AAC audio, images come out as JPEG, and all metadata is stripped. Every copy is slightly different
from its siblings, so no two outputs are byte-identical.

---

## Quick Start

1. Double-click **`Install.bat`** and approve the prompts. It installs FFmpeg, ExifTool and
   PowerShell 7, creates the folders, and starts the watcher at login.
2. Drop a video or photo into `D:\MediaPipeline\LCulk\input`.
3. Processed copies appear in `D:\MediaPipeline\LCulk\output` a moment later, and the file
   you dropped moves to `original`.

That is the whole loop. Everything below is about changing what comes out.

---

## The buttons (double-click `.bat` files)

| File | What it does |
| --- | --- |
| **`Install.bat`** | One-time setup: installs tools, creates folders, enables auto-start, starts the watcher. Safe to run again to repair the setup. |
| **`Edit Config.bat`** | Opens `config.ini` (your settings) in Notepad. |
| **`Restart Watcher.bat`** | Restarts the watcher so changes you saved in `config.ini` take effect. |
| **`Uninstall.bat`** | Stops the watcher and removes auto-start. Leaves your media files, settings, and installed tools alone. |

---

## Presets

A **preset** is a named recipe with its own drop folder. It decides how many copies each file
produces and how they are arranged.

```text
D:\MediaPipeline\<workspace>\<preset>\input
```

Copy counts are **separate for video and images**, so one folder can take a mixed batch of clips
and photos and treat each correctly. A preset that sets `VideoCopies = 0` simply ignores video.

The shipped presets:

| Preset | What it does |
| --- | --- |
| `bulk` | 20 differentiated copies of whatever you drop in |
| `video-clean` | One cleaned copy per video, images ignored |
| `image-clean` | One cleaned copy per image, video ignored |
| `image-bulk` | 100 variants from a single photo |
| `sets` | Every source file gets its own output folder holding its copies |
| `sets-batch` | Treats the whole folder as one batch, produces 10 complete sets |
| `asset-store` | Like `sets-batch`, plus a `manifest.json` describing every file |
| `video-long` | Splits long videos into segments, then makes copies of each |

**Delete any preset you do not use.** Add one by copying a block in `config.ini` and renaming it;
its folders are created on the next restart. If you only ever want "clean this up" and "make me a
lot of copies", two presets is a perfectly good setup.

### Changing how much comes out

Most of the time this is a one-line edit in `config.ini`:

```ini
[preset image-bulk]
VideoCopies = 0
ImageCopies = 100     ; <- change this
```

### Preset options

| Option | Values |
| --- | --- |
| `VideoCopies` | Outputs per source video. `0` ignores video. |
| `ImageCopies` | Outputs per source image. `0` ignores images. |
| `CopiesAlternate` | If set, consecutive files alternate between the count above and this one. |
| `Grouping` | `Flat` (one folder), `PerSource` (a folder per file), `PerSet` (complete sets) |
| `SetCount` | How many sets, when `Grouping = PerSet`. |
| `Batch` | `PerFile` processes each file as it settles. `PerGroup` waits for the whole folder. |
| `Segment` | `true` splits long videos into segments first. |
| `Manifest` | `true` writes a `manifest.json` next to the output. |
| `Normalize` | `true` converts `.mov` and `.heic` before processing. Default `true`. |
| `OnFailure` | `PreservePartial`, `DeleteFiles`, or `DeleteContainer`. |
| `Parallel` | `OverFiles` (default) or `Sequential`. |
| `Enabled` | `false` switches a preset off without deleting it. |

Any setting from the `[Video]` or `[Images]` sections can also be set on a preset, and overrides
the default for that preset only. That is how `video-long` uses a lower quality than everything else:

```ini
[preset video-long]
Segment = true
NvencCq = 28
```

---

## Workspaces

Workspaces keep different clients or projects apart. Each preset has one folder per workspace:

```text
D:\MediaPipeline\LCulk\input
D:\MediaPipeline\default\MD\input
D:\MediaPipeline\default\YL\input
```

The workspaces are `LC`, `MD`, `YL`, `PL`, and `general`. They are scanned independently and never
mix output.

---

## Folder Structure

Workspace first, then preset, so everything belonging to one client sits together. The remote
is laid out the same way.

```text
D:\MediaPipeline  LC\                      <- workspace
    video-clean      input\               <- drop files here
      output\              <- processed copies appear here
      original\            <- sources move here after success
      failed\              <- sources move here if processing fails
      archive\             <- old output is moved here
      work\                <- temporary files, segmenting only
    image-bulk\...
    sets\...
    sync\                 <- files staged for upload to the remote
  MD\...
  YL\...
  PL\...
  general\...
  logs    media-pipeline-YYYYMMDD.log     <- human-readable log
    events-YYYYMMDD.jsonl           <- machine-readable event stream
  status    watcher.json                    <- current state
  control\                          <- pause and stop flags
```

Retention deletes entries from `archive`, `original`, `failed`, `work`, `sync` and
`.sync-parts` after `AssetRetentionDays`. The `image-bulk` preset is excluded, so its assets
are kept.

Upgrading from the older preset-first layout:

```bash
pwsh -File tools\Migrate-FolderLayout.ps1 -WhatIf
```

It prints exactly what it would move and changes nothing. Re-run without `-WhatIf` to apply.

## Supported Files

Videos:

```text
.mp4, .mov, .mkv, .webm, .avi
```

Images:

```text
.jpg, .jpeg, .png, .webp, .heic
```

Image processing runs at a maximum of six files at once. Set `ImageProcessingConcurrency = auto`
to choose that limit from the CPU count. Larger manual values are capped to prevent FFmpeg from
exhausting memory on batches of full-resolution photos.

Temporary browser download files are ignored:

```text
.crdownload, .tmp, .part, .download
```

`.mov` and `.heic` are converted automatically before processing, so there is no separate convert
step and no convert folder. Drop them anywhere.

---

## How It Works

1. The watcher polls each preset's `input` folder every `PollSeconds`.
2. A file must stop changing for `StableSeconds` before it is touched, which lets browser downloads
   finish. A `PerGroup` preset waits for the whole folder to settle instead.
3. `.mov` and `.heic` sources are normalized first.
4. Video is re-encoded to H.264 MP4 with AAC audio, capped at `MaxWidth`, with a tiny random amount
   trimmed off the end of each copy. Images get a tiny random crop scaled back to the original
   dimensions. Both have all metadata stripped.
5. Outputs land in `output`, arranged according to `Grouping`.
6. The source moves to `original`, or to `failed` if anything went wrong.

Every copy differs from its siblings: video by its trim, images by their crop. Two copies of the
same source are never byte-identical.

---

## Output File Names

Output names are random and not based on the source filename. They use the iPhone Camera naming
style `IMG_####` with a random four-digit number and an uppercase extension. They contain no preset
codes, dates, variant numbers, segment numbers, set numbers, or sequence counters. Batch and set
folders use varied random word combinations.

```text
IMG_0274.MP4
IMG_4821.JPG
archive_collection_73642\
  sunny upload\
    IMG_5198.MP4
```

---

## The manifest

A preset with `Manifest = true` writes `manifest.json` into its batch folder, describing every file
it generated:

```json
{
  "schema": "heatup.assetStoreMediaManifest.v1",
  "generatedAt": "2026-08-22T09:28:03.5261057Z",
  "importRoot": ".",
  "variants": [
    {
      "familyKey": "clip_a",
      "variantKey": "clip_a__sunny_upload",
      "path": "sunny upload/IMG_5198.MP4",
      "renditionSetKey": "sunny upload",
      "generationBatchKey": "archive_collection_73642",
      "sourceOriginalName": "clip a.mov",
      "durationSeconds": 12.34,
      "sizeBytes": 1048576,
      "transformProfile": "preset_video_micro_trim",
      "metadata": { "encoder": "h264_nvenc", "trimMs": 27, "maxWidth": 1080 }
    }
  ]
}
```

Paths are relative to the manifest, so `set_folder/file` resolves next to it.

---

## Watching progress

Two machine-readable surfaces sit alongside the human-readable log.

**`logs\events-YYYYMMDD.jsonl`** is an append-only event stream, one JSON object per line. Every
event from one job carries the same `jobId`, so progress stays attributable even when several files
are processed at once:

```json
{"ts":"2026-08-22T09:27:11Z","ev":"job.variant","jobId":"eb631a92","preset":"long","workspace":"LC","file":"clip.mp4","n":4,"total":6}
```

Events are `watcher.start`, `watcher.stop`, `job.start`, `job.variant`, `job.done`, `job.failed`.

**`status\watcher.json`** is rewritten after each sweep with the process id, active encoder, every
preset's resolved options, and the queue depth of each lane.

---

## Controlling the watcher

Drop an empty file into `D:\MediaPipeline\control\`:

| File | Effect |
| --- | --- |
| `stop` | Finish the current file, then exit cleanly. |
| `pause` | Stop picking up new work. Nothing queued is lost. |
| `pause.<preset>` | Pause one preset across all workspaces. |
| `pause.<preset>.<workspace>` | Pause a single lane. |

Delete the file to resume. To retry something that failed, move it from `failed` back to `input`.

Prefer `stop` over killing the process: killing it mid-encode orphans FFmpeg, leaves partial
outputs, and strands the input file with no failed-move.

---

## The tray app

A desktop app that shows what the pipeline is doing and lets you change it without editing
files. It lives in the notification area and keeps running when you close the window; Quit is
in the tray menu.

```bash
dotnet build tray-app -c Release
```

The built `MediaPipelineTray.exe` finds the pipeline by looking for `watch-media.ps1` beside it
and walking up, then reads `PipelineRoot` from the `config.ini` next to it, so it always agrees
with the watcher about where things are.

| Tab | What it does |
| --- | --- |
| **Activity** | What is running, with real progress, plus what is queued, what finished, and what failed |
| **Presets** | Every preset's options, with help text and whether each value is overridden or inherited |
| **Settings** | The global defaults every preset inherits |
| **Uploads** | Chunked upload of large files to the remote |

It starts with Windows and lives in the notification area. **Closing the window hides it**;
the only way to exit is right-clicking the tray icon and choosing **Quit Media Pipeline**.

At Windows sign-in, the app starts directly in the notification area without opening its
window. Windows controls whether the icon stays on the taskbar or behind the `^` chevron, and
the app respects that choice. Startup can be turned off in Settings.

The status bar is always visible and can start, stop, restart, or pause the watcher. Stopping
asks the watcher to finish the file it is on rather than killing it.

It is monochrome on purpose. Green means a job finished and red means one failed, and those are
the only two places colour appears, so anything coloured is worth looking at. Work in progress
is drawn at full contrast rather than tinted, and idle and queued differ by shape.

Nothing is scraped from the text log. Progress comes from the event stream grouped by `jobId`,
state from `status\watcher.json`, and control from flag files in `control\`. Whether the
watcher is alive comes from its single-instance mutex, so a stale status file cannot fool it.

### Uploads

Large files are split into chunks, sent one at a time, and reassembled on the remote. Each
chunk carries a SHA-256 that the remote verifies before appending it, so a corrupt transfer is
caught rather than assembled in. A chunk that fails retries on its own instead of failing the
whole transfer, and cancelling keeps the local parts so the next attempt resumes.

Remote settings live in `config.ini` and default to the same values the sync scripts use:

```ini
[Upload]
RemoteDirectory = D:\MediaPipeline\sync
DeleteAfterUpload = false
ChunkSizeMB = 256
ParallelChunks = 4
```

Put a file in `<workspace>\sync` and it uploads to that workspace on the remote. The Uploads
tab lists every workspace; expand one to upload a single file, or use **Upload all** for the
whole folder. Files upload one at a time.

With `DeleteAfterUpload` on, the local file is released only after the remote copy has been read
back and confirmed the right size. Windows deletes the verified local file. The macOS app moves
it to Trash, where Finder can recover it until Trash is emptied. The setting is off by default.

---

## Changing Settings

Edit `config.ini` (or double-click **`Edit Config.bat`**), save, then double-click
**`Restart Watcher.bat`**.

`config.ini` has two kinds of section. `[Video]`, `[Images]`, `[Timing]` and `[Archive]` set the
defaults for everything. A `[preset <name>]` section overrides those defaults for one preset. A
preset only lists what it changes.

If a setting is missing or unparseable the watcher falls back to its built-in default rather than
failing, so a typo cannot stop processing.

---

# Advanced / Manual Setup

## Required Tools

- **FFmpeg** and **ffprobe** for encoding and probing
- **ExifTool** for metadata stripping
- **PowerShell 7** for parallel processing

`Install.bat` installs all three via winget. To install them manually:

```bat
winget install --exact --id Gyan.FFmpeg
winget install --exact --id OliverBetz.ExifTool
winget install --exact --id Microsoft.PowerShell
```

## Run manually

```powershell
pwsh -File watch-media.ps1
```

Useful switches:

| Switch | What it does |
| --- | --- |
| `-CheckOnly` | Creates folders, verifies the tools, and exits. |
| `-RecompressOutputs` | Re-runs the size cap on existing outputs that are over it. Useful after lowering a cap. |
| `-AsLibrary` | Loads the functions without starting the watcher. Used by parallel workers and tests. |

Only one watcher runs at a time, enforced by the `Global\MediaPipelineWatcher` mutex.

## Run silently at Windows startup

`Install.bat` registers a scheduled task named **Media Pipeline Watcher** that launches
`start-watcher-hidden.vbs` at logon. `Restart Watcher.bat` stops the running instance and starts
the task again.

## Tests

```powershell
# Behavior of every preset, fingerprinted structurally
pwsh -File tools\Test-PipelineParity.ps1 -Mode Capture -BuildCorpus
pwsh -File tools\Test-PipelineParity.ps1 -Mode Compare

# A real watcher process: folders, status, locking, pause, and clean shutdown
pwsh -File tools\Test-WatcherSmoke.ps1

# Chunking, reassembly, lane aggregation, archiving and config editing, offline
dotnet run --project tray-app\SelfTest

# A real round trip against the configured remote
dotnet run --project tray-app\SelfTest -- --live

# The tray app's lifecycle: autostart, close-to-tray, and quit
pwsh -File tools\Test-TrayLifecycle.ps1
```

The harness drives every preset against a throwaway sandbox root and records a structural
fingerprint of the output: counts, extensions, folder depth, pixel dimensions, and where sources
ended up. Output names come from a crypto RNG and can never be reproduced, so names are ignored.
Capture a baseline before a change, compare after. It never touches the real pipeline root.
