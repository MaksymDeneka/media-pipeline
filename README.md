# media-pipeline

Local Windows media-processing watcher. It watches `D:\MediaPipeline\input`, waits for browser downloads to finish, then creates processed variants per supported media file.

## Folder Structure

The script creates missing folders automatically:

```text
D:\MediaPipeline\
  input\
  output\
  original\
  failed\
  logs\
  convert\
    input\
    output\
    original\
      videos\
      images\
    failed\
  long\
    input\
    output\
    original\
    failed\
    work\
  images\
    input\
    output\
    original\
    failed\
  sets\
    input\
    output\
    original\
    failed\
  setbatch\
    input\
    output\
    original\
    failed\
```

- `input`: set your browser download folder here.
- `output`: processed variants are written here.
- `original`: source files are moved here after all 3 variants succeed.
- `failed`: source files are moved here if processing fails.
- `logs`: daily logs named like `media-pipeline-YYYYMMDD.log`.
- `convert\input`: put `.mov` or `.heic` files here to convert them into widely supported formats; other supported files are passed through to the output unchanged.
- `convert\output`: converted files are written here (`.mov` -> `.mp4`, `.heic` -> `.jpg`).
- `convert\original\videos`: source videos are moved here after conversion succeeds.
- `convert\original\images`: source images are moved here after conversion succeeds.
- `convert\failed`: source files are moved here if conversion fails.
- `long\input`: put long raw videos here for segmenting plus 3 processed variants per segment.
- `long\output`: processed long-pipeline segment variants are written here.
- `long\original`: long-pipeline source files are moved here after all segment variants succeed.
- `long\failed`: long-pipeline source files are moved here if processing fails.
- `long\work`: temporary remux/intermediate workspace; the script cleans this automatically.
- `images\input`: put images here when you want many re-encoded image variants.
- `images\output`: bulk image variants are written here.
- `images\original`: source images are moved here after all image variants succeed.
- `images\failed`: source images are moved here if bulk image processing fails.
- `sets\input`: put media files here when you want one output folder per source file.
- `sets\output`: each source file gets a random subfolder containing 10 processed copies.
- `sets\original`: source files are moved here after all 10 copies succeed.
- `sets\failed`: source files are moved here if set processing fails.
- `setbatch\input`: drop a whole group of files here to get several complete, differentiated copies of the entire group.
- `setbatch\output`: each processed batch becomes one `batch_<timestamp>_<token>` folder containing `set_01` .. `set_NN` subfolders, each holding one processed copy of every source file.
- `setbatch\original`: source files are moved here after the whole batch succeeds.
- `setbatch\failed`: every source file in the batch is moved here if batch processing fails.

## Supported Files

Videos:

```text
.mp4, .mov, .mkv, .webm, .avi
```

Images:

```text
.jpg, .jpeg, .png, .webp, .heic
```

Temporary browser download files are ignored:

```text
.crdownload, .tmp, .part, .download
```

## Required Tools

The watcher requires these commands to be available in `PATH`:

```powershell
ffmpeg
ffprobe
exiftool
```

The watcher is launched with **PowerShell 7** (`pwsh`), which enables parallel image processing. A portable build extracted to `C:\Tools\pwsh` is enough — the launcher calls `C:\Tools\pwsh\pwsh.exe` by full path, so it does not need to be on `PATH`. Under Windows PowerShell 5.1 the script still runs, but image conversions fall back to sequential.

Verify from a new PowerShell window:

```powershell
ffmpeg -version
ffprobe -version
exiftool -ver
```

### Install FFmpeg

1. Download a Windows FFmpeg build from the official FFmpeg site or a trusted Windows build provider.
2. Extract it to a stable folder, for example `C:\Tools\ffmpeg`.
3. Add the `bin` folder to your user or system `PATH`, for example:

```text
C:\Tools\ffmpeg\bin
```

Open a new PowerShell window and run:

```powershell
ffmpeg -version
ffprobe -version
```

### Install ExifTool

1. Download the Windows ExifTool executable.
2. Rename `exiftool(-k).exe` to `exiftool.exe`.
3. Put it in a stable folder, for example `C:\Tools\exiftool`.
4. Add that folder to your user or system `PATH`.

Open a new PowerShell window and run:

```powershell
exiftool -ver
```

### Install PowerShell 7

1. Download the latest stable `PowerShell-<version>-win-x64.zip` from the official PowerShell releases.
2. Extract it to `C:\Tools\pwsh`.
3. Verify:

```powershell
C:\Tools\pwsh\pwsh.exe -version
```

## Usage

Run manually:

```powershell
& "C:\Tools\pwsh\pwsh.exe" -NoProfile -ExecutionPolicy Bypass -File "D:\Projects\media-pipeline\watch-media.ps1"
```

Or start it hidden:

```powershell
D:\Projects\media-pipeline\start-watcher.bat
```

For a fully silent start with no terminal window, use:

```powershell
wscript.exe "D:\Projects\media-pipeline\start-watcher-hidden.vbs"
```

The watcher uses a polling loop:

1. Scan `D:\MediaPipeline\input` every 2 seconds.
2. Ignore unsupported files and temporary browser download files.
3. Wait until file size is stable for several seconds.
4. Wait until the file is no longer locked.
5. Process one file at a time.
6. Move successful originals to `D:\MediaPipeline\original`.
7. Move failed originals to `D:\MediaPipeline\failed`.

The default input pipeline and the video-heavy pipelines process one file at a time. The image pipelines run conversions in parallel: the convert pipeline processes multiple files at once, and the bulk image pipeline renders its per-file variants concurrently. The number of simultaneous image conversions is controlled by `$ImageProcessingConcurrency` in `watch-media.ps1` (default: up to 6, capped by CPU count). Parallel processing requires PowerShell 7.

It also scans `D:\MediaPipeline\convert\input`, converting `.mov` -> `.mp4` (stream copy) and `.heic` -> `.jpg`; any other supported media file is passed through to the output unchanged.

It also scans `D:\MediaPipeline\long\input` for longer videos, segments them, and creates 3 processed variants for each segment.

It also scans `D:\MediaPipeline\images\input` for image-only bulk processing and creates 20 re-encoded variants per source image.

It also scans `D:\MediaPipeline\sets\input` for media files and creates one output folder with 10 processed copies per source file.

It also scans `D:\MediaPipeline\setbatch\input` and, once the batch settles, turns the whole group of files into several complete sets — one folder per set, each containing a processed copy of every source file (11 sets by default).

## Browser Download Folder

Set your browser download folder to:

```text
D:\MediaPipeline\input
```

Download files from Google Drive manually in the browser. The watcher will detect completed downloads in that folder.

## Output Behavior

Every supported input in the default pipeline creates exactly 5 output files.

Videos are written as MP4 files using:

- H.264 video with `libx264`
- CRF compression, default `24`
- preset `medium`
- AAC audio at `128k`
- 8-bit `yuv420p` pixel format for broad player compatibility
- `-movflags +faststart`
- metadata stripped with FFmpeg using `-map_metadata -1`
- metadata stripped again with ExifTool
- max width `1080px`, preserving aspect ratio and avoiding upscaling

Each video variant trims a tiny random amount from the end. Playback speed and audio speed are not changed. Default trim range is `15ms` to `95ms`; very short videos use a smaller safe range or skip trimming.

Images in the default pipeline are copied into 5 random filenames in their original format where possible, then metadata is removed with ExifTool. `.heic` inputs are converted to `.png` because HEIC output is not used by this pipeline.

## Bulk Image Pipeline

Use this lane when you want many image variants from one source image.

Put source images here:

```text
D:\MediaPipeline\images\input
```

The watcher writes 20 image variants here by default:

```text
D:\MediaPipeline\images\output
```

Each output gets:

- random filename
- metadata removed with ExifTool
- FFmpeg re-encode, not a byte-for-byte copy
- tiny randomized crop and scale back to the original dimensions when the image is large enough

Supported inputs are:

```text
.jpg, .jpeg, .png, .webp, .heic
```

Output format behavior:

- `.jpg`, `.jpeg` -> `.jpg` / `.jpeg`
- `.png` -> `.png`
- `.webp` -> `.webp`
- `.heic` -> `.png`

This makes outputs different at the file and pixel level while keeping them visually close to the original. It is not a guarantee that files are impossible to detect or compare.

Successful source images move to `D:\MediaPipeline\images\original`. Failed source images move to `D:\MediaPipeline\images\failed`.

## Media Set Pipeline

Use this lane when you want every source media file to get its own output folder.

Put source media files here:

```text
D:\MediaPipeline\sets\input
```

For each input file, the watcher creates a random folder here:

```text
D:\MediaPipeline\sets\output
```

Each folder contains 10 processed copies by default.

For videos, each copy gets:

- random filename
- H.264 MP4 encode with `libx264`
- CRF compression
- AAC audio
- 8-bit `yuv420p` pixel format
- metadata removal with FFmpeg and ExifTool
- width capped at 1080px without upscaling
- tiny randomized trim from the end

For images, each copy gets:

- random filename
- FFmpeg re-encode
- metadata removal with ExifTool
- tiny randomized crop and scale back to original dimensions when the image is large enough
- `.heic` output converted to `.png`

Successful source files move to `D:\MediaPipeline\sets\original`. Failed source files move to `D:\MediaPipeline\sets\failed`.

## Batch Sets Pipeline

Use this lane when you want several complete, differentiated copies of a whole group of files at once. For example, 25 source images become 11 sets that each contain all 25 images, where every set looks the same as the originals but is byte- and pixel-different from the other sets.

This is different from the Media Set Pipeline above: that lane groups output **per source file** (one folder per image, each with N copies of that one image), while this lane groups output **per set** (N folders, each with one copy of every image).

Drop the whole group of files here:

```text
D:\MediaPipeline\setbatch\input
```

The watcher treats everything in this folder as one batch. It waits until the batch settles — no file is still being written, the file list is unchanged for a poll cycle, and nothing is locked — then writes one folder per processed batch:

```text
D:\MediaPipeline\setbatch\output\batch_YYYYMMDD_HHMMSS_<random>\
  set_01\
  set_02\
  ...
  set_11\
```

Each `set_NN` folder contains one processed copy of every source file. The default is 11 sets, controlled by `$SetBatchCount` in `watch-media.ps1`.

For images, each copy gets:

- random filename
- FFmpeg re-encode
- metadata removal with ExifTool
- tiny randomized crop and scale back to the original dimensions when the image is large enough, applied independently per set so no two sets are identical
- `.heic` converted to high-quality `.jpg` (`-q:v 2`)

For videos, each copy gets the same treatment as the other video lanes: H.264 MP4, AAC audio, metadata stripped, width capped at 1080px without upscaling, and a small randomized trim from the end — one copy per set.

Batch processing is all-or-nothing. If any file fails, the partial output folder is removed and every source file in the batch is moved to `setbatch\failed`. On success, the source files move to `setbatch\original`.

Because the whole batch is processed in one pass, the watcher is busy until it finishes; large batches (many files times many sets) can take a while.

## Convert Workflow

Use this lane to convert media from formats you cannot use into widely supported ones. It handles both videos and images. Put the source file here:

```text
D:\MediaPipeline\convert\input
```

The watcher writes one converted file here:

```text
D:\MediaPipeline\convert\output
```

Source formats and their targets:

- `.mov` -> `.mp4`
- `.heic` -> `.jpg`

Any other supported media file dropped here (for example `.jpg`, `.png`, `.webp`, or `.mp4`) is passed through to the output folder unchanged. This makes it safe to drop a mixed batch: files that need conversion are converted, and files that are already in a good format are simply moved to the output.

This does not create multiple variants. Videos are not re-encoded: `.mov` uses FFmpeg stream copy for video and audio:

```text
-map 0:v:0 -map 0:a? -dn -c copy -map_metadata -1 -movflags +faststart
```

Because the video and audio streams are copied, quality and timing should remain unchanged. Non-media data tracks from phones, such as sensor or metadata streams, are dropped because MP4 often cannot contain them.

Images are decoded and re-encoded to high-quality JPEG with `-q:v 2`, and metadata is stripped. `.heic` is the common iPhone photo format; JPEG is chosen for broad compatibility at high quality.

Converted source files are moved into `convert\original\videos` or `convert\original\images`. Passed-through files go straight to `convert\output` (they are unchanged, so no separate original copy is kept). Failed source files are moved to `convert\failed`.

## Long Video Pipeline

Use this lane for raw videos around a minute or longer when you want the script to split them into shorter clips before applying the normal 3-copy processing.

Put source videos here:

```text
D:\MediaPipeline\long\input
```

The watcher writes processed segment variants here:

```text
D:\MediaPipeline\long\output
```

Supported inputs are the same video extensions as the main pipeline. If the input is `.mov`, the script first remuxes it to a temporary MP4 using stream copy, then segments the MP4.

Segmenting defaults:

- Target segment length: `15` seconds
- Minimum segment length: `11` seconds
- Each segment creates `3` final variants

The script prefers 15-second segments, but it will not leave a tiny final tail. If the final remainder is too short, it borrows time from previous segments. For example, about 38 seconds becomes roughly:

```text
15s, 12s, 11s
```

Each segment variant then goes through the same final processing as normal videos:

- H.264 MP4 encode with `libx264`
- CRF compression
- AAC audio
- 8-bit `yuv420p` pixel format for broad player compatibility
- metadata removal with FFmpeg and ExifTool
- width capped at 1080px without upscaling
- small randomized trim from the end
- random neutral output filename

Successful source files move to `D:\MediaPipeline\long\original`. Failed source files move to `D:\MediaPipeline\long\failed`.

Output names are random and not based on the source filename:

```text
media_YYYYMMDD_HHMMSS_<random>.mp4
media_YYYYMMDD_HHMMSS_<random>.jpg
remux_YYYYMMDD_HHMMSS_<random>.mp4
long_s01_v01_YYYYMMDD_HHMMSS_<random>.mp4
image_v01_YYYYMMDD_HHMMSS_<random>.jpg
media_v01_YYYYMMDD_HHMMSS_<random>.mp4
```

## Always Run Silently At Windows Startup

Use Windows Task Scheduler with the VBS launcher for the silent setup. This avoids a terminal window and starts the watcher when you sign in.

Run this once from PowerShell:

```powershell
$action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument '"D:\Projects\media-pipeline\start-watcher-hidden.vbs"'
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit (New-TimeSpan -Seconds 0)
Register-ScheduledTask -TaskName "Media Pipeline Watcher" -Action $action -Trigger $trigger -Settings $settings -Description "Runs the local media pipeline watcher silently at logon."
```

If the task already exists, replace its launch command:

```powershell
$action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument '"D:\Projects\media-pipeline\start-watcher-hidden.vbs"'
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit (New-TimeSpan -Seconds 0)
Set-ScheduledTask -TaskName "Media Pipeline Watcher" -Action $action -Trigger $trigger -Settings $settings
Start-ScheduledTask -TaskName "Media Pipeline Watcher"
```

To remove it later:

```powershell
Unregister-ScheduledTask -TaskName "Media Pipeline Watcher" -Confirm:$false
```

Alternative simple startup-folder method:

1. Press `Win + R`.
2. Run `shell:startup`.
3. Put a shortcut to `D:\Projects\media-pipeline\start-watcher.bat` in that folder.

Task Scheduler plus `start-watcher-hidden.vbs` is preferred because `wscript.exe` can launch PowerShell with no visible console window.

## Configuration

Edit the settings at the top of `watch-media.ps1`:

```powershell
$PipelineRoot = "D:\MediaPipeline"
$RemuxInputDir = Join-Path $RemuxRootDir "input"
$RemuxOutputDir = Join-Path $RemuxRootDir "output"
$LongInputDir = Join-Path $LongRootDir "input"
$LongOutputDir = Join-Path $LongRootDir "output"
$DefaultCopiesPerFile = 5
$LongCopiesPerSegment = 3
$ImageBulkCopiesPerFile = 20
$SetCopiesPerFile = 10
$SetBatchCount = 11
$ImageBulkCropMinPermille = 5
$ImageBulkCropMaxPermille = 20
$MinTrimMs = 15
$MaxTrimMs = 95
$Crf = 24
$Preset = "medium"
$AudioBitrate = "128k"
$MaxWidth = 1080
$StableSeconds = 3
$TimeoutSeconds = 600
$PollSeconds = 2
$LongSegmentTargetSeconds = 15
$LongSegmentMinSeconds = 11
```
