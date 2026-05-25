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
```

- `input`: set your browser download folder here.
- `output`: processed variants are written here.
- `original`: source files are moved here after all 3 variants succeed.
- `failed`: source files are moved here if processing fails.
- `logs`: daily logs named like `media-pipeline-YYYYMMDD.log`.
- `convert\input`: put raw `.mov` files here when you only want MP4 remuxing.
- `convert\output`: remuxed `.mp4` files are written here.
- `convert\original`: source `.mov` files are moved here after remuxing succeeds.
- `convert\failed`: source `.mov` files are moved here if remuxing fails.
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

## Usage

Run manually:

```powershell
& "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File "D:\Projects\media-pipeline\watch-media.ps1"
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

It also scans `D:\MediaPipeline\convert\input` for `.mov` files and remuxes them to `.mp4` without re-encoding.

It also scans `D:\MediaPipeline\long\input` for longer videos, segments them, and creates 3 processed variants for each segment.

It also scans `D:\MediaPipeline\images\input` for image-only bulk processing and creates 20 re-encoded variants per source image.

It also scans `D:\MediaPipeline\sets\input` for media files and creates one output folder with 10 processed copies per source file.

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

## MOV To MP4 Remux Workflow

Use this lane for long raw `.mov` clips that you need to convert to `.mp4` before manually cutting them. Put the source file here:

```text
D:\MediaPipeline\convert\input
```

The watcher writes one MP4 here:

```text
D:\MediaPipeline\convert\output
```

This does not create 3 variants and does not re-encode the video. It uses FFmpeg stream copy for video and audio:

```text
-map 0:v:0 -map 0:a? -dn -c copy -map_metadata -1 -movflags +faststart
```

Because the video and audio streams are copied, quality and timing should remain unchanged. Non-media data tracks from phones, such as sensor or metadata streams, are dropped because MP4 often cannot contain them.

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
