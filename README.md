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
```

- `input`: set your browser download folder here.
- `output`: processed variants are written here.
- `original`: source files are moved here after all 3 variants succeed.
- `failed`: source files are moved here if processing fails.
- `logs`: daily logs named like `media-pipeline-YYYYMMDD.log`.

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

## Browser Download Folder

Set your browser download folder to:

```text
D:\MediaPipeline\input
```

Download files from Google Drive manually in the browser. The watcher will detect completed downloads in that folder.

## Output Behavior

Every supported input creates exactly 3 output files.

Videos are written as MP4 files using:

- H.264 video with `libx264`
- CRF compression, default `24`
- preset `medium`
- AAC audio at `128k`
- `-movflags +faststart`
- metadata stripped with FFmpeg using `-map_metadata -1`
- metadata stripped again with ExifTool
- max width `1080px`, preserving aspect ratio and avoiding upscaling

Each video variant trims a small random amount from the end. Playback speed and audio speed are not changed. Default trim range is `50ms` to `950ms`; very short videos use a smaller safe range or skip trimming.

Images are copied into 3 random filenames in their original format where possible, then metadata is removed with ExifTool.

Output names are random and not based on the source filename:

```text
media_YYYYMMDD_HHMMSS_<random>.mp4
media_YYYYMMDD_HHMMSS_<random>.jpg
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
```
