# media-pipeline

Local Windows media-processing watcher. It watches a folder, waits for browser downloads to finish,
then creates processed variants of each supported media file. Everything runs on your own PC.

By default it watches `D:\MediaPipeline\default\input`. You can change that folder (and many other
settings) in `config.ini` — see [Changing Settings](#changing-settings).

---

## Quick Start (recommended)

For a normal Windows PC, you do not need to touch the command line.

1. **Copy this whole folder** anywhere you like (for example `C:\media-pipeline` or your Desktop).
   Keep all the files together.
2. **Double-click `Install.bat`** and click **Yes** when Windows asks for permission.
   It will automatically:
   - install **FFmpeg**, **ExifTool**, and **PowerShell 7** (using Windows' built-in `winget`),
   - create the working folders,
   - set the watcher to start every time you sign in, and
   - start it now.
3. **Point your browser's download folder** at the input folder the installer prints at the end,
   normally:

   ```text
   D:\MediaPipeline\default\input
   ```

4. **Download or drop files** into that folder. Processed copies appear in
   `D:\MediaPipeline\default\output`, and the originals move to `D:\MediaPipeline\default\original`.

That's it. The watcher runs quietly in the background from now on.

> Requirements: Windows 10 (version 1809+) or Windows 11 with an internet connection for the
> one-time install. `winget` ships with Windows; if `Install.bat` says it's missing, open the
> Microsoft Store, install/update **App Installer**, then run `Install.bat` again.

---

## The buttons (double-click `.bat` files)

| File | What it does |
| --- | --- |
| **`Install.bat`** | One-time setup: installs tools, creates folders, enables auto-start, starts the watcher. Safe to run again to repair the setup. |
| **`Edit Config.bat`** | Opens `config.ini` (your settings) in Notepad. |
| **`Restart Watcher.bat`** | Restarts the watcher so changes you saved in `config.ini` take effect. |
| **`Uninstall.bat`** | Stops the watcher and removes auto-start. Leaves your media files, settings, and installed tools alone. |

---

## Changing Settings

All settings live in **`config.ini`**, a plain text file next to the script. To change anything:

1. Double-click **`Edit Config.bat`** (opens `config.ini` in Notepad).
2. Change a value after an `=` sign and **save**.
3. Double-click **`Restart Watcher.bat`** to apply it.

Every setting has a comment explaining what it does and its default. If `config.ini` is missing or a
value is left blank, the watcher falls back to safe built-in defaults, so you can't break it by
deleting it.

Some of the most useful settings:

| Setting | Meaning | Default |
| --- | --- | --- |
| `PipelineRoot` | Where all the watcher's folders live | `D:\MediaPipeline` |
| `DefaultPipelineAlternatingCopiesPerFile` / `DefaultPipelineMinCopiesPerFile` | How many output copies the **default** pipeline makes (these two alternate per file — set them equal for a fixed count) | `8` / `7` |
| `ImageBulkCopiesPerFile` | Variants per image in the bulk-image pipeline | `20` |
| `ImageBulkPngCompressionLevel` | PNG compression for bulk images (`1` = faster/larger, `6` = old behavior) | `1` |
| `SetCopiesPerFile` | Copies per file in the set pipeline | `10` |
| `SetBatchCount` | How many complete sets the set-batch pipeline makes | `10` |
| `AssetStoreSetCount` | How many complete sets the asset-store (manifest) pipeline makes | `15` |
| `LongCopiesPerSegment` | Variants per segment in the long-video pipeline | `3` |
| `Crf` | Video quality (lower = better quality, bigger files) | `24` |
| `Preset` | Encoder speed/size tradeoff (`fast`…`veryslow`) | `medium` |
| `MaxWidth` | Max output video width in pixels | `1080` |
| `DefaultMaxOutputSizeMB` | Size cap per default-pipeline video (`0` = off) | `8` |
| `ArchiveEnabled` / `ArchiveAgeHours` | Auto-move old outputs into `archive\` after N hours | `true` / `15` |

`config.ini` has more (audio bitrate, trim amounts, polling interval, GPU encoder tuning, etc.), each
documented in the file.

---

## Folder Structure

The watcher creates any missing folders automatically under `PipelineRoot` (default
`D:\MediaPipeline`):

```text
D:\MediaPipeline\
  default\          <- the main pipeline
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
  imageclean\
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
  assetstore\
    input\
    output\
    original\
    failed\
  archive\          <- old outputs are moved here when archiving is on
```

- `default\input`: set your browser download folder here.
- `default\output`: processed variants are written here.
- `default\original`: source files are moved here after all variants succeed.
- `default\failed`: source files are moved here if processing fails.
- `logs`: daily logs named like `media-pipeline-YYYYMMDD.log`.
- `convert\input`: put `.mov` or `.heic` files here to convert them into widely supported formats; other supported files are passed through to the output unchanged.
- `convert\output`: converted files are written here (`.mov` -> `.mp4`, `.heic` -> `.jpg`).
- `convert\original\videos`: source videos are moved here after conversion succeeds.
- `convert\original\images`: source images are moved here after conversion succeeds.
- `convert\failed`: source files are moved here if conversion fails.
- `long\input`: put long raw videos here for segmenting plus processed variants per segment.
- `long\output`: processed long-pipeline segment variants are written here.
- `long\original`: long-pipeline source files are moved here after all segment variants succeed.
- `long\failed`: long-pipeline source files are moved here if processing fails.
- `long\work`: temporary remux/intermediate workspace; the script cleans this automatically.
- `images\input`: put images here when you want many re-encoded image variants.
- `images\output`: bulk image variants are written here.
- `images\original`: source images are moved here after all image variants succeed.
- `images\failed`: source images are moved here if bulk image processing fails.
- `imageclean\input`: put images here when you want one cleaned output per source image.
- `imageclean\output`: cleaned images are written here with pure random filenames.
- `imageclean\original`: source images are moved here after cleanup succeeds.
- `imageclean\failed`: source images are moved here if cleanup fails.
- `sets\input`: put media files here when you want one output folder per source file.
- `sets\output`: each source file gets a random subfolder containing several processed copies.
- `sets\original`: source files are moved here after all copies succeed.
- `sets\failed`: source files are moved here if set processing fails.
- `setbatch\input`: drop a whole group of files here to get several complete, differentiated copies of the entire group.
- `setbatch\output`: each processed batch becomes one `bt_DD-MM-YYYY_<token>` folder containing `set_01` .. `set_NN` subfolders, each holding one processed copy of every source file.
- `setbatch\original`: source files are moved here after the whole batch succeeds.
- `setbatch\failed`: every source file in the batch is moved here if batch processing fails.
- `assetstore\input`: drop a whole group of files here to get several complete sets plus a `heatup.assetStoreMediaManifest.v1` manifest describing every generated variant.
- `assetstore\output`: each processed batch becomes one `as_DD-MM-YYYY_<token>` folder containing `set_01` .. `set_NN` subfolders and a `manifest.json` at its root.
- `assetstore\original`: source files are moved here after the whole batch (and its manifest) succeeds.
- `assetstore\failed`: every source file in the batch is moved here if batch processing fails.

> Upgrading from an older version? The default pipeline used to live directly in the root
> (`input\`, `output\`, `original\`, `failed\`). `Install.bat` automatically moves any leftover
> files from there into the new `default\` folders.

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

## How It Works

The watcher uses a polling loop:

1. Scan `default\input` every couple of seconds (`PollSeconds`).
2. Ignore unsupported files and temporary browser download files.
3. Wait until file size is stable for a few seconds (`StableSeconds`).
4. Wait until the file is no longer locked.
5. Process one file at a time.
6. Move successful originals to `default\original`.
7. Move failed originals to `default\failed`.

The default input pipeline and the video-heavy pipelines process one file at a time. The image
pipelines run conversions in parallel: the convert and image-cleanup pipelines process multiple files
at once, and the bulk image pipeline processes multiple source files at once when a batch is waiting.
For a single bulk-image source, variants are still rendered concurrently. The number of simultaneous
image conversions is controlled by `ImageProcessingConcurrency` in `config.ini` (default: up to 6,
capped by CPU count). Parallel processing requires PowerShell 7 (installed by `Install.bat`); under
Windows PowerShell 5.1 the script still runs, but image conversions fall back to sequential.

It also watches the other lanes described below.

## Output Behavior (default pipeline)

Every supported input in the default pipeline creates several output files. By default the count
alternates per file — `8` for the 1st/3rd/5th… file and `7` for the 2nd/4th/6th… — controlled by
`DefaultPipelineAlternatingCopiesPerFile` and `DefaultPipelineMinCopiesPerFile` in `config.ini`. Set
both to the same number for a fixed count.

Videos are written as MP4 files using:

- H.264 video, GPU-accelerated when possible: it picks **`h264_nvenc`** (NVIDIA, `PreferNvenc`),
  then **`h264_amf`** (AMD, `PreferAmf`), then **`libx264`** (CPU). Each GPU encoder is confirmed with
  a real test-encode at startup, so a machine that lists a GPU encoder but cannot run it falls back
  automatically instead of failing every clip.
- CRF compression, default `24` (`Crf`)
- preset `medium` (`Preset`)
- AAC audio at `128k` (`AudioBitrate`)
- 8-bit `yuv420p` pixel format for broad player compatibility
- `-movflags +faststart`
- metadata stripped with FFmpeg using `-map_metadata -1`
- metadata stripped again with ExifTool
- max width `1080px` (`MaxWidth`), preserving aspect ratio and avoiding upscaling

Each video variant trims a tiny random amount from the end. Playback speed and audio speed are not
changed. Default trim range is `15ms` to `95ms` (`MinTrimMs`/`MaxTrimMs`); very short videos use a
smaller safe range or skip trimming.

Default video outputs are capped at `8 MB` (`DefaultMaxOutputSizeMB`). When the GPU encoder
(`h264_nvenc`) is used, the first encode already applies a bitrate ceiling sized to fit the cap
(`DefaultNvencPrimaryMaxrateScale`, default `0.92` of the size-target bitrate), so most clips land
under the cap in a single pass. If a variant still exceeds the cap, the script re-encodes it from the
original source with progressively stronger compression: a quality ladder, then a reduced max width
(`DefaultSizeCapFallbackMaxWidth`, default `720px`), and finally a computed bitrate ceiling sized to
fit the target. This is the same size handling used by the long pipeline. Set
`DefaultMaxOutputSizeMB = 0` to disable it.

Images in the default pipeline are copied into the configured number of random filenames in their
original format where possible, then metadata is removed with ExifTool. `.heic` inputs are converted
to `.png` because HEIC output is not used by this pipeline.

## Bulk Image Pipeline

Use this lane when you want many image variants from one source image.

Put source images here:

```text
D:\MediaPipeline\images\input
```

The watcher writes 20 image variants here by default (`ImageBulkCopiesPerFile`):

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

Bulk PNG output uses `ImageBulkPngCompressionLevel` (default `1`) so PNG-heavy batches finish faster.
Raise it toward `6` if you prefer smaller PNG files over speed.

This makes outputs different at the file and pixel level while keeping them visually close to the
original. It is not a guarantee that files are impossible to detect or compare.

Successful source images move to `D:\MediaPipeline\images\original`. Failed source images move to
`D:\MediaPipeline\images\failed`. If a later variant fails, already completed bulk-image outputs are
left in `images\output` instead of being deleted.

## Image Cleanup Pipeline

Use this lane when you want to clean a folder of images without making duplicates. For example, 300
input images produce 300 processed output images.

Put source images here:

```text
D:\MediaPipeline\imageclean\input
```

The watcher writes one cleaned image per source image here:

```text
D:\MediaPipeline\imageclean\output
```

Each output gets the same image cleanup treatment as the bulk image pipeline:

- pure random filename like `<random>.jpg`, with no date, variant number, or source name
- metadata removed with ExifTool
- FFmpeg re-encode, not a byte-for-byte copy
- tiny randomized crop and scale back to the original dimensions when the image is large enough

Supported inputs and output format behavior match the bulk image pipeline:

- `.jpg`, `.jpeg` -> `.jpg` / `.jpeg`
- `.png` -> `.png`
- `.webp` -> `.webp`
- `.heic` -> `.png`

Successful source images move to `D:\MediaPipeline\imageclean\original`. Failed source images move to
`D:\MediaPipeline\imageclean\failed`.

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

Each folder contains 10 processed copies by default (`SetCopiesPerFile`).

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

Successful source files move to `D:\MediaPipeline\sets\original`. Failed source files move to
`D:\MediaPipeline\sets\failed`.

## Batch Sets Pipeline

Use this lane when you want several complete, differentiated copies of a whole group of files at
once. For example, 25 source images become 10 sets that each contain all 25 images, where every set
looks the same as the originals but is byte- and pixel-different from the other sets.

This is different from the Media Set Pipeline above: that lane groups output **per source file** (one
folder per image, each with N copies of that one image), while this lane groups output **per set** (N
folders, each with one copy of every image).

Drop the whole group of files here:

```text
D:\MediaPipeline\setbatch\input
```

The watcher treats everything in this folder as one batch. It waits until the batch settles — no file
is still being written, the file list is unchanged for a poll cycle, and nothing is locked — then
writes one folder per processed batch:

```text
D:\MediaPipeline\setbatch\output\bt_DD-MM-YYYY_<random>\
  set_01\
  set_02\
  ...
  set_10\
```

Each `set_NN` folder contains one processed copy of every source file. The default is 10 sets,
controlled by `SetBatchCount` in `config.ini`.

For images, each copy gets:

- random filename
- FFmpeg re-encode
- metadata removal with ExifTool
- tiny randomized crop and scale back to the original dimensions when the image is large enough, applied independently per set so no two sets are identical
- `.heic` converted to high-quality `.jpg` (`-q:v 2`)

For videos, each copy gets the same treatment as the other video lanes: H.264 MP4, AAC audio,
metadata stripped, width capped at 1080px without upscaling, and a small randomized trim from the
end — one copy per set.

Batch processing is all-or-nothing. If any file fails, the partial output folder is removed and every
source file in the batch is moved to `setbatch\failed`. On success, the source files move to
`setbatch\original`.

Because the whole batch is processed in one pass, the watcher is busy until it finishes; large
batches (many files times many sets) can take a while.

## Asset Store Manifest Pipeline

Use this lane when you want the same "several complete sets of the whole group" output as the Batch
Sets pipeline **plus** a machine-readable manifest that an importer can ingest. For example, drop 14
source videos and get 15 sets (`set_01` .. `set_15`), each containing one processed copy of every
video — 210 processed media in total — described by one `manifest.json`.

Drop the whole group of files here:

```text
D:\MediaPipeline\assetstore\input
```

The watcher treats everything in the folder as one batch (same settle logic as the Batch Sets lane:
it waits until no file is still being written, the file list is unchanged for a poll cycle, and
nothing is locked), then writes one folder per processed batch:

```text
D:\MediaPipeline\assetstore\output\as_DD-MM-YYYY_<random>\
  manifest.json
  set_01\
  set_02\
  ...
  set_15\
```

Each `set_NN` folder contains one processed copy of every source file. The number of sets is
`AssetStoreSetCount` in `config.ini` (default `15`).

Every copy is fully processed like the other lanes — videos are re-encoded to H.264 MP4 with AAC
audio, width capped at 1080px without upscaling, and metadata stripped with both FFmpeg
(`-map_metadata -1`) and ExifTool; images are re-encoded with a tiny randomized crop and metadata
removed. The one difference from the other video lanes is the **end-trim**: it is deliberately kept
to tens of milliseconds at most (`AssetStoreMinTrimMs` / `AssetStoreMaxTrimMs`, default `10`–`40` ms)
so each rendition differs at the file level without noticeably changing its length.

### The manifest

`manifest.json` is written at the root of each `as_...` batch folder using the
`heatup.assetStoreMediaManifest.v1` schema:

```json
{
  "schema": "heatup.assetStoreMediaManifest.v1",
  "generatedAt": "2026-06-04T12:00:00.000Z",
  "importRoot": ".",
  "variants": [
    {
      "familyKey": "video_001",
      "variantKey": "video_001__set_01",
      "path": "set_01/dt_v01_04-06-2026_<random>.mp4",
      "renditionSetKey": "set_01",
      "generationBatchKey": "as_04-06-2026_<random>",
      "sourceOriginalName": "video_001.mp4",
      "sourceFamilyName": "video_001",
      "durationSeconds": 1.978,
      "sizeBytes": 47130,
      "transformProfile": "asset_store_video_micro_trim",
      "generatedAt": "2026-06-04T11:58:00.000Z",
      "metadata": { "encoder": "h264_nvenc", "trimMs": 22, "maxWidth": 1080 }
    }
  ]
}
```

How the fields map to the output:

- `familyKey` — one per **source original** (the file's base name, sanitized). 14 source videos give
  14 family keys. Duplicate base names within a batch get a numeric suffix so keys stay unique.
- `variantKey` — `"<familyKey>__set_NN"`, unique within the family and across the batch (one per
  generated copy).
- `path` — relative to `importRoot`, e.g. `set_01/<file>.mp4`.
- `renditionSetKey` — which set the copy belongs to (`set_01` .. `set_NN`). This is the "15 complete
  sets" grouping; group the variants by this key to get one full set of every source.
- `generationBatchKey` — the `as_...` batch folder name.
- `durationSeconds` (videos only) and `sizeBytes` describe the produced file; `metadata` records the
  encoder and the exact trim applied.

Path resolution for a consumer: an absolute `path` is used as-is; a relative `path` resolves from
`importRoot` when present, otherwise relative to the manifest file's own folder. Asset-store
manifests always set `importRoot` to `"."`, so the relative paths resolve next to the manifest.

Processing is all-or-nothing: if any file fails, the partial batch folder is removed and every source
file is moved to `assetstore\failed`. On success (including writing the manifest) the source files
move to `assetstore\original`.

## Convert Workflow

Use this lane to convert media from formats you cannot use into widely supported ones. It handles
both videos and images. Put the source file here:

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

Any other supported media file dropped here (for example `.jpg`, `.png`, `.webp`, or `.mp4`) is passed
through to the output folder unchanged. This makes it safe to drop a mixed batch: files that need
conversion are converted, and files that are already in a good format are simply moved to the output.

This does not create multiple variants. Videos are not re-encoded: `.mov` uses FFmpeg stream copy for
video and audio:

```text
-map 0:v:0 -map 0:a? -dn -c copy -map_metadata -1 -movflags +faststart
```

Because the video and audio streams are copied, quality and timing should remain unchanged. Non-media
data tracks from phones, such as sensor or metadata streams, are dropped because MP4 often cannot
contain them.

Images are decoded and re-encoded to high-quality JPEG with `-q:v 2`, and metadata is stripped.
`.heic` is the common iPhone photo format; JPEG is chosen for broad compatibility at high quality.

Converted source files are moved into `convert\original\videos` or `convert\original\images`.
Passed-through files go straight to `convert\output` (they are unchanged, so no separate original copy
is kept). Failed source files are moved to `convert\failed`.

## Long Video Pipeline

Use this lane for raw videos around a minute or longer when you want the script to split them into
shorter clips before applying the normal multi-copy processing.

Put source videos here:

```text
D:\MediaPipeline\long\input
```

The watcher writes processed segment variants here:

```text
D:\MediaPipeline\long\output
```

Supported inputs are the same video extensions as the main pipeline. If the input is `.mov`, the
script first remuxes it to a temporary MP4 using stream copy, then segments the MP4.

Segmenting defaults (configurable in `config.ini`):

- Target segment length: `15` seconds (`LongSegmentTargetSeconds`)
- Minimum segment length: `11` seconds (`LongSegmentMinSeconds`)
- Each segment creates `3` final variants (`LongCopiesPerSegment`)

The script prefers 15-second segments, but it will not leave a tiny final tail. If the final
remainder is too short, it borrows time from previous segments. For example, about 38 seconds becomes
roughly:

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

Successful source files move to `D:\MediaPipeline\long\original`. Failed source files move to
`D:\MediaPipeline\long\failed`.

## Output File Names

Output names are random and not based on the source filename. Most pipelines put a code first
(shortened: `dt` default/media, `lg` long, `img` bulk images, `st` sets, `bt` set batch, `as` asset
store, `rx` remux, `cv` convert), then the creation date as `DD-MM-YYYY` (day and month separated by
`-`) so the date stays readable when cloud tools truncate the end of long paths. The image cleanup
pipeline is the exception: its output filenames are just random tokens plus the extension.

```text
dt_DD-MM-YYYY_<random>.mp4
dt_DD-MM-YYYY_<random>.jpg
rx_DD-MM-YYYY_<random>.mp4
cv_DD-MM-YYYY_<random>.jpg
lg_DD-MM-YYYY_s01_v01_<random>.mp4
img_v01_DD-MM-YYYY_<random>.jpg
<random>.jpg
dt_v01_DD-MM-YYYY_<random>.mp4
st_DD-MM-YYYY_<random>\
bt_DD-MM-YYYY_<random>\
```

---

# Advanced / Manual Setup

You only need this section if you are not using `Install.bat`, or you want to understand what it does.

## Required Tools

The watcher requires these commands. `Install.bat` installs all three with `winget`
(`Gyan.FFmpeg`, `OliverBetz.ExifTool`, `Microsoft.PowerShell`):

```powershell
ffmpeg
ffprobe
exiftool
```

The watcher is launched with **PowerShell 7** (`pwsh`), which enables parallel image processing. The
launchers look for `pwsh` at `C:\Program Files\PowerShell\7\pwsh.exe`, then `C:\Tools\pwsh\pwsh.exe`,
then fall back to Windows PowerShell 5.1.

Verify from a new PowerShell window:

```powershell
ffmpeg -version
ffprobe -version
exiftool -ver
```

### Manual install with winget

```powershell
winget install -e --id Gyan.FFmpeg
winget install -e --id OliverBetz.ExifTool
winget install -e --id Microsoft.PowerShell
```

### Manual install without winget

1. **FFmpeg** — download a Windows build, extract to e.g. `C:\Tools\ffmpeg`, and add
   `C:\Tools\ffmpeg\bin` to your `PATH`.
2. **ExifTool** — download the Windows executable, rename `exiftool(-k).exe` to `exiftool.exe`, put it
   in e.g. `C:\Tools\exiftool`, and add that folder to your `PATH`.
3. **PowerShell 7** — download `PowerShell-<version>-win-x64.zip`, extract to `C:\Tools\pwsh`, and
   verify with `C:\Tools\pwsh\pwsh.exe -version`.

The script also checks `C:\Tools\...` and the `winget` link folder as fallbacks, so tools resolve even
if `PATH` hasn't refreshed yet.

## Run manually

```powershell
& "C:\Program Files\PowerShell\7\pwsh.exe" -NoProfile -ExecutionPolicy Bypass -File "<app folder>\watch-media.ps1"
```

Or start it hidden:

```powershell
"<app folder>\start-watcher.bat"
```

For a fully silent start with no terminal window:

```powershell
wscript.exe "<app folder>\start-watcher-hidden.vbs"
```

Validate the setup without processing anything:

```powershell
& "C:\Program Files\PowerShell\7\pwsh.exe" -NoProfile -ExecutionPolicy Bypass -File "<app folder>\watch-media.ps1" -CheckOnly
```

## Run silently at Windows startup (manual)

`Install.bat` does this for you. To do it by hand, register a logon task that runs the hidden VBS
launcher (replace `<app folder>` with this folder's path):

```powershell
$action = New-ScheduledTaskAction -Execute "wscript.exe" -Argument '"<app folder>\start-watcher-hidden.vbs"'
$trigger = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit (New-TimeSpan -Seconds 0)
Register-ScheduledTask -TaskName "Media Pipeline Watcher" -Action $action -Trigger $trigger -Settings $settings -Description "Runs the local media pipeline watcher silently at logon."
```

To remove it later (or just use `Uninstall.bat`):

```powershell
Unregister-ScheduledTask -TaskName "Media Pipeline Watcher" -Confirm:$false
```

Only one watcher runs at a time; a system mutex (`Global\MediaPipelineWatcher`) makes duplicate
launches exit immediately.

## Configuration internals

`config.ini` is read at startup by `watch-media.ps1`. Each key maps to a script variable of the same
name; missing or unparseable keys fall back to the built-in default in the script. Changing settings
therefore only requires editing `config.ini` and restarting the watcher (`Restart Watcher.bat`) — no
code editing.
