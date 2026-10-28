# Media Pipeline for macOS

The macOS client combines the selected native sidebar window with a compact menu bar panel. It
uses SwiftUI for the interface and the self-contained C# worker for queue processing, media
conversion, archiving, and uploads. PowerShell is not part of the Mac runtime.

## What is included

- Native Activity, Saved Media, Uploads, Presets, and Settings views
- Menu bar status, live progress, pause, restart, logs, and quick access to the main window
- Apple VideoToolbox detection with libx264 fallback
- Failure notifications through macOS Notification Center
- Start at Login through `SMAppService`
- Pause and resume globally or per lane, clean stop, restart, and failed-file requeue
- Job zip creation and Zip and Upload
- Resumable chunked uploads with progress and cancellation
- An editable INI configuration that preserves preset inheritance
- A self-contained Apple silicon worker inside the app bundle

The app supports macOS 13 or newer on Apple silicon. The source is a Swift package, so it can be
opened directly in Xcode by opening `macos-app/Package.swift`.

## Saved media

Open **Saved Media** and choose **Save Media…** to keep images or videos in the app's library.
Select an item, enter a whole number from 1 to 9999, choose a workspace and preset, and click
**Make Copies**. The worker starts if needed. Activity shows progress and provides the usual
archive, upload, and retry actions.

Click **Open Output Folder** beside the output path to open the selected workspace and preset's
`output` folder in Finder, for example `general/video-clean/output`. The app creates the folder
if it does not exist yet.

Each run uses the selected preset's saved image or video transformations, normalization,
metadata removal, output grouping, and manifest settings. The entered amount is the exact
number of outputs. Alternating counts are disabled for that run, and set presets use one set.
The preset configuration stays unchanged. Save pending preset changes before making copies.

The app copies saved originals into
`~/Library/Application Support/Media Pipeline/saved-media`. You can move or delete the file
you imported and keep using its saved copy. **Remove** moves the library entry to Trash.
Already queued runs keep their own source copies and continue independently.

Requests wait in `<workspace>/<preset>/requests` under the pipeline root. Pausing or stopping
keeps unfinished requests for the next run. Failed requests retain their source and copy count
in `requests-failed`; Retry in Activity queues them again with that amount.

## Build the app on the Mac

Install Xcode's command-line tools and the .NET 8 SDK first. Then, from the repository root:

```zsh
zsh tools/Install-MacDependencies.sh
zsh tools/Build-MacApp.sh
zsh tools/Test-MacPackage.sh
```

The build creates:

```text
artifacts/Media Pipeline.app
artifacts/Media-Pipeline-macos-arm64.zip
```

The default build uses an ad-hoc signature for local use. To sign it with a Developer ID, provide
the identity below. Public distribution would also require Apple's notarization step.

```zsh
SIGN_IDENTITY="Developer ID Application: Your Name (TEAMID)" zsh tools/Build-MacApp.sh
```

Install and open the local build with:

```zsh
ditto "artifacts/Media Pipeline.app" "/Applications/Media Pipeline.app"
open "/Applications/Media Pipeline.app"
```

Start at Login works after the app is installed in `/Applications`. macOS asks for notification
permission the first time the app starts.

## Files and folders

The app keeps its editable configuration here:

```text
~/Library/Application Support/Media Pipeline/config.ini
```

Media defaults to `~/MediaPipeline`. Each workspace contains its presets:

```text
~/MediaPipeline/LC/bulk/input
~/MediaPipeline/LC/bulk/output
~/MediaPipeline/LC/bulk/original
~/MediaPipeline/LC/bulk/failed
~/MediaPipeline/LC/sync
```

The worker creates the complete folder structure on first launch. To move existing media from
Windows, copy the contents of the Windows pipeline root into `~/MediaPipeline` while preserving
this workspace-first layout. Copy any custom preset values into the Mac app's Settings and
Presets views, or replace its `config.ini` and change `PipelineRoot` to `~/MediaPipeline`.

## External tools

`Install-MacDependencies.sh` installs FFmpeg, FFprobe, ExifTool, and rclone with Homebrew. SSH is
provided by macOS. The worker looks in its bundle, `PATH`, `/opt/homebrew/bin`, and
`/usr/local/bin`, so it also works when Finder launches the app with a minimal shell environment.

Uploads still target the configured remote Windows host. The SFTP staging path and the Windows
assembly path are separate settings because the two sides use different path syntax. A path such
as `~/.ssh/key-name` is expanded to the current Mac user's home folder by the worker.

When `DeleteAfterUpload` is enabled, the worker first verifies the assembled remote file and
releases a stable local source path to the app. The app then uses Finder's native Trash operation
instead of permanently deleting the source. If Trash cannot accept the file, the upload remains
complete, the app restores the source to its original name or a noncolliding `upload-retry` name,
and the Uploads view reports the cleanup error and retained path.

On a new Mac, copy the existing rclone configuration and SSH key before uploading:

```zsh
mkdir -p ~/.config/rclone ~/.ssh
cp /path/from/old-machine/rclone.conf ~/.config/rclone/rclone.conf
cp /path/from/old-machine/your-private-key ~/.ssh/your-private-key
chmod 600 ~/.ssh/your-private-key
rclone listremotes
ssh -i ~/.ssh/your-private-key -p 2222 heatup-remote exit
```

The default configuration expects an rclone remote named `heatup-remote`. Change `RemoteName`,
`RemoteSshHost`, `RemoteSshPort`, and `RemoteSshKeyFile` in Settings if your names differ. The SSH
probe must use the same host and port as those settings. Private keys and rclone credentials are
never copied into the app bundle.

## Verification boundary

The C# core and integration suites, Windows build, native parity harness, and cross-runtime publish
are runnable from Windows. SwiftUI and the final `.app` bundle require Apple's SDK and are verified
by `swift test`, `plutil`, `codesign`, and a bundled-worker check inside `Build-MacApp.sh` and
`Test-MacPackage.sh` on the Mac.

With Command Line Tools alone, `Build-MacApp.sh` runs the standalone saved-media self-test
instead of the XCTest suite. Full Mac unit tests require Xcode. The standalone check covers
library persistence, independent originals, invalid imports, corrupt entries, and queue arguments.
If the selected SDK is missing the `SwiftUIMacros` plugin, choose an installed compatible SDK
for the build, for example:

```zsh
SDKROOT=/Library/Developer/CommandLineTools/SDKs/MacOSX26.5.sdk zsh tools/Build-MacApp.sh
```
