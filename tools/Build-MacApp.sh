#!/bin/zsh
set -euo pipefail

script_dir=${0:A:h}
repo_root=${script_dir:h}
package_dir="$repo_root/macos-app"
artifact_root="$repo_root/artifacts"
worker_output="$artifact_root/native-worker/osx-arm64"
app="$artifact_root/Media Pipeline.app"
worker_entitlements="$package_dir/Worker.entitlements"
configuration=${CONFIGURATION:-release}
sign_identity=${SIGN_IDENTITY:--}

if [[ "$(uname -s)" != "Darwin" ]]; then
    echo "Build-MacApp.sh must run on macOS."
    exit 1
fi

if ! command -v dotnet >/dev/null 2>&1; then
    for candidate in /usr/local/share/dotnet/dotnet /opt/homebrew/bin/dotnet; do
        if [[ -x "$candidate" ]]; then
            export PATH="${candidate:h}:$PATH"
            break
        fi
    done
fi

for tool in dotnet swift codesign ditto iconutil; do
    if ! command -v "$tool" >/dev/null 2>&1; then
        echo "Missing build tool: $tool"
        exit 1
    fi
done

dotnet publish "$repo_root/src/MediaPipeline.Worker/MediaPipeline.Worker.csproj" \
    -c Release \
    -r osx-arm64 \
    --self-contained true \
    -p:PublishSingleFile=true \
    -p:PublishTrimmed=false \
    -p:DebugType=None \
    -p:DebugSymbols=false \
    -o "$worker_output" \
    --nologo

swift test --package-path "$package_dir"
swift build --package-path "$package_dir" -c "$configuration" --arch arm64

binary_path=$(swift build --package-path "$package_dir" -c "$configuration" --arch arm64 --show-bin-path)
rm -rf "$app"
mkdir -p "$app/Contents/MacOS" "$app/Contents/Helpers" "$app/Contents/Resources"
cp "$binary_path/MediaPipelineMac" "$app/Contents/MacOS/MediaPipelineMac"
cp "$worker_output/media-pipeline-worker" "$app/Contents/Helpers/media-pipeline-worker"
cp "$package_dir/Resources/Info.plist" "$app/Contents/Info.plist"
cp "$package_dir/Sources/MediaPipelineMac/Resources/default-config.ini" \
    "$app/Contents/Resources/default-config.ini"

iconset="$artifact_root/AppIcon.iconset"
rm -rf "$iconset"
swift "$repo_root/tools/Generate-MacAppIcon.swift" "$iconset"
iconutil -c icns "$iconset" -o "$app/Contents/Resources/AppIcon.icns"
rm -rf "$iconset"

chmod 755 "$app/Contents/MacOS/MediaPipelineMac" "$app/Contents/Helpers/media-pipeline-worker"
codesign --force --options runtime --entitlements "$worker_entitlements" \
    --sign "$sign_identity" "$app/Contents/Helpers/media-pipeline-worker"
codesign --force --options runtime --sign "$sign_identity" "$app/Contents/MacOS/MediaPipelineMac"
codesign --force --options runtime --sign "$sign_identity" "$app"
codesign --verify --deep --strict --verbose=2 "$app"

archive="$artifact_root/Media-Pipeline-macos-arm64.zip"
rm -f "$archive"
ditto -c -k --sequesterRsrc --keepParent "$app" "$archive"

zsh "$repo_root/tools/Test-MacPackage.sh"

echo "Built: $app"
echo "Archive: $archive"
