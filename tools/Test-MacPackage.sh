#!/bin/zsh
set -euo pipefail

script_dir=${0:A:h}
repo_root=${script_dir:h}
app="$repo_root/artifacts/Media Pipeline.app"
worker="$app/Contents/Helpers/media-pipeline-worker"
main="$app/Contents/MacOS/MediaPipelineMac"

if [[ ! -d "$app" ]]; then
    echo "App bundle not found. Run tools/Build-MacApp.sh first."
    exit 1
fi

scratch=$(mktemp -d)
trap 'rm -rf "$scratch"' EXIT

plutil -lint "$app/Contents/Info.plist"
codesign --verify --deep --strict --verbose=2 "$app"
test -x "$main"
test -x "$worker"
test -f "$app/Contents/Resources/AppIcon.icns"
test -f "$app/Contents/Resources/default-config.ini"
root_entry_count=$(find "$app" -mindepth 1 -maxdepth 1 -print | wc -l | tr -d ' ')
if [[ "$root_entry_count" != "1" ]] || [[ ! -d "$app/Contents" ]]; then
    echo "The application root must contain only Contents."
    exit 1
fi
"$main" --check-resources

codesign -d --entitlements :- "$worker" > "$scratch/worker-entitlements.plist" 2>/dev/null
if [[ "$(/usr/libexec/PlistBuddy -c 'Print :com.apple.security.cs.allow-jit' \
    "$scratch/worker-entitlements.plist")" != "true" ]]; then
    echo "The worker is missing its required JIT entitlement."
    exit 1
fi

sed "s|PipelineRoot = ~/MediaPipeline|PipelineRoot = $scratch/root|" \
    "$repo_root/macos-app/Sources/MediaPipelineMac/Resources/default-config.ini" \
    > "$scratch/config.ini"

"$worker" check --config "$scratch/config.ini"
echo "macOS package checks passed."
