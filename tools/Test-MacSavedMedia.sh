#!/bin/zsh
set -euo pipefail
script_dir=${0:A:h}
repo_root=${script_dir:h}
scratch=$(mktemp -d)
trap 'rm -rf "$scratch"' EXIT
sdk=${SDKROOT:-$(xcrun --show-sdk-path)}
if command -v ffmpeg >/dev/null 2>&1; then
    ffmpeg -hide_banner -loglevel error -f lavfi -i color=c=blue:s=160x320:d=0.1 \
        -c:v libx264 -threads "${HEATUP_LOCAL_TEST_WORKERS:-2}" -pix_fmt yuv420p "$scratch/preview.mp4"
    export MEDIA_THUMBNAIL_TEST_VIDEO="$scratch/preview.mp4"
fi
swiftc -parse-as-library -swift-version 5 -j "${HEATUP_LOCAL_TEST_WORKERS:-2}" -sdk "$sdk" \
    "$repo_root/macos-app/Sources/MediaPipelineMac/Models.swift" \
    "$repo_root/macos-app/Sources/MediaPipelineMac/WorkerClient.swift" \
    "$repo_root/macos-app/Sources/MediaPipelineMac/SavedMediaStore.swift" \
    "$repo_root/macos-app/Sources/MediaPipelineMac/MediaThumbnail.swift" \
    "$repo_root/macos-app/Sources/MediaPipelineMac/ConfigurationStore.swift" \
    "$repo_root/macos-app/Sources/MediaPipelineMac/AppPaths.swift" \
    "$repo_root/macos-app/Sources/MediaPipelineMac/UploadStore.swift" \
    "$script_dir/Test-MacSavedMedia.swift" \
    -o "$scratch/saved-media-self-test"
"$scratch/saved-media-self-test"
