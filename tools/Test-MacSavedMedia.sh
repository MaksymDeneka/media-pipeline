#!/bin/zsh
set -euo pipefail
script_dir=${0:A:h}
repo_root=${script_dir:h}
scratch=$(mktemp -d)
trap 'rm -rf "$scratch"' EXIT
sdk=${SDKROOT:-$(xcrun --show-sdk-path)}
swiftc -parse-as-library -swift-version 5 -j "${HEATUP_LOCAL_TEST_WORKERS:-2}" -sdk "$sdk" \
    "$repo_root/macos-app/Sources/MediaPipelineMac/Models.swift" \
    "$repo_root/macos-app/Sources/MediaPipelineMac/WorkerClient.swift" \
    "$repo_root/macos-app/Sources/MediaPipelineMac/SavedMediaStore.swift" \
    "$script_dir/Test-MacSavedMedia.swift" \
    -o "$scratch/saved-media-self-test"
"$scratch/saved-media-self-test"
