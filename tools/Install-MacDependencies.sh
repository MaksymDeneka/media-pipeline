#!/bin/zsh
set -euo pipefail

if ! command -v brew >/dev/null 2>&1; then
    echo "Homebrew is required. Install it from https://brew.sh and run this script again."
    exit 1
fi

missing=()
command -v ffmpeg >/dev/null 2>&1 || missing+=(ffmpeg)
command -v ffprobe >/dev/null 2>&1 || missing+=(ffmpeg)
command -v exiftool >/dev/null 2>&1 || missing+=(exiftool)
command -v rclone >/dev/null 2>&1 || missing+=(rclone)

if (( ${#missing[@]} > 0 )); then
    unique=(${(u)missing})
    brew install "${unique[@]}"
fi

for tool in ffmpeg ffprobe exiftool rclone ssh; do
    if ! command -v "$tool" >/dev/null 2>&1; then
        echo "Missing required tool after installation: $tool"
        exit 1
    fi
done

remote_name=${RCLONE_REMOTE_NAME:-heatup-remote}
if ! rclone listremotes 2>/dev/null | grep -Fxq "$remote_name:"; then
    echo "Processing dependencies are ready, but uploads still need rclone remote '$remote_name'."
    echo "Copy ~/.config/rclone/rclone.conf from the existing machine or run: rclone config"
    echo "Also copy the configured SSH private key and set its permissions with chmod 600."
else
    echo "Media Pipeline dependencies and rclone remote '$remote_name' are ready."
fi
