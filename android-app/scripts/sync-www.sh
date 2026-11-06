#!/usr/bin/env bash
# Copy parent browser-game into android-app/www for Capacitor webDir.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
SRC="$(cd "$ROOT/.." && pwd)"
DEST="$ROOT/www"
rm -rf "$DEST"
mkdir -p "$DEST"
# Prefer rsync when available; else cp -a with exclusions.
if command -v rsync >/dev/null 2>&1; then
  rsync -a \
    --exclude 'android-app' \
    --exclude '.git' \
    --exclude 'node_modules' \
    --exclude 'CNAME' \
    "$SRC/" "$DEST/"
else
  # Copy top-level game files/dirs only (avoid nesting android-app).
  for name in index.html css js assets favicon.ico favicon.png favicon-16.png favicon-32.png apple-touch-icon.png README.md; do
    if [ -e "$SRC/$name" ]; then
      cp -a "$SRC/$name" "$DEST/"
    fi
  done
fi
test -f "$DEST/index.html"
echo "Synced $SRC → $DEST"
