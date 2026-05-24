#!/bin/bash
# Install OOXML-Validator CLI from GitHub Releases.
# The binary is NOT committed to this repo — re-run this script after
# `git clone` or after bumping version.json.
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

VERSION=$(node -e 'console.log(require("./version.json").version)')
echo "Installing OOXML-Validator v$VERSION ..."

OS=$(uname -s)
ARCH=$(uname -m)
case "$OS-$ARCH" in
  Darwin-arm64)        ASSET="osx-arm64.zip" ;;
  Darwin-x86_64)       ASSET="osx-x64.zip" ;;
  Linux-aarch64)       ASSET="linux-arm64.zip" ;;
  Linux-x86_64)        ASSET="linux-x64.zip" ;;
  MINGW*-x86_64|MSYS*-x86_64|CYGWIN*-x86_64) ASSET="win-x64.zip" ;;
  *)
    echo "Unsupported platform: $OS-$ARCH" >&2
    echo "See https://github.com/mikeebowen/OOXML-Validator/releases for available builds" >&2
    exit 1
    ;;
esac

URL="https://github.com/mikeebowen/OOXML-Validator/releases/download/v${VERSION}/${ASSET}"
echo "Fetching ${URL}"

TMPDIR=$(mktemp -d)
trap 'rm -rf "$TMPDIR"' EXIT

curl -fsSL "$URL" -o "$TMPDIR/$ASSET"

rm -rf bin
mkdir -p bin
unzip -q "$TMPDIR/$ASSET" -d bin

EXE=""
for cand in bin/OOXMLValidatorCLI bin/OOXMLValidatorCLI.exe; do
  if [ -f "$cand" ]; then EXE="$cand"; break; fi
done
if [ -z "$EXE" ]; then
  echo "Could not find OOXMLValidatorCLI executable in $ASSET" >&2
  ls -la bin/ >&2
  exit 1
fi

chmod +x "$EXE"
echo "Installed: $EXE"
echo "Size: $(ls -lh "$EXE" | awk '{print $5}')"

# macOS requires code-signed binaries on Apple Silicon. The upstream
# release ships unsigned binaries, so ad-hoc sign locally.
if [ "$OS" = "Darwin" ]; then
  xattr -d com.apple.quarantine "$EXE" 2>/dev/null || true
  xattr -d com.apple.provenance "$EXE" 2>/dev/null || true
  codesign --force --sign - "$EXE" >/dev/null 2>&1 || \
    echo "warn: codesign failed; binary may not run on Apple Silicon" >&2
fi

echo
echo "Smoke test:"
"$EXE" 2>&1 | head -5 || true
