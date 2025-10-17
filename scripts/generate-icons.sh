#!/usr/bin/env bash
set -euo pipefail

# Generate app icons (ICO/ICNS) and favicon from repo-root logos
# Requirements (best effort):
# - macOS: sips (built-in), iconutil (built-in)
# - ImageMagick 'convert' for ICO generation (brew install imagemagick) OR png2ico
# Inputs: ./logo.png (preferred), ./logo.svg (optional)
# Outputs:
# - deno-app/assets/icons/logo.icns
# - deno-app/assets/icons/logo.ico (if ImageMagick/pnge2ico available)
# - deno-app/static/favicon.ico
# - Temporary: ./tmp_iconset (removed at end)

ROOT_DIR=$(cd "$(dirname "$0")/.." && pwd)
LOGO_PNG="$ROOT_DIR/logo.png"
LOGO_SVG="$ROOT_DIR/logo.svg"
OUT_DIR="$ROOT_DIR/deno-app/assets/icons"
FAVICON_OUT="$ROOT_DIR/deno-app/static/favicon.ico"
ICONSET_DIR="$ROOT_DIR/tmp_iconset.iconset"

mkdir -p "$OUT_DIR"
mkdir -p "$(dirname "$FAVICON_OUT")"

# Helper: make a square PNG of given size from the best source we have
make_png_size() {
  local size=$1
  local outfile=$2
  if [[ -f "$LOGO_PNG" ]]; then
    sips -s format png -Z "$size" "$LOGO_PNG" --out "$outfile" >/dev/null
  elif [[ -f "$LOGO_SVG" ]]; then
    if command -v rsvg-convert >/dev/null 2>&1; then
      rsvg-convert -w "$size" -h "$size" "$LOGO_SVG" -o "$outfile"
    elif command -v convert >/dev/null 2>&1; then
      convert -background none -resize ${size}x${size} "$LOGO_SVG" "$outfile"
    else
      echo "ERROR: No PNG logo and cannot rasterize SVG (need rsvg-convert or ImageMagick)." >&2
      exit 1
    fi
  else
    echo "ERROR: Missing ./logo.png or ./logo.svg at repo root." >&2
    exit 1
  fi
}

# Build iconset for ICNS (macOS)
rm -rf "$ICONSET_DIR"
mkdir -p "$ICONSET_DIR"
for sz in 16 32 64 128 256 512; do
  make_png_size "$sz" "$ICONSET_DIR/icon_${sz}x${sz}.png"
  # @2x assets
  dbl=$((sz*2))
  make_png_size "$dbl" "$ICONSET_DIR/icon_${sz}x${sz}@2x.png"
done

# Create .icns
if command -v iconutil >/dev/null 2>&1; then
  iconutil -c icns "$ICONSET_DIR" -o "$OUT_DIR/logo.icns"
  echo "Created: $OUT_DIR/logo.icns"
else
  echo "WARN: iconutil not found; skipping ICNS generation" >&2
fi

# Create favicon.ico (16,32,48)
TMP_FAV16=$(mktemp -t fav16.XXXX.png)
TMP_FAV32=$(mktemp -t fav32.XXXX.png)
TMP_FAV48=$(mktemp -t fav48.XXXX.png)
make_png_size 16 "$TMP_FAV16"
make_png_size 32 "$TMP_FAV32"
make_png_size 48 "$TMP_FAV48"
if command -v convert >/dev/null 2>&1; then
  convert "$TMP_FAV16" "$TMP_FAV32" "$TMP_FAV48" "$FAVICON_OUT"
  echo "Created: $FAVICON_OUT"
elif command -v png2ico >/dev/null 2>&1; then
  png2ico "$FAVICON_OUT" "$TMP_FAV16" "$TMP_FAV32" "$TMP_FAV48"
  echo "Created: $FAVICON_OUT"
else
  echo "WARN: Neither ImageMagick 'convert' nor 'png2ico' found; cannot create favicon.ico" >&2
fi
rm -f "$TMP_FAV16" "$TMP_FAV32" "$TMP_FAV48"

# Create Windows .ico with multiple sizes if tools exist
OUT_ICO="$OUT_DIR/logo.ico"
TMP16=$(mktemp -t ico16.XXXX.png)
TMP24=$(mktemp -t ico24.XXXX.png)
TMP32=$(mktemp -t ico32.XXXX.png)
TMP48=$(mktemp -t ico48.XXXX.png)
TMP64=$(mktemp -t ico64.XXXX.png)
TMP128=$(mktemp -t ico128.XXXX.png)
make_png_size 16 "$TMP16"
make_png_size 24 "$TMP24"
make_png_size 32 "$TMP32"
make_png_size 48 "$TMP48"
make_png_size 64 "$TMP64"
make_png_size 128 "$TMP128"
if command -v convert >/dev/null 2>&1; then
  convert "$TMP16" "$TMP24" "$TMP32" "$TMP48" "$TMP64" "$TMP128" "$OUT_ICO"
  echo "Created: $OUT_ICO"
elif command -v png2ico >/dev/null 2>&1; then
  png2ico "$OUT_ICO" "$TMP16" "$TMP24" "$TMP32" "$TMP48" "$TMP64" "$TMP128"
  echo "Created: $OUT_ICO"
else
  echo "WARN: Neither ImageMagick 'convert' nor 'png2ico' found; skipping Windows .ico" >&2
fi
rm -f "$TMP16" "$TMP24" "$TMP32" "$TMP48" "$TMP64" "$TMP128"

# Cleanup
rm -rf "$ICONSET_DIR"

echo "All done. Outputs (if tools available):"
echo " - $OUT_DIR/logo.icns"
echo " - $OUT_DIR/logo.ico"
echo " - $FAVICON_OUT"
