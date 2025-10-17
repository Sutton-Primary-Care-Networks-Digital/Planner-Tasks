#!/usr/bin/env bash
set -euo pipefail

# Package the Deno macOS binary into a .app bundle and DMG
# Usage:
#   bash scripts/package-macos.sh [arm|x64]
# Defaults to arm (Apple Silicon / aarch64-apple-darwin)

ARCH=${1:-arm}
APP_NAME="Planner Tasks"
BUNDLE_ID="xyz.caerus.planner-tasks"
ROOT_DIR=$(cd "$(dirname "$0")/.." && pwd)
APP_DIR="$ROOT_DIR/deno-app"
DIST_DIR="$APP_DIR/dist"
BIN_NAME="planner-tasks-macos"

# Paths
BIN_PATH="$APP_DIR/$BIN_NAME"
ICNS_PATH="$APP_DIR/assets/icons/logo.icns"
APP_BUNDLE="$DIST_DIR/$APP_NAME.app"
CONTENTS="$APP_BUNDLE/Contents"
MACOS_DIR="$CONTENTS/MacOS"
RESOURCES_DIR="$CONTENTS/Resources"
PLIST_PATH="$CONTENTS/Info.plist"
DMG_PATH="$DIST_DIR/${APP_NAME// /_}-${ARCH}.dmg"

mkdir -p "$DIST_DIR"

# Ensure binary exists (arm or x64)
if [[ ! -f "$BIN_PATH" ]]; then
  echo "ERROR: $BIN_PATH not found. Compile first, e.g.:"
  if [[ "$ARCH" == "arm" ]]; then
    echo "  (cd deno-app && deno task compile-macos-arm)"
  else
    echo "  (cd deno-app && deno task compile-macos)"
  fi
  exit 1
fi

# Create .app bundle structure
rm -rf "$APP_BUNDLE"
mkdir -p "$MACOS_DIR" "$RESOURCES_DIR"

# Copy binary into bundle
cp "$BIN_PATH" "$MACOS_DIR/$APP_NAME"
chmod +x "$MACOS_DIR/$APP_NAME"

# Copy icon if present
if [[ -f "$ICNS_PATH" ]]; then
  cp "$ICNS_PATH" "$RESOURCES_DIR/logo.icns"
fi

# Create minimal Info.plist
cat > "$PLIST_PATH" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>CFBundleDisplayName</key>
  <string>$APP_NAME</string>
  <key>CFBundleExecutable</key>
  <string>$APP_NAME</string>
  <key>CFBundleIdentifier</key>
  <string>$BUNDLE_ID</string>
  <key>CFBundleName</key>
  <string>$APP_NAME</string>
  <key>CFBundlePackageType</key>
  <string>APPL</string>
  <key>CFBundleShortVersionString</key>
  <string>1.0.0</string>
  <key>CFBundleVersion</key>
  <string>1</string>
  <key>LSMinimumSystemVersion</key>
  <string>11.0</string>
  <key>LSApplicationCategoryType</key>
  <string>public.app-category.productivity</string>
  <key>CFBundleIconFile</key>
  <string>logo</string>
</dict>
</plist>
PLIST

# Build DMG using hdiutil
TMP_DMG_DIR=$(mktemp -d)
VOLUME_NAME="${APP_NAME// /_}"
mkdir -p "$TMP_DMG_DIR/$VOLUME_NAME"
cp -R "$APP_BUNDLE" "$TMP_DMG_DIR/$VOLUME_NAME/"

# Create DMG
hdiutil create -volname "$VOLUME_NAME" -srcfolder "$TMP_DMG_DIR/$VOLUME_NAME" -ov -format UDZO "$DMG_PATH" >/dev/null
rm -rf "$TMP_DMG_DIR"

echo "Packaged .app: $APP_BUNDLE"
echo "Created DMG: $DMG_PATH"
