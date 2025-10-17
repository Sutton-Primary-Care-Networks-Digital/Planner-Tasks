#!/bin/bash

# Microsoft Planner Task Creator - macOS App Bundle Creator
set -e

APP_NAME="Microsoft Planner Task Creator"
BUNDLE_ID="com.plannertools.taskcreator"
VERSION="1.0.0"
BUILD_DIR="./dist"
APP_DIR="${BUILD_DIR}/${APP_NAME}.app"
CONTENTS_DIR="${APP_DIR}/Contents"
MACOS_DIR="${CONTENTS_DIR}/MacOS"
RESOURCES_DIR="${CONTENTS_DIR}/Resources"

echo "🚀 Creating macOS App Bundle for Microsoft Planner Task Creator..."

# Check if binary exists
if [[ ! -f "./planner-task-creator" ]]; then
    echo "❌ Binary not found. Run 'deno compile' first."
    exit 1
fi

# Clean and create directories
rm -rf "${BUILD_DIR}"
mkdir -p "${MACOS_DIR}"
mkdir -p "${RESOURCES_DIR}"

echo "📦 Setting up app bundle structure..."

# Copy the compiled binary
cp "./planner-task-creator" "${MACOS_DIR}/planner-task-creator"
chmod +x "${MACOS_DIR}/planner-task-creator"

# Copy static files
cp -r static "${RESOURCES_DIR}/"

# Create Info.plist
echo "📝 Creating Info.plist..."
cat > "${CONTENTS_DIR}/Info.plist" << EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>launcher</string>
    <key>CFBundleIdentifier</key>
    <string>${BUNDLE_ID}</string>
    <key>CFBundleName</key>
    <string>${APP_NAME}</string>
    <key>CFBundleDisplayName</key>
    <string>${APP_NAME}</string>
    <key>CFBundleVersion</key>
    <string>${VERSION}</string>
    <key>CFBundleShortVersionString</key>
    <string>${VERSION}</string>
    <key>CFBundleInfoDictionaryVersion</key>
    <string>6.0</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.14</string>
    <key>NSHighResolutionCapable</key>
    <true/>
    <key>NSAppTransportSecurity</key>
    <dict>
        <key>NSAllowsArbitraryLoads</key>
        <true/>
    </dict>
    <key>LSUIElement</key>
    <true/>
</dict>
</plist>
EOF

# Create launcher script
echo "📜 Creating launcher script..."
cat > "${MACOS_DIR}/launcher" << 'EOF'
#!/bin/bash

# Get the app bundle path
BUNDLE_PATH="$(dirname "$(dirname "$0")")"
RESOURCES_PATH="${BUNDLE_PATH}/Resources"
BINARY_PATH="${BUNDLE_PATH}/MacOS/planner-task-creator"

# Change to resources directory so static files are accessible
cd "${RESOURCES_PATH}"

# Start the server in the background
"${BINARY_PATH}" &
SERVER_PID=$!

# Wait a moment for server to start
sleep 2

# Open the browser
open "http://localhost:8080"

echo "Microsoft Planner Task Creator is running on http://localhost:8080"
echo "Close this window to stop the application."

# Wait for the server process
wait $SERVER_PID
EOF

chmod +x "${MACOS_DIR}/launcher"

echo "✅ App bundle created at: ${APP_DIR}"

# Create DMG
echo "📦 Creating DMG installer..."
DMG_NAME="Microsoft-Planner-Task-Creator-${VERSION}.dmg"
DMG_PATH="${BUILD_DIR}/${DMG_NAME}"

# Create temporary DMG content directory
DMG_TEMP="${BUILD_DIR}/dmg_temp"
rm -rf "${DMG_TEMP}"
mkdir -p "${DMG_TEMP}"

# Copy app to DMG
cp -r "${APP_DIR}" "${DMG_TEMP}/"

# Create Applications symlink
ln -s /Applications "${DMG_TEMP}/Applications"

# Create README for DMG
cat > "${DMG_TEMP}/README.txt" << EOF
Microsoft Planner Task Creator v${VERSION}

Installation Instructions:
1. Drag "Microsoft Planner Task Creator.app" to the Applications folder
2. Launch the application from Applications
3. The app will start a local web server and open your browser
4. Access the interface at http://localhost:8080

Features:
- Upload CSV/Excel files with task data
- Authenticate with Microsoft 365 (NHS.net supported)
- Create tasks in Microsoft Planner
- Assign tasks to team members
- Bulk task creation with progress tracking

Support:
- The application requires network access to communicate with Microsoft Graph API
- Files are processed locally on your machine
- No data is sent to external servers except Microsoft's official APIs

Technical Details:
- Runs local web server on port 8080
- Supports CSV and Excel file formats
- Compatible with macOS 10.14 and later
EOF

# Create DMG
echo "🔨 Building DMG..."
hdiutil create \
    -volname "${APP_NAME}" \
    -srcfolder "${DMG_TEMP}" \
    -ov \
    -format UDZO \
    -imagekey zlib-level=9 \
    "${DMG_PATH}"

# Clean up temp directory
rm -rf "${DMG_TEMP}"

echo ""
echo "✅ Build completed successfully!"
echo ""
echo "📦 Created files:"
echo "   App Bundle: ${APP_DIR}"
echo "   DMG Installer: ${DMG_PATH}"
echo ""
echo "🚀 Installation:"
echo "   1. Double-click ${DMG_NAME} to mount"
echo "   2. Drag the app to Applications folder"
echo "   3. Launch from Applications"
echo ""
echo "⚠️  Note: The app runs a local web server on port 8080"
echo "   Access via: http://localhost:8080"
EOF