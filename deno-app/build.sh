#!/bin/bash

# Microsoft Planner Task Creator - Build Script
# Compiles the Deno application for multiple platforms

set -e

echo "🚀 Microsoft Planner Task Creator - Build Script"
echo "================================================"

# Check if Deno is installed
if ! command -v deno &> /dev/null; then
    echo "❌ Deno is not installed. Please install Deno first:"
    echo "   https://deno.land/manual/getting_started/installation"
    exit 1
fi

echo "✅ Deno version: $(deno --version | head -n1)"

# Create build directory
BUILD_DIR="build"
rm -rf "$BUILD_DIR"
mkdir -p "$BUILD_DIR"

echo ""
echo "📦 Building executables..."

# Build for current platform
echo "🔨 Building for current platform..."
deno compile --allow-net --allow-read --allow-write --allow-env --allow-run --output "$BUILD_DIR/planner-tasks" main.ts
echo "✅ Built: $BUILD_DIR/planner-tasks"

# Build for Windows
echo "🪟 Building for Windows..."
deno compile --allow-net --allow-read --allow-write --allow-env --allow-run --target x86_64-pc-windows-msvc --output "$BUILD_DIR/planner-tasks-windows.exe" main.ts
echo "✅ Built: $BUILD_DIR/planner-tasks-windows.exe"

# Build for macOS
echo "🍎 Building for macOS..."
deno compile --allow-net --allow-read --allow-write --allow-env --allow-run --target x86_64-apple-darwin --output "$BUILD_DIR/planner-tasks-macos" main.ts
echo "✅ Built: $BUILD_DIR/planner-tasks-macos"

# Build for Linux
echo "🐧 Building for Linux..."
deno compile --allow-net --allow-read --allow-write --allow-env --allow-run --target x86_64-unknown-linux-gnu --output "$BUILD_DIR/planner-tasks-linux" main.ts
echo "✅ Built: $BUILD_DIR/planner-tasks-linux"

echo ""
echo "📁 Build Summary:"
echo "=================="
ls -lh "$BUILD_DIR/"

echo ""
echo "🎉 Build complete! Executables are in the '$BUILD_DIR' directory."
echo ""
echo "To run:"
echo "  Current platform: ./$BUILD_DIR/planner-tasks"
echo "  Windows:         ./$BUILD_DIR/planner-tasks-windows.exe"  
echo "  macOS:           ./$BUILD_DIR/planner-tasks-macos"
echo "  Linux:           ./$BUILD_DIR/planner-tasks-linux"
echo ""
echo "📝 Each executable includes the complete application and can be"
echo "   distributed as a single file with no additional dependencies."