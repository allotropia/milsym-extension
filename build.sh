#!/bin/bash
# SPDX-License-Identifier: MPL-2.0

# Build script for LibreOffice extension
# Creates a .oxt file from the folder contents, excluding README and build artifacts

set -e  # Exit on any error

# Configuration
OUTPUT_NAME="milsymbol-extension.oxt"
TEMP_DIR="build_temp"

echo "Building LibreOffice extension: $OUTPUT_NAME"

# Clean up any previous build artifacts
rm -rf "$TEMP_DIR"
rm -f "$OUTPUT_NAME"

# Create temporary directory
mkdir -p "$TEMP_DIR"

echo "Copying files to temporary directory..."

# Copy all files except those we want to exclude
rsync -av \
    --exclude="README.md" \
    --exclude="build.sh" \
    --exclude=".git/" \
    --exclude=".gitignore" \
    --exclude="$TEMP_DIR/" \
    --exclude="*.oxt" \
    . "$TEMP_DIR/"

echo "Creating .oxt archive..."

# Create the .oxt file (which is just a zip file with different extension)
cd "$TEMP_DIR"
zip -r "../$OUTPUT_NAME" .
cd ..

# Clean up temporary directory
rm -rf "$TEMP_DIR"

echo "Successfully created $OUTPUT_NAME"
echo "Extension ready for installation in LibreOffice"