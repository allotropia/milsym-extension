#!/bin/bash

# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

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

echo "Processing milsymbol JavaScript files..."
cd milsymbol
./combine.sh
cd ..

# Create temporary directory
mkdir -p "$TEMP_DIR"

echo "Copying files to temporary directory..."

# Copy all files except those we want to exclude
rsync -av \
    --exclude="README.md" \
    --exclude="build.sh" \
    --exclude="milsymbol/combine.sh" \
    --exclude="milsymbol/milsymbol-3.0.3.js" \
    --exclude="milsymbol/paths-reinforced.js" \
    --exclude="milsymbol/sigint-reinforced-specialheadquarter.js" \
    --exclude="milsymbol/stack-extension.mjs" \
    --exclude="milsymbol/country-flags.js" \
    --exclude="milsymbol/convert-to-unicode.py" \
    --exclude=".*" \
    --exclude="*~" \
    --exclude="__pycache__" \
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
