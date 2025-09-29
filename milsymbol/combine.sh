#!/bin/bash

# SPDX-License-Identifier: MPL-2.0

# This script combines the contents of the milsymbol directory into a single Javascript file.
# Additionally, some statements are modified to ensure compatibility with the rhino version
# used in LibreOffice which does not support modern JS features.

set -e  # Exit on any error

# Downloaded from https://github.com/spatialillusions/milsymbol/releases/tag/v3.0.3
INPUT_FILE="milsymbol-3.0.3.js"
OUTPUT_FILE="milsymbol.js"

cp $INPUT_FILE $OUTPUT_FILE

# Replace 'const' and 'let' with 'var'
sed -i -E -e 's/\<const|let\>/var/g' $OUTPUT_FILE

# Append the final code to expose the main functionality
cat >>$OUTPUT_FILE <<\EOF
var options = {};
for (var i = 1; i < ARGUMENTS.length; ++i) {
 options[ARGUMENTS[i].Name] = ARGUMENTS[i].Value;
}
new ms.Symbol(String(ARGUMENTS[0]), options).asSVG();
EOF

echo "Successfully created $OUTPUT_FILE"

