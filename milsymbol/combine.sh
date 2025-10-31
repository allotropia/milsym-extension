#!/bin/bash

# SPDX-License-Identifier: MPL-2.0

# This script combines the contents of the milsymbol directory into a single Javascript file.
# Additionally, some statements are modified to ensure compatibility with the rhino version
# used in LibreOffice which does not support modern JS features.

set -e  # Exit on any error

# Downloaded from https://github.com/spatialillusions/milsymbol/releases/tag/v3.0.3
INPUT_FILE="milsymbol-2.2.0.js"
OUTPUT_FILE="milsymbol.js"

cp $INPUT_FILE $OUTPUT_FILE

cat country-flags.js >> $OUTPUT_FILE

# Append the final code to expose the main functionality
cat >>$OUTPUT_FILE <<\EOF
var options = {};
for (var i = 1; i < ARGUMENTS.length; ++i) {
    var name = ARGUMENTS[i].Name;
    var value = String(ARGUMENTS[i].Value);

    if (String(name) == "fill") {
        value = String(value) === "true";
    }

    options[name] = value;
}
new ms.Symbol(String(ARGUMENTS[0]), options).asSVG();
EOF

echo "Successfully created $OUTPUT_FILE"

