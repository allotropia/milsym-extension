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

sed -e 's/^export default {/const pathsReinforced = {/' \
 paths-reinforced.js >> $OUTPUT_FILE
sed -E -e 's|^import pathsReinforced from "./paths-reinforced";||' \
 -e 's/^export default function sigintReinforced/function sigintReinforced/' -e 's/\<let\>/var/g' \
 sigint-reinforced-specialheadquarter.js >> $OUTPUT_FILE
sed -E -e 's/^export default function stack/function stack/' -e 's/\<let\>/var/g' \
 stack-extension.mjs >> $OUTPUT_FILE

# Append the final code to expose the main functionality
cat >>$OUTPUT_FILE <<\EOF
const parts = ms.getSymbolParts();
parts.splice(6, 0, sigintReinforced); // Insert sigintReinforced in the begining
parts.unshift(stack); // Insert stack at the begining
ms.setSymbolParts(parts);

var options = {};
for (var i = 1; i < ARGUMENTS.length; ++i) {
	var name = ARGUMENTS[i].Name;
	var value = ARGUMENTS[i].Value;

	if (String(name) == "colorMode") {
		value = String(value);
	}

	options[name] = value;
}
new ms.Symbol(String(ARGUMENTS[0]), options).asSVG();
EOF

echo "Successfully created $OUTPUT_FILE"

