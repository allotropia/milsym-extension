#!/bin/bash

# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

# This script combines the contents of the milsymbol directory into a single Javascript file.
# Additionally, some statements are modified to ensure compatibility with the rhino version
# used in LibreOffice which does not support modern JS features.

set -e  # Exit on any error

# Downloaded from https://github.com/spatialillusions/milsymbol/releases/tag/v3.0.3
INPUT_FILE="milsymbol-3.0.3.js"
OUTPUT_FILE="milsymbol.js"

python3 convert-to-unicode.py $INPUT_FILE > $OUTPUT_FILE

# Replace 'const' and 'let' with 'var'
sed -i -E -e 's/\<const|let\>/var/g' $OUTPUT_FILE

# Replace unsafe length check with null-safe version
sed -i -e 's/e=0<E.pre.length||0<E.post.length/e=(E.pre\&\&E.pre.length>0)\|\|(E.post\&\&E.post.length>0)/' "$OUTPUT_FILE"

# Correct M.x1 and M.x2 to ensure left and right labels are fully visible.
sed -i 's/!1}))}/!1}))}if(n.L1||n.L2||n.L3||n.L4||n.L5)M.x1=M.x1-(N*0.5);if(n.R1||n.R2||n.R3||n.R4||n.R5)M.x2=M.x2+(N*0.5);/' "$OUTPUT_FILE"

sed -e 's/^export default {/const pathsReinforced = {/' \
 paths-reinforced.js >> $OUTPUT_FILE
sed -E -e 's|^import pathsReinforced from "./paths-reinforced";||' \
 -e 's/^export default function sigintReinforced/function sigintReinforced/' -e 's/\<let\>/var/g' \
 sigint-reinforced-specialheadquarter.js >> $OUTPUT_FILE
sed -E -e 's/^export default function stack/function stack/' -e 's/\<let\>/var/g' \
 stack-extension.mjs >> $OUTPUT_FILE
cat country-flags.js >> $OUTPUT_FILE

# Append the final code to expose the main functionality
cat >>$OUTPUT_FILE <<\EOF
const parts = ms.getSymbolParts();
parts.splice(6, 0, sigintReinforced); // Insert sigintReinforced in the begining
parts.unshift(stack); // Insert stack at the begining
ms.setSymbolParts(parts);

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

