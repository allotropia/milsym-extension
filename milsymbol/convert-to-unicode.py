# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import sys
f = open(sys.argv[1], mode='r', encoding='utf-8')
for c in f.read():
  if ord(c) <= 0x7F:
    print(c, end='')
  else:
    print(f'\\u{ord(c):04X}', end='')