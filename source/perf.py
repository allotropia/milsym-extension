# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

"""Timing and call counting for the extension, switched on by environment variables.

Set MILSYM_PERF=1 to print a timing tree and a counter table to stderr. The output
goes to the terminal that started soffice.

Two further switches exist to price work that LibreOffice does in response to the
extension, by leaving that work out and measuring the difference:

MILSYM_PERF_NO_WRITES=1 computes the layout but writes no shape geometry.
MILSYM_PERF_NO_TREE=1 leaves the control dialog tree empty.

A document edited with either switch set will be laid out wrongly, so they are for
measurement runs only.

MILSYM_PERF_NO_NAME_INDEX=1 makes the diagram tree behave as though the shapes did not
have a name each, which is the path taken for a document whose shapes were renamed or
copied. It is slow but correct, so it is safe to set, and it is the only way to reach
that path deliberately.
"""

import os
import sys
import time
from contextlib import contextmanager

ENABLED = os.environ.get("MILSYM_PERF") == "1"
SKIP_GEOMETRY_WRITES = os.environ.get("MILSYM_PERF_NO_WRITES") == "1"
SKIP_TREE_REBUILD = os.environ.get("MILSYM_PERF_NO_TREE") == "1"
SKIP_NAME_INDEX = os.environ.get("MILSYM_PERF_NO_NAME_INDEX") == "1"

_counters = {}
_depth = 0


def count(label, amount=1):
    """Add to the named counter. Counters are printed and cleared by the outermost timed scope."""
    if not ENABLED:
        return
    _counters[label] = _counters.get(label, 0) + amount


@contextmanager
def timed(label):
    """Time the body and print the elapsed milliseconds, indented by how deeply scopes nest.

    Usable both as a with-statement and as a decorator on the method to be timed.
    """
    if not ENABLED:
        yield
        return

    global _depth
    indent = "  " * _depth
    _depth += 1
    started = time.perf_counter()
    try:
        yield
    finally:
        _depth -= 1
        elapsed_ms = (time.perf_counter() - started) * 1000.0
        print(f"[milsym-perf] {indent}{label} {elapsed_ms:.1f} ms", file=sys.stderr)
        if _depth == 0:
            _dump_counters()


def _dump_counters():
    """Print every counter that was touched, largest first, then start a fresh tally."""
    if not _counters:
        return
    for label, value in sorted(_counters.items(), key=lambda item: -item[1]):
        print(f"[milsym-perf]   {label}: {value}", file=sys.stderr)
    _counters.clear()
