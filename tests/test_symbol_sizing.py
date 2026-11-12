# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

"""The sizing rule applied when the picture of an existing symbol shape is replaced."""

from com.sun.star.awt import Size
from utils import size_with_height_and_aspect_ratio


def size(width, height):
    return Size(Width=width, Height=height)


def test_size_stays_the_same_over_repeated_edits():
    """Alternating between drawings of different aspect ratios returns to the same size.

    A square symbol and one 1.6 times as wide, as with text labels, stand in for two
    states the user switches between by editing the symbol repeatedly.
    """
    square_drawing = size(1000, 1000)
    wide_drawing = size(1600, 1000)

    box = size(1000, 1000)
    for _ in range(10):
        box = size_with_height_and_aspect_ratio(box, wide_drawing)
        box = size_with_height_and_aspect_ratio(box, square_drawing)

    assert (box.Width, box.Height) == (1000, 1000)


def test_width_grows_when_the_drawing_gets_wider():
    box = size_with_height_and_aspect_ratio(size(1000, 1000), size(1600, 1000))

    assert (box.Width, box.Height) == (1600, 1000)


def test_height_is_kept_when_the_drawing_gets_narrower():
    box = size_with_height_and_aspect_ratio(size(1600, 1000), size(1000, 1000))

    assert (box.Width, box.Height) == (1000, 1000)


def test_a_degenerate_size_is_returned_unchanged():
    box = size(0, 1000)

    assert size_with_height_and_aspect_ratio(box, size(1000, 1000)) is box


def test_a_degenerate_drawing_leaves_the_size_unchanged():
    box = size(1000, 1000)

    assert size_with_height_and_aspect_ratio(box, size(1000, 0)) is box
