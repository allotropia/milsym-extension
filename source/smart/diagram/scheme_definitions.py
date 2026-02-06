# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

# This file incorporates work covered by the following license notice:
#   SPDX-License-Identifier: LGPL-3.0-only

"""
Color scheme definitions
Python port of SchemeDefinitions.java
"""


class SchemeDefinitions:
    """Color scheme definitions for diagrams"""

    # Color schemes as tuples of (light_color, dark_color)
    BLUE_SCHEME = (0x0099FF, 0x000077)
    AQUA_SCHEME = (65535, 0x0099FF)
    RED_SCHEME = (0xFF0000, 0x660000)
    FIRE_SCHEME = (0xFFFF00, 0xE00000)
    SUN_SCHEME = (0xFFFD45, 0xFF8000)
    GREEN_SCHEME = (65280, 0x0E4B00)
    OLIVE_SCHEME = (14677829, 0x3C4900)
    PURPLE_SCHEME = (16711935, 0x5E1B5D)
    PINK_SCHEME = (0xFFDEEC, 0xD60084)
    INDIAN_SCHEME = (0xFFEEEE, 0xCD5C5C)
    MAROON_SCHEME = (0xF9CFB5, 0xA33E03)
    BROWN_SCHEME = (0xECBA74, 0x5C2D0A)

    # Array of all color schemes
    COLOR_SCHEMES = [
        BLUE_SCHEME,
        AQUA_SCHEME,
        RED_SCHEME,
        FIRE_SCHEME,
        SUN_SCHEME,
        GREEN_SCHEME,
        OLIVE_SCHEME,
        PURPLE_SCHEME,
        PINK_SCHEME,
        INDIAN_SCHEME,
        MAROON_SCHEME,
        BROWN_SCHEME,
    ]

    NUM_OF_SCHEMES = 12

    @staticmethod
    def get_gradient_color(color_code: int, index: int, steps: int) -> int:
        """
        Get gradient color calculation
        Note: This is a simplified version - the original Java method had more complex logic
        """
        # Simplified gradient calculation
        if steps <= 1:
            return color_code

        factor = index / (steps - 1)

        # Extract RGB components
        r = (color_code >> 16) & 0xFF
        g = (color_code >> 8) & 0xFF
        b = color_code & 0xFF

        # Apply gradient factor (darken)
        r = int(r * (1.0 - factor * 0.5))
        g = int(g * (1.0 - factor * 0.5))
        b = int(b * (1.0 - factor * 0.5))

        # Combine back to single integer
        return (r << 16) | (g << 8) | b

    @staticmethod
    def get_gradient_color_with_target(
        start_color: int, target_color: int, index: int, steps: int
    ) -> int:
        """
        Get gradient color between start and target colors
        """
        if steps <= 1:
            return start_color

        factor = index / (steps - 1)

        # Extract RGB components for start color
        start_r = (start_color >> 16) & 0xFF
        start_g = (start_color >> 8) & 0xFF
        start_b = start_color & 0xFF

        # Extract RGB components for target color
        target_r = (target_color >> 16) & 0xFF
        target_g = (target_color >> 8) & 0xFF
        target_b = target_color & 0xFF

        # Interpolate
        r = int(start_r + (target_r - start_r) * factor)
        g = int(start_g + (target_g - start_g) * factor)
        b = int(start_b + (target_b - start_b) * factor)

        # Combine back to single integer
        return (r << 16) | (g << 8) | b
