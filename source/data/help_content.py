# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

"""
Help content dictionary for dialog field help buttons.

To add a new help button:
1. Add an entry to HELP_CONTENT with key matching the field name (without "btHelp" prefix)
2. Add the button to both XDL dialog files with id="btHelp<FieldName>"
3. Add the button's HelpText property to the properties files

The handler in symbol_dialog_handler.py will automatically handle any button
with an id starting with "btHelp" by looking up the field name in this dictionary.
"""

HELP_CONTENT = {
    "EngageBarText": {
        "title": "Help.EngageBarText.Title",
        "message": "Help.EngageBarText.Message",
    },
    "EvaluatRating": {
        "title": "Help.EvaluatRating.Title",
        "message": "Help.EvaluatRating.Message",
    },
}
