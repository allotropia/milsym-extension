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
        "title": "Engagement Bar Text Help",
        "message": """The engagement amplifier shall be arranged as follows: A:BBB-CC

A (1 character) - Local vs Remote Engagement:
  (blank) = Local engagement (assigned to ownship)
  R = Remote engagement (assigned outside ownship control)
  B = Both local and remote engagements

BBB (up to 3 characters) - Engagement Stage:
  ASN = Assign/Cover
  ENG = Engage
  MIF = Missile in Flight
  CF = Cease Fire
  CE = Cease Engage
  HF = Hold Fire
  TE = Terminate Engagement
  BE = Break Engagement
  MBE = Management by Exception
  M<T = MBE Less Than Threshold
  MLT = Multiple Engagements

CC (up to 2 characters) - Weapon/Asset:
  M = Missile
  BM = Ballistic Missile
  CM = Cruise Missile
  GN = Gun
  T = Torpedo
  A = Attack Aircraft
  C = Combat Air Patrol (Defensive Counter Air)
  D = Defensive Counter Air (Combat Air Patrol)
  UW = USW/ASW Engagement
  MW = Mine Warfare Engagement
  SW = Surface Warfare Engagement
  EA = Electronic Attack
  ED = Electronic Defense
  UV = Unmanned Vehicle
  CW = Close-In Weapon System
  L3 = LAMPS
  VA = Vertical Launch ASROC
  ## = Number of Engagements (02-99)""",
    },
    "EvaluatRating": {
        "title": "Evaluation Rating Help",
        "message": """A text amplifier for units, equipment and installations that consists of a one-letter reliability rating and a one-number credibility rating.

Reliability Ratings:
  A - Completely reliable
  B - Usually reliable
  C - Fairly reliable
  D - Not usually reliable
  E - Unreliable
  F - Reliability cannot be judged

Credibility Ratings:
  1 - Confirmed by other sources
  2 - Probably true
  3 - Possibly true
  4 - Doubtfully true
  5 - Improbable
  6 - Truth cannot be judged""",
    },
}
