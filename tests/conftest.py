# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

"""Stand-ins for the modules the office provides at runtime.

The modules under source/ import uno, unohelper and com.sun.star types, which exist only
inside a running office. The stand-ins here carry just enough shape for those modules to
import under plain Python, so their pure functions can be tested without an office.
"""

import sys
import types
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "source"))


class Size:
    def __init__(self, Width=0, Height=0):
        self.Width = Width
        self.Height = Height


class Point:
    def __init__(self, X=0, Y=0):
        self.X = X
        self.Y = Y


class NamedValue:
    def __init__(self, Name=None, Value=None):
        self.Name = Name
        self.Value = Value


class PropertyValue:
    def __init__(self, Name=None, Handle=0, Value=None, State=0):
        self.Name = Name
        self.Handle = Handle
        self.Value = Value
        self.State = State


def _install_module(name, **attributes):
    module = types.ModuleType(name)
    for key, value in attributes.items():
        setattr(module, key, value)
    sys.modules[name] = module
    return module


_install_module("uno", systemPathToFileUrl=lambda path: "file://" + path)
_install_module("unohelper", Base=type("Base", (), {}))
_install_module("perf", count=lambda *args: None, timed=lambda function: function)

_com = _install_module("com")
_com.sun = _install_module("com.sun")
_com.sun.star = _install_module("com.sun.star")
_com.sun.star.awt = _install_module(
    "com.sun.star.awt", Size=Size, Point=Point, XCallback=type("XCallback", (), {})
)
_com.sun.star.beans = _install_module(
    "com.sun.star.beans", NamedValue=NamedValue, PropertyValue=PropertyValue
)
_com.sun.star.document = _install_module(
    "com.sun.star.document", XUndoAction=type("XUndoAction", (), {})
)
_com.sun.star.xml = _install_module(
    "com.sun.star.xml", AttributeData=type("AttributeData", (), {})
)
