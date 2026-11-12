# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

"""Make the modules under source/ importable for the tests.

A Python with pyuno, such as the one a LibreOffice installation ships, provides the uno,
unohelper and com.sun.star modules on its own, outside a running office. On a Python
without pyuno, stand-ins are installed instead that carry just enough shape for the
source modules to import, so their pure functions can still be tested. UNO_IS_REAL says
which of the two this test run has; the tests that talk to a running office need the
real one.
"""

import shutil
import socket
import subprocess
import sys
import time
import types
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "source"))

try:
    import uno

    UNO_IS_REAL = True
except ImportError:
    UNO_IS_REAL = False


def _free_port():
    with socket.socket() as probe:
        probe.bind(("localhost", 0))
        return probe.getsockname()[1]


def _connect(port, deadline_seconds=60):
    local_context = uno.getComponentContext()
    resolver = local_context.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_context
    )
    url = f"uno:socket,host=localhost,port={port};urp;StarOffice.ComponentContext"
    deadline = time.monotonic() + deadline_seconds
    while True:
        try:
            return resolver.resolve(url)
        except Exception:
            if time.monotonic() > deadline:
                raise
            time.sleep(0.5)


@pytest.fixture(scope="session")
def office(tmp_path_factory):
    """A headless office with a scratch profile, reachable over a UNO socket.

    One office serves the whole test session, because the extension modules keep
    singletons in the test process, such as the translator, that hold on to the first
    component context they are given.
    """
    if not UNO_IS_REAL or shutil.which("soffice") is None:
        pytest.skip("needs pyuno and a soffice binary")
    profile = tmp_path_factory.mktemp("office-profile")
    oxt = Path(__file__).resolve().parent.parent / "milsymbol-extension.oxt"
    if oxt.exists():
        # The installed package supplies the dialog definitions, the images, the
        # configuration schema and the symbol script. The extension Python that the
        # tests drive is the working tree's own, imported into the test process.
        subprocess.run(
            [
                "unopkg",
                "add",
                "-f",
                "--suppress-license",
                f"-env:UserInstallation=file://{profile}",
                str(oxt),
            ],
            check=True,
            capture_output=True,
        )
    port = _free_port()
    process = subprocess.Popen(
        [
            "soffice",
            "--headless",
            "--invisible",
            "--norestore",
            f"-env:UserInstallation=file://{profile}",
            f"--accept=socket,host=localhost,port={port};urp;",
        ]
    )
    try:
        context = _connect(port)
    except Exception:
        process.kill()
        raise
    desktop = context.ServiceManager.createInstanceWithContext(
        "com.sun.star.frame.Desktop", context
    )
    yield SimpleNamespace(context=context, desktop=desktop)
    try:
        desktop.terminate()
        process.wait(timeout=10)
    except Exception:
        process.kill()


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


def _install_stand_ins():
    _install_module("uno", systemPathToFileUrl=lambda path: "file://" + path)
    _install_module("unohelper", Base=type("Base", (), {}))
    _install_module("perf", count=lambda *args: None, timed=lambda function: function)

    com = _install_module("com")
    com.sun = _install_module("com.sun")
    com.sun.star = _install_module("com.sun.star")
    com.sun.star.awt = _install_module(
        "com.sun.star.awt", Size=Size, Point=Point, XCallback=type("XCallback", (), {})
    )
    com.sun.star.beans = _install_module(
        "com.sun.star.beans", NamedValue=NamedValue, PropertyValue=PropertyValue
    )
    com.sun.star.document = _install_module(
        "com.sun.star.document", XUndoAction=type("XUndoAction", (), {})
    )
    com.sun.star.xml = _install_module(
        "com.sun.star.xml", AttributeData=type("AttributeData", (), {})
    )


if not UNO_IS_REAL:
    _install_stand_ins()
