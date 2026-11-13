# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

"""Make the modules under source/ importable, and provide the office the tests drive.

The tests require an installed office: the pyuno Python bindings for the uno and
com.sun.star modules, and the soffice and unopkg programs. The office fixture starts a
headless office with a scratch profile and connects to it over a UNO socket.
"""

import shutil
import socket
import subprocess
import sys
import time
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "source"))

try:
    import uno
except ImportError:
    raise ImportError(
        "These tests need the pyuno bindings of an installed office. Run them with the"
        " Python that has the uno module, the system Python on this machine; a tox or"
        " virtualenv Python sees it only with access to the system site-packages."
    ) from None


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
    if shutil.which("soffice") is None:
        pytest.fail("These tests need an installed office; soffice is not on the PATH.")
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
