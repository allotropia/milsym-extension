# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import platform
from symbol_dialog_handler import SymbolDialogHandler

def open_symbol_dialog(ctx, model, controller):
    dialog_provider = ctx.getServiceManager().createInstanceWithContext(
        "com.sun.star.awt.DialogProvider2", ctx)

    system = platform.system()  # 'Windows', 'Linux'
    if system == "Windows":
        dialog_file = "MilitarySymbolDlg_WINDOWS.xdl"
    else:
        dialog_file = "MilitarySymbolDlg_LINUX.xdl"

    dialog_url = f"vnd.sun.star.extension://com.collabora.milsymbol/dialog/{dialog_file}"

    try:
        handler = SymbolDialogHandler(ctx, model, controller, None)
        dialog = dialog_provider.createDialogWithHandler(dialog_url, handler)
        handler.dialog = dialog
        handler.init_dialog_controls()
        dialog.execute()
        handler.remove_temp_preview_svg()
    except Exception as e:
        print(f"Error opening symbol dialog: {e}")
