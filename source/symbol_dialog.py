# SPDX-License-Identifier: MPL-2.0

import os
import sys
from symbol_dialog_handler import SymbolDialogHandler

def open_symbol_dialog(ctx, model, controller):
    dialog_provider = ctx.getServiceManager().createInstanceWithContext(
        "com.sun.star.awt.DialogProvider2", ctx)
    dialog_url = "vnd.sun.star.extension://com.collabora.milsymbol/dialog/MilitarySymbolDlg.xdl"
    try:
        handler = SymbolDialogHandler(ctx, model, controller, None)
        dialog = dialog_provider.createDialogWithHandler(dialog_url, handler)
        handler.dialog = dialog
        handler.init_dialog_controls()
        dialog.execute()
        handler.remove_temp_preview_svg()
    except Exception as e:
        print(f"Error opening symbol dialog: {e}")