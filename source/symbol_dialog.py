# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import platform
from symbol_dialog_handler import ListEscapeKeyHandler, SymbolDialogHandler


def open_symbol_dialog(
    ctx, model, controller, sidebar_panel, selected_shape, selected_node_value
):
    dialog_provider = ctx.getServiceManager().createInstanceWithContext(
        "com.sun.star.awt.DialogProvider2", ctx
    )

    system = platform.system()  # 'Windows', 'Linux'
    if system == "Windows":
        dialog_file = "MilitarySymbolDlg_WINDOWS.xdl"
    else:
        dialog_file = "MilitarySymbolDlg_LINUX.xdl"

    dialog_url = (
        f"vnd.sun.star.extension://com.collabora.milsymbol/dialog/{dialog_file}"
    )

    try:
        handler = SymbolDialogHandler(
            ctx,
            model,
            controller,
            None,
            sidebar_panel,
            selected_shape,
            selected_node_value,
        )
        dialog = dialog_provider.createDialogWithHandler(dialog_url, handler)
        handler.dialog = dialog
        handler.init_dialog_controls()
        execute_with_list_escape(ctx, dialog, handler)
    except Exception as e:
        print(f"Error opening symbol dialog: {e}")


def execute_with_list_escape(ctx, dialog, handler):
    """Run the dialog, with Escape closing a shown dropdown tree or search result list
    instead of the dialog.

    The office gives an Escape that no control handles to the dialog, which then closes.
    A key handler on the toolkit sees each key before any window does, and takes Escape
    for itself while a list has the focus. On office versions whose toolkit has no key
    handlers, Escape closes the whole dialog.
    """
    toolkit = ctx.getServiceManager().createInstanceWithContext(
        "com.sun.star.awt.Toolkit", ctx
    )
    key_handler = ListEscapeKeyHandler(handler)
    has_key_handlers = hasattr(toolkit, "addKeyHandler")
    if has_key_handlers:
        toolkit.addKeyHandler(key_handler)
    try:
        dialog.execute()
    finally:
        if has_key_handlers:
            toolkit.removeKeyHandler(key_handler)
