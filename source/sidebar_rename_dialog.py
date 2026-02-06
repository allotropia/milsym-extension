# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import os
import unohelper

from unohelper import systemPathToFileUrl
from com.sun.star.awt import XDialogEventHandler
from com.sun.star.awt.Key import RETURN


class RenameDialog:
    def __init__(self, ctx, node, dir_path):
        self.ctx = ctx
        self.node = node
        self.favorites_dir_path = dir_path

    def run(self):
        try:
            dialog_provider = self.ctx.getServiceManager().createInstanceWithContext(
                "com.sun.star.awt.DialogProvider2", self.ctx
            )

            dialog_url = (
                "vnd.sun.star.extension://com.collabora.milsymbol/dialog/RenameDlg.xdl"
            )

            handler = RenameDlgHandler(self.ctx, self.node, self.favorites_dir_path)
            rename_dialog = dialog_provider.createDialogWithHandler(dialog_url, handler)

            name_textbox = rename_dialog.getControl("tbName")
            name_textbox.Text = self.node.getDisplayValue()

            rename_dialog.execute()
        except Exception as ex:
            print(f"Error creating rename dialog: {ex}")


class RenameDlgHandler(unohelper.Base, XDialogEventHandler):
    def __init__(self, ctx, node, dir_path):
        self.ctx = ctx
        self.node = node
        self.favorites_dir_path = dir_path
        self.symbol_name = node.getDisplayValue()

    def callHandlerMethod(self, dialog, eventObject, methodName):
        if methodName == "tbName":
            self.symbol_name = dialog.getControl(methodName).Text

            svg_path = self.get_path(self.symbol_name) + ".svg"
            invalid_button = self.symbol_name == "" or (
                self.symbol_name != self.node.getDisplayValue()
                and os.path.exists(svg_path)
            )

            dialog.getControl("btOk").getModel().State = 1 if invalid_button else 0
            return True
        elif methodName == "tbNameKeydown":
            if (
                dialog.getControl("btOk").getModel().State == 0
                and eventObject.KeyCode == RETURN
            ):
                self.set_symbol_name(dialog, self.symbol_name)
                return True
        elif methodName == "btOk":
            if dialog.getControl("btOk").getModel().State == 0:
                self.set_symbol_name(dialog, self.symbol_name)
                return True
        elif methodName == "btCancel":
            dialog.endExecute()
            return True

    def getSupportedMethodNames(self):
        return ("btOk", "btCancel", "tbName", "tbNameKeydown")

    def disposing(self, event):
        pass

    def get_path(self, symbol_name):
        category_name = self.node.getParent().getDisplayValue()
        category_path = os.path.join(self.favorites_dir_path, category_name)
        path = os.path.join(category_path, symbol_name)
        return path

    def set_symbol_name(self, dialog, new_symbol_name):
        old_symbol_name = self.node.getDisplayValue()

        if old_symbol_name == new_symbol_name:
            dialog.endExecute()
            return

        new_svg = self.get_path(new_symbol_name) + ".svg"
        new_json = self.get_path(new_symbol_name) + ".json"

        if os.path.exists(new_svg):
            dialog.endExecute()
            return

        old_svg = self.get_path(old_symbol_name) + ".svg"
        old_json = self.get_path(old_symbol_name) + ".json"

        os.rename(old_svg, new_svg)
        os.rename(old_json, new_json)

        new_svg_url = systemPathToFileUrl(new_svg)
        self.node.setNodeGraphicURL(new_svg_url)
        self.node.setDisplayValue(new_symbol_name)
        dialog.endExecute()
