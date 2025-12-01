# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import os
import uno
import json
import unohelper

from utils import insertSvgGraphic
from symbol_dialog import open_symbol_dialog

from unohelper import systemPathToFileUrl, fileUrlToSystemPath
from com.sun.star.beans import PropertyValue
from com.sun.star.awt import SystemPointer, Key, MouseButton, MenuItemStyle, Rectangle
from com.sun.star.awt import XMouseListener, XMouseMotionListener, XMenuListener, XKeyListener
from com.sun.star.awt.tree import XMutableTreeDataModel, XMutableTreeNode, XTreeControl, XTreeNode

class SidebarTree():

    def __init__(self, ctx):
        self.ctx = ctx
        self.tree_control = None
        self.favorites_dir_path = None

    def set_tree_control(self, tree_ctrl):
        self.tree_control = tree_ctrl

    def set_favorites_dir_path(self, favorites_dir_path):
        self.favorites_dir_path = favorites_dir_path

    def create_svg_file_path(self, category_name, base_name="Symbol"):
        category_dir_path = os.path.join(self.favorites_dir_path, category_name)
        os.makedirs(category_dir_path, exist_ok=True)

        counter = 1
        while True:
            svg_name = f"{base_name} {counter}"
            symbol_full_path = os.path.join(category_dir_path, f"{svg_name}.svg")
            if not os.path.exists(symbol_full_path):
                return symbol_full_path, svg_name
            counter += 1

    def create_svg_file(self, symbol_full_path, svg_data):
        with open(symbol_full_path, 'w', encoding='utf-8') as preview_file:
            preview_file.write(svg_data)

    def serialize_svg_args(self, svg_args):
        result = []
        for i, arg in enumerate(svg_args):
            if i == 0:
                result.append(arg)
            else:
                result.append({arg.Name: arg.Value})

        return result

    def create_node(self, root_node, tree_data_model, category_name, svg_data, svg_args):
        try:
            existing_category_node = None
            for i in range(root_node.getChildCount()):
                child = root_node.getChildAt(i)
                if child.getDisplayValue() == category_name:
                    existing_category_node = child
                    break

            if existing_category_node is None:
                existing_category_node = tree_data_model.createNode(category_name, True)
                root_node.appendChild(existing_category_node)

            symbol_full_path, node_name = self.create_svg_file_path(category_name)
            symbol_child = tree_data_model.createNode(node_name, False)

            self.create_svg_file(symbol_full_path, svg_data)
            svg_url = systemPathToFileUrl(symbol_full_path)
            symbol_child.setNodeGraphicURL(svg_url)

            # create JSON file to store symbol parameters
            svg_params = self.serialize_svg_args(svg_args)
            json_file_name = f"{node_name}.json"
            json_file_path = os.path.join(os.path.dirname(symbol_full_path), json_file_name)
            with open(json_file_path, 'w', encoding='utf-8') as json_file:
                json.dump(svg_params, json_file, indent=4)

            symbol_child.DataValue = svg_args

            existing_category_node.appendChild(symbol_child)
            self.tree_control.expandNode(existing_category_node)
        except Exception as e:
            print("Error creating symbol node:", e)

class TreeKeyListener(unohelper.Base, XKeyListener):
    def __init__(self, sidebar_panel, favorites_dir_path):
        self.sidebar_panel = sidebar_panel
        self.favorites_dir_path = favorites_dir_path

    def delete_selected_node(self):
        try:
            child_node = self.sidebar_panel.selected_node
            if not child_node:
                return

            parent_node = child_node.getParent()
            if not parent_node:
                return

            idx = parent_node.getIndex(child_node)
            parent_node.removeChildByIndex(idx)

            node_name = child_node.getDisplayValue()
            category_name = parent_node.getDisplayValue()
            path = os.path.join(self.favorites_dir_path, category_name, node_name)
            svg_path = f"{path}.svg"
            json_path = f"{path}.json"
            if os.path.exists(svg_path):
                os.remove(svg_path)
            if os.path.exists(json_path):
                os.remove(json_path)

            child_count = parent_node.getChildCount()
            if child_count == 0:
                root_node = parent_node.getParent()
                if root_node:
                    parent_idx = root_node.getIndex(parent_node)
                    root_node.removeChildByIndex(parent_idx)

                category_dir = os.path.dirname(svg_path)
                if os.path.exists(category_dir) and len(os.listdir(category_dir)) == 0:
                    os.rmdir(category_dir)

            self.sidebar_panel.selected_node = None

        except Exception as e:
            print("Delete node error:", e)

    def keyPressed(self, event):
        try:
            if event.KeyCode == Key.DELETE:
                self.delete_selected_node()
        except Exception as e:
            print("Delete key error:", e)

    def keyReleased(self, event):
        pass

class TreeMouseListener(unohelper.Base, XMouseListener, XMouseMotionListener):
    def __init__(self, ctx, tree_control, sidebar_panel, favorites_dir_path):
        self.ctx = ctx
        self.tree = tree_control
        self.sidebar_panel = sidebar_panel
        self.favorites_dir_path = favorites_dir_path
        self.model = tree_control.getModel()
        self.drop_allowed = False
        self.svg_data = None

        self.pointer = self.ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.awt.Pointer", self.ctx)

    def svg_data_from_url(self, file_path):
        with open(file_path, 'r', encoding='utf-8') as f:
            svg_data = f.read()
        return svg_data

    def mousePressed(self, event):
        try:
            x, y = event.X, event.Y
            node = self.tree.getNodeForLocation(x, y)
            if node and node.getChildCount() == 0 and not self.drop_allowed:
                self.sidebar_panel.selected_node = node
                svg_url = node.getNodeGraphicURL()
                file_path = fileUrlToSystemPath(svg_url)
                self.svg_data = self.svg_data_from_url(file_path)
            else:
                self.sidebar_panel.selected_node = None

        except Exception as e:
            print("Mouse pressed error:", e)

    def mouseMoved(self, event):
        try:
            # change the mouse pointer to provide visual feedback during DnD
            if self.pointer.getType() != SystemPointer.ARROW:
                self.pointer.setType(SystemPointer.ARROW)
                self.tree.getPeer().setPointer(self.pointer)
        except Exception as e:
            print("Mouse moved error:", e)

    def mouseDragged(self, event):
        try:
            # check if the mouse is in the allowed drop area
            # event.X and event.Y are relative to the TreeControl
            # the values here (-30, -80) represent thresholds
            # that determine how far the mouse has moved away from the TreeControl
            if self.sidebar_panel.selected_node and event.Buttons == MouseButton.LEFT:
                if event.X <= -30 and event.Y > -80:
                    self.drop_allowed = True
                else:
                    self.drop_allowed = False

                # change the mouse pointer to provide visual feedback during DnD
                if self.pointer.getType() != SystemPointer.MOVEDATA:
                    self.pointer.setType(SystemPointer.MOVEDATA)
                    self.tree.getPeer().setPointer(self.pointer)
        except Exception as e:
            print("Mouse dragged error:", e)

    def mouseReleased(self, event):
        try:
            node = self.sidebar_panel.selected_node
            if node and self.drop_allowed and self.svg_data:
                model = self.sidebar_panel.desktop.getCurrentComponent()
                params = node.DataValue
                insertSvgGraphic(self.ctx, model, self.svg_data, params, None, 3)
                self.drop_allowed = False

            if (event.Buttons == MouseButton.RIGHT and node and node.getChildCount() == 0):
                rect = uno.createUnoStruct("com.sun.star.awt.Rectangle")
                rect.X = event.X
                rect.Y = event.Y
                rect.Width = 0
                rect.Height = 0

                peer = self.tree.getPeer()
                popup = self._create_popup_menu()
                popup.execute(peer, rect, 0)

        except Exception as e:
            print("Mouse released error:", e)

    def mouseEntered(self, event):
        """Handle mouse entered events"""
        pass

    def mouseExited(self, event):
        """Handle mouse exited events"""
        pass

    def _create_popup_menu(self):
        sm = self.ctx.getServiceManager()
        popup = sm.createInstance("com.sun.star.awt.PopupMenu")

        popup.insertItem(1, "Edit symbol", MenuItemStyle.AUTOCHECK, 0)
        popup.insertItem(2, "Delete symbol", MenuItemStyle.AUTOCHECK, 1)

        popup.addMenuListener(PopupMenuHandler(self.ctx, self.sidebar_panel, self.favorites_dir_path))

        return popup

class PopupMenuHandler(unohelper.Base, XMenuListener):
    def __init__(self, ctx, sidebar_panel, favorites_dir_path):
        self.ctx = ctx
        self.sidebar_panel = sidebar_panel
        self.key_listener = TreeKeyListener(sidebar_panel, favorites_dir_path)

    def itemSelected(self, event):
        menu_id = event.MenuId
        if menu_id == 1: # Edit
            model = self.sidebar_panel.desktop.getCurrentComponent()
            open_symbol_dialog(self.ctx, model, None, self.sidebar_panel, None)
        elif menu_id == 2: # Delete
            self.key_listener.delete_selected_node()

    def itemActivated(self, event): pass
    def itemHighlighted(self, event): pass
    def itemDeactivated(self, event): pass
