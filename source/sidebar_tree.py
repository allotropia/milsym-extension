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
from com.sun.star.view import XSelectionChangeListener

class SidebarTree():

    def __init__(self, ctx, sidebar_panel):
        self.ctx = ctx
        self.sidebar_panel = sidebar_panel

    def create_svg_file_path(self, category_name, svg_name):
        category_dir_path = os.path.join(self.sidebar_panel.favorites_dir_path, category_name)
        os.makedirs(category_dir_path, exist_ok=True)
        symbol_full_path = os.path.join(category_dir_path, f"{svg_name}.svg")
        return symbol_full_path

    def generate_unique_name(self, parent_node, node_name, base_name="Symbol"):
        existing_names = set()
        count = parent_node.getChildCount()

        for i in range(count):
            child = parent_node.getChildAt(i)
            existing_names.add(child.getDisplayValue())

        if not node_name:
            n = 1
            while True:
                new_name = f"{base_name} {n}"
                if new_name not in existing_names:
                    return new_name
                n += 1

        m = 1
        while True:
            candidate = f"{node_name} ({m})"
            if candidate not in existing_names:
                return candidate
            m += 1

    def create_svg_file(self, symbol_full_path, svg_data):
        with open(symbol_full_path, 'w', encoding='utf-8') as preview_file:
            preview_file.write(svg_data)

    def serialize_svg_args(self, svg_args):
        result = {}
        for i, arg in enumerate(svg_args):
            if i == 0:
                result["sidc"] = arg
            else:
                result[arg.Name] = arg.Value

        return result

    def node_name_exists(self, parent_node, name: str) -> bool:
        child_count = parent_node.getChildCount()

        for i in range(child_count):
            child = parent_node.getChildAt(i)
            if child.getDisplayValue() == name:
                return True

        return False

    def get_order_index_from_json(self, json_path):
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get("order_index")

    def reorder_symbols(self, category_path):
        json_files = [
            name for name in os.listdir(category_path)
            if name.lower().endswith(".json")
        ]

        items = []
        for file_name in json_files:
            full_path = os.path.join(category_path, file_name)
            with open(full_path, "r", encoding="utf-8") as name:
                data = json.load(name)
                items.append((data.get("order_index"), file_name, data))

        items.sort(key=lambda x: x[0])

        for new_order, (_, file_name, data) in enumerate(items, start=1):
            data["order_index"] = new_order
            full_path = os.path.join(category_path, file_name)
            with open(full_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4)

    def create_node(self, root_node, tree_data_model, category_name,
                    svg_data, svg_args, is_editing, selected_node):
        try:
            if svg_data is None:
                return

            selected_node_category_name = None
            favorites_dir_path = self.sidebar_panel.favorites_dir_path

            if is_editing:
                # remove the selected node from the tree
                # and later append the newly edited node back to the tree
                parent_node = selected_node.getParent()
                child_idx = parent_node.getIndex(selected_node)
                parent_node.removeChildByIndex(child_idx)

                selected_node_category_name = parent_node.getDisplayValue()

                if parent_node.getChildCount() == 0:
                    root_node = parent_node.getParent()
                    parent_idx = root_node.getIndex(parent_node)
                    root_node.removeChildByIndex(parent_idx)

                node_name = selected_node.getDisplayValue()
                path = os.path.join(favorites_dir_path, selected_node_category_name, node_name)
                json_path = f"{path}.json"
                if os.path.exists(json_path):
                    order_index = self.get_order_index_from_json(json_path)

                # if the category name of the selected node differs from the target category,
                # not only remove the node from the tree, but also delete its associated
                # SVG and JSON files from the user's profile folder
                if selected_node_category_name != category_name:
                    count = 0
                    category_path = os.path.join(favorites_dir_path, category_name)
                    if os.path.exists(category_path):
                        count = sum(1 for name in os.listdir(category_path) if name.lower().endswith(".json"))
                    order_index = count + 1

                    svg_path = f"{path}.svg"
                    if os.path.exists(svg_path):
                        os.remove(svg_path)
                    if os.path.exists(json_path):
                        os.remove(json_path)

                    # Delete category folder when it's empty
                    category_dir = os.path.dirname(svg_path)
                    if os.path.exists(category_dir) and len(os.listdir(category_dir)) == 0:
                        os.rmdir(category_dir)

                    old_category_path = os.path.join(favorites_dir_path, selected_node_category_name)
                    if os.path.exists(old_category_path):
                        self.reorder_symbols(old_category_path)

            existing_category_node = None
            for i in range(root_node.getChildCount()):
                child = root_node.getChildAt(i)
                if child.getDisplayValue() == category_name:
                    existing_category_node = child
                    break

            if existing_category_node is None:
                existing_category_node = tree_data_model.createNode(category_name, True)
                root_node.appendChild(existing_category_node)

            if is_editing:
                node_name = selected_node.getDisplayValue()
                is_name_exists = self.node_name_exists(existing_category_node, node_name)
                if is_name_exists:
                    node_name = self.generate_unique_name(existing_category_node, node_name)
            else:
                node_name = self.generate_unique_name(existing_category_node, None)
                order_index = existing_category_node.getChildCount() + 1

            symbol_full_path = self.create_svg_file_path(category_name, node_name)
            self.create_svg_file(symbol_full_path, svg_data)
            svg_url = systemPathToFileUrl(symbol_full_path)

            symbol_child = tree_data_model.createNode(node_name, False)
            symbol_child.setNodeGraphicURL(svg_url)

            # create JSON file to store symbol parameters
            svg_params = self.serialize_svg_args(svg_args)
            svg_params["order_index"] = order_index

            json_file_name = f"{node_name}.json"
            json_file_path = os.path.join(os.path.dirname(symbol_full_path), json_file_name)
            with open(json_file_path, 'w', encoding='utf-8') as json_file:
                json.dump(svg_params, json_file, indent=4)

            symbol_child.DataValue = svg_args
            if is_editing:
                existing_category_node.insertChildByIndex(order_index-1, symbol_child)
            else:
                existing_category_node.appendChild(symbol_child)

            tree_control = self.sidebar_panel.tree_control
            tree_control.expandNode(existing_category_node)
            tree_control.select(symbol_child)
            tree_control.makeNodeVisible(symbol_child)

        except Exception as e:
            print("Error creating symbol node:", e)

class TreeSelectionChangeListener(unohelper.Base, XSelectionChangeListener):
    def __init__(self, sidebar_panel):
        self.sidebar_panel = sidebar_panel

    def selectionChanged(self, event):
        node = self.sidebar_panel.selected_node
        if node and event.Source.isEditing():
            node_name = node.getDisplayValue()
            if not node_name:
                node.setDisplayValue(self.sidebar_panel.selected_node_name)
                return

            parent_node = node.getParent()
            if not parent_node:
                return

            self.sidebar_panel.rename_symbol_files()

class TreeKeyListener(unohelper.Base, XKeyListener):
    def __init__(self, sidebar_panel, sidebar_tree):
        self.sidebar_panel = sidebar_panel
        self.sidebar_tree = sidebar_tree

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

            category_name = parent_node.getDisplayValue()
            favorites_dir_path = self.sidebar_panel.favorites_dir_path

            node_name = child_node.getDisplayValue()
            path = os.path.join(favorites_dir_path, category_name, node_name)
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

            else:
                category_path = os.path.join(favorites_dir_path, category_name)
                self.sidebar_tree.reorder_symbols(category_path)

            self.sidebar_panel.selected_node = None
            self.sidebar_panel.update_export_button_state()

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
    def __init__(self, ctx, sidebar_panel, sidebar_tree):
        self.ctx = ctx
        self.sidebar_panel = sidebar_panel
        self.sidebar_tree = sidebar_tree

    def mousePressed(self, event):
        try:
            x, y = event.X, event.Y
            tree_ctrl = event.Source
            node = tree_ctrl.getNodeForLocation(x, y)
            if node and node.getChildCount() == 0:
                if tree_ctrl.isEditing():
                    self.finalize_node_edit(tree_ctrl)
                else:
                    self.sidebar_panel.selected_node = node

                if event.ClickCount == 2:
                    self.sidebar_panel.selected_node_name = node.getDisplayValue()
                    tree_ctrl.startEditingAtNode(node)
            else:
                if tree_ctrl.isEditing():
                    self.finalize_node_edit(tree_ctrl)

        except Exception as e:
            print("Mouse pressed error:", e)

    def finalize_node_edit(self, tree_ctrl):
        selected_node = self.sidebar_panel.selected_node
        old_node_name = selected_node.getDisplayValue()

        tree_ctrl.cancelEditing()

        new_node_name = selected_node.getDisplayValue()
        if not new_node_name:
            selected_node.setDisplayValue(old_node_name)
        else:
            self.sidebar_panel.rename_symbol_files()

    def mouseMoved(self, event):
        """Handle mouse moved events"""
        pass

    def mouseDragged(self, event):
        """Handle mouse dragged events"""
        pass

    def mouseReleased(self, event):
        try:
            node = self.sidebar_panel.selected_node
            if (event.Buttons == MouseButton.RIGHT and node and node.getChildCount() == 0):
                tree_ctrl = event.Source
                tree_ctrl.select(node)

                rect = uno.createUnoStruct("com.sun.star.awt.Rectangle")
                rect.X = event.X
                rect.Y = event.Y
                rect.Width = 0
                rect.Height = 0

                peer = tree_ctrl.getPeer()
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
        popup.insertItem(2, "Rename symbol", MenuItemStyle.AUTOCHECK, 1)
        popup.insertItem(3, "Delete symbol", MenuItemStyle.AUTOCHECK, 2)

        popup.addMenuListener(PopupMenuHandler(self.ctx, self.sidebar_panel, self.sidebar_tree))

        return popup

class PopupMenuHandler(unohelper.Base, XMenuListener):
    def __init__(self, ctx, sidebar_panel, sidebar_tree):
        self.ctx = ctx
        self.sidebar_panel = sidebar_panel
        self.key_listener = TreeKeyListener(sidebar_panel, sidebar_tree)

    def itemSelected(self, event):
        menu_id = event.MenuId
        if menu_id == 1: # Edit
            node_value = self.sidebar_panel.selected_node.DataValue
            model = self.sidebar_panel.desktop.getCurrentComponent()
            open_symbol_dialog(self.ctx, model, None, self.sidebar_panel, None, node_value)
        elif menu_id == 2: # Delete
            self.sidebar_panel.rename_symbol()
        elif menu_id == 3: # Delete
            self.key_listener.delete_selected_node()

    def itemActivated(self, event): pass
    def itemHighlighted(self, event): pass
    def itemDeactivated(self, event): pass
