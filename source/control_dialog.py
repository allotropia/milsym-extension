# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import json
import os
import sys
import uno
import unohelper
import shutil

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from com.sun.star.awt import (
    KeyModifier,
    XDialogEventHandler,
    XTopWindowListener,
    XMouseListener,
    XKeyListener,
    Key,
    XWindowListener,
)
from com.sun.star.awt import MouseButton
from com.sun.star.view.SelectionType import (
    SINGLE as SELECTION_TYPE_SINGLE,
    MULTI as SELECTION_TYPE_MULTI,
)
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.datatransfer.dnd import XDragGestureListener, XDropTargetListener
from com.sun.star.datatransfer.dnd import XDragSourceListener
from com.sun.star.datatransfer.dnd.DNDConstants import ACTION_MOVE
from com.sun.star.datatransfer import XTransferable, DataFlavor
from com.sun.star.beans import NamedValue
from com.sun.star.document import XUndoAction, XUndoManager
from utils import (
    extractGraphicAttributes,
    getExtensionBasePath,
    generate_icon_svg,
    insertGraphicAttributes,
)
from unohelper import systemPathToFileUrl
import tempfile


class ControlDlgHandler(
    unohelper.Base, XDialogEventHandler, XTopWindowListener, XWindowListener
):
    buttons = ["addShape", "removeShape", "editShape"]

    def __init__(self, dialog, x_context, model):
        self.dialog = dialog
        self.x_context = x_context
        self.tree_control = None
        self._populate_tree_on_show = True
        self._parent_before_add = None
        self._temp_dir: str | None = None
        self._clipboard = None
        self._node_to_tree_item_map = {}
        self._syncing_selection = False
        self._syncing_from_tree = False
        self._is_dragging = False
        factory = self.x_context.getServiceManager().createInstanceWithContext(
            "com.sun.star.script.provider.MasterScriptProviderFactory", self.x_context
        )
        provider = factory.createScriptProvider(model)
        self.script = provider.getScript(
            "vnd.sun.star.script:milsymbol.milsymbol.js?language=JavaScript&location=user:uno_packages/"
            + getExtensionBasePath(self.x_context)
        )

    def callHandlerMethod(self, dialog, eventObject, methodName):
        if methodName == "OnAdd":
            if self.get_controller().get_diagram() is not None:
                self.get_controller().remove_selection_listener()
                # Store current selected shape before adding
                self._store_selection_before_add()

                # Get undo manager and create undo action
                undo_manager = self._get_undo_manager()
                parent_tree_item = self._parent_before_add
                added_shape = None

                # Lock undo manager to prevent Writer from creating internal undo records
                if undo_manager:
                    undo_manager.lock()
                try:
                    self.get_controller().get_diagram().add_shape()
                    self.get_controller().get_diagram().refresh_diagram()

                    added_shape = self._find_newly_added_shape(parent_tree_item)
                finally:
                    # Always unlock the undo manager
                    if undo_manager:
                        undo_manager.unlock()

                # Create and register undo action if we have an undo manager
                if undo_manager and added_shape and parent_tree_item:
                    try:
                        undo_action = AddShapeUndoAction(
                            self, added_shape, parent_tree_item
                        )
                        undo_manager.addUndoAction(undo_action)
                    except Exception as e:
                        print(f"Failed to register undo action: {e}")

                # Refresh tree after adding shape
                self.refresh_tree()
                self._select_newly_added_child()
                self.get_controller().add_selection_listener()
                if self.tree_control is not None:
                    self.tree_control.setFocus()
            return True
        elif methodName == "OnRemove":
            self.remove_selected_shape()
            return True
        elif methodName == "OnEdit":
            selected_shape = self.get_controller().get_diagram().get_last_shape()
            if selected_shape is None:
                return True

            original_attributes = extractGraphicAttributes(selected_shape)
            self.dialog.execute_properties_dialog()
            edited_attributes = extractGraphicAttributes(selected_shape)

            if original_attributes != edited_attributes:
                undo_manager = self._get_undo_manager()
                if undo_manager:
                    try:
                        undo_action = EditShapeUndoAction(
                            self, selected_shape, original_attributes, edited_attributes
                        )
                        undo_manager.addUndoAction(undo_action)
                    except Exception as e:
                        print(f"Failed to register edit undo action: {e}")

            self.get_controller().get_diagram().refresh_diagram()
            self.refresh_tree()
            if self.tree_control is not None:
                self.tree_control.setFocus()
            return True
        elif methodName == "OnDragOrbatChange":
            if self.is_drag_orbat_enabled():
                # Leave the currently active group if any
                diagram = self.get_controller().get_diagram()
                current_group_shape = diagram.get_group_shape() if diagram else None
                if current_group_shape:
                    current_group_shape.leaveGroup()
                    # Select the group shape that we just left
                    self.get_controller().set_selected_shape(current_group_shape)
            else:
                # Re-enter the group if a group shape is selected
                selected_shape = self.get_controller().get_selected_shape()
                if selected_shape and selected_shape.supportsService(
                    "com.sun.star.drawing.GroupShape"
                ):
                    selected_shape.enterGroup()
                    # Also ensure the currently selected items in the tree are selected in the document
                    selected_tree_items = self._get_selected_tree_items()
                    shapes_to_select = []
                    for tree_item in selected_tree_items:
                        shape = tree_item.get_rectangle_shape()
                        if shape:
                            shapes_to_select.append(shape)
                    if shapes_to_select:
                        self._select_shapes_in_document(shapes_to_select)

            return True
        else:
            return False

    def getSupportedMethodNames(self):
        return self.buttons

    def disposing(self, event):
        pass

    def buttonStateHandler(self, methodName):
        pass

    def _update_button_states(self):
        """Update button enabled states based on current selection.

        The Edit button is disabled when multiple items are selected,
        since editing only works with a single shape.
        """
        try:
            dialog = self.get_gui()._x_control_dialog
            if dialog is None:
                return

            edit_button = dialog.getControl("BtnEdit")
            if edit_button is None:
                return

            selected_items = self._get_selected_tree_items()
            # Enable Edit button only when 0 or 1 item is selected
            should_enable_edit = len(selected_items) <= 1
            edit_button.setEnable(should_enable_edit)
        except Exception as e:
            print(f"Error updating button states: {e}")

    def get_gui(self):
        """Get GUI reference"""
        return self.dialog

    def get_controller(self):
        """Get controller reference"""
        return self.dialog.get_controller()

    def get_drag_orbat_checkbox(self):
        """Get reference to the Drag Orbat checkbox"""
        return self.get_gui()._x_control_dialog.getControl("ChkDragOrbat")

    def is_drag_orbat_enabled(self):
        """Check if the Drag Orbat checkbox is checked"""
        return self.get_drag_orbat_checkbox().getState() == 1

    def paste_to_selected_item(self):
        """Paste clipboard contents to currently selected item.

        All copied items become children of the selected target node.
        After paste, all newly pasted items are selected.
        """
        if self._clipboard is None:
            return

        # Normalize clipboard to list format (for backward compatibility)
        clipboard_items = self._clipboard
        if not isinstance(clipboard_items, list):
            clipboard_items = [clipboard_items]

        if not clipboard_items:
            return

        try:
            if self.tree_control is None:
                return

            selection = self.tree_control.getSelection()
            if selection is None:
                return

            selected_node = None
            if hasattr(selection, "getDisplayValue"):
                # Single node selected
                selected_node = selection
            elif hasattr(selection, "__len__") and len(selection) > 0:
                # Sequence of nodes - get first one
                selected_node = selection[0]
            else:
                return

            if selected_node is None or not hasattr(selected_node, "getDisplayValue"):
                return

            node_name = selected_node.getDisplayValue()
            target_tree_item = self._node_to_tree_item_map.get(node_name)
            if target_tree_item is None:
                return

            controller = self.get_controller()
            diagram = controller.get_diagram()

            if diagram:
                controller.remove_selection_listener()

                # Get undo manager and lock to prevent Writer internal undo records
                undo_manager = self._get_undo_manager()
                if undo_manager:
                    undo_manager.lock()

                # Paste all clipboard items and track pasted shapes
                pasted_shapes = []
                any_success = False

                try:
                    for clipboard_item in clipboard_items:
                        success = diagram.paste_subtree(
                            target_tree_item, clipboard_item, self.script
                        )
                        if success:
                            any_success = True
                            # Find the newly added shape for this paste
                            pasted_shape = self._find_newly_added_shape(
                                target_tree_item
                            )
                            if pasted_shape:
                                pasted_shapes.append(pasted_shape)

                    if any_success:
                        diagram.refresh_diagram()
                finally:
                    if undo_manager:
                        undo_manager.unlock()

                if any_success:
                    self.refresh_tree()

                    # Register undo action
                    if undo_manager and pasted_shapes:
                        try:
                            undo_action = PasteShapeUndoAction(
                                self,
                                clipboard_items,
                                target_tree_item,
                                pasted_shapes,
                            )
                            undo_manager.addUndoAction(undo_action)
                        except Exception as e:
                            print(f"Failed to register paste undo action: {e}")

                    # Select all pasted shapes in the tree
                    if pasted_shapes:
                        self._select_tree_nodes_for_shapes(pasted_shapes)

                controller.add_selection_listener()

        except Exception as ex:
            print(f"Error pasting: {ex}")

    def copy_selected_item(self):
        """Copy the currently selected item(s) and their subtrees.

        When multiple items are selected, copies all of them.
        If both a parent and child are selected, only the parent is copied
        (since copying the parent includes the child in its subtree).
        """
        try:
            if self.tree_control is None:
                return

            # Get all selected tree items
            tree_items = self._get_selected_tree_items()
            if not tree_items:
                return

            # Filter out descendants (if parent and child both selected, keep only parent)
            tree_items = self._filter_out_descendants(tree_items)

            # Serialize each item and store as a list
            clipboard_items = []
            for tree_item in tree_items:
                serialized = self._serialize_tree_item(tree_item)
                if serialized is not None:
                    clipboard_items.append(serialized)

            if clipboard_items:
                self._clipboard = clipboard_items

        except Exception as ex:
            print(f"Error copying item(s): {ex}")

    def _get_selected_tree_item(self):
        """Get the tree item for the currently selected shape"""
        try:
            controller = self.get_controller()
            selection = controller._x_controller.getSelection()
            if selection and selection.getCount() > 0:
                selected_shape = selection.getByIndex(0)
                # Find the tree item for this shape
                for tree_item in self._node_to_tree_item_map.values():
                    if tree_item.get_rectangle_shape() == selected_shape:
                        return tree_item
        except Exception as e:
            print(f"Error getting selected tree item: {e}")
        return None

    def _get_selected_tree_items(self):
        """Get all tree items for currently selected nodes (supports multi-selection)"""
        tree_items = []
        try:
            if self.tree_control is None:
                return tree_items

            selection = self.tree_control.getSelection()
            if selection is None:
                return tree_items

            # Handle both single node and sequence of nodes
            nodes = []
            if hasattr(selection, "getDisplayValue"):
                # Single node selected
                nodes = [selection]
            elif hasattr(selection, "__len__"):
                # Sequence of nodes
                nodes = list(selection)
            elif hasattr(selection, "__iter__"):
                # Iterable
                nodes = list(selection)

            for node in nodes:
                if node is not None and hasattr(node, "getDisplayValue"):
                    node_name = node.getDisplayValue()
                    tree_item = self._node_to_tree_item_map.get(node_name)
                    if tree_item is not None:
                        tree_items.append(tree_item)

        except Exception as e:
            print(f"Error getting selected tree items: {e}")

        return tree_items

    def _is_node_selected(self, node):
        """Check if a tree node is currently selected"""
        try:
            if self.tree_control is None or node is None:
                return False

            selection = self.tree_control.getSelection()
            if selection is None:
                return False

            if hasattr(selection, "getDisplayValue"):
                # Single node - direct comparison
                return selection == node
            elif hasattr(selection, "__len__"):
                # Sequence of nodes
                return node in selection
            elif hasattr(selection, "__iter__"):
                # Iterable
                return node in list(selection)

            return False
        except Exception:
            return False

    def _filter_out_descendants(self, tree_items):
        """Filter out items that are descendants of other items in the list.

        If both a parent and child are selected, only keep the parent since
        removing/moving the parent affects children automatically.
        """
        if len(tree_items) <= 1:
            return tree_items

        filtered = []
        for item in tree_items:
            is_descendant = False
            for other in tree_items:
                if item != other and self._is_descendant_of(item, other):
                    is_descendant = True
                    break
            if not is_descendant:
                filtered.append(item)

        return filtered

    def _is_descendant_of(self, item, potential_ancestor):
        """Check if item is a descendant of potential_ancestor"""
        current = item.get_dad()
        while current is not None:
            if current == potential_ancestor:
                return True
            current = current.get_dad()
        return False

    def _serialize_tree_item(self, tree_item):
        """Recursively serialize a tree item and all its descendants to ClipboardItem"""
        if tree_item is None:
            return None

        shape = tree_item.get_rectangle_shape()
        if shape is None:
            return None

        attributes = extractGraphicAttributes(shape)

        # Recursively serialize all children
        children = []
        child_item = tree_item.get_first_child()
        while child_item is not None:
            child_clipboard = self._serialize_tree_item(child_item)
            if child_clipboard is not None:
                children.append(child_clipboard)
            child_item = child_item.get_first_sibling()

        return ClipboardItem(attributes, children)

    def _serialize_tree_item_only(self, tree_item):
        """Serialize only the tree item itself without its children.

        Used for undo/redo of remove operations where children are preserved
        (moved to grandparent) rather than deleted.
        """
        if tree_item is None:
            return None

        shape = tree_item.get_rectangle_shape()
        if shape is None:
            return None

        attributes = extractGraphicAttributes(shape)
        return ClipboardItem(attributes, [])

    def handle_undo(self):
        """Trigger undo via the document's undo manager"""
        undo_manager = self._get_undo_manager()
        if undo_manager and undo_manager.isUndoPossible():
            undo_manager.undo()

    def handle_redo(self):
        """Trigger redo via the document's undo manager"""
        undo_manager = self._get_undo_manager()
        if undo_manager and undo_manager.isRedoPossible():
            undo_manager.redo()

    def remove_selected_shape(self):
        """Remove the currently selected shape(s) from the diagram"""
        if self.get_controller().get_diagram() is None:
            return

        self.get_controller().remove_selection_listener()

        try:
            # Get all selected tree items
            selected_items = self._get_selected_tree_items()
            if not selected_items:
                self.get_controller().add_selection_listener()
                return

            # Collect data for undo before removal
            undo_manager = self._get_undo_manager()
            removal_data = []  # List of (serialized_data, parent_tree_item, child_shapes)

            for tree_item in selected_items:
                parent = tree_item.get_dad()
                serialized = self._serialize_tree_item_only(tree_item)

                # Collect child shapes that will be re-parented during removal
                child_shapes = []
                child_item = tree_item.get_first_child()
                while child_item is not None:
                    child_shape = child_item.get_rectangle_shape()
                    if child_shape:
                        child_shapes.append(child_shape)
                    child_item = child_item.get_first_sibling()

                if serialized and parent:
                    removal_data.append((serialized, parent, child_shapes))

            # Remove shapes (order: process items with deeper levels first to avoid
            # invalidating parent references)
            selected_items.sort(key=lambda x: -x.get_level())

            for tree_item in selected_items:
                shape = tree_item.get_rectangle_shape()
                if shape:
                    self.get_controller().set_selected_shape(shape)
                    self.get_controller().get_diagram().remove_shape()

            self.get_controller().get_diagram().refresh_diagram()
            self.refresh_tree()

            if undo_manager and removal_data:
                try:
                    undo_action = RemoveShapeUndoAction(self, removal_data)
                    undo_manager.addUndoAction(undo_action)
                except Exception as e:
                    print(f"Failed to register undo action: {e}")

        finally:
            self.get_controller().add_selection_listener()
            if self.tree_control is not None:
                self.tree_control.setFocus()

    # XTopWindowListener methods
    def windowClosing(self, event):
        """Handle window closing event"""
        if event.Source == self.get_gui()._x_control_dialog:
            self.get_gui().set_visible_control_dialog(False)

        self.cleanup()

    def windowOpened(self, event):
        """Handle window opened event"""
        if event.Source == self.get_gui()._x_control_dialog:
            # Initialize tree control when dialog opens
            self._init_tree_control()
            # Populate tree with current diagram structure
            if self._populate_tree_on_show:
                self.populate_tree()
                self._populate_tree_on_show = False
            # Add window resize listener
            self._add_resize_listener()

    def _add_resize_listener(self):
        """Add window resize listener to handle dialog resizing"""
        try:
            dialog = self.get_gui()._x_control_dialog
            if dialog is not None:
                dialog.addWindowListener(self)
        except Exception as e:
            print(f"Error adding resize listener: {e}")

    def windowClosed(self, event):
        """Handle window closed event"""
        pass

    def windowMinimized(self, event):
        """Handle window minimized event"""
        pass

    def windowNormalized(self, event):
        """Handle window normalized event"""
        pass

    def windowActivated(self, event):
        """Handle window activated event"""
        pass

    def windowDeactivated(self, event):
        """Handle window deactivated event"""
        pass

    # XWindowListener methods for resize handling
    def windowResized(self, event):
        """Handle window resize event"""
        try:
            if event.Source == self.get_gui()._x_control_dialog:
                self._resize_controls(event)
        except Exception as e:
            print(f"Error handling window resize: {e}")

    def windowMoved(self, event):
        """Handle window moved event"""
        pass

    def windowShown(self, event):
        """Handle window shown event"""
        pass

    def windowHidden(self, event):
        """Handle window hidden event"""
        pass

    def _resize_controls(self, event):
        """Resize controls based on new dialog size"""
        try:
            # Calculate margins (based on original layout)
            margin_left = 3
            margin_right = 3
            margin_top = 21
            margin_bottom = 3

            # Calculate new tree control size
            new_tree_width = (
                event.Width
                - event.LeftInset
                - event.RightInset
                - margin_left
                - margin_right
            )
            new_tree_height = (
                event.Height
                - event.TopInset
                - event.BottomInset
                - margin_top
                - margin_bottom
            )

            # Set new tree control size
            tree_model = self.tree_control.getModel()
            tree_model.Width = new_tree_width
            tree_model.Height = new_tree_height

        except Exception as e:
            print(f"Error resizing controls: {e}")

    def _init_tree_control(self):
        """Initialize the tree control and set up listeners"""
        try:
            dialog = self.get_gui()._x_control_dialog
            self.tree_control = dialog.getControl("OrbatTree")

            # Add mouse listener for tree click handling
            tree_mouse_handler = TreeMouseHandler(self)
            self.tree_control.addMouseListener(tree_mouse_handler)

            # Add key listener to detect keyboard navigation
            tree_key_handler = TreeKeyHandler(self)
            self.tree_control.addKeyListener(tree_key_handler)

            # Set up selection listener for bidirectional selection
            self._setup_selection_listener()
            # Enable drag & drop functionality
            self._setup_drag_and_drop()

            # Configure tree control properties
            tree_model = self.tree_control.getModel()
            tree_model.setPropertyValue("SelectionType", SELECTION_TYPE_MULTI)
            tree_model.setPropertyValue("RootDisplayed", True)
            tree_model.setPropertyValue("ShowsHandles", True)
            tree_model.setPropertyValue("ShowsRootHandles", True)
            tree_model.setPropertyValue("Editable", False)
        except Exception as e:
            print(f"Error initializing tree control: {e}")

    def populate_tree(self):
        """Populate tree with current diagram structure"""
        try:
            if self.tree_control is None:
                print("Tree control not available")
                return

            # Get the current diagram
            controller = self.get_controller()
            diagram = controller.get_diagram()

            if diagram is None:
                print("No diagram available")
                return

            service_manager = self.x_context.getServiceManager()
            data_model = service_manager.createInstanceWithContext(
                "com.sun.star.awt.tree.MutableTreeDataModel", self.x_context
            )

            if data_model is None:
                print("Could not create tree data model")
                return

            # Create root node with proper name
            root_node_name = "Root"  # Default fallback
            # Try to get proper root name if diagram tree is available
            try:
                # Get the current diagram
                controller = self.get_controller()
                diagram = controller.get_diagram()

                if diagram is None:
                    print("No diagram available")
                    return
                temp_diagram_tree = diagram.get_diagram_tree()
                if temp_diagram_tree:
                    temp_root_item = temp_diagram_tree.get_root_item()
                    if temp_root_item:
                        root_node_name = self._get_tree_node_display_name(
                            temp_root_item, 1
                        )
            except:
                pass

            root_node = data_model.createNode(root_node_name, True)
            data_model.setRoot(root_node)
            tree_model = self.tree_control.getModel()

            # Add icon preview before setting data_model
            diagram_tree = diagram.get_diagram_tree()
            if diagram_tree is not None:
                root_item = diagram_tree.get_root_item()
                if root_item is not None:
                    shape = root_item.get_rectangle_shape()
                    root_name = self._get_tree_node_display_name(root_item, 1)
                    self._add_icon_preview_tree_node(shape, root_name, root_node)

            tree_model.setPropertyValue("DataModel", data_model)

            diagram_tree = diagram.get_diagram_tree()
            if diagram_tree is not None:
                root_item = diagram_tree.get_root_item()
                if root_item is not None:
                    # Get the root name (should match what we used during creation)
                    root_name = self._get_tree_node_display_name(root_item, 1)

                    # Store root item in mapping for selection
                    self._node_to_tree_item_map = {}
                    self._node_to_tree_item_map.clear()
                    self._node_to_tree_item_map[root_name] = root_item

                    # Populate children of root item
                    self._populate_tree_children(data_model, root_node, root_item)
                else:
                    print("No root item found in diagram tree")
            else:
                print("No diagram tree available or invalid structure")

            # Expand all nodes in the tree to show full structure
            try:
                self._expand_all_nodes(root_node)
            except:
                pass

        except Exception as e:
            print(f"Error populating tree: {e}")

    def _populate_tree_children(self, data_model, parent_node, tree_item):
        """Populate children of a tree node from organization chart tree items"""
        try:
            if tree_item is None:
                return

            # Add children
            child_count = 0
            if tree_item.get_first_child() is not None:
                child_item = tree_item.get_first_child()
                child_num = 1
                while child_item is not None:
                    child_name = self._get_tree_node_display_name(child_item, child_num)
                    child_name = self._make_unique_display_name(child_name)
                    self._populate_tree_node(
                        data_model, parent_node, child_item, child_name
                    )

                    # Move to next sibling
                    child_item = child_item.get_first_sibling()
                    child_num += 1
                    child_count += 1

        except Exception as e:
            print(f"Error populating tree children: {e}")

    def _make_unique_display_name(self, base_name):
        """Ensure display name is unique by appending suffix if needed"""
        if not hasattr(self, "_node_to_tree_item_map"):
            return base_name

        if base_name not in self._node_to_tree_item_map:
            return base_name

        # Name exists, append suffix to make unique
        counter = 2
        while f"{base_name} #{counter}" in self._node_to_tree_item_map:
            counter += 1
        return f"{base_name} #{counter}"

    def _populate_tree_node(self, data_model, parent_node, tree_item, display_name):
        """Recursively populate tree nodes from organization chart tree items"""
        try:
            if tree_item is None:
                return

            # Create node for this tree item
            node = data_model.createNode(display_name, False)  # Start with no children
            parent_node.appendChild(node)

            # Store reference to tree item in a class attribute for later selection
            if not hasattr(self, "_node_to_tree_item_map"):
                self._node_to_tree_item_map = {}
            self._node_to_tree_item_map[display_name] = tree_item

            shape = tree_item.get_rectangle_shape()
            self._add_icon_preview_tree_node(shape, display_name, node)

            # Add children
            child_count = 0
            if tree_item.get_first_child() is not None:
                child_item = tree_item.get_first_child()
                child_num = 1
                while child_item is not None:
                    child_name = self._get_tree_node_display_name(child_item, child_num)
                    child_name = self._make_unique_display_name(child_name)
                    self._populate_tree_node(data_model, node, child_item, child_name)

                    # Move to next sibling
                    child_item = child_item.get_first_sibling()
                    child_num += 1
                    child_count += 1

            # Update node to show it has children if any were added
            if child_count > 0:
                node.setHasChildrenOnDemand(True)

        except Exception as e:
            print(f"Error populating tree node: {e}")

    def refresh_tree(self):
        """Refresh the tree structure"""
        try:
            if self.tree_control is not None:
                # Clear the node mapping before repopulating
                self._node_to_tree_item_map.clear()
                self.populate_tree()
        except Exception as e:
            print(f"Error refreshing tree: {e}")

    def _expand_all_nodes(self, node):
        """Recursively expand all nodes in the tree"""
        try:
            if node is not None:
                # Expand current node
                self.tree_control.expandNode(node)

                # Recursively expand all children
                for i in range(node.getChildCount()):
                    child_node = node.getChildAt(i)
                    self._expand_all_nodes(child_node)
        except Exception as e:
            print(f"Error expanding node: {e}")

    def _setup_drag_and_drop(self):
        """Setup drag & drop functionality for the tree control"""
        try:
            if self.tree_control is not None:
                # Create drag & drop handlers
                drag_handler = TreeDragHandler(self)
                drop_handler = TreeDropHandler(self)

                # Store handlers to prevent garbage collection
                self._drag_handler = drag_handler
                self._drop_handler = drop_handler
                peer = self.tree_control.getPeer()

                # Try to set up drag source
                toolkit = None
                try:
                    toolkit = (
                        self.x_context.getServiceManager().createInstanceWithContext(
                            "com.sun.star.awt.Toolkit", self.x_context
                        )
                    )
                    drag_gesture_recognizer = toolkit.getDragGestureRecognizer(peer)
                    drag_gesture_recognizer.addDragGestureListener(drag_handler)
                except Exception as e:
                    print(f"Could not set up drag source: {e}")

                # Try to set up drop target
                if toolkit is not None:
                    try:
                        drop_target = toolkit.getDropTarget(peer)
                        drop_target.addDropTargetListener(drop_handler)
                        drop_target.setActive(True)
                    except Exception as e:
                        print(f"Could not set up drop target: {e}")

        except Exception as e:
            print(f"Error setting up drag & drop: {e}")

    def _add_icon_preview_tree_node(self, shape, name, node):
        try:
            if shape is None:
                node.setNodeGraphicURL(
                    "vnd.sun.star.extension://com.collabora.milsymbol/img/orbat_base_small.svg"
                )
                return

            attributes = extractGraphicAttributes(shape)
            if attributes and attributes.get("MilSymCode"):
                svg_data = generate_icon_svg(self.script, attributes, 14.0)
                if svg_data:
                    svg_url = self._save_svg_to_temp_and_get_url(svg_data, name)
                    if svg_url:
                        node.setNodeGraphicURL(svg_url)
            else:
                node.setNodeGraphicURL(
                    "vnd.sun.star.extension://com.collabora.milsymbol/img/orbat_base_small.svg"
                )
        except Exception as e:
            print(f"Error adding root node icon: {e}")

    def _save_svg_to_temp_and_get_url(self, svg_data, node_identifier):
        """Save SVG to temp file and return file URL for setNodeGraphicURL()

        Args:
            svg_data: SVG content as string
            node_identifier: Unique identifier for the node (used in filename)

        Returns:
            file:// URL string or None if save fails
        """
        try:
            if not svg_data:
                return None

            safe_name = "".join(
                c if c.isalnum() or c in ("-", "_") else "_" for c in node_identifier
            )
            safe_name = safe_name[:50]

            if not hasattr(self, "_temp_dir") or self._temp_dir is None:
                self._temp_dir = tempfile.mkdtemp(prefix="orbat_icons_")

            temp_path = os.path.join(self._temp_dir, f"{safe_name}.svg")

            with open(temp_path, "w", encoding="utf-8") as f:
                f.write(svg_data)

            file_url = systemPathToFileUrl(temp_path)
            return file_url

        except Exception as e:
            print(f"Error saving SVG to temp file: {e}")
            return None

    def _get_tree_node_display_name(self, tree_item, item_number):
        """Get a meaningful display name for a tree node"""
        try:
            # Try to get shape information
            shape = tree_item.get_rectangle_shape()
            if shape is not None:
                # Try to get MilSym information
                milsym_code = ""
                attributes = extractGraphicAttributes(shape)
                if attributes and attributes.get("MilSymCode"):
                    milsym_code = attributes.get("MilSymCode").rstrip("0")

                # Try to get MilSym unique designation, headquarters and higher formation
                milsym_hier = ""
                if attributes and attributes.get("MilSymSpecialHeadquarters"):
                    milsym_hier += attributes.get("MilSymSpecialHeadquarters") + " "
                if attributes and attributes.get("MilSymHigherFormation"):
                    milsym_hier += attributes.get("MilSymHigherFormation") + " "
                if attributes and attributes.get("MilSymUniqueDesignation"):
                    milsym_hier += attributes.get("MilSymUniqueDesignation")

                # Build display name
                if milsym_code.strip() and milsym_hier.strip():
                    return f"Unit {milsym_hier} ({milsym_code})"
                elif milsym_code.strip():
                    return f"Unit ({milsym_code})"
                else:
                    return "Placeholder unit"

            return f"Unit {item_number}"

        except Exception as e:
            print(f"Error getting tree node display name: {e}")
            return f"Item {item_number}"

    def handle_tree_selection(self, selected_node):
        """Handle tree node selection - syncs tree selection to document

        Args:
            selected_node: Single tree node (XTreeNode)
        """
        try:
            if selected_node is None:
                return

            # Get the node name - handle single node only
            if not hasattr(selected_node, "getDisplayValue"):
                return

            node_name = selected_node.getDisplayValue()
            tree_item = self._node_to_tree_item_map.get(node_name)
            if tree_item:
                shape = tree_item.get_rectangle_shape()
                if shape:
                    self._syncing_selection = True
                    self._syncing_from_tree = True
                    try:
                        controller = self.get_controller()
                        try:
                            controller.set_selected_shape(shape)
                        except:
                            try:
                                controller._x_controller.select(shape)
                            except:
                                pass
                    finally:
                        self._syncing_selection = False
                        self._syncing_from_tree = False

        except Exception as e:
            print(f"Error handling tree selection: {e}")
        finally:
            self._update_button_states()

    def sync_all_selected_shapes_to_document(self):
        """Sync all selected tree items' shapes to document selection"""
        try:
            self._syncing_selection = True
            self._syncing_from_tree = True

            tree_items = self._get_selected_tree_items()
            if not tree_items:
                return

            shapes = []
            for tree_item in tree_items:
                shape = tree_item.get_rectangle_shape()
                if shape:
                    shapes.append(shape)

            if shapes:
                self._select_shapes_in_document(shapes)
        except Exception as e:
            print(f"Error syncing selected shapes: {e}")
        finally:
            self._syncing_selection = False
            self._syncing_from_tree = False
            self._update_button_states()

    def _select_shapes_in_document(self, shapes):
        """Select multiple shapes in the document"""
        try:
            if not shapes:
                return

            controller = self.get_controller()

            # Lock undo manager to prevent selection changes from creating undo records
            undo_manager = self._get_undo_manager()
            if undo_manager:
                undo_manager.lock()
            try:
                if len(shapes) == 1:
                    controller.set_selected_shape(shapes[0])
                else:
                    # Create a ShapeCollection for multiple shapes
                    service_manager = self.x_context.getServiceManager()
                    shape_collection = service_manager.createInstanceWithContext(
                        "com.sun.star.drawing.ShapeCollection", self.x_context
                    )
                    for shape in shapes:
                        shape_collection.add(shape)

                    controller._x_controller.select(shape_collection)
            finally:
                if undo_manager:
                    undo_manager.unlock()
        except Exception as e:
            print(f"Error selecting shapes in document: {e}")

    def _select_tree_nodes_for_shapes(self, shapes):
        """Select all tree nodes corresponding to the given shapes.

        Args:
            shapes: List of shapes to select in the tree view
        """
        try:
            if not self.tree_control or not shapes:
                return

            # Find matching tree nodes for each shape
            matching_nodes = []
            for shape in shapes:
                if shape is None:
                    continue
                for node_name, tree_item in self._node_to_tree_item_map.items():
                    if tree_item.get_rectangle_shape() == shape:
                        node = self._find_node_in_tree(node_name)
                        if node:
                            matching_nodes.append(node)
                        break

            if not matching_nodes:
                return

            # Select all matching nodes in tree
            # First node uses select(), subsequent nodes use addSelection()
            for i, node in enumerate(matching_nodes):
                if i == 0:
                    self.tree_control.select(node)
                else:
                    self.tree_control.addSelection(node)

            # Also select corresponding shapes in document
            self._select_shapes_in_document(shapes)

        except Exception as e:
            print(f"Error selecting tree nodes for shapes: {e}")

    def _setup_selection_listener(self):
        """Set up selection change listener for shape-to-tree selection"""
        try:
            controller = self.get_controller()
            # Create and add selection listener
            selection_listener = TreeSelectionListener(self)
            controller._x_controller.addSelectionChangeListener(selection_listener)
            # Store reference to prevent garbage collection
            self._selection_listener = selection_listener
        except Exception as e:
            print(f"Error setting up selection listener: {e}")

    def select_tree_node_for_shape(self, shape):
        """Select the tree node corresponding to the given shape"""
        try:
            if not self.tree_control:
                return

            # Find the tree item that matches this shape
            matching_node_name = None
            for node_name, tree_item in self._node_to_tree_item_map.items():
                if tree_item.get_rectangle_shape() == shape:
                    matching_node_name = node_name
                    break

            if matching_node_name:
                # Find and select the corresponding tree node
                self._select_tree_node_by_name(matching_node_name)

        except Exception as e:
            print(f"Error selecting tree node for shape: {e}")

    def sync_document_selection_to_tree(self, selection):
        """Sync document shape selection to tree view selection (supports multi-selection)

        Args:
            selection: XShapes or shape collection from the document
        """
        try:
            if not self.tree_control or not selection:
                return

            # Collect all matching tree nodes for selected shapes
            matching_nodes = []
            shape_count = selection.getCount()

            for i in range(shape_count):
                try:
                    shape = selection.getByIndex(i)
                    if shape:
                        # Find the tree node for this shape
                        for node_name, tree_item in self._node_to_tree_item_map.items():
                            if tree_item.get_rectangle_shape() == shape:
                                node = self._find_node_in_tree(node_name)
                                if node:
                                    matching_nodes.append(node)
                                break
                except Exception:
                    continue

            if not matching_nodes:
                return

            # Select all matching nodes in tree
            # First node uses select(), subsequent nodes use addSelection()
            for i, node in enumerate(matching_nodes):
                if i == 0:
                    self.tree_control.select(node)
                else:
                    self.tree_control.addSelection(node)

        except Exception as e:
            print(f"Error syncing document selection to tree: {e}")
        finally:
            self._update_button_states()

    def _find_node_in_tree(self, node_name):
        """Find a tree node by name in the current tree model

        Args:
            node_name: Display name of the node to find

        Returns:
            Tree node or None if not found
        """
        try:
            tree_model = self.tree_control.getModel()
            data_model = tree_model.getPropertyValue("DataModel")
            if data_model:
                root_node = data_model.getRoot()
                return self._find_node_by_name(root_node, node_name)
        except Exception:
            pass
        return None

    def _select_tree_node_by_name(self, node_name):
        """Select a tree node by its display name"""
        try:
            if not self.tree_control:
                return

            # Get the tree model and find the node
            tree_model = self.tree_control.getModel()
            data_model = tree_model.getPropertyValue("DataModel")

            if data_model:
                root_node = data_model.getRoot()
                target_node = self._find_node_by_name(root_node, node_name)

                if target_node:
                    # Select the node in the tree using the tree control directly
                    self.tree_control.select(target_node)

        except Exception as e:
            print(f"Error selecting tree node by name: {e}")

    def _find_node_by_name(self, node, target_name):
        """Recursively find a tree node by its display name"""
        try:
            if node.getDisplayValue() == target_name:
                return node

            # Search children
            for i in range(node.getChildCount()):
                child = node.getChildAt(i)
                result = self._find_node_by_name(child, target_name)
                if result:
                    return result

            return None
        except Exception as e:
            return None

    def move_tree_item(self, source_node_names, target_node_name, drop_position):
        """Move tree item(s) to a new position and update the diagram

        Args:
            source_node_names: Single node name (str) or list of node names
            target_node_name: Target node name to move items to
            drop_position: 'child' or 'sibling'
        """
        try:
            if not hasattr(self, "_node_to_tree_item_map"):
                return False

            # Normalize to list
            if isinstance(source_node_names, str):
                source_node_names = [source_node_names]

            target_tree_item = self._node_to_tree_item_map.get(target_node_name)
            if not target_tree_item:
                print("Could not find target tree item")
                return False

            # Get all source tree items
            source_items = []
            for name in source_node_names:
                item = self._node_to_tree_item_map.get(name)
                if item:
                    source_items.append(item)

            if not source_items:
                print("Could not find source tree items")
                return False

            # if parent selected, don't process child separately
            source_items = self._filter_out_descendants(source_items)

            for source_item in source_items:
                if source_item == target_tree_item:
                    print("Cannot move item to itself")
                    return False
                if self._is_descendant_of(target_tree_item, source_item):
                    print("Cannot move item to its descendant")
                    return False

            controller = self.get_controller()
            diagram = controller.get_diagram()

            if not diagram or not hasattr(diagram, "move_tree_item"):
                print("Diagram does not support move operations")
                return False

            # Capture original positions for undo before moving
            move_data = []
            for source_item in source_items:
                original_parent = source_item.get_dad()
                original_prev_sibling = diagram.get_diagram_tree().get_previous_sibling(
                    source_item
                )
                move_data.append((source_item, original_parent, original_prev_sibling))

            # Move each item
            all_success = True
            for source_item in source_items:
                success = diagram.move_tree_item(
                    source_item, target_tree_item, drop_position
                )
                if not success:
                    all_success = False

            if all_success:
                diagram.refresh_diagram()
                self.refresh_tree()

                undo_manager = self._get_undo_manager()
                if undo_manager and move_data:
                    try:
                        undo_action = DragDropUndoAction(
                            self, move_data, target_tree_item, drop_position
                        )
                        undo_manager.addUndoAction(undo_action)
                    except Exception as e:
                        print(f"Failed to register drag and drop undo action: {e}")

                return True
            else:
                print("Some move operations failed")
                return False

        except Exception as e:
            print(f"Error moving tree items: {e}")
            return False

    def _store_selection_before_add(self):
        """Store the currently selected tree item before adding a new shape"""
        try:
            controller = self.get_controller()
            selection = controller._x_controller.getSelection()
            if selection and selection.getCount() > 0:
                selected_shape = selection.getByIndex(0)
                # Find the tree item for this shape
                for tree_item in self._node_to_tree_item_map.values():
                    if tree_item.get_rectangle_shape() == selected_shape:
                        self._parent_before_add = tree_item
                        return
            self._parent_before_add = None
        except Exception as e:
            print(f"Error storing selection before add: {e}")
            self._parent_before_add = None

    def _select_newly_added_child(self):
        """Find and select the newly added child shape"""
        try:
            if not self._parent_before_add:
                print("No parent stored, cannot find new child")
                return

            parent_tree_item = self._parent_before_add

            # Find the parent in the current tree mapping
            parent_node_name = None
            for node_name, tree_item in self._node_to_tree_item_map.items():
                if tree_item == parent_tree_item:
                    parent_node_name = node_name
                    break

            if not parent_node_name:
                # Try to find parent by shape
                parent_shape = parent_tree_item.get_rectangle_shape()
                for node_name, tree_item in self._node_to_tree_item_map.items():
                    if tree_item.get_rectangle_shape() == parent_shape:
                        parent_node_name = node_name
                        parent_tree_item = tree_item
                        break

            if parent_tree_item and parent_tree_item.get_first_child():
                # Look for the last (newest) child
                child_item = parent_tree_item.get_first_child()
                last_child = child_item

                # Find the last sibling (most recently added)
                while child_item:
                    last_child = child_item
                    child_item = child_item.get_first_sibling()

                if last_child:
                    new_shape = last_child.get_rectangle_shape()
                    if new_shape:
                        controller = self.get_controller()
                        controller.set_selected_shape(new_shape)
                        return
        except Exception as e:
            print(f"Error selecting newly added child: {e}")

    def _get_undo_manager(self):
        """Get the document's undo manager"""
        try:
            controller = self.get_controller()
            if controller and hasattr(controller, "_x_controller"):
                # Get the document model from the controller
                model = controller._x_controller.getModel()
                if model and hasattr(model, "getUndoManager"):
                    return model.getUndoManager()
                elif hasattr(model, "UndoManager"):
                    return model.UndoManager
        except Exception as e:
            print(f"Could not get undo manager: {e}")
        return None

    def _find_newly_added_shape(self, parent_tree_item):
        """Find the shape that was just added to the parent"""
        try:
            if not parent_tree_item:
                return None

            # Look for the last (newest) child
            first_child = parent_tree_item.get_first_child()
            if first_child:
                child_item = first_child
                last_child = child_item

                # Find the last sibling (most recently added)
                while child_item:
                    last_child = child_item
                    child_item = child_item.get_first_sibling()

                if last_child:
                    return last_child.get_rectangle_shape()
        except Exception as e:
            print(f"Error finding newly added shape: {e}")
        return None

    def cleanup(self):
        """Clean up all resources before dialog disposal"""
        try:
            if (
                hasattr(self, "_selection_listener")
                and self._selection_listener is not None
            ):
                try:
                    self.get_controller()._x_controller.removeSelectionChangeListener(
                        self._selection_listener
                    )
                except Exception:
                    pass
                self._selection_listener = None

            if hasattr(self, "_node_to_tree_item_map"):
                self._node_to_tree_item_map.clear()

            if hasattr(self, "_clipboard"):
                self._clipboard = None

            if hasattr(self, "_drag_handler"):
                self._drag_handler = None
            if hasattr(self, "_drop_handler"):
                self._drop_handler = None

            self.tree_control = None

            # Clean up temp directory
            if self._temp_dir is not None:
                shutil.rmtree(self._temp_dir, ignore_errors=True)
                self._temp_dir = None

            self._populate_tree_on_show = True

        except Exception as e:
            print(f"Error during ControlDlgHandler cleanup: {e}")


class ClipboardItem:
    """Stores data for a copied tree item and its children"""

    def __init__(self, attributes, children=None):
        self.attributes = attributes  # Shape attributes (MilSymCode, etc.)
        self.children = children if children is not None else []


class EditShapeUndoAction(unohelper.Base, XUndoAction):
    """Undo action for editing a shape in the diagram"""

    def __init__(self, dialog_handler, shape, original_attributes, edited_attributes):
        self.dialog_handler = dialog_handler
        self.original_attributes = original_attributes
        self.edited_attributes = edited_attributes
        self.shape = shape
        self.Title = "Edit Shape"

    def undo(self):
        """Undo the edit by restoring original attributes"""
        try:
            self._apply_attributes(self.original_attributes)
        except Exception as e:
            print(f"Error during undo edit shape: {e}")

    def redo(self):
        """Redo the edit by restoring edited attributes"""
        try:
            self._apply_attributes(self.edited_attributes)
        except Exception as e:
            print(f"Error during redo edit shape: {e}")

    def _apply_attributes(self, attributes):
        """Apply attributes to the shape and regenerate its graphic"""
        if self.shape is None or self.dialog_handler is None:
            return
        controller = self.dialog_handler.get_controller()
        if controller is None:
            return
        controller.remove_selection_listener()

        diagram = controller.get_diagram()
        if diagram is None:
            return

        if len(attributes) == 0:
            insertGraphicAttributes(self.shape, [""])  # Empty SIDC code, no other attrs
            diagram.set_shape_properties(self.shape, diagram.DIAGRAM_SHAPE_TYPE)
            diagram.refresh_diagram()
        else:
            params = self._attributes_to_params(attributes)
            insertGraphicAttributes(self.shape, params)

            # Regenerate SVG and update graphic
            svg_data = generate_icon_svg(self.dialog_handler.script, attributes, 32.0)
            if svg_data:
                diagram.set_new_shape_properties(
                    self.shape, diagram.DIAGRAM_SHAPE_TYPE, svg_data
                )
                diagram.refresh_diagram()

        self.dialog_handler.refresh_tree()
        controller.add_selection_listener()

    def _attributes_to_params(self, attributes):
        """Convert attributes dict to params list format for insertGraphicAttributes

        Input format (from extractGraphicAttributes):
            {"MilSymCode": "...", "MilSymStack": "...", "MilSymReinforced": "...", ...}

        Output format (for insertGraphicAttributes):
            [sidc_code, NamedValue("stack", "..."), NamedValue("reinforced", "..."), ...]
        """
        params = []

        sidc_code = attributes.get("MilSymCode", "")
        params.append(sidc_code)

        # Convert remaining MilSym* attributes to NamedValue objects
        for key, value in attributes.items():
            if key == "MilSymCode":
                continue  # Already handled
            if key.startswith("MilSym"):
                # Convert "MilSymStack" -> "stack", "MilSymReinforced" -> "reinforced"
                attr_name = key[6:]  # Remove prefix
                attr_name = attr_name[0].lower() + attr_name[1:]  # lowercase first char
                params.append(NamedValue(attr_name, value))

        return params

    def disposing(self, event):
        """Handle disposing event"""
        self.dialog_handler = None
        self.shape = None
        self.original_attributes = None
        self.edited_attributes = None


class RemoveShapeUndoAction(unohelper.Base, XUndoAction):
    """Undo action for removing shape(s) from the diagram"""

    def __init__(self, dialog_handler, removal_data):
        """
        Args:
            dialog_handler: Reference to ControlDlgHandler
            removal_data: List of (serialized_data, parent_tree_item, child_shapes) tuples
                          child_shapes are the shapes that were re-parented to grandparent
        """
        self.dialog_handler = dialog_handler
        self.removal_data = removal_data
        self._restored_shapes = []

        count = len(removal_data)
        if count == 1:
            self.Title = "Remove Shape"
        else:
            self.Title = f"Remove {count} Shape(s)"

    def undo(self):
        """Undo the removal by re-adding all shapes and re-attaching children"""
        try:
            controller = self.dialog_handler.get_controller()
            diagram = controller.get_diagram()
            if diagram is None:
                return

            controller.remove_selection_listener()
            self._restored_shapes = []

            for serialized_data, parent_tree_item, child_shapes in reversed(
                self.removal_data
            ):
                # Create the removed shape (without children - they already exist)
                success = diagram.paste_subtree(
                    parent_tree_item, serialized_data, self.dialog_handler.script
                )
                if success:
                    restored = self.dialog_handler._find_newly_added_shape(
                        parent_tree_item
                    )
                    if restored:
                        self._restored_shapes.append(restored)

                        # Re-attach children to the restored shape
                        if child_shapes:
                            self._reattach_children(diagram, restored, child_shapes)

            diagram.refresh_diagram()
            self.dialog_handler.refresh_tree()
            controller.add_selection_listener()
        except Exception as e:
            print(f"Error during undo remove shape: {e}")

    def _reattach_children(self, diagram, parent_shape, child_shapes):
        """Re-attach child shapes to the restored parent shape"""
        try:
            diagram_tree = diagram.get_diagram_tree()
            if diagram_tree is None:
                return

            parent_tree_item = diagram_tree.get_tree_item(parent_shape)
            if parent_tree_item is None:
                return

            for child_shape in child_shapes:
                if child_shape is None:
                    continue
                child_tree_item = diagram_tree.get_tree_item(child_shape)
                if child_tree_item is not None:
                    diagram.move_tree_item(child_tree_item, parent_tree_item, "child")
        except Exception as e:
            print(f"Error re-attaching children: {e}")

    def redo(self):
        """Redo the removal by removing the restored shapes"""
        try:
            if not self._restored_shapes:
                return

            controller = self.dialog_handler.get_controller()
            diagram = controller.get_diagram()
            if diagram is None:
                return

            controller.remove_selection_listener()

            for shape in self._restored_shapes:
                if shape:
                    diagram.remove_shape(shape)

            diagram.refresh_diagram()
            self.dialog_handler.refresh_tree()
            self._restored_shapes = []
            controller.add_selection_listener()
        except Exception as e:
            print(f"Error during redo remove shape: {e}")

    def disposing(self, event):
        """Handle disposing event"""
        self.dialog_handler = None
        self.removal_data = None
        self._restored_shapes = None


class PasteShapeUndoAction(unohelper.Base, XUndoAction):
    """Undo action for pasting shape(s) into the diagram"""

    def __init__(
        self, dialog_handler, clipboard_data_list, parent_tree_item, pasted_shapes
    ):
        """
        Args:
            dialog_handler: Reference to ControlDlgHandler
            clipboard_data_list: List of ClipboardItem objects
            parent_tree_item: Target tree item where shapes were pasted
            pasted_shapes: List of shapes that were pasted
        """
        self.dialog_handler = dialog_handler
        self.clipboard_data_list = clipboard_data_list
        self.parent_tree_item = parent_tree_item
        self.pasted_shapes = list(pasted_shapes) if pasted_shapes else []

        count = len(self.pasted_shapes)
        if count == 1:
            self.Title = "Paste Shape"
        else:
            self.Title = f"Paste {count} Shape(s)"

    def undo(self):
        """Undo the paste by removing all pasted shapes and their subtrees"""
        try:
            if not self.pasted_shapes:
                return

            controller = self.dialog_handler.get_controller()
            diagram = controller.get_diagram()
            if diagram is None:
                return

            controller.remove_selection_listener()

            # Lock undo manager to prevent Writer from creating internal undo records
            undo_manager = self.dialog_handler._get_undo_manager()
            if undo_manager:
                undo_manager.lock()
            try:
                # Remove shapes in reverse order (last pasted first)
                for shape in reversed(self.pasted_shapes):
                    if shape is None:
                        continue
                    pasted_tree_item = self._find_tree_item_for_shape(shape)
                    if pasted_tree_item:
                        self._remove_subtree(diagram, controller, pasted_tree_item)

                # Clear references to removed shapes
                self.pasted_shapes = []

                diagram.refresh_diagram()
            finally:
                if undo_manager:
                    undo_manager.unlock()

            self.dialog_handler.refresh_tree()
            controller.add_selection_listener()
        except Exception as e:
            print(f"Error during undo paste shape: {e}")

    def _remove_subtree(self, diagram, controller, tree_item):
        """Recursively remove a tree item and all its descendants (depth-first)"""
        if tree_item is None:
            return

        # First, recursively remove all children
        child = tree_item.get_first_child()
        while child is not None:
            next_sibling = child.get_first_sibling()  # Save before removal
            self._remove_subtree(diagram, controller, child)
            child = next_sibling

        shape = tree_item.get_rectangle_shape()
        if shape:
            diagram.remove_shape(shape)

    def _find_tree_item_for_shape(self, shape):
        """Find the tree item corresponding to a shape"""
        for tree_item in self.dialog_handler._node_to_tree_item_map.values():
            if tree_item.get_rectangle_shape() == shape:
                return tree_item
        return None

    def redo(self):
        """Redo the paste by re-pasting all clipboard items"""
        try:
            controller = self.dialog_handler.get_controller()
            diagram = controller.get_diagram()
            if diagram is None:
                return

            controller.remove_selection_listener()

            # Lock undo manager to prevent Writer from creating internal undo records
            undo_manager = self.dialog_handler._get_undo_manager()
            if undo_manager:
                undo_manager.lock()
            try:
                self.pasted_shapes = []
                for clipboard_data in self.clipboard_data_list:
                    success = diagram.paste_subtree(
                        self.parent_tree_item,
                        clipboard_data,
                        self.dialog_handler.script,
                    )
                    if success:
                        pasted_shape = self.dialog_handler._find_newly_added_shape(
                            self.parent_tree_item
                        )
                        if pasted_shape:
                            self.pasted_shapes.append(pasted_shape)

                diagram.refresh_diagram()
            finally:
                if undo_manager:
                    undo_manager.unlock()

            self.dialog_handler.refresh_tree()

            # Select all re-pasted shapes
            if self.pasted_shapes:
                self.dialog_handler._select_tree_nodes_for_shapes(self.pasted_shapes)

            controller.add_selection_listener()
        except Exception as e:
            print(f"Error during redo paste shape: {e}")

    def disposing(self, event):
        """Handle disposing event"""
        self.dialog_handler = None
        self.clipboard_data_list = None
        self.parent_tree_item = None
        self.pasted_shapes = None


class AddShapeUndoAction(unohelper.Base, XUndoAction):
    """Undo action for adding a shape to the diagram"""

    def __init__(self, dialog_handler, added_shape, parent_tree_item):
        self.dialog_handler = dialog_handler
        self.added_shape = added_shape
        self.parent_tree_item = parent_tree_item
        self.Title = "Add Shape"

    def undo(self):
        """Undo the add shape operation by removing the added shape"""
        try:
            if self.added_shape and self.dialog_handler:
                controller = self.dialog_handler.get_controller()
                if controller and controller.get_diagram():
                    # Temporarily remove selection listener to avoid conflicts
                    controller.remove_selection_listener()

                    # Lock undo manager to prevent internal undo records during undo
                    undo_manager = self.dialog_handler._get_undo_manager()
                    if undo_manager:
                        undo_manager.lock()
                    try:
                        # Remove the shape using the diagram's remove method
                        controller.get_diagram().remove_shape(self.added_shape)
                        controller.get_diagram().refresh_diagram()
                    finally:
                        if undo_manager:
                            undo_manager.unlock()

                    # Clear reference to removed shape to prevent memory leak
                    self.added_shape = None

                    # Refresh the tree view
                    self.dialog_handler.refresh_tree()

                    # Select the parent shape if it still exists
                    if self.parent_tree_item:
                        parent_shape = self.parent_tree_item.get_rectangle_shape()
                        if parent_shape:
                            controller.set_selected_shape(parent_shape)

                    # Re-add selection listener
                    controller.add_selection_listener()

        except Exception as e:
            print(f"Error during undo add shape: {e}")

    def redo(self):
        """Redo the add shape operation"""
        try:
            if self.parent_tree_item and self.dialog_handler:
                controller = self.dialog_handler.get_controller()
                if controller and controller.get_diagram():
                    # Temporarily remove selection listener
                    controller.remove_selection_listener()

                    # Lock undo manager to prevent internal undo records during redo
                    undo_manager = self.dialog_handler._get_undo_manager()
                    if undo_manager:
                        undo_manager.lock()
                    try:
                        parent_shape = self.parent_tree_item.get_rectangle_shape()

                        controller.get_diagram().add_shape(parent_shape)
                        controller.get_diagram().refresh_diagram()

                        # Look up the CURRENT tree item from the diagram tree
                        current_parent_tree_item = (
                            controller.get_diagram()
                            .get_diagram_tree()
                            .get_tree_item(parent_shape)
                        )

                        # Store reference to newly added shape for subsequent undo
                        self.added_shape = self.dialog_handler._find_newly_added_shape(
                            current_parent_tree_item
                        )
                    finally:
                        if undo_manager:
                            undo_manager.unlock()

                    # Refresh tree and select newly added shape
                    self.dialog_handler.refresh_tree()
                    self.dialog_handler._select_newly_added_child()

                    # Re-add selection listener
                    controller.add_selection_listener()

        except Exception as e:
            print(f"Error during redo add shape: {e}")

    def disposing(self, event):
        """Handle disposing event"""
        self.dialog_handler = None
        self.added_shape = None
        self.parent_tree_item = None


class DragDropUndoAction(unohelper.Base, XUndoAction):
    """Undo action for moving shape(s) via drag and drop"""

    def __init__(self, dialog_handler, move_data, target_tree_item, drop_position):
        """
        Args:
            dialog_handler: Reference to ControlDlgHandler
            move_data: List of tuples (source_tree_item, original_parent, original_prev_sibling)
            target_tree_item: The target where items were moved to
            drop_position: "child" or "sibling"
        """
        self.dialog_handler = dialog_handler
        self.move_data = move_data
        self.target_tree_item = target_tree_item
        self.drop_position = drop_position

        count = len(move_data)
        if count == 1:
            self.Title = "Drag and Drop"
        else:
            self.Title = f"Drag and Drop ({count} items)"

    def undo(self):
        """Undo by moving all items back to their original positions"""
        try:
            controller = self.dialog_handler.get_controller()
            diagram = controller.get_diagram()
            if diagram is None:
                return

            controller.remove_selection_listener()

            # Restore in reverse order to maintain tree structure
            for source_item, original_parent, _ in reversed(self.move_data):
                if source_item and original_parent:
                    diagram.move_tree_item(source_item, original_parent, "child")

            diagram.refresh_diagram()
            self.dialog_handler.refresh_tree()
            controller.add_selection_listener()
        except Exception as e:
            print(f"Error during undo drag and drop: {e}")

    def redo(self):
        """Redo by moving all items to the target again"""
        try:
            controller = self.dialog_handler.get_controller()
            diagram = controller.get_diagram()
            if diagram is None:
                return

            controller.remove_selection_listener()

            for source_item, _, _ in self.move_data:
                if source_item and self.target_tree_item:
                    diagram.move_tree_item(
                        source_item, self.target_tree_item, self.drop_position
                    )

            diagram.refresh_diagram()
            self.dialog_handler.refresh_tree()
            controller.add_selection_listener()
        except Exception as e:
            print(f"Error during redo drag and drop: {e}")

    def disposing(self, event):
        """Handle disposing event"""
        self.dialog_handler = None
        self.move_data = None
        self.target_tree_item = None


class TreeKeyHandler(unohelper.Base, XKeyListener):
    """Handle keyboard events on tree control for navigation selection"""

    def __init__(self, dialog_handler):
        self.dialog_handler = dialog_handler

    def keyPressed(self, event):
        """Handle key pressed events for navigation"""
        pass

    def keyReleased(self, event):
        """Handle key released events"""
        try:
            if event.KeyCode == Key.DELETE:
                self.dialog_handler.remove_selected_shape()
                return
            elif event.KeyCode == Key.C and (event.Modifiers & KeyModifier.MOD1):
                self.dialog_handler.copy_selected_item()
                return
            elif event.KeyCode == Key.V and (event.Modifiers & KeyModifier.MOD1):
                self.dialog_handler.paste_to_selected_item()
                return
            elif event.KeyCode == Key.Z and (event.Modifiers & KeyModifier.MOD1):
                self.dialog_handler.handle_undo()
            elif event.KeyCode == Key.Y and (event.Modifiers & KeyModifier.MOD1):
                self.dialog_handler.handle_redo()

            navigation_keys = [
                Key.UP,
                Key.DOWN,
                Key.LEFT,
                Key.RIGHT,
                Key.PAGEUP,
                Key.PAGEDOWN,
                Key.HOME,
                Key.END,
            ]

            if event.KeyCode in navigation_keys:
                try:
                    tree_control = event.Source
                    selection = tree_control.getSelection()
                    if selection is not None:
                        is_shift = bool(event.Modifiers & KeyModifier.SHIFT)

                        if is_shift:
                            self.dialog_handler.sync_all_selected_shapes_to_document()
                        else:
                            selected_node = None
                            if hasattr(selection, "getDisplayValue"):
                                selected_node = selection
                            elif hasattr(selection, "__len__") and len(selection) > 0:
                                selected_node = selection[0]

                            if selected_node and hasattr(
                                selected_node, "getDisplayValue"
                            ):
                                self.dialog_handler.handle_tree_selection(selected_node)
                except Exception as e:
                    print(f"Error getting tree selection after key navigation: {e}")

        except Exception as e:
            print(f"Error handling key release: {e}")

    def disposing(self, event):
        """Handle disposing events"""
        pass


class TreeMouseHandler(unohelper.Base, XMouseListener):
    """Handle mouse events on tree control for click selection"""

    def __init__(self, dialog_handler):
        self.dialog_handler = dialog_handler

    def mousePressed(self, event):
        """Handle mouse pressed events"""
        pass

    def mouseReleased(self, event):
        """Handle mouse released events - sync selection to document shapes

        Supports:
        - Normal click: Select single item (replaces selection)
        - Shift+click: Range selection
        - Ctrl+click: Toggle item selection
        """
        try:
            if getattr(self.dialog_handler, "_is_dragging", False):
                return

            if event.Buttons == MouseButton.LEFT and event.ClickCount == 1:
                tree_control = self.dialog_handler.tree_control
                if not tree_control:
                    return

                is_shift = bool(event.Modifiers & KeyModifier.SHIFT)
                is_ctrl = bool(event.Modifiers & KeyModifier.MOD1)

                if is_shift or is_ctrl:
                    self.dialog_handler.sync_all_selected_shapes_to_document()
                else:
                    clicked_node = tree_control.getNodeForLocation(event.X, event.Y)
                    if clicked_node:
                        self.dialog_handler.handle_tree_selection(clicked_node)

        except Exception as e:
            print(f"Error in mouseReleased: {e}")

    def mouseEntered(self, event):
        """Handle mouse entered events"""
        pass

    def mouseExited(self, event):
        """Handle mouse exited events"""
        pass

    def disposing(self, event):
        """Handle disposing events"""
        pass


class TreeSelectionListener(unohelper.Base, XSelectionChangeListener):
    """Listen for shape selection changes to update tree selection"""

    def __init__(self, dialog_handler):
        self.dialog_handler = dialog_handler

    def selectionChanged(self, event):
        """Handle selection change events from the document"""
        try:
            # Skip if tree control has focus (user is navigating tree with keyboard)
            tree_control = getattr(self.dialog_handler, "tree_control", None)
            if tree_control and tree_control.hasFocus():
                return

            # Skip if we're currently syncing from tree to document
            if getattr(self.dialog_handler, "_syncing_selection", False):
                return

            # Skip if we're currently dragging
            if getattr(self.dialog_handler, "_is_dragging", False):
                return

            # Skip if sync was initiated from tree
            if getattr(self.dialog_handler, "_syncing_from_tree", False):
                return

            # Get the selected shapes and sync ALL of them to tree
            selection = event.Source.getSelection()
            if selection and selection.getCount() > 0:
                self.dialog_handler.sync_document_selection_to_tree(selection)
        except Exception as e:
            # Silently ignore selection errors to avoid spam
            pass
        finally:
            self.dialog_handler._update_button_states()

    def disposing(self, event):
        """Handle disposing events"""
        pass


class TreeDragHandler(unohelper.Base, XDragGestureListener, XDragSourceListener):
    """Handle drag operations on tree control"""

    def __init__(self, dialog_handler):
        self.dialog_handler = dialog_handler
        self.dragged_node_names = None

    def dragGestureRecognized(self, event):
        """Handle drag gesture recognition"""
        try:
            tree_control = self.dialog_handler.tree_control
            if not tree_control:
                return

            # Get the node at the drag point
            origin_node = tree_control.getNodeForLocation(
                event.DragOriginX, event.DragOriginY
            )
            if not origin_node or not hasattr(origin_node, "getDisplayValue"):
                return

            origin_name = origin_node.getDisplayValue()

            # Don't allow dragging the root node
            if origin_name == "Root" or "Diagram Structure" in origin_name:
                return

            # Check if origin node is part of current selection
            dragged_names = []
            if self.dialog_handler._is_node_selected(origin_node):
                selection = tree_control.getSelection()
                if selection is not None:
                    # Handle both single node and sequence of nodes
                    nodes = []
                    if hasattr(selection, "getDisplayValue"):
                        nodes = [selection]
                    elif hasattr(selection, "__len__"):
                        nodes = list(selection)
                    elif hasattr(selection, "__iter__"):
                        nodes = list(selection)

                    for node in nodes:
                        if node and hasattr(node, "getDisplayValue"):
                            name = node.getDisplayValue()
                            if name != "Root" and "Diagram Structure" not in name:
                                dragged_names.append(name)
            else:
                # Drag only the origin node (not in selection)
                dragged_names = [origin_name]

            if not dragged_names:
                return

            # Store for reference and set dragging flag
            self.dragged_node_names = dragged_names
            self.dialog_handler._is_dragging = True

            # Create transferable with all names
            transferable = TreeNodeTransferable(dragged_names)

            event.DragSource.startDrag(event, ACTION_MOVE, 0, 0, transferable, self)

        except Exception as e:
            print(f"Error in drag gesture: {e}")

    def dragEnter(self, event):
        """Handle drag enter"""
        pass

    def dragExit(self, event):
        """Handle drag exit"""
        pass

    def dragOver(self, event):
        """Handle drag over"""
        pass

    def dropActionChanged(self, event):
        """Handle drop action changed"""
        pass

    def dragDropEnd(self, event):
        """Handle drag drop end"""
        self.dialog_handler._is_dragging = False

        if event.DropSuccess:
            pass
        else:
            if self.dragged_node_names:
                self._restore_selection(self.dragged_node_names)

        self.dragged_node_names = None

    def _restore_selection(self, node_names):
        """Restore tree selection to the given node names"""
        try:
            tree_control = self.dialog_handler.tree_control
            if not tree_control:
                return

            tree_model = tree_control.getModel()
            data_model = tree_model.getPropertyValue("DataModel")
            if not data_model:
                return

            root_node = data_model.getRoot()

            # Find and select each node
            first = True
            for name in node_names:
                node = self.dialog_handler._find_node_by_name(root_node, name)
                if node:
                    if first:
                        tree_control.select(node)
                        first = False
                    else:
                        tree_control.addSelection(node)
        except Exception as e:
            print(f"Error restoring selection: {e}")

    def disposing(self, event):
        """Handle disposing"""
        pass


class TreeDropHandler(unohelper.Base, XDropTargetListener):
    """Handle drop operations on tree control"""

    def __init__(self, dialog_handler):
        self.dialog_handler = dialog_handler

    def drop(self, event):
        """Handle drop operation"""
        try:
            event.Source.acceptDrop(ACTION_MOVE)

            transferable = event.Transferable
            if not transferable:
                event.Source.dropComplete(False)
                return

            data_flavor = TreeNodeTransferable.get_data_flavor()
            if not transferable.isDataFlavorSupported(data_flavor):
                event.Source.dropComplete(False)
                return

            data = uno.invoke(transferable, "getTransferData", (data_flavor,))

            dragged_node_names = None
            if data:
                raw = str(data.value) if hasattr(data, "value") else str(data)
                try:
                    dragged_node_names = json.loads(raw)
                except json.JSONDecodeError:
                    # Fallback: single name (backward compatibility)
                    dragged_node_names = [raw]

            if not dragged_node_names:
                event.Source.dropComplete(False)
                return

            # Get the drop target node
            tree_control = self.dialog_handler.tree_control
            if not tree_control:
                event.Source.dropComplete(False)
                return

            target_node = tree_control.getNodeForLocation(
                event.LocationX, event.LocationY
            )
            if not target_node or not hasattr(target_node, "getDisplayValue"):
                event.Source.dropComplete(False)
                return

            target_node_name = target_node.getDisplayValue()

            # Don't allow dropping on root or on any of the dragged nodes
            if (
                target_node_name == "Root"
                or "Diagram Structure" in target_node_name
                or target_node_name in dragged_node_names
            ):
                event.Source.rejectDrop()
                return

            drop_position = "child"

            # Perform the move operation with all dragged nodes
            success = self.dialog_handler.move_tree_item(
                dragged_node_names, target_node_name, drop_position
            )

            event.Source.dropComplete(success)

        except Exception as e:
            print(f"Error handling drop: {e}")
            try:
                event.Source.dropComplete(False)
            except:
                pass

    def dragEnter(self, event):
        """Handle drag enter over drop target"""
        # Accept the drag if it contains our data type
        try:
            if event.SourceActions & ACTION_MOVE:
                event.Source.acceptDrag(ACTION_MOVE)
            else:
                event.Source.rejectDrag()
        except Exception as e:
            print(f"Error in drag enter: {e}")
            event.Source.rejectDrag()

    def dragExit(self, event):
        """Handle drag exit from drop target"""
        pass

    def dragOver(self, event):
        """Handle drag over drop target"""
        # Continue to accept the drag
        try:
            if event.SourceActions & ACTION_MOVE:
                event.Source.acceptDrag(ACTION_MOVE)
            else:
                event.Source.rejectDrag()
        except Exception as e:
            print(f"Error in drag over: {e}")
            event.Source.rejectDrag()

    def dropActionChanged(self, event):
        """Handle drop action changed"""
        pass

    def disposing(self, event):
        """Handle disposing"""
        pass


class TreeNodeTransferable(unohelper.Base, XTransferable):
    """Transferable data for tree node drag & drop"""

    def __init__(self, node_names):
        # Accept single name (str) or list of names
        if isinstance(node_names, str):
            self.node_names = [node_names]
        else:
            self.node_names = list(node_names)
        self._data_flavor = self._create_data_flavor()

    def getTransferData(self, flavor):
        """Get transfer data for the given flavor"""
        if self.isDataFlavorSupported(flavor):
            # Return as JSON string for multi-node support
            return json.dumps(self.node_names)
        return None

    def getTransferDataFlavors(self):
        """Get available data flavors"""
        return (self._data_flavor,)

    def isDataFlavorSupported(self, flavor):
        """Check if data flavor is supported"""
        return (
            flavor.MimeType == self._data_flavor.MimeType
            and flavor.HumanPresentableName == self._data_flavor.HumanPresentableName
        )

    def _create_data_flavor(self):
        """Create the data flavor for tree node data"""
        flavor = DataFlavor()
        flavor.MimeType = "application/x-treenode"
        flavor.HumanPresentableName = "Tree Node"
        flavor.DataType = uno.getTypeByName("string")
        return flavor

    @staticmethod
    def get_data_flavor():
        """Get the data flavor for tree node data"""
        flavor = DataFlavor()
        flavor.MimeType = "application/x-treenode"
        flavor.HumanPresentableName = "Tree Node"
        flavor.DataType = uno.getTypeByName("string")
        return flavor
