# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import os
import sys
import uno
import unohelper
import shutil

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from com.sun.star.awt import KeyModifier, XDialogEventHandler, XTopWindowListener, XMouseListener, XKeyListener, Key, XWindowListener
from com.sun.star.awt import MouseButton
from com.sun.star.view.SelectionType import SINGLE as SELECTION_TYPE_SINGLE
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.datatransfer.dnd import XDragGestureListener, XDropTargetListener
from com.sun.star.datatransfer.dnd import XDragSourceListener
from com.sun.star.datatransfer.dnd.DNDConstants import ACTION_MOVE
from com.sun.star.datatransfer import XTransferable, DataFlavor
from com.sun.star.beans import NamedValue
from utils import extractGraphicAttributes, getExtensionBasePath, generate_icon_svg
from unohelper import systemPathToFileUrl
import tempfile


class ControlDlgHandler(unohelper.Base, XDialogEventHandler, XTopWindowListener, XWindowListener):
    buttons = ["addShape", "removeShape", "editShape"]

    def __init__(self, dialog, x_context, model):
        self.dialog = dialog
        self.x_context = x_context
        self.tree_control = None
        self._populate_tree_on_show = True
        self._parent_before_add = None
        self._temp_dir: str | None = None
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
                # Add shape and get reference to newly created shape
                self.get_controller().get_diagram().add_shape()
                self.get_controller().get_diagram().refresh_diagram()
                # Refresh tree after adding shape
                self.refresh_tree()
                self._select_newly_added_child()
                self.get_controller().add_selection_listener()
            return True
        elif methodName == "OnRemove":
            self.remove_selected_shape()
            return True
        elif methodName == "OnEdit":
            self.dialog.execute_properties_dialog()
            self.get_controller().get_diagram().refresh_diagram()
            # Refresh tree after editing (properties might affect display)
            self.refresh_tree()
            return True
        else:
            return False

    def getSupportedMethodNames(self):
        return self.buttons

    def disposing(self, event):
        pass

    def buttonStateHandler(self, methodName):
        pass

    def get_gui(self):
        """Get GUI reference"""
        return self.dialog

    def get_controller(self):
        """Get controller reference"""
        return self.dialog.get_controller()

    def paste_to_selected_item(self):
        """Paste clipboard contents to currently selected item"""

        if self._clipboard is None:
            return

        try:
            selected_node = self.tree_control.getSelection()
            if selected_node is None:
                return

            node_name = selected_node.getDisplayValue()
            target_tree_item = self._node_to_tree_item_map.get(node_name)
            if target_tree_item is None:
                return

            controller = self.get_controller()
            diagram = controller.get_diagram()

            if diagram:
                controller.remove_selection_listener()
                success = diagram.paste_subtree(
                    target_tree_item, self._clipboard, self.script
                )
                if success:
                    diagram.refresh_diagram()
                    self.refresh_tree()
                controller.add_selection_listener()

        except Exception as ex:
            print(f"Error pasting: {ex}")

    def copy_selected_item(self):
        """Copy the currently selected item and its subtree"""
        try:
            if self.tree_control is None:
                return

            selected_node = self.tree_control.getSelection()
            if selected_node is None:
                return

            node_name = selected_node.getDisplayValue()
            tree_item = self._node_to_tree_item_map.get(node_name)
            if tree_item is None:
                return

            self._clipboard = self._serialize_tree_item(tree_item)

        except Exception as ex:
            print(f"Error copying item: {ex}")

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

    def remove_selected_shape(self):
        """Remove the currently selected shape from the diagram"""
        if self.get_controller().get_diagram() is not None:
            self.get_controller().remove_selection_listener()
            if self.get_controller().get_group_type() == self.get_controller().ORGANIGROUP:
                org_chart = self.get_controller().get_diagram()
                if org_chart.is_error_in_tree():
                    self.get_gui().ask_user_for_repair(org_chart)
                else:
                    self.get_controller().get_diagram().remove_shape()
                    self.get_controller().get_diagram().refresh_diagram()
                    # Refresh tree after removing shape
                    self.refresh_tree()
            else:
                self.get_controller().get_diagram().remove_shape()
                self.get_controller().get_diagram().refresh_diagram()
                # Refresh tree after removing shape
                self.refresh_tree()
            self.get_controller().add_selection_listener()

    # XTopWindowListener methods
    def windowClosing(self, event):
        """Handle window closing event"""
        if event.Source == self.get_gui()._x_control_dialog:
            self.get_gui().set_visible_control_dialog(False)

        # Delete tmp preview directory
        if self._temp_dir is None:
            return

        shutil.rmtree(self._temp_dir, ignore_errors=True)
        self._temp_dir = None

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
            new_tree_width = event.Width - event.LeftInset - event.RightInset - margin_left - margin_right
            new_tree_height = event.Height - event.TopInset - event.BottomInset - margin_top - margin_bottom

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
            tree_model.setPropertyValue("SelectionType", SELECTION_TYPE_SINGLE)
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
                    "com.sun.star.awt.tree.MutableTreeDataModel", self.x_context)

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
                        root_node_name = self._get_tree_node_display_name(temp_root_item, 1)
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
                    self._populate_tree_node(data_model, parent_node, child_item, child_name)

                    # Move to next sibling
                    child_item = child_item.get_first_sibling()
                    child_num += 1
                    child_count += 1

        except Exception as e:
            print(f"Error populating tree children: {e}")

    def _populate_tree_node(self, data_model, parent_node, tree_item, display_name):
        """Recursively populate tree nodes from organization chart tree items"""
        try:
            if tree_item is None:
                return

            # Create node for this tree item
            node = data_model.createNode(display_name, False)  # Start with no children
            parent_node.appendChild(node)

            # Store reference to tree item in a class attribute for later selection
            if not hasattr(self, '_node_to_tree_item_map'):
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
            if attributes and "MilSymCode" in attributes:
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
                shape_name = ""

                # Try to get shape name
                try:
                    shape_name = shape.getName()
                except:
                    pass

                # Try to get shape text/string content
                shape_text = ""
                try:
                    shape_text = shape.getString()
                except:
                    try:
                        text_obj = shape.getText()
                        if text_obj:
                            shape_text = text_obj.getString()
                    except:
                        pass

                # Try to get level information
                level_info = ""
                try:
                    level = tree_item.get_level()
                    level_info = f" (L{level})"
                except:
                    pass

                # Build display name
                if shape_text and shape_text.strip():
                    return f"{shape_text.strip()}{level_info}"
                elif shape_name and shape_name.strip():
                    return f"Shape: {shape_name}{level_info}"
                else:
                    return f"Item {item_number}{level_info}"

            return f"Node {item_number}"

        except Exception as e:
            print(f"Error getting tree node display name: {e}")
            return f"Item {item_number}"

    def handle_tree_selection(self, selected_node):
        """Handle tree node selection"""
        try:
            if selected_node is not None:
                node_name = selected_node.getDisplayValue()

                # Get the tree item associated with this node
                tree_item = self._node_to_tree_item_map.get(node_name)
                if tree_item:
                    shape = tree_item.get_rectangle_shape()
                    if shape:
                        # Select the shape in the document
                        controller = self.get_controller()
                        try:
                            controller.set_selected_shape(shape)
                        except:
                            try:
                                controller._x_controller.select(shape)
                            except:
                                print("Could not select shape through controller")

        except Exception as e:
            print(f"Error handling tree selection: {e}")

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
                    # Select the node in the tree
                    try:
                        selection = self.tree_control.getSelection()
                        selection.select(target_node)
                    except:
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

    def move_tree_item(self, source_node_name, target_node_name, drop_position):
        """Move a tree item to a new position and update the diagram"""
        try:
            if not hasattr(self, "_node_to_tree_item_map"):
                return False

            source_tree_item = self._node_to_tree_item_map.get(source_node_name)
            target_tree_item = self._node_to_tree_item_map.get(target_node_name)

            if not source_tree_item or not target_tree_item:
                print(f"Could not find tree items for move operation")
                return False

            # Get the diagram and perform the move operation
            controller = self.get_controller()
            diagram = controller.get_diagram()

            if diagram and hasattr(diagram, "move_tree_item"):
                # Call diagram's move method
                success = diagram.move_tree_item(
                    source_tree_item, target_tree_item, drop_position
                )
                if success:
                    # Refresh the diagram and tree view
                    diagram.refresh_diagram()
                    self.refresh_tree()
                    return True
                else:
                    print(f"Failed to move tree item in diagram")
                    return False
            else:
                print(f"Diagram does not support move operations")
                return False

        except Exception as e:
            print(f"Error moving tree item: {e}")
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

class ClipboardItem:
    """Stores data for a copied tree item and its children"""
    def __init__(self, attributes, children=None):
        self.attributes = attributes    # Shape attributes (MilSymCode, etc.)
        self.children = children if children is not None else []

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
                    selection = event.Source.getSelection()
                    if selection:
                        self.dialog_handler.handle_tree_selection(selection)
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
        """Handle mouse released events - detect single-clicks for shape selection"""
        try:
            if event.Buttons == MouseButton.LEFT and event.ClickCount == 1:
                # Single-click detected, get selected node and select shape
                tree_control = self.dialog_handler.tree_control
                if tree_control:
                    try:
                        selected_node = tree_control.getNodeForLocation(
                            event.X, event.Y
                        )
                        if selected_node:
                            self.dialog_handler.handle_tree_selection(selected_node)

                    except Exception as inner_e:
                        print(f"Error getting selected node: {inner_e}")
        except Exception as e:
            print(f"Error handling tree single-click: {e}")

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
            # Get the selected shapes
            selection = event.Source.getSelection()
            if selection and selection.getCount() > 0:
                # Get the first selected shape
                selected_shape = selection.getByIndex(0)
                if selected_shape:
                    # Update tree selection to match
                    self.dialog_handler.select_tree_node_for_shape(selected_shape)
        except Exception as e:
            # Silently ignore selection errors to avoid spam
            pass

    def disposing(self, event):
        """Handle disposing events"""
        pass


class TreeDragHandler(unohelper.Base, XDragGestureListener, XDragSourceListener):
    """Handle drag operations on tree control"""

    def __init__(self, dialog_handler):
        self.dialog_handler = dialog_handler
        self.dragged_node_name = None

    def dragGestureRecognized(self, event):
        """Handle drag gesture recognition"""
        try:
            tree_control = self.dialog_handler.tree_control
            if tree_control:
                # Get the node at the drag point
                try:
                    node = tree_control.getNodeForLocation(
                        event.DragOriginX, event.DragOriginY
                    )
                    if node and hasattr(node, "getDisplayValue"):
                        self.dragged_node_name = node.getDisplayValue()

                        # Don't allow dragging the root node
                        if (
                            self.dragged_node_name == "Root"
                            or "Diagram Structure" in self.dragged_node_name
                        ):
                            return

                        # Create transferable data
                        transferable = TreeNodeTransferable(self.dragged_node_name)

                        # Start drag operation
                        event.DragSource.startDrag(
                            event, ACTION_MOVE, 0, 0, transferable, self
                        )
                        print(f"Started drag for node: {self.dragged_node_name}")

                except Exception as e:
                    print(f"Error getting drag node: {e}")
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
        if event.DropSuccess:
            print(f"Drag operation completed successfully")
        else:
            print(f"Drag operation failed")

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
            # Accept the drop
            event.Source.acceptDrop(ACTION_MOVE)

            # Get the transferable data
            transferable = event.Transferable
            if transferable:
                # Get the dragged node name
                try:
                    data_flavor = TreeNodeTransferable.get_data_flavor()
                    if transferable.isDataFlavorSupported(data_flavor):
                        data = transferable.getTransferData(data_flavor)  # Doesn't work
                        data = uno.invoke(
                            transferable, "getTransferData", (data_flavor,)
                        )  # Doesn't work either
                        print("Received drop data:", data)
                        dragged_node_name = None
                        if data:
                            # Extract string from uno.Any if needed
                            if hasattr(data, "value"):
                                dragged_node_name = str(data.value)
                            else:
                                dragged_node_name = str(data)

                        if dragged_node_name is None:
                            event.Source.dropComplete(False)
                            return

                        # Get the drop target node
                        tree_control = self.dialog_handler.tree_control
                        if tree_control:
                            target_node = tree_control.getNodeForLocation(
                                event.LocationX, event.LocationY
                            )
                            if target_node and hasattr(target_node, "getDisplayValue"):
                                target_node_name = target_node.getDisplayValue()

                                # Don't allow dropping on the root node or on itself
                                if (
                                    target_node_name == "Root"
                                    or "Diagram Structure" in target_node_name
                                    or dragged_node_name == target_node_name
                                ):
                                    event.Source.rejectDrop()
                                    return

                                # Determine drop position (before, after, or as child)
                                # For simplicity, we'll treat all drops as "move as sibling"
                                drop_position = "child"

                                # Perform the move operation
                                success = self.dialog_handler.move_tree_item(
                                    dragged_node_name, target_node_name, drop_position
                                )

                                if success:
                                    event.Source.dropComplete(True)
                                else:
                                    event.Source.dropComplete(False)
                                return

                        event.Source.dropComplete(False)

                except Exception as e:
                    print(f"Error processing drop data: {e}")
                    event.Source.dropComplete(False)
            else:
                event.Source.dropComplete(False)

        except Exception as e:
            print(f"Error in drop operation: {e}")
            event.Source.rejectDrop()

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

    def __init__(self, node_name):
        self.node_name = node_name
        self._data_flavor = self._create_data_flavor()

    def getTransferData(self, flavor):
        """Get transfer data for the given flavor"""
        if self.isDataFlavorSupported(flavor):
            # Return the node name as a string
            return self.node_name
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
