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

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from symbol_dialog import open_symbol_dialog
from com.sun.star.awt import XDialogEventHandler, XTopWindowListener, XMouseListener
from com.sun.star.awt import WindowDescriptor, WindowAttribute, MouseButton
from com.sun.star.awt.WindowClass import MODALTOP
from com.sun.star.view.SelectionType import SINGLE as SELECTION_TYPE_SINGLE

class Gui:
    def __init__(self, controller, x_context, x_frame):
        """Initialize GUI"""
        self._controller = controller
        self._x_context = x_context
        self._x_frame = x_frame
        self._x_control_dialog = None
        self._is_visible_control_dialog = False
        self._o_listener = None

    def get_controller(self):
        """Get controller reference"""
        return self._controller

    def execute_properties_dialog(self):
        ctx = self._x_context
        model = self._x_frame.getController().getModel()
        controller = self.get_controller()
        selected_shape = controller.get_diagram().get_last_shape()
        open_symbol_dialog(ctx, model, controller, None, selected_shape, None)

    def set_visible_control_dialog(self, visible: bool):
        """Set visibility of control dialog"""
        new_diagram_id = self.get_controller().get_diagram().get_diagram_id()

        # Need to create new controlDialog when a new diagram selected (not same GUI panels of diagrams)
        if ((self.get_controller().get_last_diagram_type() != -1 or
             self.get_controller().get_last_diagram_id() != -1) and
            (self.get_controller().get_last_diagram_type() != self.get_controller().get_diagram_type() or
             self.get_controller().get_last_diagram_id() != new_diagram_id)):
            if self._x_control_dialog is not None:
                self._x_control_dialog.setVisible(False)
            self.create_control_dialog()

        if self._x_control_dialog is None:
            self.create_control_dialog()

        if self._x_control_dialog is not None:
            self._x_control_dialog.setVisible(visible)
            if visible:
                self._is_visible_control_dialog = True
                self._x_control_dialog.setFocus()
                # Refresh tree when dialog becomes visible
                if self._o_listener:
                    self._o_listener.refresh_tree()
            else:
                self._is_visible_control_dialog = False

        self.get_controller().set_last_diagram_type(self.get_controller().get_diagram_type())
        self.get_controller().set_last_diagram_id(new_diagram_id)

    def create_control_dialog(self):
        """Create control dialog"""
        try:
            dialog_provider = self._x_context.getServiceManager().createInstanceWithContext("com.sun.star.awt.DialogProvider2", self._x_context)
            s_dialog_url = "vnd.sun.star.extension://com.collabora.milsymbol/dialog/ControlDlg.xdl"

            self._o_listener = ControlDlgHandler(self)
            self._x_control_dialog = dialog_provider.createDialogWithHandler(s_dialog_url, self._o_listener)

            if self._x_control_dialog is not None:
                self._x_control_dialog.addTopWindowListener(self._o_listener)

        except Exception as ex:
            print(f"Error creating control dialog: {ex}")

    def close_and_dispose_control_dialog(self):
        """Close and dispose control dialog"""
        if self._x_control_dialog is not None:
            self._x_control_dialog.removeTopWindowListener(self._o_listener)
            self._x_control_dialog.setVisible(False)
            x_comp = self._x_control_dialog
            if x_comp is not None:
                x_comp.dispose()

        self._x_control_dialog = None

    def is_visible_control_dialog(self):
        """Check if control dialog is visible"""
        return self._is_visible_control_dialog

    def enable_control_dialog_window(self, enable: bool):
        """Enable or disable control dialog window"""
        if self._x_control_dialog is not None:
            self._x_control_dialog.setEnable(enable)

    def set_focus_control_dialog(self):
        """Set focus to control dialog"""
        if self._x_control_dialog is not None:
            self._x_control_dialog.setFocus()

    def enable_and_set_focus_control_dialog(self):
        """Enable and set focus to control dialog"""
        self.enable_control_dialog_window(True)
        self.set_focus_control_dialog()

    def get_control_dialog_window(self):
        """Get control dialog window"""
        return self._x_control_dialog

    def show_message_box(self, s_title: str, s_message: str):
        """Show message box dialog"""
        try:
            o_toolkit = self._x_context.getServiceManager().createInstanceWithContext("com.sun.star.awt.Toolkit", self._x_context)
            x_toolkit = o_toolkit

            if self._x_frame is not None and x_toolkit is not None:
                a_descriptor = WindowDescriptor()
                a_descriptor.Type = MODALTOP
                a_descriptor.WindowServiceName = "infobox"
                a_descriptor.ParentIndex = -1
                a_descriptor.Parent = self._x_frame.getContainerWindow()
                a_descriptor.WindowAttributes = WindowAttribute.BORDER | WindowAttribute.MOVEABLE | WindowAttribute.CLOSEABLE

                x_message_box = x_toolkit.createWindow(a_descriptor)
                if x_message_box is not None:
                    x_message_box.CaptionText = s_title
                    x_message_box.MessageText = s_message
                    self.enable_control_dialog_window(False)
                    x_message_box.execute()
                    x_component = x_message_box
                    if x_component is not None:
                        x_component.dispose()
                    self.enable_control_dialog_window(True)
                    self.set_focus_control_dialog()
        except Exception as ex:
            print(f"Error showing message box: {ex}")

    def get_dialog_property_value(self, dialog_name: str, property_name: str) -> str:
        """Get dialog property value from resource file"""
        result = None
        x_resources = None
        m_res_root_url = self.get_package_location() + "/dialog/"

        try:
            args = (m_res_root_url, True, self.get_locale(), dialog_name, '', uno.Any("com.sun.star.task.XInteractionHandler", None))
            x_resources = uno.invoke(self._x_context.ServiceManager, 'createInstanceWithArgumentsAndContext', ('com.sun.star.resource.StringResourceWithLocation', args, self._x_context))
        except Exception as ex:
            print(f"Error creating string resource: {ex}")

        # Map properties
        if x_resources is not None:
            ids = x_resources.getResourceIDs()
            for resource_id in ids:
                if property_name in resource_id:
                    result = x_resources.resolveString(resource_id)

        return result

    def get_package_location(self) -> str:
        """Get package location from package information provider"""
        location = None
        try:
            x_name_access = self._x_context
            x_pip = x_name_access.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
            location = x_pip.getPackageLocation("com.collabora.milsymbol")
        except Exception as ex:
            print(f"Error getting package location: {ex}")
        return location

    def get_locale(self):
        """Get locale from configuration provider"""
        locale = None
        try:
            x_mcf = self._x_context.getServiceManager()
            o_configuration_provider = x_mcf.createInstanceWithContext(
                "com.sun.star.configuration.ConfigurationProvider",
                self._x_context
            )
            x_localizable = o_configuration_provider
            locale = x_localizable.getLocale()
        except Exception as ex:
            print(f"Error getting locale: {ex}")
        return locale

class ControlDlgHandler(unohelper.Base, XDialogEventHandler, XTopWindowListener):
    buttons = ["addShape", "removeShape", "editShape"]

    def __init__(self, dialog):
        self.dialog = dialog
        self.tree_control = None
        self._populate_tree_on_show = True

    def callHandlerMethod(self, dialog, eventObject, methodName):
        if methodName == "OnAdd":
            if self.get_controller().get_diagram() is not None:
                self.get_controller().remove_selection_listener()
                if self.get_controller().get_group_type() == self.get_controller().ORGANIGROUP:
                    org_chart = self.get_controller().get_diagram()
                    if org_chart.is_error_in_tree():
                        self.get_gui().ask_user_for_repair(org_chart)
                    else:
                        self.get_controller().get_diagram().add_shape()
                        self.get_controller().get_diagram().refresh_diagram()
                        # Refresh tree after adding shape
                        self.refresh_tree()
                else:
                    self.get_controller().get_diagram().add_shape()
                    self.get_controller().get_diagram().refresh_diagram()
                    # Refresh tree after adding shape
                    self.refresh_tree()
                self.get_controller().add_selection_listener()
                self.get_controller().set_text_field_of_control_dialog()
            return True
        elif methodName == "OnRemove":
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
                self.get_controller().set_text_field_of_control_dialog()
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

    # XTopWindowListener methods
    def windowClosing(self, event):
        """Handle window closing event"""
        if event.Source == self.get_gui()._x_control_dialog:
            self.get_gui().set_visible_control_dialog(False)

    def windowOpened(self, event):
        """Handle window opened event"""
        if event.Source == self.get_gui()._x_control_dialog:
            # Initialize tree control when dialog opens
            self._init_tree_control()
            # Populate tree with current diagram structure
            if self._populate_tree_on_show:
                self.populate_tree()
                self._populate_tree_on_show = False

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

    def _init_tree_control(self):
        """Initialize the tree control"""
        try:
            dialog = self.get_gui()._x_control_dialog
            if dialog is not None:
                # Get the tree control from the dialog
                self.tree_control = dialog.getControl("OrbatTree")
                if self.tree_control is not None:
                    print("Tree control initialized successfully")
                else:
                    print("Warning: Could not find OrbatTree control in dialog")
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

            # Get or create the tree data model
            ctx = self.get_gui()._x_context
            service_manager = ctx.getServiceManager()

            # Create a new tree data model
            try:
                data_model = service_manager.createInstanceWithContext(
                    "com.sun.star.awt.tree.MutableTreeDataModel", ctx)

                if data_model is None:
                    print("Could not create tree data model")
                    return

                # Create root node
                root_node = data_model.createNode("Diagram Structure", True)
                data_model.setRoot(root_node)

                # Set the data model to the tree control
                tree_model = self.tree_control.getModel()
                tree_model.setPropertyValue("DataModel", data_model)
                tree_model.setPropertyValue("SelectionType", SELECTION_TYPE_SINGLE)
                tree_model.setPropertyValue("RootDisplayed", True)
                tree_model.setPropertyValue("ShowsHandles", True)
                tree_model.setPropertyValue("ShowsRootHandles", True)
                tree_model.setPropertyValue("Editable", False)

                # Add mouse listener for tree click handling
                tree_mouse_handler = TreeMouseHandler(self)
                self.tree_control.addMouseListener(tree_mouse_handler)

            except Exception as e:
                print(f"Error creating tree data model: {e}")
                return

            # Get diagram tree structure
            diagram_tree = None
            if hasattr(diagram, 'get_diagram_tree'):
                diagram_tree = diagram.get_diagram_tree()
            elif hasattr(diagram, '_diagram_tree'):
                diagram_tree = diagram._diagram_tree

            if diagram_tree is not None and hasattr(diagram_tree, 'get_root_item'):
                root_item = diagram_tree.get_root_item()
                if root_item is not None:
                    # Update root node display name
                    root_name = self._get_tree_node_display_name(root_item, 1)
                    if not root_name or root_name == "Item 1":
                        root_name = "Root"
                    root_node.setDisplayValue(root_name)

                    # Populate children of root item
                    self._populate_tree_children(data_model, root_node, root_item)
                else:
                    print("No root item found in diagram tree")
            else:
                print("No diagram tree available or invalid structure")

            # Expand the tree to show structure
            if hasattr(self.tree_control, 'expandNode') and root_node:
                try:
                    self.tree_control.expandNode(root_node)
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
            if hasattr(tree_item, 'get_first_child') and tree_item.get_first_child() is not None:
                child_item = tree_item.get_first_child()
                child_num = 1
                while child_item is not None:
                    child_name = self._get_tree_node_display_name(child_item, child_num)
                    self._populate_tree_node(data_model, parent_node, child_item, child_name)

                    # Move to next sibling
                    child_item = child_item.get_first_sibling() if hasattr(child_item, 'get_first_sibling') else None
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

            # Add children
            child_count = 0
            if hasattr(tree_item, 'get_first_child') and tree_item.get_first_child() is not None:
                child_item = tree_item.get_first_child()
                child_num = 1
                while child_item is not None:
                    child_name = self._get_tree_node_display_name(child_item, child_num)
                    self._populate_tree_node(data_model, node, child_item, child_name)

                    # Move to next sibling
                    child_item = child_item.get_first_sibling() if hasattr(child_item, 'get_first_sibling') else None
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
                if hasattr(self, '_node_to_tree_item_map'):
                    self._node_to_tree_item_map.clear()
                self.populate_tree()
        except Exception as e:
            print(f"Error refreshing tree: {e}")

    def _get_tree_node_display_name(self, tree_item, item_number):
        """Get a meaningful display name for a tree node"""
        try:
            # Try to get shape information
            if hasattr(tree_item, 'get_rectangle_shape'):
                shape = tree_item.get_rectangle_shape()
                if shape is not None:
                    shape_name = ""

                    # Try to get shape name
                    try:
                        if hasattr(shape, 'getName'):
                            shape_name = shape.getName()
                    except:
                        pass

                    # Try to get shape text/string content
                    shape_text = ""
                    try:
                        if hasattr(shape, 'getString'):
                            shape_text = shape.getString()
                        elif hasattr(shape, 'getText'):
                            text_obj = shape.getText()
                            if text_obj and hasattr(text_obj, 'getString'):
                                shape_text = text_obj.getString()
                    except:
                        pass

                    # Try to get level information
                    level_info = ""
                    try:
                        if hasattr(tree_item, 'get_level'):
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
            if selected_node is not None and hasattr(selected_node, 'getDisplayValue'):
                node_name = selected_node.getDisplayValue()

                # Get the tree item associated with this node
                if hasattr(self, '_node_to_tree_item_map'):
                    tree_item = self._node_to_tree_item_map.get(node_name)
                    if tree_item and hasattr(tree_item, 'get_rectangle_shape'):
                        shape = tree_item.get_rectangle_shape()
                        if shape:
                            # Select the shape in the document
                            controller = self.get_controller()
                            if hasattr(controller, 'set_selected_shape'):
                                controller.set_selected_shape(shape)
                            elif hasattr(controller, '_x_controller'):
                                # Try to select through document controller
                                try:
                                    controller._x_controller.select(shape)
                                except:
                                    print("Could not select shape through controller")
                            print(f"Selected shape for node: {node_name}")
                    else:
                        print(f"No tree item found for node: {node_name}")

        except Exception as e:
            print(f"Error handling tree selection: {e}")


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
                        selected_node = tree_control.getNodeForLocation(event.X, event.Y)
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
