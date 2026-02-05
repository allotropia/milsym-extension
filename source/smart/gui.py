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

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from symbol_dialog import open_symbol_dialog
from control_dialog import ControlDlgHandler
from .. utils import get_package_location

from com.sun.star.awt import WindowDescriptor, WindowAttribute
from com.sun.star.awt.WindowClass import MODALTOP

class Gui:
    # Class-level variables to ensure only one dialog exists globally
    _global_control_dialog = None
    _global_control_dlg_listener = None
    # Track if user explicitly closed the dialog this session
    _user_closed_dialog = (
        False
    )

    def __init__(self, controller, x_context, x_frame):
        """Initialize GUI"""
        self._controller = controller
        self._x_context = x_context
        self._x_frame = x_frame
        self._x_control_dialog = None
        self._is_visible_control_dialog = False
        self._control_dlg_listener = None

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

        # Check if we need to recreate dialog for a different diagram
        need_new_dialog = False
        if Gui._global_control_dlg_listener is not None:
            try:
                global_controller = Gui._global_control_dlg_listener.get_controller()
                if global_controller != self.get_controller():
                    need_new_dialog = True  # Different document, need new dialog
            except Exception:
                need_new_dialog = True  # Can't get controller, recreate dialog

        if not need_new_dialog:
            if ((self.get_controller().get_last_diagram_type() != -1 or
                self.get_controller().get_last_diagram_id() != -1) and
                (self.get_controller().get_last_diagram_type() != self.get_controller().get_diagram_type() or
                self.get_controller().get_last_diagram_id() != new_diagram_id)):
                need_new_dialog = True

        if Gui._global_control_dialog is None or need_new_dialog:
            if need_new_dialog:
                self.close_and_dispose_control_dialog()
            self.create_control_dialog()

        # Sync instance reference from global
        self._x_control_dialog = Gui._global_control_dialog
        self._control_dlg_listener = Gui._global_control_dlg_listener

        if self._x_control_dialog is not None:
            if visible:
                self._x_control_dialog.setVisible(True)
                self._is_visible_control_dialog = True
                self._x_control_dialog.setFocus()
                # Refresh tree when dialog becomes visible
                if self._control_dlg_listener:
                    self._control_dlg_listener.refresh_tree()
            else:
                self._x_control_dialog.setVisible(False)
                self._is_visible_control_dialog = False

        self.get_controller().set_last_diagram_type(self.get_controller().get_diagram_type())
        self.get_controller().set_last_diagram_id(new_diagram_id)

    def create_control_dialog(self):
        """Create control dialog - ensures only one instance is created globally"""
        # If a global dialog exists, dispose it first to ensure clean handler binding
        if Gui._global_control_dialog is not None:
            self._dispose_global_dialog()

        try:
            dialog_provider = self._x_context.getServiceManager().createInstanceWithContext("com.sun.star.awt.DialogProvider2", self._x_context)
            s_dialog_url = "vnd.sun.star.extension://com.collabora.milsymbol/dialog/ControlDlg.xdl"

            # Create handler first
            model = self._x_frame.getController().getModel()
            new_listener = ControlDlgHandler(self, self._x_context, model)

            # Create dialog with handler to ensure proper binding
            new_dialog = dialog_provider.createDialogWithHandler(s_dialog_url, new_listener)

            if new_dialog is not None:
                new_dialog.addTopWindowListener(new_listener)

                # Store in both global and instance variables
                Gui._global_control_dialog = new_dialog
                Gui._global_control_dlg_listener = new_listener
                self._x_control_dialog = new_dialog
                self._control_dlg_listener = new_listener

        except Exception as ex:
            print(f"Error creating control dialog: {ex}")
            # Reset state on error to allow retry
            Gui._global_control_dialog = None
            Gui._global_control_dlg_listener = None
            self._x_control_dialog = None
            self._control_dlg_listener = None

    def _dispose_global_dialog(self):
        """Helper method to dispose the global dialog"""
        if Gui._global_control_dialog is not None:
            try:
                if Gui._global_control_dlg_listener is not None:
                    if hasattr(Gui._global_control_dlg_listener, 'cleanup'):
                        Gui._global_control_dlg_listener.cleanup()
                    Gui._global_control_dialog.removeTopWindowListener(Gui._global_control_dlg_listener)
                Gui._global_control_dialog.setVisible(False)

                # Dispose through XComponent interface
                x_component = Gui._global_control_dialog.queryInterface(uno.getTypeByName("com.sun.star.lang.XComponent"))
                if x_component is not None:
                    x_component.dispose()
                else:
                    Gui._global_control_dialog.dispose()
            except Exception as ex:
                print(f"Error disposing global dialog: {ex}")
            finally:
                Gui._global_control_dialog = None
                Gui._global_control_dlg_listener = None

    def close_and_dispose_control_dialog(self):
        """Close and dispose control dialog"""
        # Clear all references to prevent recreation during disposal
        self._x_control_dialog = None
        self._control_dlg_listener = None
        self._is_visible_control_dialog = False

        # Dispose the global dialog
        self._dispose_global_dialog()

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
        m_res_root_url = get_package_location(self._x_context) + "/dialog/"

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

