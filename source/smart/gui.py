# SPDX-License-Identifier: MPL-2.0

import unohelper

from com.sun.star.awt import XDialogEventHandler, XTopWindowListener
from com.sun.star.awt import WindowDescriptor, WindowAttribute
from com.sun.star.awt.WindowClass import MODALTOP

class Gui:
    def __init__(self, controller, x_context, x_frame):
        """Initialize GUI"""
        self._controller = controller
        self._x_context = x_context
        self._x_frame = x_frame
        self._x_control_dialog = None
        self._x_control_dialog_window = None
        self._x_control_dialog_top_window = None
        self._is_visible_control_dialog = False
        self._o_listener = None

    def get_controller(self):
        """Get controller reference"""
        return self._controller

    def execute_properties_dialog(self):
        """Execute properties dialog of the current diagram"""
        # TODO: Open the "edit symbol" dialog

    def set_visible_control_dialog(self, visible: bool):
        """Set visibility of control dialog"""
        new_diagram_id = self.get_controller().get_diagram().get_diagram_id()

        # Need to create new controlDialog when a new diagram selected (not same GUI panels of diagrams)
        if ((self.get_controller().get_last_diagram_type() != -1 or
             self.get_controller().get_last_diagram_id() != -1) and
            (self.get_controller().get_last_diagram_type() != self.get_controller().get_diagram_type() or
             self.get_controller().get_last_diagram_id() != new_diagram_id)):
            if self._x_control_dialog_window is not None:
                self._x_control_dialog_window.setVisible(False)
            self.create_control_dialog()

        if self._x_control_dialog_window is None:
            self.create_control_dialog()

        if self._x_control_dialog_window is not None:
            self._x_control_dialog_window.setVisible(visible)
            if visible:
                self._is_visible_control_dialog = True
                self._x_control_dialog_window.setFocus()
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
                self._x_control_dialog_window = self._x_control_dialog
                self._x_control_dialog_top_window = self._x_control_dialog
                self._x_control_dialog_top_window.addTopWindowListener(self._o_listener)

        except Exception as ex:
            print(f"Error creating control dialog: {ex}")

    def close_and_dispose_control_dialog(self):
        """Close and dispose control dialog"""
        if self._x_control_dialog_top_window is not None:
            self._x_control_dialog_top_window.removeTopWindowListener(self._o_listener)

        if self._x_control_dialog_window is not None:
            self._x_control_dialog_window.setVisible(False)
            x_comp = self._x_control_dialog_window
            if x_comp is not None:
                x_comp.dispose()

        self._x_control_dialog_top_window = None
        self._x_control_dialog_window = None
        self._x_control_dialog = None

    def is_visible_control_dialog(self):
        """Check if control dialog is visible"""
        return self._is_visible_control_dialog

    def enable_control_dialog_window(self, enable: bool):
        """Enable or disable control dialog window"""
        if self._x_control_dialog_window is not None:
            self._x_control_dialog_window.setEnable(enable)

    def set_focus_control_dialog(self):
        """Set focus to control dialog"""
        if self._x_control_dialog_window is not None:
            self._x_control_dialog_window.setFocus()

    def enable_and_set_focus_control_dialog(self):
        """Enable and set focus to control dialog"""
        self.enable_control_dialog_window(True)
        self.set_focus_control_dialog()

    def get_control_dialog_window(self):
        """Get control dialog window"""
        return self._x_control_dialog_window

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
            args = (m_res_root_url, True, self.get_locale(), dialog_name, '', None)
            x_resources = self._x_context.ServiceManager.createInstanceWithArgumentsAndContext('com.sun.star.resource.StringResourceWithLocation', args, self._x_context)
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

    def callHandlerMethod(self, dialog, eventObject, methodName):
        print("callHandlerMethod ", methodName)
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
                else:
                    self.get_controller().get_diagram().add_shape()
                    self.get_controller().get_diagram().refresh_diagram()
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
                else:
                    self.get_controller().get_diagram().remove_shape()
                    self.get_controller().get_diagram().refresh_diagram()
                self.get_controller().add_selection_listener()
                self.get_controller().set_text_field_of_control_dialog()
            return True
        elif methodName == "OnEdit":
            self.get_controller().get_diagram().show_edit_dialog()
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
        if event.Source == self.get_gui()._x_control_dialog_top_window:
            self.get_gui().set_visible_control_dialog(False)

    def windowOpened(self, event):
        """Handle window opened event"""
        pass

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