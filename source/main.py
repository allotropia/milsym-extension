# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import sys
import uno
import unohelper
import officehelper
import os

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from symbol_dialog import open_symbol_dialog
from smart.controller import Controller
from smart.diagram.data_of_diagram import DataOfDiagram
from sidebar import SidebarFactory
from utils import is_orbat_feature_enabled

from com.sun.star.task import XJobExecutor, XJob
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.util import XCloseListener
from com.sun.star.ui import XContextMenuInterceptor
from com.sun.star.ui.ContextMenuInterceptorAction import EXECUTE_MODIFIED
from com.sun.star.ui.ContextMenuInterceptorAction import IGNORED
from com.sun.star.frame import XDispatchProvider, XDispatch
from com.sun.star.lang import XInitialization
from translator import translate


class DocumentCloseListener(unohelper.Base, XCloseListener):
    def __init__(self, model, frame):
        self.model = model
        self.frame = frame

    def notifyClosing(self, event):
        ListenerRegistry.instance().unregister(self.model)
        ControllerManager().remove_controller(self.frame)

    def queryClosing(self, event, getsOwnership):
        """Clean up BEFORE document closes and shapes are destroyed.

        This is critical for preventing memory leaks and crashes. We must clear
        all Python references to UNO shapes before LibreOffice destroys the
        underlying C++ objects.
        """
        try:
            # First, clear all shape references via the controller
            controller = ControllerManager().get_controller_for_frame(self.frame)
            if controller is not None:
                controller.dispose_diagram()

            # Then clear the undo stack
            if hasattr(self.model, "getUndoManager"):
                undo_manager = self.model.getUndoManager()
                if undo_manager is not None:
                    undo_manager.clear()
        except Exception as e:
            print(f"Error in queryClosing cleanup: {e}")
        return True

    def disposing(self, event):
        pass


class ListenerRegistry:
    _instance = None

    def __init__(self):
        self._map = {}
        self._registered = set()
        self._interceptors = {}
        self.selected_shape = None

    @classmethod
    def instance(cls):
        if cls._instance is None:
            cls._instance = ListenerRegistry()
        return cls._instance

    def has(self, xcontroller):
        return xcontroller in self._registered

    def register(self, model, xcontroller, listener):
        self._registered.add(xcontroller)
        self._map[model] = (xcontroller, listener)

    def register_interceptor(self, xcontroller, interceptor):
        self._interceptors[xcontroller] = interceptor

    def unregister(self, model):
        if model in self._map:
            xcontroller, listener = self._map[model]
            try:
                xcontroller.removeSelectionChangeListener(listener)
                print("Removed selection listener from controller")
            except Exception as e:
                print("Failed removing listener:", e)

            del self._map[model]

    def update_selected_shape(self, selection):
        self.selected_shape = selection

    def get_selected_shape(self):
        return self.selected_shape

    def clear_selected_shape(self):
        self.selected_shape = None


class ControllerManager:
    """Manages Controller instances for different frames"""

    _instance = None
    _controllers = {}

    def __new__(cls, ctx=None):
        if cls._instance is None:
            cls._instance = super(ControllerManager, cls).__new__(cls)
            cls._instance.ctx = ctx
            cls._instance.orbat_enabled = is_orbat_feature_enabled(ctx)
        return cls._instance

    def get_or_create_controller(self, ctx, frame):
        """Get existing controller or create new one for the frame"""
        if frame not in self._controllers:
            # Check if document has existing smart diagrams
            model = frame.getController().getModel()
            if self.document_has_smart_diagrams(model):
                controller = Controller(None, ctx, frame)
                self._controllers[frame] = controller
                return controller
        return self._controllers.get(frame)

    def remove_controller(self, frame):
        """Remove controller for frame and dispose it properly"""
        if frame in self._controllers:
            controller = self._controllers[frame]
            try:
                controller.dispose()
            except Exception as e:
                print(f"Error disposing controller: {e}")
            del self._controllers[frame]

    def get_controller_for_frame(self, frame):
        """Get controller for a given frame, if one exists"""
        return self._controllers.get(frame)

    def document_has_smart_diagrams(self, model):
        """Check if document contains smart diagram shapes"""
        if not self.orbat_enabled:
            return False

        try:
            # Check different document types for shapes
            if model.supportsService("com.sun.star.text.TextDocument"):
                draw_page = model.getDrawPage()
                return self.check_shapes_for_diagrams(draw_page)
            elif model.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
                sheets = model.getSheets()
                for i in range(sheets.getCount()):
                    sheet = sheets.getByIndex(i)
                    draw_page = sheet.getDrawPage()
                    if self.check_shapes_for_diagrams(draw_page):
                        return True
            elif model.supportsService(
                "com.sun.star.presentation.PresentationDocument"
            ) or model.supportsService("com.sun.star.drawing.DrawingDocument"):
                draw_pages = model.getDrawPages()
                for i in range(draw_pages.getCount()):
                    draw_page = draw_pages.getByIndex(i)
                    if self.check_shapes_for_diagrams(draw_page):
                        return True
        except Exception as e:
            print(f"Error checking for smart diagrams: {e}")
        return False

    def check_shapes_for_diagrams(self, draw_page):
        """Check if a draw page contains smart diagram shapes"""
        try:
            for i in range(draw_page.getCount()):
                shape = draw_page.getByIndex(i)
                if hasattr(shape, "getName"):
                    shape_name = shape.getName()
                    if shape_name.startswith("OrbatDiagram"):
                        return True
        except Exception as e:
            print(f"Error checking shapes: {e}")
        return False


class ContextMenuInterceptor(unohelper.Base, XContextMenuInterceptor):
    """Intercepts context menu to add 'Edit Military Symbol' and 'Edit Orbat' options"""

    def __init__(self, ctx):
        self.ctx = ctx
        self.orbat_enabled = is_orbat_feature_enabled(ctx)

    def notifyContextMenuExecute(self, event):
        """Called when a context menu is about to be displayed"""
        if not self.orbat_enabled:
            return IGNORED

        try:
            orbat_shape = self._get_orbat_group_shape(event)
            if orbat_shape is not None:
                menu_container = event.ActionTriggerContainer
                self._insert_edit_orbat_menu_item(menu_container)
                return EXECUTE_MODIFIED

            shape = ListenerRegistry.instance().get_selected_shape()

            if shape is None:
                return IGNORED

            try:
                attrs = shape.UserDefinedAttributes
                if not attrs or not attrs.hasByName("MilSymCode"):
                    return IGNORED
            except:
                return IGNORED

            menu_container = event.ActionTriggerContainer
            self._insert_menu_item(menu_container)

            return EXECUTE_MODIFIED

        except Exception as e:
            print(f"ContextMenuInterceptor error: {e}")
            return IGNORED

    def _get_orbat_group_shape(self, event):
        """Check if the selected shape is an ORBAT group shape"""
        try:
            controller = event.Selection
            if controller is None:
                return None

            selection = (
                controller.getSelection()
                if hasattr(controller, "getSelection")
                else controller
            )

            if selection is None:
                return None

            # Check if it's a single shape or a collection
            shape = None
            if selection.supportsService("com.sun.star.drawing.GroupShape"):
                shape = selection
            elif selection.supportsService("com.sun.star.drawing.Shapes"):
                if selection.getCount() == 1:
                    shape = selection.getByIndex(0)
                    if not shape.supportsService("com.sun.star.drawing.GroupShape"):
                        return None

            if shape is None:
                return None

            shape_name = shape.getName() if hasattr(shape, "getName") else ""
            if shape_name.startswith("OrbatDiagram"):
                return shape

            return None
        except Exception as e:
            print(f"Error checking for ORBAT group: {e}")
            return None

    def _insert_edit_orbat_menu_item(self, menu_container):
        """Insert 'Edit Orbat' menu item"""
        try:
            menu_item = menu_container.createInstance("com.sun.star.ui.ActionTrigger")

            menu_text = translate(self.ctx, "ContextMenu.EditOrbat")
            menu_item.setPropertyValue("Text", menu_text)
            menu_item.setPropertyValue(
                "CommandURL", "service:com.collabora.milsymbol.do?editOrbat"
            )

            separator = menu_container.createInstance(
                "com.sun.star.ui.ActionTriggerSeparator"
            )

            menu_container.insertByIndex(0, separator)
            menu_container.insertByIndex(0, menu_item)
        except Exception as e:
            print(f"_insert_edit_orbat_menu_item error: {e}")

    def _insert_menu_item(self, menu_container):
        """Insert 'Edit Military Symbol' menu item"""
        try:
            menu_item = menu_container.createInstance("com.sun.star.ui.ActionTrigger")

            menu_text = translate(self.ctx, "ContextMenu.EditMilitarySymbol")
            menu_item.setPropertyValue("Text", menu_text)
            menu_item.setPropertyValue(
                "CommandURL", "service:com.collabora.milsymbol.do?symbolDialog"
            )

            separator = menu_container.createInstance(
                "com.sun.star.ui.ActionTriggerSeparator"
            )

            menu_container.insertByIndex(0, separator)
            menu_container.insertByIndex(0, menu_item)
        except Exception as e:
            print(f"_insert_menu_item error: {e}")


class StartupJob(unohelper.Base, XJob, XSelectionChangeListener):
    """Job that runs on document events to initialize controllers"""

    def __init__(self, ctx):
        self.ctx = ctx

    def execute(self, args):
        """Execute method called by LibreOffice Job framework"""
        self.initialize_controllers()

    def selectionChanged(self, event):
        try:
            ListenerRegistry.instance().clear_selected_shape()
            selection = event.Source.getSelection()
            if selection is None:
                return

            if selection.supportsService("com.sun.star.text.TextGraphicObject"):
                shape = selection
                try:
                    if (
                        hasattr(shape, "UserDefinedAttributes")
                        and shape.UserDefinedAttributes
                    ):
                        ListenerRegistry.instance().update_selected_shape(shape)
                except Exception:
                    pass
            elif selection.supportsService("com.sun.star.drawing.Shapes"):
                if selection.getCount() > 0:
                    shape = selection.getByIndex(0)
                    try:
                        if (
                            hasattr(shape, "UserDefinedAttributes")
                            and shape.UserDefinedAttributes
                        ):
                            ListenerRegistry.instance().update_selected_shape(shape)
                    except Exception:
                        pass

        except Exception as e:
            print(f"Error document selection listener: {e}")

    def disposing(self, event):
        """Handle disposing event from XEventListener"""
        pass

    def initialize_controllers(self):
        """Initialize controllers for documents with smart diagrams"""
        try:
            controller_manager = ControllerManager(self.ctx)
            desktop = self.ctx.getServiceManager().createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx
            )

            # Check all open documents
            frames = desktop.getFrames()
            if frames.getCount() > 0:
                for i in range(frames.getCount()):
                    frame = frames.getByIndex(i)

                    xcontroller = frame.getController()
                    if xcontroller:
                        registry = ListenerRegistry.instance()
                        if not registry.has(xcontroller):
                            model = xcontroller.getModel()
                            xcontroller.addSelectionChangeListener(self)
                            ListenerRegistry.instance().register(
                                model, xcontroller, self
                            )
                            model.addCloseListener(DocumentCloseListener(model, frame))
                            interceptor = ContextMenuInterceptor(self.ctx)
                            xcontroller.registerContextMenuInterceptor(interceptor)
                            ListenerRegistry.instance().register_interceptor(
                                xcontroller, interceptor
                            )

                    try:
                        controller_manager.get_or_create_controller(self.ctx, frame)
                    except Exception as e:
                        print(f"Error initializing controller for frame {i}: {e}")
            else:
                print("StartupJob: No open documents found")
        except Exception as e:
            print(f"Error in StartupJob initialization: {e}")


# The MainJob is a UNO component derived from unohelper.Base class
# and also the XJobExecutor, the implemented interface
class MainJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx, orbat_enabled=True):
        self.ctx = ctx
        self.orbat_enabled = orbat_enabled
        # handling different situations (inside LibreOffice or other process)
        try:
            self.sm = ctx.getServiceManager()
            self.desktop = XSCRIPTCONTEXT.getDesktop()
        except NameError:
            self.sm = ctx.ServiceManager
            self.desktop = self.ctx.getServiceManager().createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx
            )

    def trigger(self, args):
        desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx
        )
        self.model = desktop.getCurrentComponent()

        if args == "symbolDialog":
            selected_shape = ListenerRegistry.instance().get_selected_shape()
            open_symbol_dialog(self.ctx, self.model, None, None, selected_shape, None)
        if self.orbat_enabled and args == "orgChart":
            self.onOrgChart()
        if self.orbat_enabled and args == "editOrbat":
            self.onEditOrbat()

    def onEditOrbat(self):
        """Open the ORBAT dialog for the currently selected ORBAT group"""
        frame = self.desktop.getCurrentFrame()
        controller_manager = ControllerManager(self.ctx)

        if frame in controller_manager._controllers:
            controller = controller_manager._controllers[frame]
        else:
            controller = controller_manager.get_or_create_controller(self.ctx, frame)
            if controller is None:
                return

        xcontroller = frame.getController()
        selection = xcontroller.getSelection()
        if selection is None:
            return

        shape = None
        if selection.supportsService("com.sun.star.drawing.GroupShape"):
            shape = selection
        elif selection.supportsService("com.sun.star.drawing.Shapes"):
            if selection.getCount() == 1:
                shape = selection.getByIndex(0)

        if shape is None:
            return

        shape_name = shape.getName() if hasattr(shape, "getName") else ""
        if not shape_name.startswith("OrbatDiagram"):
            return

        diagram_name = shape_name.split("-", 1)[0]
        diagram_id = int("".join(c for c in diagram_name if c.isdigit()) or "0")

        controller.set_group_type(controller.ORGANIGROUP)
        controller.set_diagram_type(controller.ORGANIGRAM)
        controller.instantiate_diagram()
        controller._last_diagram_name = diagram_name
        controller.get_diagram().init_diagram(diagram_id)
        controller.get_diagram().init_properties()

        controller._gui.set_visible_control_dialog(True)

    def onOrgChart(self):
        """Create a simple organization chart"""
        frame = self.desktop.getCurrentFrame()

        controller_manager = ControllerManager(self.ctx)

        if frame in controller_manager._controllers:
            controller = controller_manager._controllers[frame]
        else:
            controller = Controller(None, self.ctx, frame)
            controller_manager._controllers[frame] = controller

            xcontroller = frame.getController()
            model = xcontroller.getModel()
            registry = ListenerRegistry.instance()
            if not registry.has(xcontroller):
                model.addCloseListener(DocumentCloseListener(model, frame))

        # Create hierarchical data
        data = DataOfDiagram()
        data.add(0, "")  # Level 0 (root)
        data.add(1, "")  # Level 1
        data.add(1, "")  # Level 1
        data.add(1, "")  # Level 1

        # Create the diagram and show the dialog
        controller.create_diagram(data)
        controller._gui.set_visible_control_dialog(True)

    def initialize_controllers_for_open_documents(self):
        """Initialize controllers for all currently open documents that contain smart diagrams"""
        try:
            # Get all open frames
            frames = self.desktop.getFrames()
            for i in range(frames.getCount()):
                frame = frames.getByIndex(i)
                try:
                    model = frame.getController().getModel()
                    if model and self.controller_manager.document_has_smart_diagrams(
                        model
                    ):
                        self.controller_manager.get_or_create_controller(
                            self.ctx, frame
                        )
                except Exception as e:
                    print(f"Error initializing controller for frame {i}: {e}")
        except Exception as e:
            print(f"Error initializing controllers for open documents: {e}")


class Dispatcher(unohelper.Base, XDispatch):
    """Dispatch handler for com.collabora.milsymbol protocol URLs.

    Added to control menu item enabled/disabled state. Not dynamic,
    only gets called once.
    """

    def __init__(self, ctx, orbat_enabled=True):
        self.ctx = ctx
        self.orbat_enabled = orbat_enabled
        self.job = MainJob(self.ctx, orbat_enabled)

    def dispatch(self, url, args):
        self.job.trigger(url.Path)

    def addStatusListener(self, listener, url):
        event = uno.createUnoStruct("com.sun.star.frame.FeatureStateEvent")
        event.Source = self
        event.FeatureURL = url
        event.IsEnabled = True
        if url.Path == 'orgChart' and not self.orbat_enabled:
            event.IsEnabled = False
        event.Requery = False
        listener.statusChanged(event)

    def removeStatusListener(self, listener, url):
        pass


class ProtocolHandler(unohelper.Base, XDispatchProvider, XInitialization):
    """Protocol handler for com.collabora.milsymbol protocol URLs.

    Registered via ProtocolHandler.xcu. Used to insert XDispatch objects that control
    menu item state via the dispatch framework.
    """

    def __init__(self, ctx):
        self.ctx = ctx
        self._frame = None
        self.orbat_enabled = is_orbat_feature_enabled(ctx)

    def initialize(self, args):
        if args:
            self._frame = args[0]

    def queryDispatch(self, url, target_frame_name, search_flags):
        if url.Protocol == "com.collabora.milsymbol:":
            # create our custom dispatcher, which can disable menu items
            return Dispatcher(self.ctx, self.orbat_enabled)
        return None

    def queryDispatches(self, requests):
        return [
            self.queryDispatch(r.FeatureURL, r.FrameName, r.SearchFlags)
            for r in requests
        ]


# Starting from Python IDE
def main():
    try:
        ctx = XSCRIPTCONTEXT
    except NameError:
        ctx = officehelper.bootstrap()
        if ctx is None:
            print("ERROR: Could not bootstrap default Office.")
            sys.exit(1)
    job = MainJob(ctx)
    job.trigger("symbolDialog")


# Starting from command line
if __name__ == "__main__":
    main()


# pythonloader loads a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    MainJob,  # UNO object class
    "com.collabora.milsymbol.do",  # implementation name (customize for yourself)
    ("com.sun.star.task.Job",),
)  # implemented services (only 1)

g_ImplementationHelper.addImplementation(
    SidebarFactory,
    "com.collabora.milsymbol.TacticalSymbolsFactory",
    ("com.sun.star.ui.UIElementFactory",),
)

g_ImplementationHelper.addImplementation(
    StartupJob,
    "com.collabora.milsymbol.StartupJob",
    ("com.sun.star.task.Job",),
)

g_ImplementationHelper.addImplementation(
    ProtocolHandler,
    "com.collabora.milsymbol.ProtocolHandler",
    ("com.sun.star.frame.ProtocolHandler",),
)

