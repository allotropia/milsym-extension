# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import sys
import unohelper
import officehelper
import os
import uno

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from symbol_dialog import open_symbol_dialog
from utils import insertSvgGraphic, getExtensionBasePath
from smart.controller import Controller
from smart.diagram.data_of_diagram import DataOfDiagram
from sidebar import SidebarFactory

from com.sun.star.beans import NamedValue
from com.sun.star.task import XJobExecutor, XJob
from com.sun.star.view import XSelectionChangeListener
from com.sun.star.util import XCloseListener


class DocumentCloseListener(unohelper.Base, XCloseListener):
    def __init__(self, model):
        self.model = model

    def notifyClosing(self, event):
        ListenerRegistry.instance().unregister(self.model)

    def queryClosing(self, event, getsOwnership):
        return True

    def disposing(self, event):
        pass


class ListenerRegistry:
    _instance = None

    def __init__(self):
        self._map = {}
        self._registered = set()
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
        """Remove controller for frame"""
        if frame in self._controllers:
            try:
                self._controllers[frame].remove_selection_listener()
            except:
                pass
            del self._controllers[frame]

    def document_has_smart_diagrams(self, model):
        """Check if document contains smart diagram shapes"""
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
            elif (model.supportsService("com.sun.star.presentation.PresentationDocument") or
                  model.supportsService("com.sun.star.drawing.DrawingDocument")):
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
                if hasattr(shape, 'getName'):
                    shape_name = shape.getName()
                    if shape_name.startswith("OrganizationDiagram"):
                        return True
        except Exception as e:
            print(f"Error checking shapes: {e}")
        return False


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
                if shape.UserDefinedAttributes:
                    ListenerRegistry.instance().update_selected_shape(shape)
            elif selection.supportsService("com.sun.star.drawing.Shapes"):
                shape = selection.getByIndex(0)
                if shape.UserDefinedAttributes:
                    ListenerRegistry.instance().update_selected_shape(shape)

        except Exception as e:
            print(f"Error document selection listener: {e}")

    def initialize_controllers(self):
        """Initialize controllers for documents with smart diagrams"""
        try:
            controller_manager = ControllerManager(self.ctx)
            desktop = self.ctx.getServiceManager().createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx)

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
                            ListenerRegistry.instance().register(model, xcontroller, self)
                            model.addCloseListener(DocumentCloseListener(model))

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
    def __init__(self, ctx):
        self.ctx = ctx
        # handling different situations (inside LibreOffice or other process)
        try:
            self.sm = ctx.getServiceManager()
            self.desktop = XSCRIPTCONTEXT.getDesktop()
        except NameError:
            self.sm = ctx.ServiceManager
            self.desktop = self.ctx.getServiceManager().createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx)

    def trigger(self, args):
        desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx)
        self.model = desktop.getCurrentComponent()

        if args == "symbolDialog":
            selected_shape = ListenerRegistry.instance().get_selected_shape()
            open_symbol_dialog(self.ctx, self.model, None, None, selected_shape, None)
        if args == "orgChart":
            self.onOrgChart()

    def onOrgChart(self):
        """Create a simple organization chart"""

        # Create controller and GUI instances (stub implementations)
        controller = Controller(None, self.ctx, self.desktop.getCurrentFrame())

        # Create hierarchical data
        data = DataOfDiagram()
        data.add(0, "")  # Level 0 (root)
        data.add(1, "")  # Level 1
        data.add(1, "")  # Level 1
        data.add(1, "")  # Level 1

        # Create the diagram
        controller.create_diagram(data)

    def initialize_controllers_for_open_documents(self):
        """Initialize controllers for all currently open documents that contain smart diagrams"""
        try:
            # Get all open frames
            frames = self.desktop.getFrames()
            for i in range(frames.getCount()):
                frame = frames.getByIndex(i)
                try:
                    model = frame.getController().getModel()
                    if model and self.controller_manager.document_has_smart_diagrams(model):
                        self.controller_manager.get_or_create_controller(self.ctx, frame)
                except Exception as e:
                    print(f"Error initializing controller for frame {i}: {e}")
        except Exception as e:
            print(f"Error initializing controllers for open documents: {e}")


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
    ("com.sun.star.task.Job",), )  # implemented services (only 1)

g_ImplementationHelper.addImplementation(
    SidebarFactory,
    "com.collabora.milsymbol.TacticalSymbolsFactory",
    ("com.sun.star.ui.UIElementFactory",), )

g_ImplementationHelper.addImplementation(
    StartupJob,
    "com.collabora.milsymbol.StartupJob",
    ("com.sun.star.task.Job",), )
