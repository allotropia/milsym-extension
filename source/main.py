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

from com.sun.star.awt import Size, Point
from com.sun.star.beans import NamedValue, PropertyValue
from com.sun.star.task import XJobExecutor, XJob





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


class StartupJob(unohelper.Base, XJob):
    """Job that runs on document events to initialize controllers"""

    def __init__(self, ctx):
        self.ctx = ctx

    def execute(self, args):
        """Execute method called by LibreOffice Job framework"""
        self.initialize_controllers()

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
            open_symbol_dialog(self.ctx, self.model, None, None)
        if args == "testSymbol":
            self.insertSymbol(self.model, "sfgpewrh--mt")
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

        # Create the diagram (this calls stub methods)
        controller.create_diagram(data)

    def insertSymbol(self, model, code):
        factory = self.ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.script.provider.MasterScriptProviderFactory", self.ctx)
        provider = factory.createScriptProvider(model)
        script = provider.getScript(
            "vnd.sun.star.script:milsymbol.milsymbol.js?language=JavaScript&location=user:uno_packages/" +
            getExtensionBasePath(self.ctx))

        args = (
            code,
            NamedValue("size", 35.0),
            NamedValue("quantity", 200.0),
            NamedValue("staffComments", "FOR REINFORCEMENTS"),
            NamedValue("additionalInformation", "ADDED SUPPORT FOR JJ"),
            NamedValue("direction", (750.0 * 360.0) / 6400.0),
            NamedValue("type", "MACHINE GUN"),
            NamedValue("dtg", "30140000ZSEP97"),
            NamedValue("location", "0900000.0E570306.0N")
        )

        try:
            result = script.invoke(args, (), ())
            # Assuming the result contains SVG data
            if result and len(result) > 0:
                insertSvgGraphic(self.ctx, model,
                                 str(result[0]),args)

        except Exception as e:
            print(f"Error executing script: {e}")
            return

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
