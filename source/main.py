# SPDX-License-Identifier: MPL-2.0

import sys
import unohelper
import officehelper
import os
import uno

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from symbol_dialog import open_symbol_dialog
from utils import insertSvgGraphic
from smart.controller import Controller
from smart.diagram.data_of_diagram import DataOfDiagram

from com.sun.star.awt import Size, Point
from com.sun.star.beans import NamedValue, PropertyValue
from com.sun.star.task import XJobExecutor
#from com.sun.star.text import TextContentAnchorType


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
            open_symbol_dialog(self.ctx, self.model, None)
        if args == "testSymbol":
            self.insertSymbol(self.model, "sfgpewrh--mt")
        if args == "orgChart":
            self.onOrgChart()

    def onOrgChart(self):
        """Create a simple organization chart"""
        print("Creating simple organization chart...")

        # Create controller and GUI instances (stub implementations)
        controller = Controller(None, self.ctx, self.desktop.getCurrentFrame())

        # Create hierarchical data
        data = DataOfDiagram()
        data.add(0, "CEO")                    # Level 0 (root)
        data.add(1, "Chief Technology Officer")  # Level 1
        data.add(1, "Chief Financial Officer")   # Level 1
        data.add(1, "Chief Marketing Officer")   # Level 1
        data.add(2, "Development Manager")       # Level 2 (under CTO)
        data.add(2, "QA Manager")               # Level 2 (under CTO)
        data.add(3, "Senior Developer")         # Level 3 (under Dev Manager)
        data.add(3, "Junior Developer")         # Level 3 (under Dev Manager)

        print(f"Created data with {data.size()} items")
        data.print_data()

        # Create the diagram (this calls stub methods)
        print("\nCreating diagram...")
        controller.create_diagram(data)

    def insertSymbol(self, model, code):
        factory = self.ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.script.provider.MasterScriptProviderFactory", self.ctx)
        provider = factory.createScriptProvider(model)
        script = provider.getScript(
            "vnd.sun.star.script:milsymbol.milsymbol.js?language=JavaScript&location=user:uno_packages/milsymbol-extension.oxt")

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
