# SPDX-License-Identifier: MPL-2.0

import sys
import unohelper
import officehelper
import tempfile
import os
import uno

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from dialog_handler import DialogHandler
from smart.controller import Controller
from smart.gui import Gui
from smart.diagram.organizationcharts.orgchart.orgchart import OrgChart
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
            self.onSymbolDialog()
        if args == "testSymbol":
            self.insertSymbol(self.model, "sfgpewrh--mt")
        if args == "orgChart":
            self.onOrgChart()

    def onSymbolDialog(self):
        dialog_provider = self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.awt.DialogProvider2", self.ctx)
        dialog_url = "vnd.sun.star.extension://com.collabora.milsymbol/dialog/MilitarySymbolDlg.xdl"
        try:
            handler = DialogHandler(self)
            dialog = dialog_provider.createDialogWithHandler(dialog_url, handler)
            handler.dialog = dialog
            handler.initialize_dialog_controls(dialog)

            dialog.Model.Step = 1

            dialog.execute()
        except Exception as e:
            pass

    def onOrgChart(self):
        """Create a simple organization chart"""
        print("Creating simple organization chart...")

        # Create controller and GUI instances (stub implementations)
        controller = Controller(None, self.ctx, self.desktop.getCurrentFrame())
        gui = Gui()
        
        # Create organization chart
        #org_chart = OrgChart(controller, gui, None)
        
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
                self.insertSvgGraphic(model, str(result[0]))

        except Exception as e:
            print(f"Error executing script: {e}")
            return

    def insertSvgGraphic(self, model, svg_data):
        graphic_provider = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.graphic.GraphicProvider", self.ctx)

        with tempfile.NamedTemporaryFile(mode='w', suffix='.svg', delete=False) as temp_file:
            temp_file.write(svg_data)
            temp_svg_path = temp_file.name
        try:
            from unohelper import systemPathToFileUrl
            svg_url = systemPathToFileUrl(temp_svg_path)

            media_properties = (PropertyValue("URL", 0, svg_url, 0),)
            graphic = graphic_provider.queryGraphic(media_properties)

            shape_factory = model.createInstance("com.sun.star.drawing.GraphicObjectShape")
            shape_factory.setPropertyValue("Graphic", graphic)

            # TODO: Make this dependent on actual SVG dimensions
            # (parse width and height properties of the svg element)
            shape_size = Size()
            shape_size.Height = 930
            shape_size.Width = 4000
            shape_factory.setSize(shape_size)

            # TODO: Set default anchoring for text documents
            #try:
            #    shape_factory.setPropertyValue("AnchorType", TextContentAnchorType.AT_PARAGRAPH)
            #except:
            #    pass  # Not all document types support anchoring

            # Writer
            if model.supportsService("com.sun.star.text.TextDocument"):
                controller = model.getCurrentController()
                view_cursor_supplier = controller
                cursor = view_cursor_supplier.getViewCursor()
                text = cursor.getText()
                text.insertTextContent(cursor, shape_factory, True)
            # Calc
            elif model.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
                current_selection = model.getCurrentSelection()
                try:
                    cell_position = current_selection.getPropertyValue("Position")
                    shape_factory.setPosition(cell_position)
                except:
                    # Default position if we can't get cell position
                    default_pos = Point()
                    default_pos.X = 1000
                    default_pos.Y = 1000
                    shape_factory.setPosition(default_pos)
                controller = model.getCurrentController()
                active_sheet = controller.getActiveSheet()
                draw_page_supplier = active_sheet
                draw_page = draw_page_supplier.getDrawPage()
                draw_page.add(shape_factory)
            # Impress/Draw
            elif (model.supportsService("com.sun.star.presentation.PresentationDocument") or 
                    model.supportsService("com.sun.star.drawing.DrawingDocument")):
                controller = model.getCurrentController()
                current_page = controller.getCurrentPage()
                current_page.add(shape_factory)
            else:
                print("Unsupported document type for graphic insertion")
        finally:
            # Clean up temporary file
            try:
                os.unlink(temp_svg_path)
            except:
                pass

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
    job.trigger("hello")


# Starting from command line
if __name__ == "__main__":
    main()


# pythonloader loads a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    MainJob,  # UNO object class
    "com.collabora.milsymbol.do",  # implementation name (customize for yourself)
    ("com.sun.star.task.Job",), )  # implemented services (only 1)
