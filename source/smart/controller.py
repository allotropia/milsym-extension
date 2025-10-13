"""
Controller class for LibreOffice extension
Python port of Controller.java
"""

import uno
import unohelper
from com.sun.star.awt import Point, Size
from com.sun.star.beans import Property, PropertyValue, UnknownPropertyException, XPropertySet
from com.sun.star.container import NoSuchElementException, XEnumeration, XEnumerationAccess, XNamed
from com.sun.star.document import XExporter, XFilter
from com.sun.star.drawing import XDrawPage, XDrawPages, XDrawPagesSupplier, XDrawView, XShape, XShapes
from com.sun.star.frame import XController, XFrame, XStorable
from com.sun.star.lang import EventObject, IllegalArgumentException, IndexOutOfBoundsException
from com.sun.star.lang import Locale, WrappedTargetException, XComponent, XLocalizable
from com.sun.star.lang import XMultiComponentFactory, XServiceInfo
from com.sun.star.text import XText, XTextContent, XTextRange
from com.sun.star.uno import XComponentContext
from com.sun.star.view import XSelectionChangeListener, XSelectionSupplier

from smart.diagram.organizationcharts.orgchart.orgchart import OrgChart

class Controller(unohelper.Base, XSelectionChangeListener):
    """Controller class for LibreOffice extension"""

    # Group types
    ORGANIGROUP = 0
    RELATIONGROUP = 1
    PROCESSGROUP = 2
    LISTGROUP = 3
    MATRIXGROUP = 4

    # Diagram types
    NOTDIAGRAM = -1
    SIMPLEORGANIGRAM = 0
    HORIZONTALORGANIGRAM = 1
    TABLEHIERARCHYDIAGRAM = 2
    ORGANIGRAM = 3
    VENNDIAGRAM = 10
    CYCLEDIAGRAM = 11
    PYRAMIDDIAGRAM = 12
    TARGETDIAGRAM = 13
    CONTINUOUSBLOCKPROCESS = 20
    STAGGEREDPROCESS = 21
    BENDINGPROCESS = 22
    UPWARDARROWPROCESS = 23

    def __init__(self, smart_ph, x_context, x_frame):
        """Initialize Controller"""
        self._smart_ph = smart_ph
        self._x_context = x_context
        self._x_frame = x_frame
        self._x_controller = x_frame.getController()
        self._gui = None
        self._x_selection_supplier = None

        self._last_diagram_name = ""
        self._diagram = None
        self._diagram_type = None
        self._group_type = None
        self._last_diagram_type = -1
        self._last_diagram_id = -1

        self.set_gui()
        self.add_selection_listener()

    def is_smart_diagram_shape(self, shape_name):
        """Check if shape is a SmART diagram shape"""
        return (shape_name.startswith("SimpleOrganizationDiagram") or
                shape_name.startswith("TableHierarchyDiagram") or
                shape_name.startswith("HorizontalOrganizationDiagram") or
                shape_name.startswith("OrganizationDiagram") or
                shape_name.startswith("VennDiagram") or
                shape_name.startswith("PyramidDiagram") or
                shape_name.startswith("CycleDiagram") or
                shape_name.startswith("TargetDiagram") or
                shape_name.startswith("ContinuousBlockProcess") or
                shape_name.startswith("StaggeredProcess") or
                shape_name.startswith("BendingProcess") or
                shape_name.startswith("UpwardArrowProcess"))

    def is_smart_organigram_shape(self, shape_name):
        """Check if shape is a SmART organigram shape"""
        return (shape_name.startswith("SimpleOrganizationDiagram") or
                shape_name.startswith("TableHierarchyDiagram") or
                shape_name.startswith("HorizontalOrganizationDiagram") or
                shape_name.startswith("OrganizationDiagram"))

    def set_new_size(self):
        """Set new diagram size"""
        self.get_diagram().increase_size_prop()
        width = 0
        height = 0
        x_pos = 0
        y_pos = 0

        if self.get_diagram().get_size_prop() == self.get_diagram().UD_SIZE:
            width = self.get_diagram().get_ud_width_prop()
            height = self.get_diagram().get_ud_height_prop()
            x_pos = self.get_diagram().get_ud_x_pos_prop()
            y_pos = self.get_diagram().get_ud_y_pos_prop()

        if self.get_diagram().get_size_prop() == self.get_diagram().FULL_SIZE:
            s = self.get_diagram().get_group_shape_size()
            self.get_diagram().set_ud_width_prop(s.Width)
            self.get_diagram().set_ud_height_prop(s.Height)
            p = self.get_diagram().get_group_shape_pos()
            self.get_diagram().set_ud_x_pos_prop(p.X)
            self.get_diagram().set_ud_y_pos_prop(p.Y)
            width = (self.get_diagram().page_props.Width -
                    self.get_diagram().page_props.BorderLeft -
                    self.get_diagram().page_props.BorderRight)
            height = (self.get_diagram().page_props.Height -
                     self.get_diagram().page_props.BorderTop -
                     self.get_diagram().page_props.BorderBottom)
            x_pos = self.get_diagram().page_props.BorderLeft
            y_pos = self.get_diagram().page_props.BorderTop

        self.get_diagram().set_group_shape_size_and_pos(width, height, x_pos, y_pos)
        self.get_diagram().refresh_diagram()

    def get_number_of_pages(self):
        """Get number of pages in document"""
        o_document = self._x_frame.getController().getModel()
        pages_supplier = o_document
        pages = pages_supplier.getDrawPages()
        return pages.getCount()

    def get_smart_ph(self):
        """Get SmartProtocolHandler"""
        return self._smart_ph

    def get_diagram(self):
        """Get current diagram"""
        return self._diagram

    def set_null_diagram(self):
        """Set diagram to None"""
        if self._diagram is not None:
            self._diagram = None

    def set_last_diagram_name(self, name):
        """Set last diagram name"""
        self._last_diagram_name = name

    def set_group_type(self, d_type):
        """Set group type"""
        self._group_type = d_type

    def get_group_type(self):
        """Get group type"""
        return self._group_type

    def set_diagram_type(self, d_type):
        """Set diagram type"""
        self._diagram_type = d_type

    def get_diagram_type(self):
        """Get diagram type"""
        return self._diagram_type

    def set_last_diagram_type(self, d_type):
        """Set last diagram type"""
        self._last_diagram_type = d_type

    def get_last_diagram_type(self):
        """Get last diagram type"""
        return self._last_diagram_type

    def set_last_diagram_id(self, id_val):
        """Set last diagram ID"""
        self._last_diagram_id = id_val

    def get_last_diagram_id(self):
        """Get last diagram ID"""
        return self._last_diagram_id

    def add_selection_listener(self):
        """Add selection change listener"""
        if self._x_selection_supplier is None:
            self._x_selection_supplier = self._x_controller
        if self._x_selection_supplier is not None:
            self._x_selection_supplier.addSelectionChangeListener(self)

    def remove_selection_listener(self):
        """Remove selection change listener"""
        if self._x_selection_supplier is not None:
            self._x_selection_supplier.removeSelectionChangeListener(self)

    def get_gui(self):
        """Get GUI instance"""
        return self._gui

    def set_gui(self):
        """Initialize GUI if needed"""
        if self._gui is None:
            if self._x_context is not None and self._x_frame is not None:
                # self._gui = Gui(self, self._x_context, self._x_frame)  # TODO: Import Gui class
                pass

    def execute_gallery_dialog(self):
        """Execute gallery dialog"""
        if self._gui is not None:
            return self._gui.execute_gallery_dialog()
        return 0

    def get_current_page(self):
        """Get current draw page"""
        return self._x_controller.getCurrentPage()

    def get_location(self):
        """Get current locale"""
        locale = None
        try:
            x_mcf = self._x_context.getServiceManager()
            o_configuration_provider = x_mcf.createInstanceWithContext(
                "com.sun.star.configuration.ConfigurationProvider", self._x_context)
            x_localizable = o_configuration_provider
            locale = x_localizable.getLocale()
        except Exception as ex:
            print(f"Error getting locale: {ex}")
        return locale

    def get_current_diagram_id(self):
        """Get current diagram ID from selected shape"""
        name = self.get_diagram().get_shape_name(self.get_selected_shape())
        s = ""
        i = 0

        # Skip non-digit characters
        while i < len(name) and not name[i].isdigit():
            i += 1

        # Collect digits until dash
        while i < len(name) and name[i] != '-':
            s += name[i]
            i += 1

        return int(s) if s else 0

    def get_shape_id(self, name):
        """Get shape ID from name"""
        s = ""
        i = 0

        # Skip to dash
        while i < len(name) and name[i] != '-':
            i += 1

        # Skip non-digit characters after dash
        while i < len(name) and not name[i].isdigit():
            i += 1

        # Collect digits
        while i < len(name) and name[i].isdigit():
            s += name[i]
            i += 1

        return int(s) if s else 0

    def get_selected_shape(self):
        """Get currently selected shape"""
        try:
            selection = self._x_selection_supplier.getSelection()
            if selection and selection.getCount() > 0:
                return selection.getByIndex(0)
        except Exception:
            pass
        return None

    def create_diagram(self, data=None):
        print("Controller::create_diagram called")
        """Create diagram with optional data"""
        self.remove_selection_listener()
        self.instantiate_diagram()

        if self.get_diagram() is not None:
            if data is not None:
                self.get_diagram().create_diagram(data)
            else:
                self.get_diagram().create_diagram()

            # Initialize object tree in organigrams
            if self.get_group_type() == self.ORGANIGROUP:
                self.get_diagram().init_diagram()

            #self.get_gui().set_visible_control_dialog(True)
            #if self.get_gui().is_shown_tips():
            #    self.get_gui().show_tips_message_box()

        self.add_selection_listener()

    def is_shape_service(self, obj):
        """Check if object is a shape service"""
        is_shape = False
        if obj is not None:
            try:
                # In Python UNO, check for supported services
                if hasattr(obj, 'supportsService'):
                    if obj.supportsService("com.sun.star.drawing.Shape"):
                        is_shape = True
                    if obj.supportsService("com.sun.star.drawing.GroupShape"):
                        return False
            except Exception:
                pass
        return is_shape

    def create_diagram_from_list(self):
        """Create diagram from selected text list"""
        self.remove_selection_listener()

        if (self.get_selected_shape() is not None and
            self.get_selected_shapes().getCount() == 1 and
            self.is_shape_service(self.get_selected_shape())):

            try:
                x_text = self.get_selected_shape()
                text_content = x_text.getString() if hasattr(x_text, 'getString') else ""

                if not text_content or text_content.strip() == "":
                    title = self.get_gui().get_dialog_property_value("Strings2", "Strings2.CouldnotCreateDiagram.Title.Label")
                    message = self.get_gui().get_dialog_property_value("Strings2", "Strings2.CouldnotCreateDiagram.Message.Label")
                    self.get_gui().show_message_box(title, message)
                else:
                    # TODO: Implement DataOfDiagram creation from text
                    # This would parse the text content and create diagram data
                    pass

            except Exception as ex:
                print(f"Error creating diagram from list: {ex}")

        self.add_selection_listener()

    def instantiate_diagram(self):
        """Instantiate diagram based on type"""
        self._diagram = OrgChart(self, self.get_gui(), self._x_frame)

    def get_selected_shapes(self):
        """Get selected shapes collection"""
        try:
            return self._x_selection_supplier.getSelection()
        except Exception:
            return None

    def set_selected_shape(self, obj):
        """Set the selected shape"""
        try:
            self._x_selection_supplier.select(obj)
        except Exception as ex:
            print(f"Error setting selected shape: {ex}")

    def selectionChanged(self, event):
        """Handle selection change events - XSelectionChangeListener implementation"""
        selected_shape = self.get_selected_shape()

        if selected_shape is None:
            self.disappear_control_dialog()
        else:
            # Get shape name
            selected_shape_name = ""
            try:
                if hasattr(selected_shape, 'getName'):
                    selected_shape_name = selected_shape.getName()
            except Exception:
                selected_shape_name = ""

            # Listen for clicks on diagrams
            if self.is_smart_diagram_shape(selected_shape_name):
                new_diagram_name = selected_shape_name.split("-", 1)[0]

                # If the previous selected item is not in the same diagram,
                # need to instantiate the new diagram
                if (self._last_diagram_name == "" or
                    self._last_diagram_name != new_diagram_name):

                    # Set diagram types based on shape name
                    if selected_shape_name.startswith("OrganizationDiagram"):
                        self.set_group_type(self.ORGANIGROUP)
                        self.set_diagram_type(self.ORGANIGRAM)
                        # OrgChartTree.LASTHORLEVEL = -1  # TODO: Implement when OrgChartTree is available

                    self.instantiate_diagram()
                    self._last_diagram_name = new_diagram_name

                    if self.get_diagram() is not None:
                        self.get_diagram().init_diagram()
                        # self.get_diagram().init_properties()  # TODO: Implement when available

                    # if self.get_gui() is not None:
                    #     self.get_gui().set_visible_control_dialog(True)  # TODO: Implement when GUI is available

                # Handle organization chart shape selection
                org_chart_shapes = ["OrganizationDiagram", "SimpleOrganizationDiagram",
                                   "HorizontalOrganizationDiagram", "TableHierarchyDiagram"]
                is_org_chart = any(selected_shape_name.startswith(shape) for shape in org_chart_shapes)

                if (is_org_chart and selected_shape_name.endswith("RectangleShape0")):
                    if self.get_diagram() is not None:
                        self.get_diagram().select_shapes()
                        pass

                # GUI control logic (commented out until GUI is implemented)
                # if self.get_gui() is not None:
                #     if not self.get_gui().is_visible_control_dialog():
                #         self.get_gui().set_visible_control_dialog(True)
                #
                #     if self.is_only_simple_item_selected():
                #         self.get_gui().enable_text_field_of_control_dialog(True)
                #     else:
                #         self.get_gui().enable_text_field_of_control_dialog(False)
                #
                #     if self.get_gui().is_visible_control_dialog():
                #         self.get_gui().set_focus_control_dialog()

            else:
                self.disappear_control_dialog()

            self.set_text_field_of_control_dialog()

    def disappear_control_dialog(self):
        """Hide control dialog"""
        # TODO: Implement when GUI is available
        # if self.get_gui() is not None:
        #     self.get_gui().set_visible_control_dialog(False)
        pass

    def is_only_simple_item_selected(self):
        """Check if only simple item is selected"""
        # TODO: Implement logic to check if only simple item is selected
        return True

    def set_text_field_of_control_dialog(self):
        """Set text field of control dialog"""
        # TODO: Implement when GUI is available
        pass

    # TODO: Add remaining methods from the Java implementation
    # - create_diagram methods for different diagram types
    # - test methods
    # - utility methods
    # - event handling methods


# pythonloader loads a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    Controller,  # UNO object class
    "com.collabora.milsymbol.Controller",  # implementation name (customize for yourself)
    ("com.sun.star.view.XSelectionChangeListener",), )  # implemented services (only 1)
