# SPDX-License-Identifier: MPL-2.0
# This file incorporates work covered by the following license notice:
#   SPDX-License-Identifier: LGPL-3.0-only

"""
Base Diagram class - stub implementation
Python port of Diagram.java
"""

from abc import ABC, abstractmethod

from com.sun.star.awt import Point, Size

class Diagram(ABC):
    """Base diagram class - simplified version of the Java Diagram class"""

    # Connection types
    CONN_LINE = 0
    CONN_CURVE = 1

    # Color modes
    BASE_COLORS_MODE = 0
    FIRST_COLORSCHEME_MODE_VALUE = 10

    # Base colors array (simplified)
    _BASE_COLORS = [
        0xff0000, 0x00ff00, 0x0000ff, 0xffff00,
        0xff00ff, 0x00ffff, 0x800000, 0x008000
    ]

    def __init__(self, controller, gui, x_frame):
        self._controller = controller
        self._gui = gui
        self._x_frame = x_frame
        self._x_model = x_frame.getController().getModel()
        self._x_draw_page = None
        self._x_shapes = None
        self._diagram_id = -1
        self._color_mode_prop = self.BASE_COLORS_MODE
        self._style_prop = 0

        # Drawing area properties
        self._draw_area_width = 0
        self._draw_area_height = 0

        # Page properties (simplified)
        class PageProps:
            def __init__(self):
                self.border_left = 1000
                self.border_top = 1000
                self.border_right = 1000
                self.border_bottom = 1000

        self._page_props = PageProps()

    def get_controller(self):
        """Get controller reference"""
        return self._controller

    def get_gui(self):
        """Get GUI reference"""
        return self._gui

    def get_shapes(self):
        """Get shapes collection"""
        return self._x_shapes

    def set_draw_area(self):
        """Set drawing area dimensions"""
        # Simplified implementation
        self._draw_area_width = 10000  # Default width
        self._draw_area_height = 7000  # Default height

    def create_shape(self, shape_type: str, shape_id: int, x: int = 0, y: int = 0, width: int = 0, height: int = 0):
        print(f"Creating {shape_type} with ID {shape_id} at ({x}, {y}) size ({width}, {height})")
        x_shape = None
        try:
            # Create shape using LibreOffice service manager
            x_shape = self._x_model.createInstance(f"com.sun.star.drawing.{shape_type}")

            # Set shape name
            shape_name = f"{self.get_diagram_type_name()}{self._diagram_id}-{shape_type}{shape_id}"
            x_shape.setName(shape_name)

            # Set position and size
            if x > 0 and y > 0:
                x_shape.setPosition(Point(X=x, Y=y))
            if width > 0 and height > 0:
                x_shape.setSize(Size(Width=width, Height=height))
        except Exception as ex:
            print(f"Error creating shape: {ex}")

        return x_shape

    def set_text_of_shape(self, shape, text: str):
        """Set text content of a shape"""
        print(f"Setting text '{text}' on shape")
        shape.setString(text)

    def set_move_protect_of_shape(self, shape):
        """Set move protection on a shape"""
        try:
            shape.setPropertyValue("MoveProtect", True)
        except Exception as ex:
            print(f"Error setting move protection: {ex}")

    def set_color_prop(self, color: int):
        """Set color property"""
        # not needed
        print(f"Setting color to {hex(color)}")

    def set_shape_properties(self, shape, shape_type: str):
        """Set shape properties"""
        try:
            if shape_type == "BaseShape":
                if self.is_text_fit_prop():
                    shape.setPropertyValue("TextFitToSize", 1)  # PROPORTIONAL
                else:
                    shape.setPropertyValue("TextFitToSize", 0)  # NONE
                    shape.setPropertyValue("CharHeight", 40.0)

            elif shape_type == "RectangleShape":
                if self.is_modify_colors_prop():
                    self.set_color_settings_of_shape(shape)

                # Corner radius settings
                if self.get_rounded_prop() == 0:  # NULL_ROUNDED
                    shape.setPropertyValue("CornerRadius", 0)
                elif self.get_rounded_prop() == 1:  # MEDIUM_ROUNDED
                    shape.setPropertyValue("CornerRadius", 500)  # CORNER_RADIUS2
                elif self.get_rounded_prop() == 2:  # EXTRA_ROUNDED
                    shape.setPropertyValue("CornerRadius", 1000)  # CORNER_RADIUS3

                # Line/outline settings
                if self.is_outline_prop():
                    shape.setPropertyValue("LineStyle", 1)  # SOLID
                    shape.setPropertyValue("LineWidth", self.get_shapes_line_width_prop())
                else:
                    shape.setPropertyValue("LineStyle", 0)  # NONE

                # Shadow settings
                if self.is_shadow_prop():
                    shape.setPropertyValue("Shadow", True)
                    shape.setPropertyValue("ShadowXDistance", 100)  # SHADOW_DIST1
                    shape.setPropertyValue("ShadowYDistance", -100)  # -SHADOW_DIST1
                    shape.setPropertyValue("ShadowTransparence", 50)  # SHADOW_TRANSP

                    # Determine shadow color
                    shadow_color = -1
                    try:
                        fill_style = shape.getPropertyValue("FillStyle")
                        if fill_style == 1:  # SOLID
                            shadow_color = int(shape.getPropertyValue("FillColor"))
                        else:
                            gradient = shape.getPropertyValue("FillGradient")
                            start_color = gradient.StartColor
                            end_color = gradient.EndColor
                            shadow_color = min(start_color, end_color)
                    except Exception:
                        pass

                    if shadow_color == -1:
                        shadow_color = 8421504  # Default gray
                    shape.setPropertyValue("ShadowColor", shadow_color)
                else:
                    shape.setPropertyValue("Shadow", False)

                self.set_font_properties_of_shape(shape)

            elif shape_type == "ConnectorShape":
                if self.is_text_fit_prop():
                    shape.setPropertyValue("TextFitToSize", 1)  # PROPORTIONAL
                else:
                    shape.setPropertyValue("TextFitToSize", 0)  # NONE
                self.set_connector_shape_line_props(shape)

        except Exception as ex:
            print(f"Error setting shape properties: {ex}")

    # Stub methods for property checking - to be implemented
    def is_text_fit_prop(self):
        return False

    def is_modify_colors_prop(self):
        return False

    def get_rounded_prop(self):
        return 0

    def is_outline_prop(self):
        return True

    def get_shapes_line_width_prop(self):
        return 100

    def is_shadow_prop(self):
        return False

    def set_color_settings_of_shape(self, shape):
        pass

    def set_font_properties_of_shape(self, shape):
        pass

    def set_connector_shape_line_props(self, shape):
        try:
            shape.setPropertyValue("LineColor", self.get_connector_color_prop())

            if self.is_shown_connectors_prop():
                shape.setPropertyValue("LineStyle", 1)  # SOLID
            else:
                shape.setPropertyValue("LineStyle", 0)  # NONE

            # Connector types
            conn_type = self.get_connector_type_prop()
            if conn_type == 0:  # CONN_STANDARD
                shape.setPropertyValue("EdgeKind", 0)  # STANDARD
            elif conn_type == self.CONN_LINE:
                shape.setPropertyValue("EdgeKind", 1)  # LINE
            elif conn_type == 2:  # CONN_STRAIGHT
                shape.setPropertyValue("EdgeKind", 2)  # LINES
            elif conn_type == self.CONN_CURVE:
                shape.setPropertyValue("EdgeKind", 3)  # CURVE

            # Arrow settings
            if self.is_connector_start_arrow_prop():
                shape.setPropertyValue("LineStartName", "Arrow")
            else:
                shape.setPropertyValue("LineStartName", "")

            if self.is_connector_end_arrow_prop():
                shape.setPropertyValue("LineEndName", "Arrow")
            else:
                shape.setPropertyValue("LineEndName", "")

            # Line width and arrow sizing
            line_width = self.get_connectors_line_width_prop()
            shape.setPropertyValue("LineWidth", line_width)

            arrow_width = 400
            if line_width == 200:
                arrow_width = 600
            elif line_width >= 300:
                arrow_width = int(line_width * 2.5)

            shape.setPropertyValue("LineStartWidth", arrow_width)
            shape.setPropertyValue("LineEndWidth", arrow_width)

        except Exception as ex:
            print(f"Error setting connector properties: {ex}")

    # Stub methods for connector properties
    def get_connector_color_prop(self):
        return 0x000000  # Black

    def is_shown_connectors_prop(self):
        return True

    def get_connector_type_prop(self):
        return 0  # STANDARD

    def is_connector_start_arrow_prop(self):
        return False

    def is_connector_end_arrow_prop(self):
        return True

    def get_connectors_line_width_prop(self):
        return 100

    def set_connector_shape_props(self, connector_shape, start_shape, start_conn_pos: int, end_shape, end_conn_pos: int):
        """Set connector shape properties"""
        try:
            # Direct property access in Python UNO
            connector_shape.setPropertyValue("StartShape", start_shape)
            connector_shape.setPropertyValue("EndShape", end_shape)
            connector_shape.setPropertyValue("StartGluePointIndex", start_conn_pos)
            connector_shape.setPropertyValue("EndGluePointIndex", end_conn_pos)
            connector_shape.setPropertyValue("LineColor", self.get_connector_color_prop())

            # Set line properties
            self.set_connector_shape_line_props(connector_shape)

            # Text fit settings
            if self.is_text_fit_prop():
                connector_shape.setPropertyValue("TextFitToSize", 1)  # PROPORTIONAL
            else:
                connector_shape.setPropertyValue("TextFitToSize", 0)  # NONE

        except Exception as ex:
            print(f"Error setting connector shape properties: {ex}")

    def refresh_diagram(self):
        """Refresh the diagram display"""
        self.get_diagram_tree().refresh()

    def get_shape_name(self, shape):
        """Get the name of a shape"""
        try:
            return shape.getName()
        except Exception:
            return ""

    def init_diagram(self):
        """Initialize diagram"""
        try:
            print("2 Initializing diagram...")
            x_curr_shape = None
            curr_shape_name = ""
            self._x_draw_page = self.get_controller().get_current_page()
            print("2 self._x_draw_page:", self._x_draw_page)
            self._diagram_id = self.get_controller().get_current_diagram_id()
            print("2 self._diagram_id:", self._diagram_id)
            s_diagram_id = str(self._diagram_id)
            print("2 count: ", self._x_draw_page.getCount())

            for i in range(self._x_draw_page.getCount()):
                x_curr_shape = self._x_draw_page.getByIndex(i)
                curr_shape_name = self.get_shape_name(x_curr_shape)

                if (s_diagram_id in curr_shape_name and
                    curr_shape_name.startswith(self.get_diagram_type_name()) and
                    curr_shape_name.endswith("GroupShape")):

                    # Query interfaces for XShapes and XShape
                    self._x_shapes = x_curr_shape
                    self._x_group_shape = x_curr_shape

        except Exception as ex:
            print(f"Error in init_diagram: {ex}")

    def create_diagram(self, data):
        """Create diagram from data"""
        import random

        try:
            self._x_draw_page = self.get_controller().get_current_page()
            self._diagram_id = int(random.random() * 10000)

            # set diagramName in the Controller object
            diagram_name = self.get_diagram_type_name() + str(self._diagram_id)
            self.get_controller().set_last_diagram_name(diagram_name)

            # set new PageProps object with data of page
            # width, height, borderLeft, borderRight, borderTop, borderBottom
            self.adjust_page_props()

            # get minimum of width and height
            self.set_group_size()

            # Create group shape
            self._x_group_shape = self._x_model.createInstance("com.sun.star.drawing.GroupShape")
            print("Created group shape:", self._x_group_shape)

            # Set name for the group shape
            self._x_group_shape.setName(self.get_diagram_type_name() + str(self._diagram_id) + "-GroupShape")

            # Add to draw page
            self._x_draw_page.add(self._x_group_shape)

            # Get XShapes interface for the group
            self._x_shapes = self._x_group_shape

        except Exception as ex:
            print(f"Error in create_diagram: {ex}")

    def adjust_page_props(self):
        """Adjust page properties"""
        width = 0
        height = 0
        border_left = 0
        border_right = 0
        border_top = 0
        border_bottom = 0

        try:
            # In Python UNO, draw page has direct access to property methods
            x_page_properties = self._x_draw_page

            width = int(x_page_properties.getPropertyValue("Width"))
            height = int(x_page_properties.getPropertyValue("Height"))
            border_left = int(x_page_properties.getPropertyValue("BorderLeft"))
            border_right = int(x_page_properties.getPropertyValue("BorderRight"))
            border_top = int(x_page_properties.getPropertyValue("BorderTop"))
            border_bottom = int(x_page_properties.getPropertyValue("BorderBottom"))

            # Ensure minimum border values
            if border_left < 1000:
                border_left = 1000
            if border_right < 1000:
                border_right = 1000
            if border_top < 1000:
                border_top = 1000
            if border_bottom < 1000:
                border_bottom = 1000

        except Exception as ex:
            print(f"Error in adjust_page_props: {ex}")
            # Set default values on error
            width = 21000  # Default A4 width
            height = 29700  # Default A4 height
            border_left = border_right = border_top = border_bottom = 1000

        # Update page properties
        self._page_props.Width = width
        self._page_props.Height = height
        self._page_props.BorderLeft = border_left
        self._page_props.BorderRight = border_right
        self._page_props.BorderTop = border_top
        self._page_props.BorderBottom = border_bottom

    def set_group_size(self):
        """Set group size based on page properties"""
        # Calculate available drawing area
        if self._x_draw_page is None:
            self._x_draw_page = self.get_controller().get_current_page()
        if self._page_props is None:
            self.adjust_page_props()
        if self._page_props is not None:
            self._draw_area_width = (self._page_props.Width -
                                     self._page_props.BorderLeft -
                                     self._page_props.BorderRight)
            self._draw_area_height = (self._page_props.Height -
                                      self._page_props.BorderTop -
                                      self._page_props.BorderBottom)

    @abstractmethod
    def get_diagram_type_name(self) -> str:
        """Get diagram type name - to be implemented by subclasses"""
