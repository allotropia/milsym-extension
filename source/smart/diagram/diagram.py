# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

# This file incorporates work covered by the following license notice:
#   SPDX-License-Identifier: LGPL-3.0-only

"""
Base Diagram class - stub implementation
Python port of Diagram.java
"""
import sys
import unohelper
import officehelper
import os
import uno

# Add parent directory to path to import utils
base_dir = os.path.dirname(os.path.dirname(os.path.dirname(__file__)))
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from utils import parse_svg_dimensions

from abc import ABC, abstractmethod
from com.sun.star.beans import PropertyValue
from com.sun.star.awt import Point, Size
from com.sun.star.text.TextContentAnchorType import AT_PARAGRAPH
from com.sun.star.xml import AttributeData

class Diagram(ABC):
    """Base diagram class - simplified version of the Java Diagram class"""

    # Shape types
    DIAGRAM_SHAPE_TYPE = "GraphicObjectShape"  # start shape and child shapes
    DIAGRAM_BASE_SHAPE_TYPE = "RectangleShape" # base control shape
    CONNECTOR_SHAPE = "ConnectorShape"         # connector shape

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

    def __init__(self, controller, gui, x_frame, x_context):
        self._controller = controller
        self._gui = gui
        self._x_frame = x_frame
        self._x_context = x_context
        self._x_model = x_frame.getController().getModel()
        self._x_controller = x_frame.getController()
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

        self.page_props = PageProps()

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
        """Set drawing area dimensions - to be overridden in subclasses"""
        pass

    def get_group_shape(self):
        """Get group shape"""
        if self._x_group_shape is not None:
            return self._x_group_shape
        return None

    def get_group_shape_pos(self):
        """Get group shape position"""
        if self._x_group_shape is not None:
            return self._x_group_shape.getPosition()
        return None

    def set_group_shape_size_and_pos(self, width: int, height: int, x_pos: int, y_pos: int):
        """Set group shape size and position"""
        if self._x_group_shape is not None:
            try:
                self._x_group_shape.setSize(Size(width, height))
                self._x_group_shape.setPosition(Point(x_pos, y_pos))
            except Exception as ex:
                print(f"Error setting group shape size and position: {ex}")

    def create_shape(self, shape_type: str, shape_id: int, x: int = 0, y: int = 0, width: int = 0, height: int = 0):
        x_shape = None
        try:
            # Create shape using LibreOffice service manager
            x_shape = self._x_model.createInstance(f"com.sun.star.drawing.{shape_type}")

            # Set shape name
            shape_name = f"{self.get_diagram_type_name()}{self._diagram_id}-{shape_type}"
            if shape_type != self.DIAGRAM_BASE_SHAPE_TYPE:
                shape_name += f"{shape_id}"
            x_shape.setName(shape_name)

            # Set position and size
            if x > 0 and y > 0:
                x_shape.setPosition(Point(X=x, Y=y))
            if width > 0 and height > 0:
                x_shape.setSize(Size(Width=width, Height=height))
        except Exception as ex:
            print(f"Error creating shape: {ex}")

        return x_shape

    def remove_shape(self):
        """Remove shape - to be overridden in subclasses"""
        pass

    def show_edit_dialog(self):
        """Show edit dialog - to be overridden in subclasses"""
        pass

    def set_move_protect_of_shape(self, shape):
        """Set move/resize protection on a shape"""
        try:
            shape.setPropertyValue("MoveProtect", True)
            shape.setPropertyValue("SizeProtect", True)
        except Exception as ex:
            print(f"Error setting move protection: {ex}")

    def set_color_prop(self, color: int):
        """Set color property"""
        # not needed
        pass

    def get_last_shape(self):
        return self.get_controller().get_selected_shape()

    # set SVG data from symbol dialog
    def set_svg_data(self, svg_data):
        x_selected_shapes = self.get_controller().get_selected_shapes()
        last_shape = None
        for i in range(x_selected_shapes.getCount()):
            x_shape = x_selected_shapes.getByIndex(i)
            if x_shape is not None:
                last_shape = x_shape
                self.set_new_shape_properties(x_shape, self.DIAGRAM_SHAPE_TYPE, svg_data)

        return last_shape

    def set_new_shape_properties(self, shape, shape_type: str, svg_data):
        """Set shape properties"""
        try:
            if shape_type == self.DIAGRAM_SHAPE_TYPE:
                graphic_provider = self._x_context.ServiceManager.createInstanceWithContext("com.sun.star.graphic.GraphicProvider", self._x_context)
                pipe = self._x_context.ServiceManager.createInstanceWithContext("com.sun.star.io.Pipe", self._x_context)
                pipe.writeBytes(uno.ByteSequence(svg_data.encode('utf-8')))
                pipe.flush()
                pipe.closeOutput()
                media_properties = (PropertyValue("InputStream", 0, pipe, 0),)
                graphic = graphic_provider.queryGraphic(media_properties)
                shape.setPropertyValue("Graphic", graphic)

                # Parse SVG dimensions and resize shape to maintain proportions
                size = parse_svg_dimensions(svg_data)
                shape.setSize(size)

        except Exception as ex:
            print(f"Error setting shape properties: {ex}")

    def set_shape_properties(self, shape, shape_type: str):
        """Set shape properties"""
        try:
            if shape_type == self.DIAGRAM_SHAPE_TYPE:
                svg_url = "vnd.sun.star.extension://com.collabora.milsymbol/img/orbat_base.svg"
                graphic_provider = self._x_context.ServiceManager.createInstanceWithContext("com.sun.star.graphic.GraphicProvider", self._x_context)
                media_properties = (PropertyValue("URL", 0, svg_url, 0),)
                graphic = graphic_provider.queryGraphic(media_properties)
                shape.setPropertyValue("Graphic", graphic)

                self.set_font_properties_of_shape(shape)

            elif shape_type == self.CONNECTOR_SHAPE:
                if self.is_text_fit_prop():
                    shape.setPropertyValue("TextFitToSize", 1)  # PROPORTIONAL
                else:
                    shape.setPropertyValue("TextFitToSize", 0)  # NONE
                self.set_connector_shape_line_props(shape)

        except Exception as ex:
            print(f"Error setting shape properties: {ex}")

    def remove_shape_from_group(self, x_shape):
        """Remove shape from the group"""
        if self._x_shapes is not None:
            self._x_shapes.remove(x_shape)

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
        return False

    def get_connectors_line_width_prop(self):
        return 40

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

    def get_diagram_id(self) -> int:
        """Get diagram ID"""
        return self._diagram_id

    def init_diagram(self, diagram_id=None):
        """Initialize diagram"""
        try:
            x_curr_shape = None
            curr_shape_name = ""
            self._x_draw_page = self.get_controller().get_current_page()

            if diagram_id is not None and diagram_id != 0:
                self._diagram_id = diagram_id
            else:
                _current_diagram_id = self.get_controller().get_current_diagram_id()
                if _current_diagram_id != 0:
                    self._diagram_id = _current_diagram_id

            s_diagram_id = str(self._diagram_id)

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

    def init_properties(self):
        """Initialize diagram properties - to be overridden in subclasses"""
        pass

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

            # Set name for the group shape
            self._x_group_shape.setName(self.get_diagram_type_name() + str(self._diagram_id) + "-GroupShape")

            # Add to draw page
            self.add_group_shape_to_draw_page()

            # Get XShapes interface for the group
            self._x_shapes = self._x_group_shape

        except Exception as ex:
            print(f"Error in create_diagram: {ex}")

    def add_group_shape_to_draw_page(self):
        if self._x_model.supportsService("com.sun.star.text.TextDocument"): # Writer
            self._x_group_shape.setPropertyValue("AnchorType", AT_PARAGRAPH)
            cursor = self._x_model.getText().createTextCursor()
            cursor.getText().insertTextContent(cursor, self._x_group_shape, False)
        elif self._x_model.supportsService("com.sun.star.sheet.SpreadsheetDocument"): # Calc
            self._x_draw_page.getDrawPage().add(self._x_group_shape)
        else: # Impress/Draw
            self._x_draw_page.add(self._x_group_shape)

    def adjust_page_props(self):
        """Adjust page properties"""
        # Default values - use A4 size
        width = 21000  # Default A4 width
        height = 29700  # Default A4 height
        border_left = 1000
        border_right = 1000
        border_top = 1000
        border_bottom = 1000

        # Writer does not provide a page size via the draw page
        if not self._x_model.supportsService("com.sun.star.text.TextDocument"):
            try:
                width = int(self._x_draw_page.getPropertyValue("Width"))
                height = int(self._x_draw_page.getPropertyValue("Height"))
                border_left = int(self._x_draw_page.getPropertyValue("BorderLeft"))
                border_right = int(self._x_draw_page.getPropertyValue("BorderRight"))
                border_top = int(self._x_draw_page.getPropertyValue("BorderTop"))
                border_bottom = int(self._x_draw_page.getPropertyValue("BorderBottom"))

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

        # Update page properties
        self.page_props.Width = width
        self.page_props.Height = height
        self.page_props.BorderLeft = border_left
        self.page_props.BorderRight = border_right
        self.page_props.BorderTop = border_top
        self.page_props.BorderBottom = border_bottom

    def set_group_size(self):
        """Set group size based on page properties"""
        # Calculate available drawing area
        if self._x_draw_page is None:
            self._x_draw_page = self.get_controller().get_current_page()
        if self.page_props is None:
            self.adjust_page_props()
        if self.page_props is not None:
            self._draw_area_width = (self.page_props.Width -
                                     self.page_props.BorderLeft -
                                     self.page_props.BorderRight)
            self._draw_area_height = (self.page_props.Height -
                                      self.page_props.BorderTop -
                                      self.page_props.BorderBottom)

    @abstractmethod
    def get_diagram_type_name(self) -> str:
        """Get diagram type name - to be implemented by subclasses"""
