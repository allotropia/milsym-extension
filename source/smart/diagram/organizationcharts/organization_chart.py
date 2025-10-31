# SPDX-License-Identifier: MPL-2.0
# This file incorporates work covered by the following license notice:
#   SPDX-License-Identifier: LGPL-3.0-only

"""
Organization Chart base class
Python port of OrganizationChart.java
"""

from abc import abstractmethod

# Import base classes
from ..diagram import Diagram

from com.sun.star.awt import Size, Point

#from com.sun.star.drawing import LineStyle
# from com.sun.star.drawing.ConnectorType import STANDARD as CONN_STANDARD_VALUE
# from com.sun.star.drawing.ConnectorType import LINE as CONN_LINE_VALUE
# from com.sun.star.drawing.ConnectorType import LINES as CONN_LINES_VALUE
# from com.sun.star.drawing.ConnectorType import CURVE as CONN_CURVE_VALUE


class OrganizationChart(Diagram):
    """Base class for organization charts"""

    # Hierarchy types
    UNDERLING = 0
    ASSOCIATE = 1

    # Style constants
    DEFAULT = 0
    WITHOUT_OUTLINE = 1
    NOT_ROUNDED = 2
    WITH_SHADOW = 3

    GREEN_DARK = 4
    GREEN_BRIGHT = 5
    BLUE_DARK = 6
    BLUE_BRIGHT = 7
    PURPLE_DARK = 8
    PURPLE_BRIGHT = 9
    ORANGE_DARK = 10
    ORANGE_BRIGHT = 11
    YELLOW_DARK = 12
    YELLOW_BRIGHT = 13

    BLUE_SCHEME = 14
    AQUA_SCHEME = 15
    RED_SCHEME = 16
    FIRE_SCHEME = 17
    SUN_SCHEME = 18
    GREEN_SCHEME = 19
    OLIVE_SCHEME = 20
    PURPLE_SCHEME = 21
    PINK_SCHEME = 22
    INDIAN_SCHEME = 23
    MAROON_SCHEME = 24
    BROWN_SCHEME = 25
    USER_DEFINE = 26

    FIRST_COLORTHEMEGRADIENT_STYLE_VALUE = 4
    FIRST_COLORSCHEME_STYLE_VALUE = 14

    def __init__(self, controller, gui, x_frame, x_context):
        super().__init__(controller, gui, x_frame, x_context)

        # Rates of measure of group shape (e.g.: 10:6)
        self._group_width = 0
        self._group_height = 0

        # Rates of measure of rectangles (e.g.: WIDTH:HORSPACE 2:1, HEIGHT:VERSPACE 4:3)
        self._shape_width = 0
        self._hor_space = 0
        self._shape_height = 0
        self._ver_space = 0

        # Horizontal offset in the draw page if needed
        self._half_diff = 0

        # Item hierarchy type in diagram
        self._new_item_h_type = self.UNDERLING

        # Hidden root element flag
        self._is_hidden_root_element = False

        self.set_default_props()

    def is_hidden_root_element_prop(self) -> bool:
        """Check if root element is hidden"""
        return self._is_hidden_root_element

    def set_hidden_root_element_prop(self, is_hidden: bool):
        """Set hidden root element property"""
        self._is_hidden_root_element = is_hidden

    def init_root_element_hidden_property(self):
        """Initialize root element hidden property"""
        if self.get_diagram_tree().get_root_item() is not None:
            self._is_hidden_root_element = self.get_diagram_tree().get_root_item().is_hidden_element()

    def is_color_scheme_style(self, style: int) -> bool:
        """Check if style is a color scheme style"""
        return style in [
            self.BLUE_SCHEME, self.AQUA_SCHEME, self.RED_SCHEME, self.FIRE_SCHEME,
            self.SUN_SCHEME, self.GREEN_SCHEME, self.OLIVE_SCHEME, self.PURPLE_SCHEME,
            self.PINK_SCHEME, self.INDIAN_SCHEME, self.MAROON_SCHEME, self.BROWN_SCHEME
        ]

    def get_color_mode_of_scheme_style(self, style: int) -> int:
        """Get color mode of scheme style"""
        return style - self.FIRST_COLORSCHEME_STYLE_VALUE + Diagram.FIRST_COLORSCHEME_MODE_VALUE

    def is_color_theme_gradient_style(self, style: int) -> bool:
        """Check if style is a color theme gradient style"""
        return style in [
            self.GREEN_DARK, self.GREEN_BRIGHT, self.BLUE_DARK, self.BLUE_BRIGHT,
            self.PURPLE_DARK, self.PURPLE_BRIGHT, self.ORANGE_DARK, self.ORANGE_BRIGHT,
            self.YELLOW_DARK, self.YELLOW_BRIGHT
        ]

    def get_color_mode_of_theme_gradient_style(self, style: int) -> int:
        """Get color mode of theme gradient style"""
        return style - self.FIRST_COLORTHEMEGRADIENT_STYLE_VALUE + 1  # BASE_COLORS_WITH_GRADIENT_MODE

    def get_user_define_style_value(self) -> int:
        """Get user define style value"""
        return self.USER_DEFINE

    def set_default_props(self):
        """Set default properties"""
        # Stub implementation

    def get_shape_width(self) -> int:
        """Get shape width ratio"""
        return self._shape_width

    def get_hor_space(self) -> int:
        """Get horizontal space ratio"""
        return self._hor_space

    def get_shape_height(self) -> int:
        """Get shape height ratio"""
        return self._shape_height

    def get_ver_space(self) -> int:
        """Get vertical space ratio"""
        return self._ver_space

    def get_shape_name(self, shape) -> str:
        """Get shape name"""
        try:
            if hasattr(shape, 'getName'):
                return shape.getName()
            return ""
        except Exception:
            return ""

    def set_hor_level_of_control_shape(self, control_shape, level: int):
        """Set horizontal level of control shape"""
        if control_shape is not None:
            text = control_shape.getString()
            if text == "" or ":" not in text:
                control_shape.setString(f"LastHorLevel:{level}")
            else:
                is_already_defined = False
                a_str = text.split(":")
                for i in range(len(a_str) - 1):
                    if a_str[i] == "LastHorLevel":
                        a_str[i + 1] = str(level)
                        is_already_defined = True

                if is_already_defined:
                    text = ""
                    for i in range(len(a_str)):
                        text += a_str[i]
                        if i < len(a_str) - 1:
                            text += ":"
                else:
                    text += f":LastHorLevel:{level}"

                control_shape.setString(text)
            control_shape.MoveProtect = True
            control_shape.Visible = False

    def get_hor_level_of_control_shape(self, control_shape) -> int:
        """Get horizontal level of control shape"""
        if control_shape is not None:
            text = control_shape.getString()
            s_number = ""
            if ":" not in text:
                s_number = text
            else:
                a_str = text.split(":")
                for i in range(len(a_str) - 1):
                    if a_str[i] == "LastHorLevel":
                        s_number = a_str[i + 1]

            if s_number == "":
                return -1
            else:
                return int(s_number)
        return -1

    def remove_hor_level_props_of_control_shape(self, control_shape):
        """Remove horizontal level properties of control shape"""
        if control_shape is not None:
            text = control_shape.getString()
            if "LastHorLevel:" in text:
                new_text = ""
                a_str = text.split(":")
                for i in range(0, len(a_str) - 1, 2):
                    if a_str[i] != "LastHorLevel":
                        new_text += a_str[i] + ":" + a_str[i + 1] + ":"
                new_text = new_text[:-1]  # Remove trailing colon
                control_shape.setString(new_text)

    def set_control_shape_props(self, control_shape):
        """Set control shape properties"""
        control_shape.MoveProtect = True
        control_shape.Visible = False

    def set_color_mode_and_style_of_control_shape(self, control_shape):
        """Set color mode and style of control shape"""
        # Stub implementation

    def get_top_shape_id(self) -> int:
        """Get top shape ID"""
        i_top_shape_id = -1
        x_curr_shape = None
        curr_shape_name = ""
        shape_id = 0

        try:
            for i in range(self._x_shapes.getCount()):
                x_curr_shape = self._x_shapes.getByIndex(i)
                curr_shape_name = self.get_shape_name(x_curr_shape)
                if "RectangleShape" in curr_shape_name:
                    shape_id = self.get_controller().get_shape_id(curr_shape_name)
                    if shape_id > i_top_shape_id:
                        i_top_shape_id = shape_id
        except Exception as ex:
            print(f"Error getting top shape ID: {ex}")

        return i_top_shape_id

    def set_draw_area(self):
        """Set draw area for organization chart with shadow allowance"""

        try:
            # Allow horizontal place for shadow properties
            self.page_props.BorderTop += (Diagram.SHADOW_DIST1 + 100)
            self._draw_area_width -= (2 * Diagram.SHADOW_DIST1 + 100)
            self._draw_area_height -= (Diagram.SHADOW_DIST1 + 100)

            origin_gs_width = self._draw_area_width

            if (self._draw_area_width / self._group_width) <= (self._draw_area_height / self._group_height):
                self._draw_area_height = self._draw_area_width * self._group_height // self._group_width
            else:
                self._draw_area_width = self._draw_area_height * self._group_width // self._group_height

            # Set new size of group shape for organigram
            self._x_group_shape.setSize(Size(self._draw_area_width, self._draw_area_height))

            self._half_diff = 0
            if origin_gs_width > self._draw_area_width:
                self._half_diff = (origin_gs_width - self._draw_area_width) // 2

            self._half_diff += Diagram.SHADOW_DIST1
            self._x_group_shape.setPosition(
                Point(self.page_props.BorderLeft + self._half_diff, self.page_props.BorderTop)
            )
        except Exception as ex:
            print(f"Error setting draw area: {ex}")

    def clear_empty_diagram_and_recreate(self):
        """Clear empty diagram and recreate"""
        try:
            if self._x_shapes is not None:
                x_shape = None
                for i in range(self._x_shapes.getCount()):
                    x_shape = self._x_shapes.getByIndex(i)
                    if x_shape is not None:
                        self._x_shapes.remove(x_shape)
            self.create_diagram(1)
        except Exception as ex:
            print(f"Error clearing and recreating diagram: {ex}")

    def select_shapes(self):
        """Select all shapes in the organization chart"""
        self.get_controller().set_selected_shape(self._x_shapes)

    def init_properties(self):
        """Initialize properties from control and root shapes"""
        x_control_shape = self.get_diagram_tree().get_control_shape()
        x_root_shape = self.get_diagram_tree().get_root_item().get_rectangle_shape()
        if x_control_shape is not None and x_root_shape is not None:
            self.init_properties_from_shapes(x_control_shape, x_root_shape)
            self.init_root_element_hidden_property()

    def init_properties_from_shapes(self, x_control_shape, x_root_shape):
        """Initialize properties from control and root shapes"""
        self.set_default_props()
        # self.init_color_mode_and_style()

        #try:
            # Get root shape properties
            # if self.is_simple_color_mode():
            #     fill_color = x_root_shape.getPropertyValue("FillColor")
            #     self.set_color_prop(fill_color)

            # if self.is_gradient_color_mode():
            #     a_gradient = x_root_shape.getPropertyValue("FillGradient")
            #     self.set_start_color_prop(a_gradient.StartColor)
            #     self.set_end_color_prop(a_gradient.EndColor)
            #     if a_gradient.Angle == 900:
            #         self.set_gradient_direction_prop(Diagram.HORIZONTAL)

            # if self.is_color_theme_gradient_mode():
            #     self.set_color_theme_gradient_colors()
            #     # self.set_shapes_line_width_prop(Diagram.LINE_WIDTH200)
            #     self.set_rounded_prop(Diagram.NULL_ROUNDED)

            # Handle style property
            # if self.get_style_prop() == OrganizationChart.DEFAULT:
            #     pass

            # if self.get_style_prop() == OrganizationChart.WITHOUT_OUTLINE:
            #     self.set_outline_prop(False)

            # if self.get_style_prop() == OrganizationChart.NOT_ROUNDED:
            #     self.set_rounded_prop(Diagram.NULL_ROUNDED)

            # if self.get_style_prop() == OrganizationChart.WITH_SHADOW:
            #     self.set_shadow_prop(True)

            # if self.get_style_prop() == OrganizationChart.USER_DEFINE:
            #     line_style = x_root_shape.getPropertyValue("LineStyle")
            #     if line_style == LineStyle.NONE:
            #         self.set_outline_prop(False)

            #     line_width = x_root_shape.getPropertyValue("LineWidth")
            #     self.set_shapes_line_width_prop(line_width)

            #     corner_radius = x_root_shape.getPropertyValue("CornerRadius")
            #     if corner_radius < 200:
            #         self.set_rounded_prop(Diagram.NULL_ROUNDED)
            #     elif corner_radius < 600:
            #         self.set_rounded_prop(Diagram.MEDIUM_ROUNDED)
            #     else:
            #         self.set_rounded_prop(Diagram.EXTRA_ROUNDED)

            #     shadow = x_root_shape.getPropertyValue("Shadow")
            #     if shadow:
            #         self.set_shadow_prop(True)

            # self.set_font_property_values()

            # Get text color from root shape
            # x_text = x_root_shape
            # x_text_cursor = x_text.createTextCursor()
            # text_color = x_text_cursor.getPropertyValue("CharColor")
            # self.set_text_color_prop(text_color)

            # """ if self.is_shown_connectors_prop():
            #     x_conn_shape = self.get_roots_connector()

            #     connectors_line_width = x_conn_shape.getPropertyValue("LineWidth")
            #     self.set_connectors_line_width_prop(connectors_line_width)

            #     connector_color = x_conn_shape.getPropertyValue("LineColor")
            #     self.set_connector_color_prop(connector_color)

            #     edge_kind = x_conn_shape.getPropertyValue("EdgeKind")
            #     if edge_kind.value == CONN_STANDARD_VALUE:
            #         self.set_connector_type_prop(Diagram.CONN_STANDARD)
            #     elif edge_kind.value == CONN_LINE_VALUE:
            #         self.set_connector_type_prop(Diagram.CONN_LINE)
            #     elif edge_kind.value == CONN_LINES_VALUE:
            #         self.set_connector_type_prop(Diagram.CONN_STRAIGHT)
            #     elif edge_kind.value == CONN_CURVE_VALUE:
            #         self.set_connector_type_prop(Diagram.CONN_CURVED)

            #     line_start_name = x_conn_shape.getPropertyValue("LineStartName")
            #     if line_start_name == "Arrow":
            #         self.set_connector_start_arrow_prop(True)

            #     line_end_name = x_conn_shape.getPropertyValue("LineEndName")
            #     if line_end_name == "Arrow":
            #         self.set_connector_end_arrow_prop(True) """

        #except Exception as ex:
        #    print(f"Error initializing properties: {ex}")

    @abstractmethod
    def init_diagram_tree(self, diagram_tree):
        """Initialize diagram tree - to be implemented by subclasses"""

    @abstractmethod
    def get_diagram_tree(self):
        """Get diagram tree - to be implemented by subclasses"""

    @abstractmethod
    def add_shape(self):
        """Add shape - to be implemented by subclasses"""

    def remove_shape(self, x_selected_shape=None):
        """Remove shape from organization chart"""
        if x_selected_shape is None:
            # No specific shape provided, remove all selected shapes
            x_selected_shapes = self.get_controller().get_selected_shapes()
            x_shape = None
            try:
                if x_selected_shapes is not None:
                    for i in range(x_selected_shapes.getCount()):
                        x_shape = x_selected_shapes.getByIndex(i)
                        if x_shape is not None:
                            self.remove_shape(x_shape)
            except Exception as ex:
                print(f"Error removing shapes: {ex}")
        else:
            # Remove specific shape
            if x_selected_shape is not None:
                selected_shape_name = x_selected_shape.getName()
                if "GraphicObjectShape" in selected_shape_name and "RectangleShape0" not in selected_shape_name:

                    if selected_shape_name.endswith("RectangleShape1"):
                        title = self.get_gui().get_dialog_property_value("Strings", "ShapeRemoveError.Title")
                        message = self.get_gui().get_dialog_property_value("Strings", "ShapeRemoveError.Message")
                        self.get_gui().show_message_box(title, message)
                    else:
                        # Clear everything under the item in the tree
                        selected_item = self.get_diagram_tree().get_tree_item(x_selected_shape)

                        no_item = False
                        dad_item = selected_item.get_dad()

                        if selected_item == dad_item.get_first_child():
                            if selected_item.get_first_sibling() is not None:
                                dad_item.set_first_child(selected_item.get_first_sibling())
                            else:
                                dad_item.set_first_child(None)
                                no_item = True
                        else:
                            previous_sibling = self.get_diagram_tree().get_previous_sibling(selected_item)
                            if previous_sibling is not None:
                                if selected_item.get_first_sibling() is not None:
                                    previous_sibling.set_first_sibling(selected_item.get_first_sibling())
                                else:
                                    previous_sibling.set_first_sibling(None)

                        x_dad_shape = selected_item.get_dad().get_rectangle_shape()

                        if selected_item.get_first_child() is not None:
                            selected_item.get_first_child().remove_items()

                        x_conn_shape = self.get_diagram_tree().get_dad_connector_shape(x_selected_shape)
                        if x_conn_shape is not None:
                            self.get_diagram_tree().remove_from_connectors(x_conn_shape)
                            self._x_shapes.remove(x_conn_shape)

                        self.get_diagram_tree().remove_from_rectangles(x_selected_shape)
                        self._x_shapes.remove(x_selected_shape)
                        self.set_null_selected_item(selected_item)

                        if (self.is_hidden_root_element_prop() and
                            self.get_diagram_tree().get_root_item().get_rectangle_shape() == x_dad_shape):
                            if no_item:
                                self.get_diagram_tree().get_root_item().hide_element()
                                self.set_hidden_root_element_prop(False)
                                self.get_controller().set_selected_shape(x_dad_shape)
                            else:
                                self.get_controller().set_selected_shape(
                                    self.get_diagram_tree().get_root_item().get_first_child().get_rectangle_shape()
                                )
                        else:
                            self.get_controller().set_selected_shape(x_dad_shape)

    def create_diagram(self, data=None):
        """Create diagram - base implementation"""
        print("Creating organization chart diagram...")
        super().create_diagram(data)

    def set_null_selected_item(self, item):
        """Set all references of the item to None"""
        item.set_dad(None)
        item.set_first_child(None)
        item.set_first_sibling(None)

    def show_property_dialog(self):
        """Show property dialog and apply changes"""
        self.get_gui().enable_control_dialog_window(False)
        exec_result = self.get_gui().execute_properties_dialog()
        if exec_result == 1:
            # TODO: Update shape
            pass

        self.get_gui().enable_and_set_focus_control_dialog()

    # Color arrays (simplified versions of the Java arrays)

    # Base orange colors
    _LO_ORANGES = [0xffc000, 0xff8000, 0xff4000, 0xff0000]

    # Base colors for organization charts
    _ORG_CHART_COLORS = [
        0x4f81bd, 0x9cbb58, 0xf79646, 0x8064a2,
        0x4bacc6, 0xc0504d, 0x1f497d, 0x17365d
    ]

    # Color matrix for organization charts (5 colors x 5 levels)
    _LO_COLORS_2 = [
        [0xffc000, 0xffd966, 0xffe699, 0xfff2cc, 0xfffdf4],  # Yellow series
        [0x70ad47, 0x9dc268, 0xc5e0b4, 0xe2efda, 0xf2f8f0],  # Green series
        [0x4472c4, 0x8db4e2, 0xc5d9f1, 0xe1ecf7, 0xf4f8fd],  # Blue series
        [0xff6600, 0xff9933, 0xffcc99, 0xffe6cc, 0xfff3e6],  # Orange series
        [0x7030a0, 0x9966cc, 0xccb3ff, 0xe6d9ff, 0xf3ecff]   # Purple series
    ]