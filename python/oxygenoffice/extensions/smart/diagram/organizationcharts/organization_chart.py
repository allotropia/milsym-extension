"""
Organization Chart base class
Python port of OrganizationChart.java
"""

from typing import Any, Optional, TYPE_CHECKING
from abc import ABC, abstractmethod

# Import base classes
from ..diagram import Diagram
from ..data_of_diagram import DataOfDiagram
from ..scheme_definitions import SchemeDefinitions

import uno
from com.sun.star.drawing import XShape
from com.sun.star.frame import XFrame


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
    
    def __init__(self, controller, gui, x_frame):
        super().__init__(controller, gui, x_frame)
        
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
        pass
    
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
        # In real implementation, would use UNO API:
        # x_named = UnoRuntime.queryInterface(XNamed.class, shape)
        # return x_named.getName()
        return f"Shape{id(shape)}"  # Stub
    
    def set_hor_level_of_control_shape(self, control_shape, level: int):
        """Set horizontal level of control shape"""
        # Stub implementation
        pass
    
    def get_hor_level_of_control_shape(self, control_shape) -> int:
        """Get horizontal level of control shape"""
        # Stub implementation
        return 2
    
    def remove_hor_level_props_of_control_shape(self, control_shape):
        """Remove horizontal level properties of control shape"""
        # Stub implementation
        pass
    
    def set_control_shape_props(self, control_shape):
        """Set control shape properties"""
        # Stub implementation
        pass
    
    def set_color_mode_and_style_of_control_shape(self, control_shape):
        """Set color mode and style of control shape"""
        # Stub implementation
        pass
    
    def get_top_shape_id(self) -> int:
        """Get top shape ID"""
        # Stub implementation
        return 10
    
    def clear_empty_diagram_and_recreate(self):
        """Clear empty diagram and recreate"""
        # Stub implementation
        pass
    
    @abstractmethod
    def init_diagram_tree(self, diagram_tree):
        """Initialize diagram tree - to be implemented by subclasses"""
        pass
    
    @abstractmethod
    def get_diagram_tree(self):
        """Get diagram tree - to be implemented by subclasses"""
        pass
    
    @abstractmethod
    def add_shape(self):
        """Add shape - to be implemented by subclasses"""
        pass
    
    def create_diagram(self, data=None):
        """Create diagram - base implementation"""
        super().init_diagram()
    
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