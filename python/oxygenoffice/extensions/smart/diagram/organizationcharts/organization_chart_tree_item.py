"""
Organization Chart Tree Item base class
Python port of OrganizationChartTreeItem.java
"""

from typing import Any, Optional, TYPE_CHECKING
from abc import ABC

import uno
from com.sun.star.drawing import XShape
from com.sun.star.awt import Point, Size

class OrganizationChartTreeItem(ABC):
    """Base class for organization chart tree items"""
    
    # Static class variables
    _max_level = -1
    
    def __init__(self, diagram_tree, dad=None, item=None):
        self._diagram_tree = diagram_tree
        self._dad = dad
        self._first_child = None
        self._first_sibling = None
        self._level = -1
        self._pos = -1.0
        
        if item is not None:
            # Copy constructor
            self._x_rectangle_shape = item._x_rectangle_shape
            self._rectangle_name = self._diagram_tree.get_org_chart().get_shape_name(self._x_rectangle_shape) if self._x_rectangle_shape else ""
        else:
            self._x_rectangle_shape = None
            self._rectangle_name = ""
    
    def __init_with_shape(self, diagram_tree, x_shape, dad):
        """Constructor with shape"""
        self._diagram_tree = diagram_tree
        self._x_rectangle_shape = x_shape
        self._rectangle_name = self._diagram_tree.get_org_chart().get_shape_name(x_shape) if x_shape else ""
        self._dad = dad
        self._first_child = None
        self._first_sibling = None
        self._level = -1
        self._pos = -1.0
    
    def hide_element(self):
        """Hide the element by setting fill and line style to none"""
        try:
            if self._x_rectangle_shape and self._diagram_tree.get_org_chart().is_hidden_root_element_prop():
                # In real implementation, would use UNO API to set properties:
                # xBaseProps.setPropertyValue("FillStyle", FillStyle.NONE)
                # xBaseProps.setPropertyValue("LineStyle", LineStyle.NONE)
                print("Hiding element")
        except Exception as ex:
            print(f"Error hiding element: {ex}")
    
    def is_hidden_element(self) -> bool:
        """Check if element is hidden"""
        try:
            if self._x_rectangle_shape:
                # In real implementation, would check FillStyle property
                # xProps = UnoRuntime.queryInterface(XPropertySet.class, shape)
                # style = xProps.getPropertyValue("FillStyle")
                # return style.getValue() == FillStyle.NONE_value
                return False
        except Exception as ex:
            print(f"Error checking if element is hidden: {ex}")
        return False
    
    def convert_tree_items(self, tree_item):
        """Convert tree items - to be overridden in subclasses"""
        pass
    
    def set_diagram_tree(self, diagram_tree):
        """Set diagram tree reference"""
        self._diagram_tree = diagram_tree
    
    def init_tree_items(self):
        """Initialize tree items - to be overridden in subclasses"""
        pass
    
    def set_positions_of_items(self):
        """Set positions of items - to be overridden in subclasses"""
        pass
    
    def set_pos_of_rect(self):
        """Set position of rectangle - to be overridden in subclasses"""
        pass
    
    def set_measure_props(self):
        """Set measure properties - to be overridden in subclasses"""
        pass
    
    def is_dad(self) -> bool:
        """Check if this item has a parent"""
        return self._dad is not None
    
    def get_dad(self):
        """Get parent item"""
        return self._dad
    
    def set_dad(self, dad):
        """Set parent item"""
        self._dad = dad
    
    def is_first_child(self) -> bool:
        """Check if this item has children"""
        return self._first_child is not None
    
    def get_first_child(self):
        """Get first child"""
        return self._first_child
    
    def set_first_child(self, child):
        """Set first child"""
        self._first_child = child
    
    def get_first_sibling(self):
        """Get first sibling"""
        return self._first_sibling
    
    def set_first_sibling(self, sibling):
        """Set first sibling"""
        self._first_sibling = sibling
    
    def is_first_sibling(self) -> bool:
        """Check if this item has siblings"""
        return self._first_sibling is not None
    
    def get_last_sibling(self):
        """Get last sibling in chain"""
        if self.is_first_sibling():
            return self.get_first_sibling().get_last_sibling()
        else:
            return self
    
    def get_last_child(self):
        """Get last child"""
        if self.is_first_child():
            return self.get_first_child().get_last_sibling()
        else:
            return None
    
    def get_rectangle_shape(self):
        """Get rectangle shape"""
        return self._x_rectangle_shape
    
    def get_position(self):
        """Get position of shape"""
        if self._x_rectangle_shape:
            # In real implementation: return self._x_rectangle_shape.getPosition()
            return None
        return None
    
    def set_position(self, point):
        """Set position of shape"""
        if self._x_rectangle_shape:
            # In real implementation: self._x_rectangle_shape.setPosition(point)
            pass
    
    def get_size(self):
        """Get size of shape"""
        if self._x_rectangle_shape:
            # In real implementation: return self._x_rectangle_shape.getSize()
            return None
        return None
    
    def set_size(self, size):
        """Set size of shape"""
        try:
            if self._x_rectangle_shape:
                # In real implementation: self._x_rectangle_shape.setSize(size)
                pass
        except Exception as ex:
            print(f"Error setting size: {ex}")
    
    def get_diagram_tree(self):
        """Get diagram tree reference"""
        return self._diagram_tree
    
    def set_pos(self, pos: float):
        """Set position - to be overridden in subclasses"""
        pass
    
    def get_pos(self) -> float:
        """Get position"""
        return self._pos
    
    def set_level(self, level: int):
        """Set level"""
        self._level = level
        if self._level > OrganizationChartTreeItem._max_level:
            OrganizationChartTreeItem._max_level = self._level
    
    def get_level(self) -> int:
        """Get level"""
        return self._level
    
    def get_previous_sibling(self, tree_item):
        """Get previous sibling of specified item"""
        previous_sibling = None
        
        if self._first_child is not None:
            item = self._first_child.get_previous_sibling(tree_item)
            if item is not None:
                previous_sibling = item
        
        if self.get_first_sibling() == tree_item:
            previous_sibling = self
        
        if self._first_sibling is not None:
            item = self._first_sibling.get_previous_sibling(tree_item)
            if item is not None:
                previous_sibling = item
        
        return previous_sibling
    
    def display(self):
        """Display the item - calls set_pos_of_rect and recurses to children"""
        self.set_pos_of_rect()
        
        if self._first_child is not None:
            self._first_child.display()
        
        if self._first_sibling is not None:
            self._first_sibling.display()
    
    def increase_pos_in_branch(self, diff: float):
        """Increase position in branch by diff amount"""
        self._pos += diff
        
        if self._first_child is not None:
            self._first_child.increase_pos_in_branch(diff)
        
        if self._first_sibling is not None:
            self._first_sibling.increase_pos_in_branch(diff)
    
    def print_tree(self):
        """Print tree structure for debugging"""
        if self._first_child is not None:
            self._first_child.print_tree()
        
        if self._first_sibling is not None:
            self._first_sibling.print_tree()
    
    def get_deep_of_tree_branch(self, tree_item) -> int:
        """Get depth of tree branch"""
        if tree_item._first_child is None:
            return 0
        else:
            max_depth = 0
            item = tree_item._first_child
            while item is not None:
                depth = self.get_deep_of_tree_branch(item)
                if depth > max_depth:
                    max_depth = depth
                item = item._first_sibling
            return max_depth + 1
    
    def get_deep_of_item(self) -> int:
        """Get depth of this item from root"""
        if self.is_dad():
            return self.get_dad().get_deep_of_item() + 1
        else:
            return 0
    
    def get_number_of_items_in_branch(self, tree_item) -> int:
        """Get number of items in branch"""
        if tree_item._first_child is None:
            return 1
        else:
            count = 1
            item = tree_item._first_child
            while item is not None:
                count += self.get_number_of_items_in_branch(item)
                item = item._first_sibling
            return count