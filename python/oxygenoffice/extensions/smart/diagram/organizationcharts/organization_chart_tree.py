"""
Organization Chart Tree base class
Python port of OrganizationChartTree.java
"""

from typing import Any, List, Optional, TYPE_CHECKING
from abc import ABC, abstractmethod


import uno
from com.sun.star.drawing import XShape, XShapes
from com.sun.star.awt import Point, Size

class OrganizationChartTree(ABC):
    """Base class for organization chart trees"""
    
    def __init__(self, org_chart, diagram_tree=None):
        self._org_chart = org_chart
        self._x_shapes = org_chart.get_shapes()
        self._x_control_shape = None
        self._x_root_shape = None
        self._root_item = None
        self._selected_item = None
        
        if diagram_tree is None:
            # New tree
            self._rectangle_list = []
            self._connector_list = []
        else:
            # Copy from existing tree
            self._rectangle_list = diagram_tree._rectangle_list
            self._connector_list = diagram_tree._connector_list
            self._x_control_shape = diagram_tree._x_control_shape
            
            # Remove horizontal level properties if not organigram
            if self.get_org_chart().get_controller().get_diagram_type() != 0:  # Controller.ORGANIGRAM
                self.get_org_chart().remove_hor_level_props_of_control_shape(self._x_control_shape)
            
            self._x_root_shape = diagram_tree._x_root_shape
    
    @abstractmethod
    def init_tree_items(self):
        """Initialize tree items - to be implemented by subclasses"""
        pass
    
    @abstractmethod
    def get_first_child_shape(self, x_dad_shape):
        """Get first child shape - to be implemented by subclasses"""
        pass
    
    @abstractmethod
    def get_last_child_shape(self, x_dad_shape):
        """Get last child shape - to be implemented by subclasses"""
        pass
    
    @abstractmethod
    def get_first_sibling_shape(self, x_base_shape, dad):
        """Get first sibling shape - to be implemented by subclasses"""
        pass
    
    @abstractmethod
    def refresh(self):
        """Refresh tree - to be implemented by subclasses"""
        pass
    
    def get_org_chart(self):
        """Get organization chart reference"""
        return self._org_chart
    
    def get_root_item(self):
        """Get root item"""
        return self._root_item
    
    def set_control_shape(self, control_shape):
        """Set control shape"""
        self._x_control_shape = control_shape
    
    def get_control_shape(self):
        """Get control shape"""
        return self._x_control_shape
    
    def get_control_shape_pos(self):
        """Get control shape position"""
        if self._x_control_shape:
            # In real implementation: return self._x_control_shape.getPosition()
            return Point()  # Stub
        return None
    
    def get_control_shape_size(self):
        """Get control shape size"""
        if self._x_control_shape:
            # In real implementation: return self._x_control_shape.getSize()
            return Size()  # Stub
        return None
    
    def add_to_rectangles(self, shape):
        """Add shape to rectangles list"""
        self._rectangle_list.append(shape)
    
    def add_to_connectors(self, shape):
        """Add shape to connectors list"""
        self._connector_list.append(shape)
    
    def set_lists(self):
        """Set up lists from existing shapes"""
        # This would analyze existing shapes and populate lists
        pass
    
    def set_tree(self):
        """Set up tree structure"""
        # This would build the tree structure from shapes
        pass
    
    def get_tree_item(self, shape):
        """Get tree item for a given shape"""
        # This would search through the tree to find the item with the given shape
        # Simplified stub implementation
        return None
    
    def get_start_shape_of_connector(self, connector_shape):
        """Get start shape of connector"""
        # In real implementation, would query connector properties
        return None
    
    def get_end_shape_of_connector(self, connector_shape):
        """Get end shape of connector"""
        # In real implementation, would query connector properties
        return None
    
    def refresh_connector_props(self):
        """Refresh connector properties - can be overridden by subclasses"""
        pass