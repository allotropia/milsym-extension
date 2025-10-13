"""
OrgChart Tree Item class
Python port of OrgChartTreeItem.java
"""

from typing import List, Optional, TYPE_CHECKING

from ..organization_chart_tree_item import OrganizationChartTreeItem

import uno
from com.sun.star.drawing import XShape
from com.sun.star.awt import Point, Size

class OrgChartTreeItem(OrganizationChartTreeItem):
    """Organization chart tree item implementation"""
    
    # Static class variables
    _max_positions: List[float] = []
    _max_branch_positions: List[float] = []
    _max_pos = -1.0
    
    # Static measure variables
    _hor_space = 0
    _ver_space = 0
    _shape_width = 0
    _shape_height = 0
    _group_pos_x = 0
    _group_pos_y = 0
    
    def __init__(self, diagram_tree, dad_or_shape=None, item_or_dad=None, level=None, pos=None):
        """
        Multiple constructor patterns:
        1. OrgChartTreeItem(diagram_tree, dad, item) - copy constructor
        2. OrgChartTreeItem(diagram_tree, shape, dad, level, pos) - new item constructor
        """
        if level is not None and pos is not None:
            # Constructor with shape, dad, level, pos
            super().__init__(diagram_tree, item_or_dad, None)
            self._x_rectangle_shape = dad_or_shape
            self._rectangle_name = self._diagram_tree.get_org_chart().get_shape_name(dad_or_shape) if dad_or_shape else ""
            self.set_level(level)
            self.set_pos(pos)
        else:
            # Copy constructor
            super().__init__(diagram_tree, dad_or_shape, item_or_dad)
    
    @staticmethod
    def init_static_members():
        """Initialize static members"""
        OrganizationChartTreeItem._max_level = -1
        OrgChartTreeItem._max_pos = -1.0
        OrgChartTreeItem._max_positions = [-1.0] * 100
        OrgChartTreeItem._max_branch_positions = [-1.0] * 100
    
    def convert_tree_items(self, tree_item):
        """Convert tree items from another tree"""
        if tree_item.is_first_child():
            self._first_child = OrgChartTreeItem(self.get_diagram_tree(), self, tree_item.get_first_child())
            self._first_child.convert_tree_items(tree_item.get_first_child())
        
        if tree_item.is_first_sibling():
            self._first_sibling = OrgChartTreeItem(self.get_diagram_tree(), self.get_dad(), tree_item.get_first_sibling())
            self._first_sibling.convert_tree_items(tree_item.get_first_sibling())
    
    def set_pos(self, pos: float):
        """Set position and update max positions"""
        self._pos = pos
        if self._pos > OrgChartTreeItem._max_positions[self._level]:
            OrgChartTreeItem._max_positions[self._level] = self._pos
        if self._pos > OrgChartTreeItem._max_pos:
            OrgChartTreeItem._max_pos = self._pos
    
    def init_tree_items(self):
        """Initialize tree items recursively"""
        # Avoid circular import by accessing LAST_HOR_LEVEL through diagram tree
        last_hor_level = getattr(self._diagram_tree, 'LAST_HOR_LEVEL', 2)
        
        # Get first child shape
        x_first_child_shape = self.get_diagram_tree().get_first_child_shape(self._x_rectangle_shape)
        if x_first_child_shape is not None:
            first_child_level = self._level + 1
            first_child_pos = 0.0
            
            if first_child_level <= last_hor_level:
                first_child_pos = OrgChartTreeItem._max_positions[first_child_level] + 1.0
            else:
                first_child_pos = self._pos + 0.5
            
            self._first_child = OrgChartTreeItem(
                self.get_diagram_tree(), x_first_child_shape, self, first_child_level, first_child_pos
            )
            self._first_child.init_tree_items()
        
        # Handle branch positioning for items at LAST_HOR_LEVEL
        if self._level == last_hor_level:
            deep = self.get_number_of_items_in_branch(self)
            if deep > 2:
                max_pos_in_level = OrgChartTreeItem._max_branch_positions[self._level + deep - 1]
                if self._pos < max_pos_in_level + 0.5:
                    if self.is_first_child():
                        self.get_first_child().increase_pos_in_branch(max_pos_in_level + 0.5 - self._pos)
                        self.set_pos(max_pos_in_level + 0.5)
            self.set_max_pos_of_branch()
        
        # Get first sibling shape
        x_first_sibling_shape = self.get_diagram_tree().get_first_sibling_shape(self._x_rectangle_shape, self._dad)
        if x_first_sibling_shape is not None:
            first_sibling_level = self._level
            first_sibling_pos = self._pos + 1.0
            
            if first_sibling_level > last_hor_level:
                first_sibling_pos = self._pos
                first_sibling_level = self._level + self.get_number_of_items_in_branch(self)
            
            self._first_sibling = OrgChartTreeItem(
                self.get_diagram_tree(), x_first_sibling_shape, self._dad, first_sibling_level, first_sibling_pos
            )
            self._first_sibling.init_tree_items()
        
        # Position adjustment for horizontal levels
        if (self._level <= last_hor_level and 
            self._dad is not None and 
            self._dad.get_first_child() == self):
            
            new_pos = 0.0
            if self.is_first_sibling():
                new_pos = (OrgChartTreeItem._max_positions[self._level] + self._pos) / 2
            else:
                new_pos = self._pos
            
            if new_pos > self._dad.get_pos():
                self._dad.set_pos(new_pos)
            if new_pos < self._dad.get_pos():
                self.increase_pos_in_branch(self._dad.get_pos() - new_pos)
    
    def set_positions_of_items(self):
        """Set positions of items recursively"""
        last_hor_level = getattr(self._diagram_tree, 'LAST_HOR_LEVEL', 2)
        
        if self._first_child is not None:
            first_child_level = self._level + 1
            first_child_pos = 0.0
            
            if first_child_level <= last_hor_level:
                first_child_pos = OrgChartTreeItem._max_positions[first_child_level] + 1.0
            
            if first_child_level > last_hor_level:
                first_child_pos = self._pos + 0.5
            
            self._first_child.set_level(first_child_level)
            self._first_child.set_pos(first_child_pos)
            self._first_child.set_positions_of_items()
        
        # Handle branch positioning
        if self._level == last_hor_level:
            deep = self.get_number_of_items_in_branch(self)
            if deep > 2:
                max_pos_in_level = OrgChartTreeItem._max_branch_positions[self._level + deep - 1]
                if self._pos < max_pos_in_level + 0.5:
                    if self.is_first_child():
                        self.get_first_child().increase_pos_in_branch(max_pos_in_level + 0.5 - self._pos)
                        self.set_pos(max_pos_in_level + 0.5)
            self.set_max_pos_of_branch()
        
        if self._first_sibling is not None:
            first_sibling_level = self._level
            first_sibling_pos = self._pos + 1.0
            
            if first_sibling_level > last_hor_level:
                first_sibling_pos = self._pos
                first_sibling_level = self._level + self.get_number_of_items_in_branch(self)
            
            self._first_sibling.set_level(first_sibling_level)
            self._first_sibling.set_pos(first_sibling_pos)
            self._first_sibling.set_positions_of_items()
        
        # Position adjustment
        if (self._level <= last_hor_level and 
            self._dad is not None and 
            self._dad.get_first_child() == self):
            
            new_pos = 0.0
            if self.is_first_sibling():
                new_pos = (OrgChartTreeItem._max_positions[self._level] + self._pos) / 2
            else:
                new_pos = self._pos
            
            if new_pos > self._dad.get_pos():
                self._dad.set_pos(new_pos)
            if new_pos < self._dad.get_pos():
                self.increase_pos_in_branch(self._dad.get_pos() - new_pos)
    
    def set_max_pos_of_branch(self):
        """Set max position of branch"""
        last_hor_level = getattr(self._diagram_tree, 'LAST_HOR_LEVEL', 2)
        
        # Copy max positions to branch positions
        OrgChartTreeItem._max_branch_positions = OrgChartTreeItem._max_positions.copy()
        
        local_max = -1.0
        for i in range(len(OrgChartTreeItem._max_branch_positions)):
            if i > last_hor_level:
                if OrgChartTreeItem._max_branch_positions[i] > local_max:
                    local_max = OrgChartTreeItem._max_branch_positions[i]
                if OrgChartTreeItem._max_branch_positions[i] < local_max:
                    OrgChartTreeItem._max_branch_positions[i] = local_max
    
    def set_measure_props(self):
        """Set measure properties"""
        hidden_element_num = 0
        if self.get_diagram_tree().get_org_chart().is_hidden_root_element_prop():
            hidden_element_num = 1
        
        base_shape_size = self.get_diagram_tree().get_control_shape_size()
        base_shape_width = OrgChartTreeItem._shape_width = base_shape_size.Width if base_shape_size else 1000
        base_shape_height = OrgChartTreeItem._shape_height = base_shape_size.Height if base_shape_size else 1000
        
        OrgChartTreeItem._hor_space = OrgChartTreeItem._ver_space = 0
        
        if OrgChartTreeItem._max_pos > 0:
            org_chart = self.get_diagram_tree().get_org_chart()
            hor_unit = int(base_shape_width / (
                OrgChartTreeItem._max_pos * (org_chart.get_shape_width() + org_chart.get_hor_space()) + 
                org_chart.get_shape_width()
            ))
            OrgChartTreeItem._shape_width = hor_unit * org_chart.get_shape_width()
            OrgChartTreeItem._hor_space = hor_unit * org_chart.get_hor_space()
        
        if OrganizationChartTreeItem._max_level > 0:
            org_chart = self.get_diagram_tree().get_org_chart()
            ver_unit = base_shape_height // (
                (OrganizationChartTreeItem._max_level - hidden_element_num) * 
                (org_chart.get_shape_height() + org_chart.get_ver_space()) + 
                org_chart.get_shape_height()
            )
            OrgChartTreeItem._shape_height = ver_unit * org_chart.get_shape_height()
            OrgChartTreeItem._ver_space = ver_unit * org_chart.get_ver_space()
        
        control_shape_pos = self.get_diagram_tree().get_control_shape_pos()
        OrgChartTreeItem._group_pos_x = control_shape_pos.X if control_shape_pos else 0
        OrgChartTreeItem._group_pos_y = control_shape_pos.Y if control_shape_pos else 0
    
    def set_pos_of_rect(self):
        """Set position of rectangle"""
        x_coord = OrgChartTreeItem._group_pos_x + int((OrgChartTreeItem._shape_width + OrgChartTreeItem._hor_space) * self.get_pos())
        y_coord = OrgChartTreeItem._group_pos_y + (OrgChartTreeItem._shape_height + OrgChartTreeItem._ver_space) * self.get_level()
        
        if self.get_diagram_tree().get_org_chart().is_hidden_root_element_prop():
            if self == self.get_diagram_tree().get_root_item():
                y_coord = OrgChartTreeItem._group_pos_y - 10
            else:
                y_coord = OrgChartTreeItem._group_pos_y + (OrgChartTreeItem._shape_height + OrgChartTreeItem._ver_space) * (self.get_level() - 1)
        
        last_hor_level = getattr(self._diagram_tree, 'LAST_HOR_LEVEL', 2)
        if self._level > last_hor_level:
            self.set_position(Point(X=int(x_coord + OrgChartTreeItem._shape_width * 0.1), Y=y_coord))
            self.set_size(Size(Width=int(OrgChartTreeItem._shape_width * 0.9), Height=OrgChartTreeItem._shape_height))
        else:
            self.set_position(Point(X=x_coord, Y=y_coord))
            self.set_size(Size(Width=OrgChartTreeItem._shape_width, Height=OrgChartTreeItem._shape_height))