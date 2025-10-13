# SPDX-License-Identifier: MPL-2.0
# This file incorporates work covered by the following license notice:
#   SPDX-License-Identifier: LGPL-3.0-only

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
            return self._x_control_shape.getPosition()
        return None

    def get_control_shape_size(self):
        """Get control shape size"""
        if self._x_control_shape:
            return self._x_control_shape.getSize()
        return None

    def add_to_rectangles(self, shape):
        """Add shape to rectangles list"""
        self._rectangle_list.append(shape)

    def add_to_connectors(self, shape):
        """Add shape to connectors list"""
        self._connector_list.append(shape)

    def clear_lists(self):
        """Clear rectangle and connector lists"""
        if self._rectangle_list is not None:
            self._rectangle_list.clear()
        if self._connector_list is not None:
            self._connector_list.clear()

    def set_lists(self):
        """Set up lists from existing shapes"""
        try:
            self.clear_lists()
            curr_shape = None
            curr_shape_name = ""

            for i in range(self._x_shapes.getCount()):
                curr_shape = self._x_shapes.getByIndex(i)
                curr_shape_name = self.get_org_chart().get_shape_name(curr_shape)

                if "RectangleShape" in curr_shape_name:
                    if curr_shape_name.endswith("RectangleShape0"):
                        self.set_control_shape(curr_shape)
                    else:
                        self.add_to_rectangles(curr_shape)

                if "ConnectorShape" in curr_shape_name:
                    self.add_to_connectors(curr_shape)

        except Exception as ex:
            print(f"Error setting lists: {ex}")

    def set_root_item(self):
        """Set root item, return number of roots (if number is not 1, then there is an error)"""
        num_of_roots = 0
        # Search for root shape
        for rectangle_shape in self._rectangle_list:
            is_root = True
            for conn_shape in self._connector_list:
                if rectangle_shape == self.get_end_shape_of_connector(conn_shape):
                    is_root = False

            if is_root:
                num_of_roots += 1
                if self._x_root_shape is None:
                    self._x_root_shape = rectangle_shape
                else:
                    if rectangle_shape.getPosition().Y < self._x_root_shape.getPosition().Y:
                        self._x_root_shape = rectangle_shape

        return num_of_roots

    def set_tree(self):
        """Set up tree structure"""
        self._x_root_shape = None
        error = self.set_root_item()

        if self._x_root_shape is None or error > 1:
            title = self.get_org_chart().get_gui().get_dialog_property_value("Strings", "RoutShapeError.Title")
            message = self.get_org_chart().get_gui().get_dialog_property_value("Strings", "RoutShapeError.Message")
            self.get_org_chart().get_gui().show_message_box(title, message)
        else:
            self.init_tree_items()

    def get_tree_item(self, shape):
        """Get tree item for a given shape"""
        if self._x_root_shape is not None:
            if shape == self._x_root_shape:
                return self._root_item
            self._root_item.search_item(shape)
        return self._selected_item

    def get_start_shape_of_connector(self, connector_shape):
        """Get start shape of connector"""
        start_shape = None
        try:
            start_shape = connector_shape.getPropertyValue("StartShape")
        except Exception as ex:
            print(f"Error getting start shape: {ex}")
        return start_shape

    def get_end_shape_of_connector(self, connector_shape):
        """Get end shape of connector"""
        end_shape = None
        try:
            end_shape = connector_shape.getPropertyValue("EndShape")
        except Exception as ex:
            print(f"Error getting end shape: {ex}")
        return end_shape

    def refresh_connector_props(self):
        """Refresh connector properties - can be overridden by subclasses"""
        pass