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
Organization Chart Tree Item base class
Python port of OrganizationChartTreeItem.java
"""

from abc import ABC

from com.sun.star.drawing.FillStyle import GRADIENT, NONE as FILL_STYLE_NONE, SOLID
from com.sun.star.drawing.LineStyle import NONE as LINE_STYLE_NONE, SOLID as LINE_STYLE_SOLID

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
            x_conn_props = None

            if self.get_diagram_tree().get_org_chart().is_hidden_root_element_prop():
                self._x_rectangle_shape.setPropertyValue("FillStyle", FILL_STYLE_NONE)
                self._x_rectangle_shape.setPropertyValue("LineStyle", LINE_STYLE_NONE)

                if (self.get_diagram_tree().get_org_chart().get_controller().get_diagram_type() !=
                    self.get_diagram_tree().get_org_chart().get_controller().TABLEHIERARCHYDIAGRAM):
                    for x_conn_shape in self.get_diagram_tree().connector_list:
                        if self._x_rectangle_shape == self.get_diagram_tree().get_start_shape_of_connector(x_conn_shape):
                            if x_conn_shape is not None:
                                x_conn_shape.setPropertyValue("LineStyle", LINE_STYLE_NONE)

                if (self.get_diagram_tree().get_org_chart().get_controller().get_selected_shape() ==
                    self._x_rectangle_shape):
                    self.get_diagram_tree().get_org_chart().get_controller().set_selected_shape(
                        self.get_first_child().get_rectangle_shape()
                    )
            else:
                if self.get_diagram_tree().get_org_chart().is_any_gradient_color_mode():
                    self._x_rectangle_shape.setPropertyValue("FillStyle", GRADIENT)
                else:
                    self._x_rectangle_shape.setPropertyValue("FillStyle", SOLID)

                if self.get_diagram_tree().get_org_chart().is_outline_prop():
                    self._x_rectangle_shape.setPropertyValue("LineStyle", LINE_STYLE_SOLID)
                else:
                    self._x_rectangle_shape.setPropertyValue("LineStyle", LINE_STYLE_NONE)

                if (self.get_diagram_tree().get_org_chart().get_controller().get_diagram_type() !=
                    self.get_diagram_tree().get_org_chart().get_controller().TABLEHIERARCHYDIAGRAM):
                    for x_conn_shape in self.get_diagram_tree().connector_list:
                        if self._x_rectangle_shape == self.get_diagram_tree().get_start_shape_of_connector(x_conn_shape):
                            if x_conn_shape is not None:
                                x_conn_shape.setPropertyValue("LineStyle", LINE_STYLE_SOLID)

        except Exception as ex:
            print(f"Error hiding element: {ex}")

    def is_hidden_element(self) -> bool:
        """Check if element is hidden"""
        try:
            if self._x_rectangle_shape is not None:
                fill_style = self._x_rectangle_shape.getPropertyValue("FillStyle")
                if fill_style == FILL_STYLE_NONE:
                    return True
        except Exception as ex:
            print(f"Error checking if element is hidden: {ex}")
        return False

    def convert_tree_items(self, tree_item):
        """Convert tree items - to be overridden in subclasses"""

    def set_diagram_tree(self, diagram_tree):
        """Set diagram tree reference"""
        self._diagram_tree = diagram_tree

    def init_tree_items(self):
        """Initialize tree items - to be overridden in subclasses"""

    def set_positions_of_items(self):
        """Set positions of items - to be overridden in subclasses"""

    def set_pos_of_rect(self):
        """Set position of rectangle - to be overridden in subclasses"""

    def set_measure_props(self):
        """Set measure properties - to be overridden in subclasses"""

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
        return self._x_rectangle_shape.getPosition()

    def set_position(self, point):
        """Set position of shape"""
        self._x_rectangle_shape.setPosition(point)

    def get_size(self):
        """Get size of shape"""
        return self._x_rectangle_shape.getSize()

    def set_size(self, size):
        """Set size of shape"""
        try:
            self._x_rectangle_shape.setSize(size)
        except Exception as ex:
            print(f"Error setting size: {ex}")

    def get_diagram_tree(self):
        """Get diagram tree reference"""
        return self._diagram_tree

    def set_pos(self, pos: float):
        """Set position - to be overridden in subclasses"""

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

    def search_item(self, x_shape):
        """Search for item with matching shape"""
        if self._first_child is not None:
            self._first_child.search_item(x_shape)
        if x_shape == self._x_rectangle_shape:
            self.get_diagram_tree().set_selected_item(self)
        if self._first_sibling is not None:
            self._first_sibling.search_item(x_shape)

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

    def set_properties(self):
        """Set properties recursively for tree items"""
        if self._first_child is not None:
            self._first_child.set_properties()
        self.get_diagram_tree().get_org_chart().set_shape_properties(self._x_rectangle_shape, Diagram.DIAGRAM_SHAPE_TYPE)
        if self._first_sibling is not None:
            self._first_sibling.set_properties()

    def remove_items(self):
        """Remove items recursively from tree"""
        if self.is_first_child():
            self._first_child.remove_items()

        x_conn_shape = self.get_diagram_tree().get_dad_connector_shape(self._x_rectangle_shape)
        if x_conn_shape is not None:
            self.get_diagram_tree().remove_from_connectors(x_conn_shape)
            self.get_diagram_tree().get_org_chart().remove_shape_from_group(x_conn_shape)

        self.get_diagram_tree().remove_from_rectangles(self._x_rectangle_shape)
        self.get_diagram_tree().get_org_chart().remove_shape_from_group(self._x_rectangle_shape)

        if self.is_first_sibling():
            self._first_sibling.remove_items()

    def increase_descendants_pos_num(self, diff: int):
        """Increase descendants position number"""
        self._first_child.increase_pos_in_branch(diff)

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
