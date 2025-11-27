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
OrgChart Tree Item class
Python port of OrgChartTreeItem.java
"""

from typing import List

from ..organization_chart_tree_item import OrganizationChartTreeItem

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
        last_hor_level = getattr(self._diagram_tree, 'LAST_HOR_LEVEL', 2)

        x_first_child_shape = self.get_diagram_tree().get_first_child_shape(self._x_rectangle_shape)
        if x_first_child_shape is not None:
            first_child_level = self._level + 1
            first_child_pos = 0.0
            if first_child_level <= last_hor_level:
                first_child_pos = OrgChartTreeItem._max_positions[first_child_level] + 1.0
            else:
                first_child_pos = self._pos + 0.5
            self._first_child = OrgChartTreeItem(self.get_diagram_tree(), x_first_child_shape, self, first_child_level, first_child_pos)
            self._first_child.init_tree_items()

        if self._level == last_hor_level:
            deep = self.get_number_of_items_in_branch(self)
            if deep > 2:
                max_pos_in_level = OrgChartTreeItem._max_branch_positions[self._level + deep - 1]
                if self._pos < max_pos_in_level + 0.5:
                    if self.is_first_child():
                        self.get_first_child().increase_pos_in_branch(max_pos_in_level + 0.5 - self._pos)
                        self.set_pos(max_pos_in_level + 0.5)
            self.set_max_pos_of_branch()

        x_first_sibling_shape = self.get_diagram_tree().get_first_sibling_shape(self._x_rectangle_shape, self._dad)
        if x_first_sibling_shape is not None:
            first_sibling_level = self._level
            first_sibling_pos = self._pos + 1.0

            if first_sibling_level > last_hor_level:
                first_sibling_pos = self._pos
                first_sibling_level = self._level + self.get_number_of_items_in_branch(self)

            self._first_sibling = OrgChartTreeItem(self.get_diagram_tree(), x_first_sibling_shape, self._dad, first_sibling_level, first_sibling_pos)
            self._first_sibling.init_tree_items()

        if self._level <= last_hor_level and self._dad is not None and self._dad.get_first_child() == self:
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

        # Use fixed dimensions instead of scaling to fit available space
        org_chart = self.get_diagram_tree().get_org_chart()

        # Set fixed shape dimensions (convert to appropriate units)
        OrgChartTreeItem._shape_width = org_chart.get_shape_width() * 1000
        OrgChartTreeItem._shape_height = org_chart.get_shape_height() * 1000
        OrgChartTreeItem._hor_space = org_chart.get_hor_space() * 1000
        OrgChartTreeItem._ver_space = org_chart.get_ver_space() * 1000

        control_shape_pos = self.get_diagram_tree().get_control_shape_pos()
        OrgChartTreeItem._group_pos_x = control_shape_pos.X if control_shape_pos else 0
        OrgChartTreeItem._group_pos_y = control_shape_pos.Y if control_shape_pos else 0

    def set_pos_of_rect(self):
        """Set position of rectangle"""
        x_coord = OrgChartTreeItem._group_pos_x + int((OrgChartTreeItem._shape_width + OrgChartTreeItem._hor_space) * self.get_pos())
        last_hor_level = getattr(self._diagram_tree, 'LAST_HOR_LEVEL', 2)

        # Use smaller vertical spacing for levels beyond horizontal threshold (vertical stacking)
        if self._level > last_hor_level:
            # For vertically stacked levels, use much smaller vertical spacing
            vertical_spacing = OrgChartTreeItem._ver_space // 4  # Reduce to 1/4 of normal spacing
            # Calculate y position with reduced spacing for vertical levels
            base_y = OrgChartTreeItem._group_pos_y + (OrgChartTreeItem._shape_height + OrgChartTreeItem._ver_space) * last_hor_level
            vertical_offset = (OrgChartTreeItem._shape_height + vertical_spacing) * (self.get_level() - last_hor_level)
            y_coord = base_y + vertical_offset
        else:
            # Normal horizontal level spacing
            y_coord = OrgChartTreeItem._group_pos_y + (OrgChartTreeItem._shape_height + OrgChartTreeItem._ver_space) * self.get_level()

        if self.get_diagram_tree().get_org_chart().is_hidden_root_element_prop():
            if self == self.get_diagram_tree().get_root_item():
                y_coord = OrgChartTreeItem._group_pos_y - 10
            else:
                if self.get_level() > last_hor_level:
                    # Apply same reduced spacing logic for hidden root
                    vertical_spacing = OrgChartTreeItem._ver_space // 4
                    base_y = OrgChartTreeItem._group_pos_y + (OrgChartTreeItem._shape_height + OrgChartTreeItem._ver_space) * (last_hor_level - 1)
                    vertical_offset = (OrgChartTreeItem._shape_height + vertical_spacing) * (self.get_level() - last_hor_level)
                    y_coord = base_y + vertical_offset
                else:
                    y_coord = OrgChartTreeItem._group_pos_y + (OrgChartTreeItem._shape_height + OrgChartTreeItem._ver_space) * (self.get_level() - 1)

        # Calculate size based on graphic aspect ratio while fitting within default bounds
        calculated_width, calculated_height = self._calculate_size_for_aspect_ratio()

        if self._level > last_hor_level:
            self.set_position(Point(X=int(x_coord + calculated_width * 0.1), Y=y_coord))
            self.set_size(Size(Width=int(calculated_width * 0.9), Height=calculated_height))
        else:
            self.set_position(Point(X=x_coord, Y=y_coord))
            self.set_size(Size(Width=calculated_width, Height=calculated_height))

    def _calculate_size_for_aspect_ratio(self):
        """Calculate size that maintains the graphic's aspect ratio while fitting within default bounds"""
        default_width = OrgChartTreeItem._shape_width
        default_height = OrgChartTreeItem._shape_height

        try:
            if self._x_rectangle_shape.Graphic:
                # Get the original graphic size
                graphic_size = self._x_rectangle_shape.Graphic.SizePixel
                if graphic_size.Height > 0 and graphic_size.Width > 0:
                    # Calculate aspect ratio
                    aspect_ratio = graphic_size.Width / graphic_size.Height

                    # Calculate dimensions that fit within default bounds
                    width_from_height = int(default_height * aspect_ratio)
                    height_from_width = int(default_width / aspect_ratio)

                    # Choose the scaling that fits within both width and height constraints
                    if width_from_height <= default_width:
                        # Height is the limiting factor
                        return width_from_height, default_height
                    else:
                        # Width is the limiting factor
                        return default_width, height_from_width
        except Exception as ex:
            print(f"Could not get graphic aspect ratio: {ex}")

        # Fallback to default dimensions if we can't get aspect ratio
        return default_width, default_height