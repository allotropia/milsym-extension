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
OrgChart Tree class
Python port of OrgChartTree.java
"""

from ..organization_chart_tree import OrganizationChartTree
from .orgchart_tree_item import OrgChartTreeItem


class OrgChartTree(OrganizationChartTree):
    """Organization chart tree implementation"""

    LAST_HOR_LEVEL = 2

    def __init__(self, organigram, control_shape_or_tree=None, root_item_shape=None):
        """
        Multiple constructor patterns:
        1. OrgChartTree(organigram) - basic constructor
        2. OrgChartTree(organigram, control_shape, root_item_shape) - with shapes
        3. OrgChartTree(organigram, diagram_tree) - copy from existing tree
        """
        if root_item_shape is not None:
            # Constructor with control and root shapes
            super().__init__(organigram)
            self.set_control_shape(control_shape_or_tree)
            OrgChartTreeItem.init_static_members()
            self.add_to_rectangles(root_item_shape)
            self._root_item = OrgChartTreeItem(self, root_item_shape, None, 0, 0.0)
        elif control_shape_or_tree is not None and hasattr(control_shape_or_tree, 'get_root_item'):
            # Constructor with existing diagram tree
            super().__init__(organigram, control_shape_or_tree)
            OrgChartTreeItem.init_static_members()
            self.get_org_chart().set_hor_level_of_control_shape(self.get_control_shape(), OrgChartTree.LAST_HOR_LEVEL)
            self._root_item = OrgChartTreeItem(self, None, control_shape_or_tree.get_root_item())
            self._root_item.set_level(0)
            self._root_item.set_pos(0.0)
            self._root_item.convert_tree_items(control_shape_or_tree.get_root_item())
        else:
            # Basic constructor
            super().__init__(organigram)

    def set_last_hor_level(self, level: int):
        """Set last horizontal level"""
        OrgChartTree.LAST_HOR_LEVEL = level
        self.get_org_chart().set_hor_level_of_control_shape(self._x_control_shape, level)

    def init_tree_items(self):
        """Initialize tree items"""
        OrgChartTreeItem.init_static_members()
        self._root_item = OrgChartTreeItem(self, self._x_root_shape, None, 0, 0)
        self._root_item.init_tree_items()

    def get_first_child_shape(self, x_dad_shape):
        """Get first child shape based on position"""
        # The structure of diagram changes below second level
        level = self.get_tree_item(x_dad_shape).get_level() + 1 if self.get_tree_item(x_dad_shape) else 1
        x_pos = -1
        y_pos = -1
        x_child_shape = None
        x_first_child_shape = None

        for x_conn_shape in self._connector_list:
            if x_dad_shape == self.get_start_shape_of_connector(x_conn_shape):
                x_child_shape = self.get_end_shape_of_connector(x_conn_shape)

                if level <= OrgChartTree.LAST_HOR_LEVEL:
                    # Horizontal layout - find leftmost child
                    child_pos = x_child_shape.getPosition() if hasattr(x_child_shape, 'getPosition') else None
                    if child_pos and (x_pos == -1 or child_pos.X < x_pos):
                        x_pos = child_pos.X
                        x_first_child_shape = x_child_shape
                else:
                    # Vertical layout - find topmost child
                    child_pos = x_child_shape.getPosition() if hasattr(x_child_shape, 'getPosition') else None
                    if child_pos and (y_pos == -1 or child_pos.Y < y_pos):
                        y_pos = child_pos.Y
                        x_first_child_shape = x_child_shape

        return x_first_child_shape

    def get_last_child_shape(self, x_dad_shape):
        """Get last child shape based on position"""
        level = self.get_tree_item(x_dad_shape).get_level() + 1 if self.get_tree_item(x_dad_shape) else 1
        x_pos = -1
        y_pos = -1
        x_child_shape = None
        x_last_child_shape = None

        for x_conn_shape in self._connector_list:
            if x_dad_shape == self.get_start_shape_of_connector(x_conn_shape):
                x_child_shape = self.get_end_shape_of_connector(x_conn_shape)

                if level <= OrgChartTree.LAST_HOR_LEVEL:
                    # Horizontal layout - find rightmost child
                    child_pos = x_child_shape.getPosition() if hasattr(x_child_shape, 'getPosition') else None
                    if child_pos and (x_pos == -1 or child_pos.X > x_pos):
                        x_pos = child_pos.X
                        x_last_child_shape = x_child_shape
                else:
                    # Vertical layout - find bottommost child
                    child_pos = x_child_shape.getPosition() if hasattr(x_child_shape, 'getPosition') else None
                    if child_pos and (y_pos == -1 or child_pos.Y > y_pos):
                        y_pos = child_pos.Y
                        x_last_child_shape = x_child_shape

        return x_last_child_shape

    def get_first_sibling_shape(self, x_base_shape, dad):
        """Get first sibling shape after base shape"""
        if dad is None or dad.get_rectangle_shape() is None:
            return None

        level = dad.get_level() + 1
        x_dad_shape = dad.get_rectangle_shape()
        x_sibling_shape = None
        x_first_sibling_shape = None
        base_shape_pos = x_base_shape.getPosition() if hasattr(x_base_shape, 'getPosition') else None

        if not base_shape_pos:
            return None

        x_pos = -1
        y_pos = -1

        for x_conn_shape in self._connector_list:
            if x_dad_shape == self.get_start_shape_of_connector(x_conn_shape):
                x_sibling_shape = self.get_end_shape_of_connector(x_conn_shape)
                sibling_pos = x_sibling_shape.getPosition() if hasattr(x_sibling_shape, 'getPosition') else None

                if not sibling_pos:
                    continue

                if level <= OrgChartTree.LAST_HOR_LEVEL:
                    # Horizontal layout - find next sibling to the right
                    if sibling_pos.X > base_shape_pos.X:
                        if x_pos == -1 or sibling_pos.X < x_pos:
                            x_pos = sibling_pos.X
                            x_first_sibling_shape = x_sibling_shape
                else:
                    # Vertical layout - find next sibling below
                    if sibling_pos.Y > base_shape_pos.Y:
                        if y_pos == -1 or sibling_pos.Y < y_pos:
                            y_pos = sibling_pos.Y
                            x_first_sibling_shape = x_sibling_shape

        return x_first_sibling_shape

    def refresh(self):
        """Refresh the tree"""
        OrgChartTreeItem.init_static_members()
        self._root_item.set_level(0)
        self._root_item.set_pos(0.0)
        self._root_item.set_positions_of_items()
        self._root_item.set_measure_props()
        self._root_item.display()

    def refresh_connector_props(self):
        """Refresh connector properties"""
        for x_conn_shape in self._connector_list:
            x_shape = self.get_end_shape_of_connector(x_conn_shape)
            tree_item = self.get_tree_item(x_shape)

            if tree_item:
                level = tree_item.get_level()
                start_pos = 2  # Bottom connection point

                if level <= OrgChartTree.LAST_HOR_LEVEL:
                    end_pos = 0  # Top connection point
                else:
                    end_pos = 3  # Left connection point

                self.get_org_chart().set_connector_shape_props(x_conn_shape, start_pos, end_pos)

    def get_end_glue_point_index(self, x_conn_shape) -> int:
        """Get end glue point index of connector"""
        try:
            # In real implementation, would use UNO API:
            # x_props = UnoRuntime.queryInterface(XPropertySet.class, x_conn_shape)
            # return AnyConverter.toInt(x_props.getPropertyValue("EndGluePointIndex"))
            return 0  # Stub implementation
        except Exception as ex:
            print(f"Error getting end glue point index: {ex}")
            return -1
