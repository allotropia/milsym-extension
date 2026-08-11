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

from perf import count, timed

from ..organization_chart_tree import OrganizationChartTree
from .orgchart_tree_item import OrgChartTreeItem


class OrgChartTree(OrganizationChartTree):
    """Organization chart tree implementation"""

    LAST_HOR_LEVEL = 1

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
        elif control_shape_or_tree is not None and hasattr(
            control_shape_or_tree, "get_root_item"
        ):
            # Constructor with existing diagram tree
            super().__init__(organigram, control_shape_or_tree)
            OrgChartTreeItem.init_static_members()
            self._root_item = OrgChartTreeItem(
                self, None, control_shape_or_tree.get_root_item()
            )
            self._root_item.set_level(0)
            self._root_item.set_pos(0.0)
            self._root_item.convert_tree_items(control_shape_or_tree.get_root_item())
        else:
            # Basic constructor
            super().__init__(organigram)

    @timed("init_tree_items")
    def init_tree_items(self):
        """Initialize tree items"""
        OrgChartTreeItem.init_static_members()
        self._root_item = OrgChartTreeItem(self, self._x_root_shape, None, 0, 0)
        self._root_item.init_tree_items()

    def _level_below(self, x_dad_shape):
        """Level of the children of this shape, counting the root as level zero."""
        dad_item = self.get_tree_item(x_dad_shape)
        return dad_item.get_level() + 1 if dad_item else 1

    def _extreme_child_shape(self, x_dad_shape, want_last):
        """Return the child of this shape that sits furthest along the axis of its level.

        Children of the upper levels are laid out side by side, so they are ordered by
        their distance from the left edge. Below that they are stacked, so they are
        ordered by their distance from the top.
        """
        children = self.get_child_rect_names(self.name_of_shape(x_dad_shape))
        if not children:
            return None

        use_x = self._level_below(x_dad_shape) <= OrgChartTree.LAST_HOR_LEVEL
        chosen_name = None
        chosen_distance = None

        for child_name in children:
            position = self.get_rect_position(child_name)
            if position is None:
                continue
            distance = position[0] if use_x else position[1]
            if (
                chosen_distance is None
                or (want_last and distance > chosen_distance)
                or (not want_last and distance < chosen_distance)
            ):
                chosen_distance = distance
                chosen_name = child_name

        if chosen_name is None:
            # Nothing is known about where these children sit, which is the case for a
            # diagram that has only just been drawn. They were added in order, so the
            # order they appear in the group says which is first and which is last.
            chosen_name = children[-1] if want_last else children[0]

        return self.get_shape_by_name(chosen_name)

    def get_first_child_shape(self, x_dad_shape):
        """Get first child shape based on position"""
        if self._names_are_unique:
            return self._extreme_child_shape(x_dad_shape, want_last=False)
        return self._extreme_child_shape_by_searching(x_dad_shape, want_last=False)

    def get_last_child_shape(self, x_dad_shape):
        """Get last child shape based on position"""
        if self._names_are_unique:
            return self._extreme_child_shape(x_dad_shape, want_last=True)
        return self._extreme_child_shape_by_searching(x_dad_shape, want_last=True)

    def _extreme_child_shape_by_searching(self, x_dad_shape, want_last):
        """Find the outermost child by asking every connector which shapes it joins."""
        level = self._level_below(x_dad_shape)
        use_x = level <= OrgChartTree.LAST_HOR_LEVEL
        chosen_shape = None
        chosen_distance = None

        for x_conn_shape in self._connector_list:
            if x_dad_shape != self.get_start_shape_of_connector(x_conn_shape):
                continue

            x_child_shape = self.get_end_shape_of_connector(x_conn_shape)
            child_pos = (
                x_child_shape.getPosition()
                if hasattr(x_child_shape, "getPosition")
                else None
            )
            if not child_pos:
                continue

            distance = child_pos.X if use_x else child_pos.Y
            if (
                chosen_distance is None
                or (want_last and distance > chosen_distance)
                or (not want_last and distance < chosen_distance)
            ):
                chosen_distance = distance
                chosen_shape = x_child_shape

        return chosen_shape

    def get_first_sibling_shape(self, x_base_shape, dad):
        """Get first sibling shape after base shape"""
        if dad is None or dad.get_rectangle_shape() is None:
            return None

        if self._names_are_unique:
            return self._first_sibling_shape_from_indexes(x_base_shape, dad)
        return self._first_sibling_shape_by_searching(x_base_shape, dad)

    def _first_sibling_shape_from_indexes(self, x_base_shape, dad):
        """Return the next shape along from the base shape among the children of dad."""
        base_name = self.name_of_shape(x_base_shape)
        base_position = self.get_rect_position(base_name)
        if base_position is None:
            # Fall back to the order the children were added in, as above
            siblings = self.get_child_rect_names(dad.get_rectangle_name())
            if base_name not in siblings:
                return None
            following = siblings[siblings.index(base_name) + 1 :]
            return self.get_shape_by_name(following[0]) if following else None

        use_x = dad.get_level() + 1 <= OrgChartTree.LAST_HOR_LEVEL
        base_distance = base_position[0] if use_x else base_position[1]

        chosen_name = None
        chosen_distance = None
        for sibling_name in self.get_child_rect_names(dad.get_rectangle_name()):
            position = self.get_rect_position(sibling_name)
            if position is None:
                continue
            distance = position[0] if use_x else position[1]
            if distance <= base_distance:
                continue
            if chosen_distance is None or distance < chosen_distance:
                chosen_distance = distance
                chosen_name = sibling_name

        return self.get_shape_by_name(chosen_name) if chosen_name else None

    def _first_sibling_shape_by_searching(self, x_base_shape, dad):
        """Find the next sibling by asking every connector which shapes it joins."""
        level = dad.get_level() + 1
        x_dad_shape = dad.get_rectangle_shape()
        x_sibling_shape = None
        x_first_sibling_shape = None
        base_shape_pos = (
            x_base_shape.getPosition() if hasattr(x_base_shape, "getPosition") else None
        )

        if not base_shape_pos:
            return None

        x_pos = -1
        y_pos = -1

        for x_conn_shape in self._connector_list:
            if x_dad_shape == self.get_start_shape_of_connector(x_conn_shape):
                x_sibling_shape = self.get_end_shape_of_connector(x_conn_shape)
                sibling_pos = (
                    x_sibling_shape.getPosition()
                    if hasattr(x_sibling_shape, "getPosition")
                    else None
                )

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

    def recompute_levels_and_positions(self):
        """Work out the level and the layout position of every item from the tree shape.

        An item is told its level when it is made, but its level really follows from
        where it ends up in the tree, and code that builds a tree node by node reads the
        level of the node it added last to decide where the next one goes. Nothing here
        is written to the shapes, so it costs nothing to call while a tree is being
        built.
        """
        OrgChartTreeItem.init_static_members()
        self._root_item.set_level(0)
        self._root_item.set_pos(0.0)
        self._root_item.set_positions_of_items()

    @timed("tree refresh")
    def refresh(self):
        """Refresh the tree"""
        with timed("set_positions_of_items"):
            self.recompute_levels_and_positions()
        self._root_item.set_measure_props()
        with timed("display"):
            self._root_item.display()

    @timed("refresh_connector_props")
    def refresh_connector_props(self):
        """Refresh connector properties when tree structure has changed"""
        for x_conn_shape in self._connector_list:
            connector_name = self.name_of_shape(x_conn_shape)
            if self._names_are_unique:
                recorded_start, recorded_end, recorded_start_glue, recorded_end_glue = (
                    self.get_connector_ends(connector_name)
                )
                current_end_shape = self.get_shape_by_name(recorded_end)
            else:
                recorded_start = recorded_end = None
                recorded_start_glue = recorded_end_glue = None
                current_end_shape = self.get_end_shape_of_connector(x_conn_shape)

            if not current_end_shape:
                continue

            # Get the tree item for the end shape (child)
            child_tree_item = self.get_tree_item(current_end_shape)
            if not child_tree_item:
                continue

            # Get the correct parent from the tree structure
            parent_tree_item = (
                child_tree_item.get_dad()
                if hasattr(child_tree_item, "get_dad")
                else None
            )
            if not parent_tree_item:
                continue

            expected_start_shape = (
                parent_tree_item.get_rectangle_shape()
                if hasattr(parent_tree_item, "get_rectangle_shape")
                else None
            )

            level = child_tree_item.get_level()
            start_pos = 2  # Bottom connection point

            if level <= OrgChartTree.LAST_HOR_LEVEL:
                end_pos = 0  # Top connection point
            else:
                end_pos = 3  # Left connection point

            # Writing the ends of a connector makes the office reroute it, so leave alone
            # the connectors that already join the shapes they should at the right points
            if self._names_are_unique and (
                recorded_start == self.name_of_shape(expected_start_shape)
                and recorded_end == self.name_of_shape(current_end_shape)
                and recorded_start_glue == start_pos
                and recorded_end_glue == end_pos
            ):
                count("connector: rewiring skipped")
                continue

            self.get_org_chart().set_connector_shape_props(
                x_conn_shape,
                expected_start_shape,
                start_pos,
                current_end_shape,
                end_pos,
            )
