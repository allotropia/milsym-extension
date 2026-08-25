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
        # Where each column starts, as measured by measure_columns. The list holds
        # (layout position, x offset) pairs in left to right order, and the dictionary
        # holds the x offset of each column head item.
        self._column_starts = []
        self._column_x_by_head = {}

        # Where each row of the side by side levels starts, one y offset per level, and
        # the y offset of each stacked item measured from the top of its column head.
        self._row_starts = []
        self._stacked_y_by_item = {}

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

    def measure_columns(self):
        """Work out where each column and each shape in it sits, from the sizes of the
        shapes the tree holds.

        A column is a shape on the last side by side level together with the shapes
        stacked below it. The shapes differ in size, because decorations such as echelon
        markers or text labels make a symbol larger. Each column is measured as its
        widest shape and the next column starts after that width plus the gap. Down a
        column each shape starts below the one before it, so a taller symbol moves the
        shapes below it further down. The rows above the columns are each as tall as
        their tallest shape.
        """
        self._column_starts = []
        self._column_x_by_head = {}
        self._row_starts = []
        self._stacked_y_by_item = {}

        heads = []
        row_heights = [0] * OrgChartTree.LAST_HOR_LEVEL
        pending = [self._root_item]
        while pending:
            item = pending.pop()
            if item is None:
                continue
            if item.get_level() == OrgChartTree.LAST_HOR_LEVEL:
                heads.append(item)
            elif item.get_level() < OrgChartTree.LAST_HOR_LEVEL:
                height = item._calculate_size_for_aspect_ratio()[1]
                if height > row_heights[item.get_level()]:
                    row_heights[item.get_level()] = height
                pending.append(item.get_first_child())
            pending.append(item.get_first_sibling())

        y_offset = 0
        for height in row_heights:
            self._row_starts.append(y_offset)
            y_offset += height + OrgChartTreeItem._ver_space
        self._row_starts.append(y_offset)

        heads.sort(key=lambda head: head.get_pos())

        x_offset = 0.0
        gap = OrgChartTreeItem.column_gap()
        for head in heads:
            self._column_starts.append((head.get_pos(), x_offset))
            self._column_x_by_head[head] = x_offset
            column_width, y_offset_by_item = head.measure_column()
            self._stacked_y_by_item.update(y_offset_by_item)
            x_offset += column_width + gap

    def horizontal_offset_of(self, item):
        """The x distance from the left edge of the diagram to the shape of this item.

        Column heads sit where measure_columns placed their column. A stacked shape sits
        at its column start plus its indent. A shape above the columns sits between the
        columns around its layout position, in proportion to where that position falls
        between theirs.
        """
        unit = OrgChartTreeItem.horizontal_pos_unit()

        if not self._column_starts:
            return item.get_pos() * unit

        level = item.get_level()

        if level > OrgChartTree.LAST_HOR_LEVEL:
            head = item.get_dad()
            while head is not None and head.get_level() > OrgChartTree.LAST_HOR_LEVEL:
                head = head.get_dad()
            head_x = self._column_x_by_head.get(head)
            if head_x is not None:
                return head_x + (item.get_pos() - head.get_pos()) * unit
            return item.get_pos() * unit

        if level == OrgChartTree.LAST_HOR_LEVEL:
            head_x = self._column_x_by_head.get(item)
            if head_x is not None:
                return head_x
            return item.get_pos() * unit

        pos = item.get_pos()
        previous_pos, previous_x = self._column_starts[0]
        if pos <= previous_pos:
            return previous_x
        for column_pos, column_x in self._column_starts[1:]:
            if pos <= column_pos:
                fraction = (pos - previous_pos) / (column_pos - previous_pos)
                return previous_x + fraction * (column_x - previous_x)
            previous_pos, previous_x = column_pos, column_x
        return previous_x

    def vertical_offset_of(self, item):
        """The y distance from the top of the diagram to the shape of this item.

        The rows above and including the column heads start where measure_columns put
        them. A stacked shape sits below the top of its column head by its measured
        offset, so it clears every shape above it in the column whatever their heights.
        """
        level = item.get_level()
        last_hor_level = OrgChartTree.LAST_HOR_LEVEL

        if level <= last_hor_level and level < len(self._row_starts):
            return self._row_starts[level]

        if level > last_hor_level:
            y_offset = self._stacked_y_by_item.get(item)
            if y_offset is not None and len(self._row_starts) > last_hor_level:
                return self._row_starts[last_hor_level] + y_offset

        # Not measured: every shape gets the same room, one full row per side by side
        # level and a quarter of the vertical space between stacked shapes
        row_step = OrgChartTreeItem._shape_height + OrgChartTreeItem._ver_space
        if level > last_hor_level:
            stacked_step = (
                OrgChartTreeItem._shape_height + OrgChartTreeItem._ver_space // 4
            )
            return row_step * last_hor_level + stacked_step * (level - last_hor_level)
        return row_step * level

    def hidden_root_row_shift(self):
        """The y distance the diagram moves up by when the root element is hidden: the
        height of the top row together with the space below it."""
        if len(self._row_starts) > 1:
            return self._row_starts[1]
        return OrgChartTreeItem._shape_height + OrgChartTreeItem._ver_space

    @timed("tree refresh")
    def refresh(self):
        """Refresh the tree"""
        with timed("set_positions_of_items"):
            self.recompute_levels_and_positions()
        self._root_item.set_measure_props()
        with timed("measure_columns"):
            self.measure_columns()
        with timed("display"):
            self._root_item.display()

    @timed("refresh_connector_props")
    def refresh_connector_props(self):
        """Refresh connector properties when tree structure has changed"""
        # The record holds what the extension last wrote. It describes the document except
        # where an undo has put an earlier state back, so the ends are read from the office
        # on the first pass after that, and taken from the record on every other pass. On
        # the 273 symbol document reading them costs 55 ms, which is ten times what the
        # whole pass costs without.
        ask_the_office = (
            not self._names_are_unique or self.connector_ends_may_have_changed()
        )

        for x_conn_shape in self._connector_list:
            connector_name = self.name_of_shape(x_conn_shape)
            if self._names_are_unique:
                recorded_start, recorded_end, recorded_start_glue, recorded_end_glue = (
                    self.get_connector_ends(connector_name)
                )
            else:
                recorded_start = recorded_end = None
                recorded_start_glue = recorded_end_glue = None

            if ask_the_office:
                current_end_shape = self.get_end_shape_of_connector(x_conn_shape)
            else:
                current_end_shape = self.get_shape_by_name(recorded_end)
                if current_end_shape is None:
                    # The record says nothing about this connector, so the office is the
                    # only place the shape it ends on can be found
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

            start_pos, end_pos = self.get_org_chart().connector_glue_positions(
                expected_start_shape, child_tree_item.get_level()
            )

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

        if ask_the_office:
            # Every end has been read, and any that had moved was written back, so the
            # record describes the document again
            self._connector_ends_may_have_changed = False
