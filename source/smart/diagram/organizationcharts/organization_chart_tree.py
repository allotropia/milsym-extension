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
Organization Chart Tree base class
Python port of OrganizationChartTree.java
"""

from abc import ABC, abstractmethod

from com.sun.star.awt import Point

from perf import SKIP_NAME_INDEX, count, timed

from ..diagram import Diagram


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
            self._reset_indexes()
        else:
            # Copy from existing tree
            self._rectangle_list = diagram_tree._rectangle_list
            self._connector_list = diagram_tree._connector_list
            self._x_control_shape = diagram_tree._x_control_shape

            # Share the lookup tables with the tree being copied from, the same way the
            # rectangle and connector lists are shared, so both trees see one topology
            self._shape_by_name = diagram_tree._shape_by_name
            self._position_by_rect_name = diagram_tree._position_by_rect_name
            self._size_by_rect_name = diagram_tree._size_by_rect_name
            self._control_shape_pos = diagram_tree._control_shape_pos
            self._may_read_geometry = diagram_tree._may_read_geometry
            self._start_name_by_connector_name = (
                diagram_tree._start_name_by_connector_name
            )
            self._end_name_by_connector_name = diagram_tree._end_name_by_connector_name
            self._glue_by_connector_name = diagram_tree._glue_by_connector_name
            self._connector_name_by_child_name = (
                diagram_tree._connector_name_by_child_name
            )
            self._child_names_by_rect_name = diagram_tree._child_names_by_rect_name
            self._item_by_rect_name = diagram_tree._item_by_rect_name
            self._names_are_unique = diagram_tree._names_are_unique

            # Remove horizontal level properties if not organigram
            if (
                self.get_org_chart().get_controller().get_diagram_type() != 0
            ):  # Controller.ORGANIGRAM
                self.get_org_chart().remove_hor_level_props_of_control_shape(
                    self._x_control_shape
                )

            self._x_root_shape = diagram_tree._x_root_shape

    def _reset_indexes(self):
        """Empty every lookup table and assume the shape names can key them again."""
        # Name of a rectangle or a connector, to the shape itself
        self._shape_by_name = {}
        # Name of a rectangle, to its position on the page as an (x, y) pair, and to its
        # size as a (width, height) pair. Both hold what was last read or written, and a
        # name is missing until one of those has happened.
        self._position_by_rect_name = {}
        self._size_by_rect_name = {}
        # Where the control shape sits, as an (x, y) pair, or None while unknown
        self._control_shape_pos = None
        # False while the office must not be asked where a shape is. Asking works out the
        # bounding box of the group, which recomputes the route of every connector in it,
        # and that needs a lock a menu command does not hold.
        self._may_read_geometry = True
        # Name of a connector, to the name of the rectangle each of its ends is glued to
        self._start_name_by_connector_name = {}
        self._end_name_by_connector_name = {}
        # True while the document may join the connectors up differently from the two tables
        # above, which is the case once an undo has put an earlier state of it back
        self._connector_ends_may_have_changed = False
        # Name of a connector, to its (start, end) glue point indexes
        self._glue_by_connector_name = {}
        # Name of a rectangle, to the connector that ends on it
        self._connector_name_by_child_name = {}
        # Name of a rectangle, to the names of the rectangles hanging below it
        self._child_names_by_rect_name = {}
        # Name of a rectangle, to the tree item that carries it
        self._item_by_rect_name = {}
        # False once a shape is found with no name, or with a name another shape in the
        # same group already uses. The lookups then give way to searching the shapes.
        self._names_are_unique = not SKIP_NAME_INDEX

    def name_of_shape(self, shape):
        """Return the name of a shape, or an empty string when it cannot be read."""
        if shape is None:
            return ""
        try:
            return shape.getName()
        except Exception:
            return ""

    @timed("build_indexes")
    def build_indexes(self, read_geometry=True):
        """Fill the lookup tables from the rectangles and connectors found in the group.

        One pass over the rectangles and one over the connectors, so the work grows with
        the number of shapes rather than with their product.

        Pass read_geometry as False when the diagram has just been drawn. Asking a shape
        inside a group where it is makes the office work out the bounding box of the whole
        group, which recomputes the route of every connector in it and announces the
        change, and announcing a change needs the drawing lock. A menu command reaches
        this code without that lock, because the office lets go of it before handing a
        command to an extension. Right after drawing a diagram the routes are all out of
        date, so that is exactly when the question is dangerous to ask. Where the geometry
        is not read, what an earlier pass learned is kept, the order shapes appear in the
        group stands in for their order on the page, and the sizes and positions that are
        still unknown are learned as they are written.
        """
        self._may_read_geometry = read_geometry or not self._connector_list

        # What an earlier pass learned about where the shapes are is kept unless the
        # caller means to measure them again, because a pass that may not ask the office
        # has no other way to know it.
        if read_geometry:
            self._position_by_rect_name = {}
            self._size_by_rect_name = {}
            self._control_shape_pos = None

        self._shape_by_name = {}
        self._start_name_by_connector_name = {}
        self._end_name_by_connector_name = {}
        self._glue_by_connector_name = {}
        self._connector_name_by_child_name = {}
        self._child_names_by_rect_name = {}
        self._item_by_rect_name = {}
        self._names_are_unique = not SKIP_NAME_INDEX
        # The ends are read from the office further down, so they describe the document
        self._connector_ends_may_have_changed = False

        for shape in self._rectangle_list + self._connector_list:
            name = self.name_of_shape(shape)
            if not name or name in self._shape_by_name:
                self._names_are_unique = False
                continue
            self._shape_by_name[name] = shape

        if self._x_control_shape is not None:
            control_name = self.name_of_shape(self._x_control_shape)
            if control_name:
                self._shape_by_name.setdefault(control_name, self._x_control_shape)

        if not self._names_are_unique:
            # Two shapes answering to one name make a position or a size held against
            # that name belong to either of them, so drop what was learned
            self._position_by_rect_name = {}
            self._size_by_rect_name = {}
            print(
                "Milsymbol: the shapes in this diagram do not have unique names, "
                "falling back to searching them"
            )
            return

        self._forget_geometry_of_departed_shapes()

        # Learn the origin while it is safe to ask, which is either because the caller
        # said so or because the group holds no connector yet
        self.get_control_shape_pos()

        if read_geometry:
            for shape in self._rectangle_list:
                name = self.name_of_shape(shape)
                try:
                    count("shape: read position and size")
                    position = shape.getPosition()
                    size = shape.getSize()
                except Exception:
                    continue
                self._position_by_rect_name[name] = (position.X, position.Y)
                self._size_by_rect_name[name] = (size.Width, size.Height)

        for connector in self._connector_list:
            connector_name = self.name_of_shape(connector)
            start_name = self.name_of_shape(
                self.get_start_shape_of_connector(connector)
            )
            end_name = self.name_of_shape(self.get_end_shape_of_connector(connector))

            self._start_name_by_connector_name[connector_name] = start_name
            self._end_name_by_connector_name[connector_name] = end_name
            self._glue_by_connector_name[connector_name] = self._read_glue_points(
                connector
            )

            if end_name:
                self._connector_name_by_child_name[end_name] = connector_name
            if start_name and end_name:
                self._child_names_by_rect_name.setdefault(start_name, []).append(
                    end_name
                )

    def _forget_geometry_of_departed_shapes(self):
        """Drop the positions and sizes held against names the group no longer holds.

        A name that has left comes back on the next shape the diagram makes, and that
        shape sits somewhere else, so what was learned about the shape that is gone would
        keep the new one from being placed.
        """
        for name in list(self._position_by_rect_name):
            if name not in self._shape_by_name:
                del self._position_by_rect_name[name]
        for name in list(self._size_by_rect_name):
            if name not in self._shape_by_name:
                del self._size_by_rect_name[name]

    def _read_glue_points(self, connector):
        """Return the (start, end) glue point indexes a connector is attached at."""
        try:
            return (
                connector.getPropertyValue("StartGluePointIndex"),
                connector.getPropertyValue("EndGluePointIndex"),
            )
        except Exception:
            return (None, None)

    def get_child_rect_names(self, rect_name):
        """Return the names of the rectangles hanging below the named rectangle."""
        return self._child_names_by_rect_name.get(rect_name, [])

    def get_child_connector_shapes(self, rect_name, x_rect_shape):
        """Return the connectors that run from a rectangle down to its children.

        Both the name and the shape are needed. The name reaches the lookup tables where
        the names can be trusted, and the shape is what the connectors are compared
        against where they cannot, which is the case this is asked in for a rectangle whose
        name another shape in the group also answers to.
        """
        if self._names_are_unique:
            shapes = []
            for child_name in self.get_child_rect_names(rect_name):
                connector_name = self._connector_name_by_child_name.get(child_name)
                connector = self._shape_by_name.get(connector_name)
                if connector is not None:
                    shapes.append(connector)
            return shapes

        if x_rect_shape is None:
            return []

        return [
            connector
            for connector in self._connector_list
            if x_rect_shape == self.get_start_shape_of_connector(connector)
        ]

    def get_rect_position(self, rect_name):
        """Return the last known position of the named rectangle as an (x, y) pair."""
        return self._position_by_rect_name.get(rect_name)

    def get_shape_by_name(self, name):
        """Return the shape with this name, or None when the group holds no such shape."""
        return self._shape_by_name.get(name)

    def has_trustworthy_names(self):
        """Say whether every shape in the group has a name of its own."""
        return self._names_are_unique

    def get_all_shape_names(self):
        """Names of every rectangle and connector the group holds."""
        return list(self._shape_by_name.keys())

    def knows_every_shape_in_the_group(self):
        """Say whether the lists account for every shape the group holds.

        The lists are built by walking the group, and the extension keeps them in step as
        it adds shapes and takes them away. A shape that arrived any other way, which is
        what an undo putting a deleted symbol back does, sits in the group and in neither
        list, and then the lists describe less than the diagram holds.
        """
        try:
            known = len(self._rectangle_list) + len(self._connector_list)
            if self._x_control_shape is not None:
                known += 1
            return known == self._x_shapes.getCount()
        except Exception:
            return False

    def get_tree_item_by_name(self, rect_name):
        """Return the tree item carrying the named rectangle, without touching the office."""
        return self._item_by_rect_name.get(rect_name)

    def register_item(self, item):
        """Record a tree item so that its rectangle can be looked up by name."""
        name = item.get_rectangle_name() if item is not None else ""
        if name:
            self._item_by_rect_name[name] = item

    def unregister_item(self, item):
        """Forget a tree item whose rectangle has left the diagram."""
        name = item.get_rectangle_name() if item is not None else ""
        self._item_by_rect_name.pop(name, None)

    def note_connector_ends(
        self, connector, start_shape, end_shape, start_glue, end_glue
    ):
        """Record which rectangles a connector now joins, and where it is glued to them."""
        connector_name = self.name_of_shape(connector)
        if not connector_name:
            return

        start_name = self.name_of_shape(start_shape)
        end_name = self.name_of_shape(end_shape)

        previous_start = self._start_name_by_connector_name.get(connector_name)
        previous_end = self._end_name_by_connector_name.get(connector_name)

        if previous_start and previous_end:
            siblings = self._child_names_by_rect_name.get(previous_start)
            if siblings is not None and previous_end in siblings:
                siblings.remove(previous_end)
        if previous_end and self._connector_name_by_child_name.get(previous_end) == (
            connector_name
        ):
            self._connector_name_by_child_name.pop(previous_end, None)

        self._start_name_by_connector_name[connector_name] = start_name
        self._end_name_by_connector_name[connector_name] = end_name
        self._glue_by_connector_name[connector_name] = (start_glue, end_glue)

        if end_name:
            self._connector_name_by_child_name[end_name] = connector_name
        if start_name and end_name:
            children = self._child_names_by_rect_name.setdefault(start_name, [])
            if end_name not in children:
                children.append(end_name)

    def forget_connector_ends(self):
        """Say that the document may join the connectors up differently from the record.

        The record itself is kept, because the tree is laid out from the shape each connector
        joins to which one, and the pass that joins them up again asks the office instead of
        trusting it until it has read every end once more.
        """
        self._connector_ends_may_have_changed = True

    def connector_ends_may_have_changed(self):
        """Say whether the recorded ends still have to be checked against the document."""
        return self._connector_ends_may_have_changed

    def get_connector_ends(self, connector_name):
        """Return the (start name, end name, start glue, end glue) recorded for a connector."""
        start_glue, end_glue = self._glue_by_connector_name.get(
            connector_name, (None, None)
        )
        return (
            self._start_name_by_connector_name.get(connector_name),
            self._end_name_by_connector_name.get(connector_name),
            start_glue,
            end_glue,
        )

    def note_rect_position(self, rect_name, x, y):
        """Record where a rectangle now sits, so that sibling order stays right.

        Nothing is recorded while the names cannot be trusted. Two shapes answering to one
        name share the entry, and the second of them would then be left where it is on the
        strength of where the first was put.
        """
        if rect_name and self._names_are_unique:
            self._position_by_rect_name[rect_name] = (x, y)

    def forget_geometry(self):
        """Forget where the shapes were put, so that the next layout writes them again."""
        self._position_by_rect_name.clear()
        self._size_by_rect_name.clear()

    def forget_graphic_aspect_ratios(self):
        """Forget the proportions of every symbol's picture, so that each is measured again."""
        pending = [self._root_item] if self._root_item is not None else []
        while pending:
            item = pending.pop()
            item.forget_graphic_aspect_ratio()
            if item.get_first_child() is not None:
                pending.append(item.get_first_child())
            if item.get_first_sibling() is not None:
                pending.append(item.get_first_sibling())

    def forget_geometry_of(self, rect_name):
        """Forget where one rectangle was put, so that the next layout writes it again."""
        if not rect_name:
            return
        self._position_by_rect_name.pop(rect_name, None)
        self._size_by_rect_name.pop(rect_name, None)

    def note_rect_size(self, rect_name, width, height):
        """Record how big a rectangle now is.

        Nothing is recorded while the names cannot be trusted, for the same reason the
        positions are not.
        """
        if rect_name and self._names_are_unique:
            self._size_by_rect_name[rect_name] = (width, height)

    def get_rect_size(self, rect_name):
        """Return the last known size of the named rectangle as a (width, height) pair."""
        return self._size_by_rect_name.get(rect_name)

    @abstractmethod
    def init_tree_items(self):
        """Initialize tree items - to be implemented by subclasses"""

    @abstractmethod
    def get_first_child_shape(self, x_dad_shape):
        """Get first child shape - to be implemented by subclasses"""

    @abstractmethod
    def get_last_child_shape(self, x_dad_shape):
        """Get last child shape - to be implemented by subclasses"""

    @abstractmethod
    def get_first_sibling_shape(self, x_base_shape, dad):
        """Get first sibling shape - to be implemented by subclasses"""

    @abstractmethod
    def refresh(self):
        """Refresh tree - to be implemented by subclasses"""

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
        """Where the control shape sits, which is the origin the layout is measured from.

        Answered from what was read when the diagram was opened, so that laying a diagram
        out never has to ask the office where a shape is. Returns None when the origin was
        never learned, and the layout then starts from the corner of the group.
        """
        if self._control_shape_pos is not None:
            return Point(X=self._control_shape_pos[0], Y=self._control_shape_pos[1])

        if self._x_control_shape is None or not self._may_read_geometry:
            return None

        count("shape: read control shape position")
        position = self._x_control_shape.getPosition()
        self._control_shape_pos = (position.X, position.Y)
        return position

    def update_origin(self):
        """Read again where the control shape group shape sits, and
        move what is kept about the other shapes along with it.

        The kept origin is a page coordinate. Dragging the group, or undoing such a
        drag, carries the control shape and every other shape the same distance, and
        what is kept then describes where the shapes stood before the move. Reading the
        origin again and shifting the kept positions by the distance it moved makes them
        describe the shapes as they stand now, so the next layout writes only what the
        edit itself changes.

        This asks the office where a shape is. That is safe while nothing has been
        changed yet: every connector route is current, so working out the bounding box of
        the group has nothing to route again and nothing to announce.

        """
        if self._x_control_shape is None:
            return
        try:
            count("shape: read control shape position")
            position = self._x_control_shape.getPosition()
        except Exception as ex:
            print(f"Error reading the control shape position: {ex}")
            return

        origin = (position.X, position.Y)
        kept_origin = self._control_shape_pos
        self._control_shape_pos = origin
        if kept_origin is None or kept_origin == origin:
            return

        dx = origin[0] - kept_origin[0]
        dy = origin[1] - kept_origin[1]
        for name, (x, y) in list(self._position_by_rect_name.items()):
            self._position_by_rect_name[name] = (x + dx, y + dy)

    def add_to_rectangles(self, shape):
        """Add shape to rectangles list"""
        self._rectangle_list.append(shape)
        name = self.name_of_shape(shape)
        if not name or name in self._shape_by_name:
            self._names_are_unique = False
            return
        self._shape_by_name[name] = shape

    def remove_from_rectangles(self, shape):
        """Remove shape from rectangles list"""
        name = self.name_of_shape(shape)
        self._remove_shape_from_list(self._rectangle_list, shape, name)

        parent_name = None
        connector_name = self._connector_name_by_child_name.pop(name, None)
        if connector_name is not None:
            parent_name = self._start_name_by_connector_name.get(connector_name)
        if parent_name:
            siblings = self._child_names_by_rect_name.get(parent_name)
            if siblings is not None and name in siblings:
                siblings.remove(name)

        self._shape_by_name.pop(name, None)
        self._position_by_rect_name.pop(name, None)
        self._size_by_rect_name.pop(name, None)
        self._child_names_by_rect_name.pop(name, None)
        self._item_by_rect_name.pop(name, None)

    def add_to_connectors(self, shape):
        """Add shape to connectors list"""
        self._connector_list.append(shape)
        name = self.name_of_shape(shape)
        if not name or name in self._shape_by_name:
            self._names_are_unique = False
            return
        self._shape_by_name[name] = shape

    def remove_from_connectors(self, shape):
        """Remove shape from connectors list"""
        name = self.name_of_shape(shape)
        self._remove_shape_from_list(self._connector_list, shape, name)

        start_name = self._start_name_by_connector_name.pop(name, None)
        end_name = self._end_name_by_connector_name.pop(name, None)
        self._glue_by_connector_name.pop(name, None)
        self._shape_by_name.pop(name, None)

        if end_name and self._connector_name_by_child_name.get(end_name) == name:
            self._connector_name_by_child_name.pop(end_name, None)
        if start_name and end_name:
            siblings = self._child_names_by_rect_name.get(start_name)
            if siblings is not None and end_name in siblings:
                siblings.remove(end_name)

    def _remove_shape_from_list(self, shapes, shape, name):
        """Drop a shape from one of the lists, matching on name where names can be trusted.

        Matching on name avoids comparing office objects, which costs a call across the
        bridge for every entry the list holds.
        """
        if self._names_are_unique and name:
            for index, candidate in enumerate(shapes):
                if self.name_of_shape(candidate) == name:
                    del shapes[index]
                    return
        try:
            shapes.remove(shape)
        except ValueError:
            pass

    def clear_lists(self):
        """Clear rectangle and connector lists"""
        if self._rectangle_list is not None:
            self._rectangle_list.clear()
        if self._connector_list is not None:
            self._connector_list.clear()
        self._reset_indexes()

    @timed("set_lists")
    def set_lists(self, read_geometry=True):
        """Set up lists from existing shapes"""
        try:
            # Emptying the tables loses where the shapes are, and a pass that may not
            # measure cannot learn it again, so hand it to build_indexes to keep or drop
            kept_positions = self._position_by_rect_name
            kept_sizes = self._size_by_rect_name
            kept_control_shape_pos = self._control_shape_pos

            self.clear_lists()

            self._position_by_rect_name = kept_positions
            self._size_by_rect_name = kept_sizes
            self._control_shape_pos = kept_control_shape_pos

            curr_shape = None
            curr_shape_name = ""

            for i in range(self._x_shapes.getCount()):
                curr_shape = self._x_shapes.getByIndex(i)
                curr_shape_name = self.get_org_chart().get_shape_name(curr_shape)
                role = Diagram.shape_role(curr_shape_name)

                if role == Diagram.DIAGRAM_BASE_SHAPE_TYPE:
                    self.set_control_shape(curr_shape)
                elif role == Diagram.DIAGRAM_SHAPE_TYPE:
                    self.add_to_rectangles(curr_shape)
                elif role == Diagram.CONNECTOR_SHAPE:
                    self.add_to_connectors(curr_shape)

            self.build_indexes(read_geometry)

        except Exception as ex:
            print(f"Error setting lists: {ex}")

    @timed("set_root_item")
    def set_root_item(self):
        """Set root item, return number of roots (if number is not 1, then there is an error)"""
        if self._names_are_unique:
            return self._set_root_item_from_indexes()
        return self._set_root_item_by_searching()

    def _set_root_item_from_indexes(self):
        """Pick the root as the topmost rectangle that no connector ends on."""
        num_of_roots = 0
        best_y = None

        for rectangle_shape in self._rectangle_list:
            name = self.name_of_shape(rectangle_shape)
            if name in self._connector_name_by_child_name:
                continue

            num_of_roots += 1
            position = self._position_by_rect_name.get(name, (0, 0))
            if self._x_root_shape is None or position[1] < best_y:
                self._x_root_shape = rectangle_shape
                best_y = position[1]

        return num_of_roots

    def _set_root_item_by_searching(self):
        """Pick the root by asking every connector where it ends, for unnamed shapes."""
        num_of_roots = 0
        for rectangle_shape in self._rectangle_list:
            is_root = True
            for conn_shape in self._connector_list:
                if rectangle_shape == self.get_end_shape_of_connector(conn_shape):
                    is_root = False
                    break

            if is_root:
                num_of_roots += 1
                if self._x_root_shape is None:
                    self._x_root_shape = rectangle_shape
                else:
                    if (
                        rectangle_shape.getPosition().Y
                        < self._x_root_shape.getPosition().Y
                    ):
                        self._x_root_shape = rectangle_shape

        return num_of_roots

    @timed("set_tree")
    def set_tree(self):
        """Set up tree structure"""
        self._x_root_shape = None
        error = self.set_root_item()

        if self._x_root_shape is None or error > 1:
            title = (
                self.get_org_chart()
                .get_gui()
                .get_dialog_property_value("Strings", "RoutShapeError.Title")
            )
            message = (
                self.get_org_chart()
                .get_gui()
                .get_dialog_property_value("Strings", "RoutShapeError.Message")
            )
            self.get_org_chart().get_gui().show_message_box(title, message)
        else:
            self.init_tree_items()

    def get_tree_item(self, shape):
        """Get tree item for a given shape"""
        count("tree: get_tree_item")
        if shape is None:
            return None

        if self._names_are_unique:
            item = self._item_by_rect_name.get(self.name_of_shape(shape))
            if item is not None:
                self._selected_item = item
            return item

        if self._root_item is None:
            return None
        if self._x_root_shape is not None and shape == self._x_root_shape:
            return self._root_item
        # search_item() sets self._selected_item as a side effect
        self._selected_item = None
        self._root_item.search_item(shape)
        return self._selected_item

    def get_start_shape_of_connector(self, connector_shape):
        """Get start shape of connector"""
        start_shape = None
        count("connector: read StartShape")
        try:
            start_shape = connector_shape.getPropertyValue("StartShape")
        except Exception as ex:
            print(f"Error getting start shape: {ex}")
        return start_shape

    def get_end_shape_of_connector(self, connector_shape):
        """Get end shape of connector"""
        end_shape = None
        count("connector: read EndShape")
        try:
            end_shape = connector_shape.getPropertyValue("EndShape")
        except Exception as ex:
            print(f"Error getting end shape: {ex}")
        return end_shape

    def get_dad_connector_shape(self, x_rect_shape):
        """Get connector shape that connects to this rectangle shape"""
        if self._names_are_unique:
            connector_name = self._connector_name_by_child_name.get(
                self.name_of_shape(x_rect_shape)
            )
            return self._shape_by_name.get(connector_name) if connector_name else None

        for x_conn_shape in self._connector_list:
            if x_rect_shape == self.get_end_shape_of_connector(x_conn_shape):
                return x_conn_shape
        return None

    def refresh_connector_props(self):
        """Refresh connector properties - can be overridden by subclasses"""

    def set_selected_item(self, tree_item):
        """Set the selected tree item"""
        self._selected_item = tree_item

    def get_previous_sibling(self, tree_item):
        """Get previous sibling of given tree item"""
        return self._root_item.get_previous_sibling(tree_item)
