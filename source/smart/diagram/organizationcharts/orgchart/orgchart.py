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
OrgChart class - Main organization chart implementation
Python port of OrgChart.java
"""

from utils import generate_icon_svg, get_recorded_symbol_size_px, locked_controllers
from perf import count
from ...diagram import Diagram
from ..organization_chart import OrganizationChart
from .orgchart_tree import OrgChartTree
from .orgchart_tree_item import OrgChartTreeItem, STACKED_CHANNEL_FRACTION

from com.sun.star.awt import Point
from com.sun.star.drawing import GluePoint2
from com.sun.star.lang import IndexOutOfBoundsException
from com.sun.star.drawing.EscapeDirection import VERTICAL as ESCAPE_VERTICAL
from com.sun.star.drawing.EscapeDirection import DOWN as ESCAPE_DOWN
from com.sun.star.drawing.EscapeDirection import UP as ESCAPE_UP
from com.sun.star.drawing.EscapeDirection import LEFT as ESCAPE_LEFT
from com.sun.star.drawing.Alignment import CENTER as ALIGNMENT_CENTER

# A shape starts with four builtin glue points, indices 0 to 3, one on the middle of each
# edge: 0 top, 1 right, 2 bottom, 3 left. The first user defined point gets index 4.
BUILTIN_GLUE_POINT_COUNT = 4
GLUE_TOP = 0
GLUE_BOTTOM = 2
GLUE_LEFT = 3

# The user defined glue points of a shape that shows a milsymbol drawing. They are placed
# from where the frame outline and the anchor of the drawing sit within the picture.
# ANCHOR_START_GLUE is where every connector to a child leaves: the anchor of the symbol,
# moved down to the bottom of the frame when the anchor lies inside it. For a
# headquarters the anchor is the end of the staff, so the children hang off the staff.
# ANCHOR_TOP_GLUE is where a connector arrives on a child that hangs below its parent: the
# top of the frame, above the anchor. ANCHOR_LEFT_GLUE is where a connector arrives on a
# stacked child: the middle of the frame's left edge.
ANCHOR_START_GLUE = 4
ANCHOR_TOP_GLUE = 5
ANCHOR_LEFT_GLUE = 6

def relative_glue_position(x_fraction, y_fraction):
    """The Position of a relative glue point aligned at the centre of its shape, for a point
    given as fractions of the shape's width and height with the top left corner at (0, 0).

    A relative glue point measures its position in hundredths of a percent of the shape's
    width and height, offset from the alignment point. The centre lies at (0, 0) and the
    edges at -5000 and 5000.
    """
    return Point(
        X=int(round((x_fraction - 0.5) * 10000)),
        Y=int(round((y_fraction - 0.5) * 10000)),
    )


# Where on the bottom edge of a shape without symbol geometry the connectors to its
# stacked children leave: a small fraction of the width in from the left edge, so that
# the point stays left of the children whatever the shape's width.
STACKED_GLUE_POSITION = relative_glue_position(STACKED_CHANNEL_FRACTION, 1.0)


def anchor_glue_point_positions(geometry):
    """The Position and Escape of each user defined glue point of a shape with the given
    SymbolGeometry, keyed by glue point index."""
    return {
        ANCHOR_START_GLUE: (
            relative_glue_position(
                geometry.anchor_x(), max(geometry.anchor_y(), geometry.frame_bottom())
            ),
            ESCAPE_DOWN,
        ),
        ANCHOR_TOP_GLUE: (
            relative_glue_position(geometry.anchor_x(), geometry.frame_top()),
            ESCAPE_UP,
        ),
        ANCHOR_LEFT_GLUE: (
            relative_glue_position(geometry.frame_left(), geometry.frame_centre_y()),
            ESCAPE_LEFT,
        ),
    }


def _same_glue_point(glue, position, escape):
    """Whether a glue point already is a relative, centre aligned point at this position
    with this escape direction."""
    return (
        glue.IsRelative
        and glue.PositionAlignment == ALIGNMENT_CENTER
        and glue.Escape == escape
        and glue.Position.X == position.X
        and glue.Position.Y == position.Y
    )


def _make_glue_point(position, escape):
    glue = GluePoint2()
    glue.IsRelative = True
    glue.Position = position
    glue.Escape = escape
    glue.PositionAlignment = ALIGNMENT_CENTER
    glue.IsUserDefined = True
    return glue


def _replace_glue_point(glue_points, index, glue):
    """Replace the glue point at this index and say whether the point now is the one given.

    The office applies the replacement and then raises IndexOutOfBoundsException all the
    same, so the exception says nothing about the outcome; the point is read back instead.
    """
    try:
        glue_points.replaceByIndex(index, glue)
    except IndexOutOfBoundsException:
        pass
    replaced = glue_points.getByIndex(index)
    return _same_glue_point(replaced, glue.Position, glue.Escape)


def update_anchor_glue_points(shape, geometry):
    """Keep the three user defined glue points of a shape that shows a milsymbol drawing
    at the places that follow from the drawing's SymbolGeometry.

    The points are relative to the shape's size, so a shape that is only resized keeps
    them without a write. A point already at its place is left alone. A shape that has
    fewer user defined points than three, for example one made before these points were
    introduced with only the stacked start point at index 4, gets the missing ones added.
    """
    try:
        glue_points = shape.getGluePoints()
        wanted = anchor_glue_point_positions(geometry)
        for index in (ANCHOR_START_GLUE, ANCHOR_TOP_GLUE, ANCHOR_LEFT_GLUE):
            position, escape = wanted[index]
            if index < glue_points.getCount():
                glue = glue_points.getByIndex(index)
                if _same_glue_point(glue, position, escape):
                    continue
                count("shape: glue point moved")
                if not _replace_glue_point(
                    glue_points, index, _make_glue_point(position, escape)
                ):
                    print(f"Warning: glue point {index} of a shape could not be moved")
            else:
                count("shape: glue point added")
                glue_points.insertByIndex(index, _make_glue_point(position, escape))
    except Exception as ex:
        print(f"Error setting the anchor glue points: {ex}")


def update_stacked_glue_point(shape, add_if_missing=False):
    """Keep the glue point the connectors to stacked children start on at its place on
    the bottom edge of a shape without symbol geometry.

    A shape that carries no user defined point yet is given one when add_if_missing says
    so; otherwise, and when the point cannot be added, the index of the builtin bottom
    centre point is returned instead. Returns the index of the point to start on.
    """
    try:
        glue_points = shape.getGluePoints()
        curr_count = glue_points.getCount()

        if curr_count > ANCHOR_START_GLUE:
            glue = glue_points.getByIndex(ANCHOR_START_GLUE)
            if not _same_glue_point(glue, STACKED_GLUE_POSITION, ESCAPE_VERTICAL):
                count("shape: glue point moved")
                _replace_glue_point(
                    glue_points,
                    ANCHOR_START_GLUE,
                    _make_glue_point(STACKED_GLUE_POSITION, ESCAPE_VERTICAL),
                )
            return ANCHOR_START_GLUE

        if not add_if_missing:
            return GLUE_BOTTOM

        count("shape: glue point added")
        glue_points.insertByIndex(
            curr_count, _make_glue_point(STACKED_GLUE_POSITION, ESCAPE_VERTICAL)
        )
        return curr_count
    except Exception as ex:
        print(f"Error setting the stacked connector glue point: {ex}")
        return GLUE_BOTTOM


class OrgChart(OrganizationChart):
    """Organization chart implementation"""

    def __init__(self, controller, gui, x_frame, x_context):
        super().__init__(controller, gui, x_frame, x_context)
        self._diagram_tree = None

        # Set specific dimensions for org chart
        self._group_width = 10
        self._group_height = 6
        self._shape_width = 2
        self._hor_space = 1
        self._shape_height = 1
        self._ver_space = 1

    def init_diagram_tree(self, diagram_tree):
        """Initialize diagram tree"""
        super().init_diagram()
        self._diagram_tree = OrgChartTree(self, diagram_tree)

    def get_diagram_tree(self):
        """Get diagram tree"""
        return self._diagram_tree

    def get_diagram_type_name(self) -> str:
        """Get diagram type name"""
        return "OrbatDiagram"

    def create_diagram(self, datas):
        """Create diagram from data"""
        if isinstance(datas, int):
            # Create simple diagram with n shapes
            self._create_diagram_with_count(datas)
            return

        # Create diagram from DataOfDiagram
        if not datas.is_empty():
            super().create_diagram(datas)
            is_root_item = datas.is_one_first_level_data()

            if not is_root_item:
                datas.increase_levels()

            if self._x_draw_page is not None and self._x_shapes is not None:
                self.set_draw_area()

                # Create base control shape
                x_base_shape = self.create_shape(
                    Diagram.DIAGRAM_BASE_SHAPE_TYPE,
                    0,
                    self.page_props.border_left,
                    self.page_props.border_top,
                )
                self._x_shapes.add(x_base_shape)
                self.set_control_shape_props(x_base_shape)
                self.set_color_mode_and_style_of_control_shape(x_base_shape)

                # Create start shape
                x_start_shape = self.create_shape(
                    Diagram.DIAGRAM_SHAPE_TYPE,
                    1,
                    self.page_props.border_left,
                    self.page_props.border_top,
                )
                self._x_shapes.add(x_start_shape)

                self.set_move_protect_of_shape(x_start_shape)
                self.set_color_prop(self._LO_ORANGES[2])
                self.set_shape_properties(x_start_shape, Diagram.DIAGRAM_SHAPE_TYPE)

                if x_start_shape is not None:
                    self.get_controller().set_selected_shape(x_start_shape)

                # The shapes are being drawn now, so their geometry must not be read
                # back out of the office while this runs
                self.init_diagram(read_geometry=False)

                # Initialize diagram tree
                if self._diagram_tree is None:
                    self._diagram_tree = OrgChartTree(self, x_base_shape, x_start_shape)
                dad_item = self._diagram_tree.get_root_item()
                new_tree_item = None
                last_tree_item = dad_item
                size = datas.size()
                i_root = 1 if is_root_item else 0
                i_color = 0

                # Create all shapes and tree items
                for i in range(i_root, size):
                    x_shape = self.create_shape(
                        Diagram.DIAGRAM_SHAPE_TYPE, i + (2 - i_root)
                    )
                    self._x_shapes.add(x_shape)
                    self.set_move_protect_of_shape(x_shape)

                    # Set color based on level
                    if i > i_root and datas.get(i).get_level() == 1:
                        i_color += 1
                    i_color %= 5

                    i_color_level = datas.get(i).get_level()
                    if i_color_level > 4:
                        i_color_level = 4

                    self.set_color_prop(self._LO_COLORS_2[i_color][i_color_level])
                    self.set_shape_properties(x_shape, Diagram.DIAGRAM_SHAPE_TYPE)
                    self._diagram_tree.add_to_rectangles(x_shape)

                    # Determine parent item based on level
                    if last_tree_item.get_level() == datas.get(i).get_level():
                        pass  # Same level
                    elif last_tree_item.get_level() < datas.get(i).get_level():
                        dad_item = last_tree_item  # Child of previous item
                    else:
                        # Go up levels to find parent
                        lev = dad_item.get_level() + 1 - datas.get(i).get_level()
                        for j in range(lev):
                            dad_item = dad_item.get_dad()

                    # Create connector shape
                    x_connector_shape = self.create_shape(
                        Diagram.CONNECTOR_SHAPE, i + (2 - i_root)
                    )
                    self._x_shapes.add(x_connector_shape)
                    self.set_move_protect_of_shape(x_connector_shape)

                    start_conn_pos, end_shape_conn_pos = self.connector_glue_positions(
                        dad_item, None, dad_item.get_level() + 1
                    )

                    self.set_connector_shape_props(
                        x_connector_shape,
                        dad_item.get_rectangle_shape(),
                        start_conn_pos,
                        x_shape,
                        end_shape_conn_pos,
                    )
                    self._diagram_tree.add_to_connectors(x_connector_shape)

                    # Create tree item and link to tree
                    new_tree_item = OrgChartTreeItem(
                        self._diagram_tree, x_shape, dad_item, 0, 0.0
                    )

                    if last_tree_item.get_level() == datas.get(i).get_level():
                        last_tree_item.set_first_sibling(new_tree_item)
                    elif last_tree_item.get_level() < datas.get(i).get_level():
                        if not dad_item.is_first_child():
                            dad_item.set_first_child(new_tree_item)
                    else:
                        dad_item.get_last_child().set_first_sibling(new_tree_item)

                    last_tree_item = new_tree_item

                    # The next turn of this loop reads the level of the item just added
                    # to decide whether the one after it is a child or a sibling
                    self._diagram_tree.recompute_levels_and_positions()

                # Handle root visibility
                if not is_root_item:
                    self.get_controller().set_selected_shape(
                        self._diagram_tree.get_root_item()
                        .get_last_child()
                        .get_rectangle_shape()
                    )
                    self.set_hidden_root_element_prop(True)
                    self.get_diagram_tree().get_root_item().hide_element()
                else:
                    i_color += 1
                    i_color %= 5
                    self.set_color_prop(self._LO_COLORS_2[i_color][1])
                    self.get_controller().set_selected_shape(
                        self._diagram_tree.get_root_item().get_rectangle_shape()
                    )

                self.refresh_diagram()

                # Writer states the position of a shape in a group relative to the
                # bounding box of that group, and the box only reaches its final extent
                # once every shape has been placed. The shapes placed early therefore sit
                # in a frame that no longer applies. Place them all a second time, now
                # that the box has settled, so that they agree with each other.
                self._diagram_tree.forget_geometry()
                self.refresh_diagram()

    def _create_diagram_with_count(self, n: int):
        """Create diagram with n simple shapes"""
        if self._x_draw_page is not None and self._x_shapes is not None and n > 0:
            self.set_draw_area()

            # Create base control shape
            x_base_shape = self.create_shape(
                Diagram.DIAGRAM_BASE_SHAPE_TYPE,
                0,
                self.page_props.border_left + self._half_diff,
                self.page_props.border_top,
                self._draw_area_width,
                self._draw_area_height,
            )
            self._x_shapes.add(x_base_shape)
            self.set_control_shape_props(x_base_shape)
            self.set_color_mode_and_style_of_control_shape(x_base_shape)

            # Use fixed dimensions - don't scale shapes to fit available space
            if n > 1:
                # Use fixed shape dimensions instead of scaling
                shape_width = self._shape_width * 1000  # Convert to appropriate units
                shape_height = self._shape_height * 1000  # Convert to appropriate units
                hor_space = self._hor_space * 1000  # Convert to appropriate units
                ver_space = self._ver_space * 1000  # Convert to appropriate units
            else:
                # For single shape, still use fixed dimensions
                shape_width = self._shape_width * 1000
                shape_height = self._shape_height * 1000
                hor_space = 0
                ver_space = 0

            # Create start shape (root)
            x_coord = (
                self.page_props.border_left
                + self._half_diff
                + self._draw_area_width // 2
                - shape_width // 2
            )
            y_coord = self.page_props.border_top

            x_start_shape = self.create_shape(
                Diagram.DIAGRAM_SHAPE_TYPE,
                1,
                x_coord,
                y_coord,
                shape_width,
                shape_height,
            )
            self._x_shapes.add(x_start_shape)
            self.set_move_protect_of_shape(x_start_shape)
            self.set_color_prop(self._ORG_CHART_COLORS[0])
            self.set_shape_properties(x_start_shape, Diagram.DIAGRAM_SHAPE_TYPE)

            # Create child shapes
            x_coord = self.page_props.border_left + self._half_diff
            y_coord = self.page_props.border_top + shape_height + ver_space
            x_selected_shape = None

            for i in range(2, n + 1):
                x_rect_shape = self.create_shape(
                    Diagram.DIAGRAM_SHAPE_TYPE,
                    i,
                    x_coord + (shape_width + hor_space) * (i - 2),
                    y_coord,
                    shape_width,
                    shape_height,
                )
                self._x_shapes.add(x_rect_shape)
                self.set_move_protect_of_shape(x_rect_shape)
                self.set_color_prop(self._ORG_CHART_COLORS[(i - 1) % 8])
                self.set_shape_properties(x_rect_shape, Diagram.DIAGRAM_SHAPE_TYPE)

                # Create connector
                x_connector_shape = self.create_shape(Diagram.CONNECTOR_SHAPE, i)
                self._x_shapes.add(x_connector_shape)
                self.set_move_protect_of_shape(x_connector_shape)
                self.set_connector_shape_props(
                    x_connector_shape, x_start_shape, 2, x_rect_shape, 0
                )

                if i == 2 and x_rect_shape is not None:
                    x_selected_shape = x_rect_shape

            # Set selected shape
            if n == 1 and x_start_shape is not None:
                self.get_controller().set_selected_shape(x_start_shape)
            elif x_selected_shape is not None:
                self.get_controller().set_selected_shape(x_selected_shape)
                shape_id = self.get_controller().get_shape_id(
                    self.get_shape_name(x_selected_shape)
                )
                self.set_color_prop(self._ORG_CHART_COLORS[(shape_id - 1) % 8])

    def init_diagram(self, diagram_id=None, read_geometry=True, group_shape=None):
        """Initialize diagram"""
        super().init_diagram(diagram_id, group_shape)

        if self._diagram_tree is None:
            self._diagram_tree = OrgChartTree(self)

        self._diagram_tree.set_lists(read_geometry)
        self._diagram_tree.set_tree()

    def update_anchor_glue_points(self, shape, geometry):
        """Keep the user defined glue points of a shape with a milsymbol drawing at the
        places that follow from its SymbolGeometry."""
        update_anchor_glue_points(shape, geometry)

    def update_stacked_glue_point(self, shape, add_if_missing=False):
        """Keep the stacked start glue point of a shape without symbol geometry at its
        place. Returns the index of the point the connectors to stacked children start on.
        """
        return update_stacked_glue_point(shape, add_if_missing)

    def connector_glue_positions(self, parent_item, child_item, child_level):
        """The glue points the connector from a parent to a child at this level runs
        between, as the pair of start and end glue point indices.

        The items are tree items; either may be None, and the child's level is passed
        separately because an item that is still being added does not know its level yet.

        On a shape with symbol geometry the connector leaves from the anchor point and
        arrives at the top of the frame, or at the left edge of the frame for a
        stacked child, and the glue points for that are put in place here. A shape
        without geometry keeps the builtin points: a child on the side by side levels
        hangs below its parent, so the connector runs from the parent's bottom centre to
        the child's top centre. A stacked child is entered on its left edge, and the
        connector starts on a point on the parent's bottom edge that lies left of every
        stacked child whatever the width of the parent's picture, so the connector
        routes downward and then right without doubling back.
        """
        stacked = child_level > OrgChartTree.LAST_HOR_LEVEL

        parent_shape = None
        parent_geometry = None
        if parent_item is not None:
            parent_shape = parent_item.get_rectangle_shape()
            parent_geometry = parent_item.get_symbol_geometry()
        if parent_geometry is not None and parent_shape is not None:
            update_anchor_glue_points(parent_shape, parent_geometry)
            start = ANCHOR_START_GLUE
        elif stacked and parent_shape is not None:
            start = update_stacked_glue_point(parent_shape, True)
        else:
            start = GLUE_BOTTOM

        child_geometry = None
        if child_item is not None:
            child_shape = child_item.get_rectangle_shape()
            child_geometry = child_item.get_symbol_geometry()
            if child_geometry is not None and child_shape is not None:
                update_anchor_glue_points(child_shape, child_geometry)
        if child_geometry is not None:
            end = ANCHOR_LEFT_GLUE if stacked else ANCHOR_TOP_GLUE
        else:
            end = GLUE_LEFT if stacked else GLUE_TOP

        return start, end

    def refresh_diagram(self):
        """Lay the diagram out again and join the connectors up at the points that follow
        from the pictures the shapes show now.

        A connector is made before its child gets a picture, so it starts out on the
        builtin glue points. Once the layout has placed the anchor glue points, the
        connectors are moved onto them here; a connector that already joins the right
        points is left alone.
        """
        with locked_controllers(self._x_model):
            super().refresh_diagram()
            if self._diagram_tree is not None:
                self._diagram_tree.refresh_connector_props()

    def paste_subtree(self, target_tree_item, clipboard_item, script=None):
        """Paste copied subtree as children of target item"""
        if self._diagram_tree is None:
            return False
        if target_tree_item is None or clipboard_item is None:
            return False

        self.update_origin()
        try:
            self._paste_script = script
            self._paste_item_recursive(target_tree_item, clipboard_item)
            return True
        except Exception as ex:
            print(f"Error pasting subtree: {ex}")
            return False
        finally:
            self._paste_script = None

    def _calculate_actual_level(self, tree_item):
        """Calculate actual tree level by traversing up to root via _dad chain"""
        level = 0
        current = tree_item
        while current is not None and current.get_dad() is not None:
            level += 1
            current = current.get_dad()
        return level

    def _paste_item_recursive(self, parent_tree_item, clipboard_item):
        """Recursively paste a ClipboardItem and its children"""
        top_shape_id = self.get_top_shape_id() + 1
        x_new_shape = self.create_shape(Diagram.DIAGRAM_SHAPE_TYPE, top_shape_id)
        self._x_shapes.add(x_new_shape)
        self._diagram_tree.add_to_rectangles(x_new_shape)

        if self._paste_script and "MilSymCode" in clipboard_item.attributes:
            # The drawing is made at the frame size recorded on the copied symbol, so it
            # agrees with the size attribute that is copied onto the new shape below
            svg_data = generate_icon_svg(
                self._paste_script,
                clipboard_item.attributes,
                get_recorded_symbol_size_px(clipboard_item.attributes, self._x_context),
            )
            if svg_data:
                self.set_new_shape_properties(
                    x_new_shape, Diagram.DIAGRAM_SHAPE_TYPE, svg_data
                )
                self._copy_attributes_to_shape(x_new_shape, clipboard_item.attributes)
            else:
                self.set_shape_properties(x_new_shape, Diagram.DIAGRAM_SHAPE_TYPE)
        else:
            svg_data = None
            self.set_shape_properties(x_new_shape, Diagram.DIAGRAM_SHAPE_TYPE)

        new_tree_item = OrgChartTreeItem(
            self._diagram_tree, x_new_shape, parent_tree_item, 0, 0.0
        )
        if svg_data:
            # The picture was set before this item existed, so the item is told about it
            # here and never reads it back from the office
            new_tree_item.note_symbol_svg(svg_data)

        if not parent_tree_item.is_first_child():
            parent_tree_item.set_first_child(new_tree_item)
        else:
            last_child = parent_tree_item.get_last_child()
            if last_child is not None:
                last_child.set_first_sibling(new_tree_item)

        self.set_move_protect_of_shape(x_new_shape)

        x_connector_shape = self.create_shape(Diagram.CONNECTOR_SHAPE, top_shape_id)
        self._x_shapes.add(x_connector_shape)
        self.set_move_protect_of_shape(x_connector_shape)
        self._diagram_tree.add_to_connectors(x_connector_shape)

        # Calculate actual level by traversing up the tree (parent's level + 1)
        new_item_level = self._calculate_actual_level(parent_tree_item) + 1
        start_conn_pos, end_shape_conn_pos = self.connector_glue_positions(
            parent_tree_item, new_tree_item, new_item_level
        )

        self.set_connector_shape_props(
            x_connector_shape,
            parent_tree_item.get_rectangle_shape(),
            start_conn_pos,
            x_new_shape,
            end_shape_conn_pos,
        )

        for child_clipboard in clipboard_item.children:
            self._paste_item_recursive(new_tree_item, child_clipboard)

        return new_tree_item

    def _copy_attributes_to_shape(self, shape, attributes):
        """Copy attributes to shape's UserDefinedAttributes"""
        from com.sun.star.xml import AttributeData

        attribute_hash = shape.UserDefinedAttributes
        user_attrs = AttributeData()
        for name, value in attributes.items():
            user_attrs.Type = "CDATA"
            user_attrs.Value = str(value)
            attribute_hash[name] = user_attrs
        shape.setPropertyValue("UserDefinedAttributes", attribute_hash)

    def add_shape(self, x_selected_shape=None):
        """Add new shape to diagram

        Args:
            x_selected_shape: Optional shape to use as parent. If None, uses current selection.
        """
        if self._diagram_tree is not None:
            self.update_origin()
            if x_selected_shape is None:
                x_selected_shape = self.get_controller().get_selected_shape()

            if x_selected_shape is not None:
                # Get shape name (in real implementation, would use UNO API)
                selected_shape_name = self.get_shape_name(x_selected_shape)

                role = Diagram.shape_role(selected_shape_name)
                if role == Diagram.DIAGRAM_SHAPE_TYPE:
                    selected_item = self._diagram_tree.get_tree_item(x_selected_shape)

                    # Can't be associate of root item
                    if (
                        selected_item.get_dad() is None
                        and self._new_item_h_type == self.ASSOCIATE
                    ):
                        title = self.get_gui().get_dialog_property_value(
                            "Strings", "ItemAddError.Title"
                        )
                        message = self.get_gui().get_dialog_property_value(
                            "Strings", "ItemAddError.Message"
                        )
                        self.get_gui().show_message_box(title, message)
                    else:
                        top_shape_id = self.get_top_shape_id()

                        if top_shape_id <= 0:
                            self.clear_empty_diagram_and_recreate()
                        else:
                            top_shape_id += 1
                            x_rectangle_shape = self.create_shape(
                                Diagram.DIAGRAM_SHAPE_TYPE, top_shape_id
                            )
                            self._x_shapes.add(x_rectangle_shape)
                            self._diagram_tree.add_to_rectangles(x_rectangle_shape)

                            new_tree_item = None
                            dad_item = None

                            if self._new_item_h_type == self.UNDERLING:
                                # Add as child
                                dad_item = selected_item
                                new_tree_item = OrgChartTreeItem(
                                    self._diagram_tree,
                                    x_rectangle_shape,
                                    dad_item,
                                    0,
                                    0.0,
                                )

                                if not dad_item.is_first_child():
                                    dad_item.set_first_child(new_tree_item)
                                else:
                                    x_previous_child = (
                                        self._diagram_tree.get_last_child_shape(
                                            x_selected_shape
                                        )
                                    )
                                    if x_previous_child is not None:
                                        previous_item = (
                                            self._diagram_tree.get_tree_item(
                                                x_previous_child
                                            )
                                        )
                                        if previous_item is not None:
                                            previous_item.set_first_sibling(
                                                new_tree_item
                                            )

                            elif self._new_item_h_type == self.ASSOCIATE:
                                # Add as sibling
                                dad_item = selected_item.get_dad()
                                new_tree_item = OrgChartTreeItem(
                                    self._diagram_tree,
                                    x_rectangle_shape,
                                    dad_item,
                                    0,
                                    0.0,
                                )

                                if selected_item.is_first_sibling():
                                    new_tree_item.set_first_sibling(
                                        selected_item.get_first_sibling()
                                    )
                                selected_item.set_first_sibling(new_tree_item)

                            # Set shape properties
                            self.set_move_protect_of_shape(x_rectangle_shape)
                            self.set_shape_properties(
                                x_rectangle_shape, Diagram.DIAGRAM_SHAPE_TYPE
                            )

                            # Create connector if not root level
                            if top_shape_id > 1:
                                x_connector_shape = self.create_shape(
                                    Diagram.CONNECTOR_SHAPE, top_shape_id
                                )
                                self._x_shapes.add(x_connector_shape)
                                self.set_move_protect_of_shape(x_connector_shape)
                                self._diagram_tree.add_to_connectors(x_connector_shape)

                                x_start_shape = None
                                start_item = None
                                new_item_level = 0
                                if self._new_item_h_type == self.UNDERLING:
                                    start_item = selected_item
                                    new_item_level = selected_item.get_level() + 1
                                elif self._new_item_h_type == self.ASSOCIATE:
                                    start_item = selected_item.get_dad()
                                    new_item_level = selected_item.get_level()
                                if start_item is not None:
                                    x_start_shape = start_item.get_rectangle_shape()

                                start_conn_pos, end_shape_conn_pos = (
                                    self.connector_glue_positions(
                                        start_item, new_tree_item, new_item_level
                                    )
                                )

                                self.set_connector_shape_props(
                                    x_connector_shape,
                                    x_start_shape,
                                    start_conn_pos,
                                    x_rectangle_shape,
                                    end_shape_conn_pos,
                                )

                                # Handle hidden root element
                                if self.is_hidden_root_element_prop():
                                    if (
                                        self.get_diagram_tree()
                                        .get_root_item()
                                        .get_rectangle_shape()
                                        == x_start_shape
                                    ):
                                        self.get_diagram_tree().get_root_item().hide_element()
