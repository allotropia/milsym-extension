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

from utils import generate_icon_svg, get_recorded_symbol_size_px
from perf import count
from ...diagram import Diagram
from ..organization_chart import OrganizationChart
from .orgchart_tree import OrgChartTree
from .orgchart_tree_item import OrgChartTreeItem

from com.sun.star.awt import Point
from com.sun.star.drawing import GluePoint2
from com.sun.star.drawing.EscapeDirection import DOWN as ESCAPE_DOWN

# Where on the bottom edge of a shape the connectors to its stacked children leave, as a
# fraction of the shape's width on the relative glue point scale of 0 to 10000. On very
# wide shapes the point moves further left than this, so that it stays left of the
# children.
STACKED_GLUE_FRACTION = 2500


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
                        dad_item.get_rectangle_shape(), dad_item.get_level() + 1
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

    def init_diagram(self, diagram_id=None, read_geometry=True):
        """Initialize diagram"""
        super().init_diagram(diagram_id)

        if self._diagram_tree is None:
            self._diagram_tree = OrgChartTree(self)

        self._diagram_tree.set_lists(read_geometry)
        self._diagram_tree.set_tree()

    def _stacked_glue_fraction(self, shape):
        """Where on the bottom edge of this shape the connectors to its stacked children
        leave, as a fraction of the shape's width on the relative glue point scale of 0
        to 10000.

        A quarter of the width looks right on a shape of ordinary width. The stacked
        children sit at least half a horizontal layout unit right of the shape's left
        edge, and on a very wide shape a quarter of the width would reach past them, so
        the offset is capped at half of that distance. The downward line then stays left
        of the children, with room to turn, however wide the shape's picture is.
        """
        fraction = STACKED_GLUE_FRACTION
        cap = OrgChartTreeItem.horizontal_pos_unit() // 4
        item = (
            self._diagram_tree.get_tree_item(shape) if self._diagram_tree else None
        )
        if item is not None and cap > 0:
            width = item._calculate_size_for_aspect_ratio()[0]
            if width > 0:
                fraction = min(fraction, cap * 10000 // width)
        return fraction

    def update_stacked_glue_point(self, shape, add_if_missing=False):
        """Keep the glue point the connectors to stacked children start on at its place
        on the bottom edge of the shape, where its place follows from the shape's width.

        A shape starts with the four builtin glue points, indices 0 to 3, one on the
        middle of each edge, so the first user defined point gets index 4. Returns that
        index. A shape that carries no such point yet is given one when add_if_missing
        says so; otherwise, and when the point cannot be added, the index of the builtin
        bottom center point is returned instead.
        """
        try:
            fraction = self._stacked_glue_fraction(shape)
            glue_points = shape.getGluePoints()

            if glue_points.getCount() > 4:
                glue = glue_points.getByIndex(4)
                if glue.Position.X != fraction:
                    count("shape: glue point moved")
                    glue.Position = Point(X=fraction, Y=10000)
                    glue_points.replaceByIndex(4, glue)
                return 4

            if not add_if_missing:
                return 2

            count("shape: glue point added")
            glue = GluePoint2()
            glue.IsRelative = True
            glue.Position = Point(X=fraction, Y=10000)
            glue.Escape = ESCAPE_DOWN
            glue.IsUserDefined = True
            glue_points.insertByIndex(glue_points.getCount(), glue)
            return 4
        except Exception as ex:
            print(f"Error setting the stacked connector glue point: {ex}")
            return 2

    def connector_glue_positions(self, parent_shape, child_level):
        """The glue points the connector to a child at this level runs between.

        Returns the pair of start and end glue point indices. A child on the side by
        side levels hangs below its parent, so the connector runs from the parent's
        bottom center to the child's top center. A stacked child is entered on its left
        edge, and the connector starts on a point on the parent's bottom edge that lies
        left of every stacked child whatever the width of the parent's picture, so the
        connector routes downward and then right without doubling back.
        """
        if child_level > OrgChartTree.LAST_HOR_LEVEL:
            if parent_shape is not None:
                return self.update_stacked_glue_point(parent_shape, True), 3
            return 2, 3
        return 2, 0

    def paste_subtree(self, target_tree_item, clipboard_item, script=None):
        """Paste copied subtree as children of target item"""
        if self._diagram_tree is None:
            return False
        if target_tree_item is None or clipboard_item is None:
            return False

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
            self.set_shape_properties(x_new_shape, Diagram.DIAGRAM_SHAPE_TYPE)

        new_tree_item = OrgChartTreeItem(
            self._diagram_tree, x_new_shape, parent_tree_item, 0, 0.0
        )

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
            parent_tree_item.get_rectangle_shape(), new_item_level
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
            if x_selected_shape is None:
                x_selected_shape = self.get_controller().get_selected_shape()

            if x_selected_shape is not None:
                # Get shape name (in real implementation, would use UNO API)
                selected_shape_name = self.get_shape_name(x_selected_shape)

                if (
                    Diagram.DIAGRAM_SHAPE_TYPE in selected_shape_name
                    and Diagram.DIAGRAM_BASE_SHAPE_TYPE not in selected_shape_name
                ):
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
                                new_item_level = 0

                                if self._new_item_h_type == self.UNDERLING:
                                    x_start_shape = selected_item.get_rectangle_shape()
                                    new_item_level = selected_item.get_level() + 1
                                elif self._new_item_h_type == self.ASSOCIATE:
                                    x_start_shape = (
                                        selected_item.get_dad().get_rectangle_shape()
                                    )
                                    new_item_level = selected_item.get_level()

                                start_conn_pos, end_shape_conn_pos = (
                                    self.connector_glue_positions(
                                        x_start_shape, new_item_level
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
