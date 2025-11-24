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
from ...diagram import Diagram
from ..organization_chart import OrganizationChart
from .orgchart_tree import OrgChartTree
from .orgchart_tree_item import OrgChartTreeItem


class OrgChart(OrganizationChart):
    """Organization chart implementation"""

    def __init__(self, controller, gui, x_frame, x_context):
        super().__init__(controller, gui, x_frame, x_context)
        self._diagram_tree = None

        # Set specific dimensions for org chart
        self._group_width = 10
        self._group_height = 6
        self._shape_width = 4
        self._hor_space = 1
        self._shape_height = 4
        self._ver_space = 3

    def init_diagram_tree(self, diagram_tree):
        """Initialize diagram tree"""
        #breakpoint()
        super().init_diagram()
        self._diagram_tree = OrgChartTree(self, diagram_tree)

    def get_diagram_tree(self):
        """Get diagram tree"""
        return self._diagram_tree

    def get_diagram_type_name(self) -> str:
        """Get diagram type name"""
        return "OrganizationDiagram"

    def create_diagram(self, datas):
        """Create diagram from data"""
        if isinstance(datas, int):
            # Create simple diagram with n shapes
            self._create_diagram_with_count(datas)
            return

        # Create diagram from DataOfDiagram
        last_hor_level = OrgChartTree.LAST_HOR_LEVEL

        if not datas.is_empty():
            super().create_diagram(datas)
            is_root_item = datas.is_one_first_level_data()

            if not is_root_item:
                datas.increase_levels()

            if self._x_draw_page is not None and self._x_shapes is not None:
                self.set_draw_area()

                # Create base control shape
                x_base_shape = self.create_shape(
                    Diagram.DIAGRAM_BASE_SHAPE_TYPE, 0,
                    self.page_props.border_left, self.page_props.border_top,
                    self._draw_area_width, self._draw_area_height
                )
                self._x_shapes.add(x_base_shape)
                self.set_control_shape_props(x_base_shape)
                self.set_color_mode_and_style_of_control_shape(x_base_shape)

                OrgChartTree.LAST_HOR_LEVEL = 10000
                self.set_hor_level_of_control_shape(x_base_shape, 10000)

                # Create start shape
                x_start_shape = self.create_shape(
                    Diagram.DIAGRAM_SHAPE_TYPE, 1,
                    self.page_props.border_left, self.page_props.border_top,
                    self._draw_area_width, self._draw_area_height
                )
                self._x_shapes.add(x_start_shape)

                self.set_move_protect_of_shape(x_start_shape)
                self.set_color_prop(self._LO_ORANGES[2])
                self.set_shape_properties(x_start_shape, Diagram.DIAGRAM_SHAPE_TYPE)

                if x_start_shape is not None:
                    self.get_controller().set_selected_shape(x_start_shape)

                self.init_diagram()

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
                    x_shape = self.create_shape(Diagram.DIAGRAM_SHAPE_TYPE, i + (2 - i_root))
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
                    x_connector_shape = self.create_shape(Diagram.CONNECTOR_SHAPE, i + (2 - i_root))
                    self._x_shapes.add(x_connector_shape)
                    self.set_move_protect_of_shape(x_connector_shape)

                    end_shape_conn_pos = 0
                    if dad_item.get_level() + 1 > last_hor_level:
                        end_shape_conn_pos = 3

                    self.set_connector_shape_props(
                        x_connector_shape,
                        dad_item.get_rectangle_shape(), 2,
                        x_shape, end_shape_conn_pos
                    )
                    self._diagram_tree.add_to_connectors(x_connector_shape)

                    # Create tree item and link to tree
                    new_tree_item = OrgChartTreeItem(self._diagram_tree, x_shape, dad_item, 0, 0.0)

                    if last_tree_item.get_level() == datas.get(i).get_level():
                        last_tree_item.set_first_sibling(new_tree_item)
                    elif last_tree_item.get_level() < datas.get(i).get_level():
                        if not dad_item.is_first_child():
                            dad_item.set_first_child(new_tree_item)
                    else:
                        dad_item.get_last_child().set_first_sibling(new_tree_item)

                    last_tree_item = new_tree_item
                    self.refresh_diagram()

                # Handle root visibility
                if not is_root_item:
                    self.get_controller().set_selected_shape(
                        self._diagram_tree.get_root_item().get_last_child().get_rectangle_shape()
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

                OrgChartTree.LAST_HOR_LEVEL = last_hor_level
                self.set_hor_level_of_control_shape(x_base_shape, last_hor_level)
                self.refresh_diagram()

    def _create_diagram_with_count(self, n: int):
        """Create diagram with n simple shapes"""
        if self._x_draw_page is not None and self._x_shapes is not None and n > 0:
            self.set_draw_area()

            # Create base control shape
            x_base_shape = self.create_shape(
                Diagram.DIAGRAM_BASE_SHAPE_TYPE, 0,
                self.page_props.border_left + self._half_diff, self.page_props.border_top,
                self._draw_area_width, self._draw_area_height
            )
            self._x_shapes.add(x_base_shape)
            self.set_control_shape_props(x_base_shape)
            self.set_color_mode_and_style_of_control_shape(x_base_shape)

            OrgChartTree.LAST_HOR_LEVEL = 2
            self.set_hor_level_of_control_shape(x_base_shape, 2)

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
            x_coord = (self.page_props.border_left + self._half_diff +
                      self._draw_area_width // 2 - shape_width // 2)
            y_coord = self.page_props.border_top

            x_start_shape = self.create_shape(
                Diagram.DIAGRAM_SHAPE_TYPE, 1, x_coord, y_coord, shape_width, shape_height
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
                    Diagram.DIAGRAM_SHAPE_TYPE, i,
                    x_coord + (shape_width + hor_space) * (i - 2), y_coord,
                    shape_width, shape_height
                )
                self._x_shapes.add(x_rect_shape)
                self.set_move_protect_of_shape(x_rect_shape)
                self.set_color_prop(self._ORG_CHART_COLORS[(i - 1) % 8])
                self.set_shape_properties(x_rect_shape, Diagram.DIAGRAM_SHAPE_TYPE)

                # Create connector
                x_connector_shape = self.create_shape(Diagram.CONNECTOR_SHAPE, i)
                self._x_shapes.add(x_connector_shape)
                self.set_move_protect_of_shape(x_connector_shape)
                self.set_connector_shape_props(x_connector_shape, x_start_shape, 2, x_rect_shape, 0)

                if i == 2 and x_rect_shape is not None:
                    x_selected_shape = x_rect_shape

            # Set selected shape
            if n == 1 and x_start_shape is not None:
                self.get_controller().set_selected_shape(x_start_shape)
            elif x_selected_shape is not None:
                self.get_controller().set_selected_shape(x_selected_shape)
                shape_id = self.get_controller().get_shape_id(self.get_shape_name(x_selected_shape))
                self.set_color_prop(self._ORG_CHART_COLORS[(shape_id - 1) % 8])

    def init_diagram(self):
        """Initialize diagram"""
        super().init_diagram()

        if self._diagram_tree is None:
            self._diagram_tree = OrgChartTree(self)

        self._diagram_tree.set_lists()

        if OrgChartTree.LAST_HOR_LEVEL == -1:
            OrgChartTree.LAST_HOR_LEVEL = self.get_hor_level_of_control_shape(
                self._diagram_tree.get_control_shape()
            )
        else:
            self.set_hor_level_of_control_shape(
                self._diagram_tree.get_control_shape(),
                OrgChartTree.LAST_HOR_LEVEL
            )

        self._diagram_tree.set_tree()

    def add_shape(self):
        """Add new shape to diagram"""
        if self._diagram_tree is not None:
            x_selected_shape = self.get_controller().get_selected_shape()

            if x_selected_shape is not None:
                # Get shape name (in real implementation, would use UNO API)
                selected_shape_name = self.get_shape_name(x_selected_shape)

                if (Diagram.DIAGRAM_SHAPE_TYPE in selected_shape_name and
                    Diagram.DIAGRAM_BASE_SHAPE_TYPE not in selected_shape_name):

                    selected_item = self._diagram_tree.get_tree_item(x_selected_shape)

                    # Can't be associate of root item
                    if selected_item.get_dad() is None and self._new_item_h_type == self.ASSOCIATE:
                        title = self.get_gui().get_dialog_property_value("Strings", "ItemAddError.Title")
                        message = self.get_gui().get_dialog_property_value("Strings", "ItemAddError.Message")
                        self.get_gui().show_message_box(title, message)
                    else:
                        top_shape_id = self.get_top_shape_id()

                        if top_shape_id <= 0:
                            self.clear_empty_diagram_and_recreate()
                        else:
                            top_shape_id += 1
                            x_rectangle_shape = self.create_shape(Diagram.DIAGRAM_SHAPE_TYPE, top_shape_id)
                            self._x_shapes.add(x_rectangle_shape)
                            self._diagram_tree.add_to_rectangles(x_rectangle_shape)

                            new_tree_item = None
                            dad_item = None

                            if self._new_item_h_type == self.UNDERLING:
                                # Add as child
                                dad_item = selected_item
                                new_tree_item = OrgChartTreeItem(
                                    self._diagram_tree, x_rectangle_shape, dad_item, 0, 0.0
                                )

                                if not dad_item.is_first_child():
                                    dad_item.set_first_child(new_tree_item)
                                else:
                                    x_previous_child = self._diagram_tree.get_last_child_shape(x_selected_shape)
                                    if x_previous_child is not None:
                                        previous_item = self._diagram_tree.get_tree_item(x_previous_child)
                                        if previous_item is not None:
                                            previous_item.set_first_sibling(new_tree_item)

                            elif self._new_item_h_type == self.ASSOCIATE:
                                # Add as sibling
                                dad_item = selected_item.get_dad()
                                new_tree_item = OrgChartTreeItem(
                                    self._diagram_tree, x_rectangle_shape, dad_item, 0, 0.0
                                )

                                if selected_item.is_first_sibling():
                                    new_tree_item.set_first_sibling(selected_item.get_first_sibling())
                                selected_item.set_first_sibling(new_tree_item)

                            # Set shape properties
                            self.set_move_protect_of_shape(x_rectangle_shape)
                            self.set_shape_properties(x_rectangle_shape, Diagram.DIAGRAM_SHAPE_TYPE)

                            # Create connector if not root level
                            if top_shape_id > 1:
                                x_connector_shape = self.create_shape(Diagram.CONNECTOR_SHAPE, top_shape_id)
                                self._x_shapes.add(x_connector_shape)
                                self.set_move_protect_of_shape(x_connector_shape)
                                self._diagram_tree.add_to_connectors(x_connector_shape)

                                x_start_shape = None
                                end_shape_conn_pos = 0

                                if self._new_item_h_type == self.UNDERLING:
                                    x_start_shape = selected_item.get_rectangle_shape()
                                    if selected_item.get_level() + 1 > OrgChartTree.LAST_HOR_LEVEL:
                                        end_shape_conn_pos = 3
                                elif self._new_item_h_type == self.ASSOCIATE:
                                    x_start_shape = selected_item.get_dad().get_rectangle_shape()
                                    if selected_item.get_level() > OrgChartTree.LAST_HOR_LEVEL:
                                        end_shape_conn_pos = 3

                                self.set_connector_shape_props(
                                    x_connector_shape, x_start_shape, 2,
                                    x_rectangle_shape, end_shape_conn_pos
                                )

                                # Handle hidden root element
                                if self.is_hidden_root_element_prop():
                                    if (self.get_diagram_tree().get_root_item().get_rectangle_shape() ==
                                        x_start_shape):
                                        self.get_diagram_tree().get_root_item().hide_element()
