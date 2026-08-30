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

from utils import (
    get_default_symbol_height_cm,
    parse_svg_dimensions,
    parse_svg_symbol_geometry,
    read_shape_svg,
)
from perf import count

from ..organization_chart_tree_item import OrganizationChartTreeItem

from com.sun.star.awt import Point, Size


def stacked_column_offsets(channel_x, gap, unit, stacked):
    """Work out where the shapes stacked below a column head sit horizontally.

    channel_x is the x distance from the column start to the connector channel that runs
    down from the head's anchor. gap is the empty x distance kept between the channel and
    the nearest left edge of a stacked shape. unit is the x distance one layout position
    stands for. stacked is a list of (pos_delta, width, overhang) tuples, one per stacked
    shape: its layout position relative to the head, its width, and the distance from its
    left edge to the left edge of its octagon.

    The octagons of the column line up: every shape is placed so that its octagon's left
    edge sits on one line, shifted right by its own pos_delta for nesting. That line is
    far enough from the channel for the widest overhang in the column, so the left labels
    of every symbol stay clear of the channel.

    Returns the list of x offsets from the column start, in the order of stacked, and the
    right edge of the widest shape among them (0 when stacked is empty).
    """
    widest_overhang = max((overhang for _, _, overhang in stacked), default=0)
    octagon_left = channel_x + gap + widest_overhang
    x_offsets = []
    right_edge = 0
    for pos_delta, width, overhang in stacked:
        x_offset = octagon_left + pos_delta * unit - overhang
        x_offsets.append(x_offset)
        if x_offset + width > right_edge:
            right_edge = x_offset + width
    return x_offsets, right_edge


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

    # Height of the symbol frame in 1/100 mm, as configured. None before the first
    # layout pass has read it.
    _configured_symbol_height = None

    # The gap between one column and the next, and the indent of the shapes stacked
    # below a column head, are widened by this factor to give the columns more air.
    HORIZONTAL_STEP_FACTOR = 1.5

    def __init__(
        self, diagram_tree, dad_or_shape=None, item_or_dad=None, level=None, pos=None
    ):
        """
        Multiple constructor patterns:
        1. OrgChartTreeItem(diagram_tree, dad, item) - copy constructor
        2. OrgChartTreeItem(diagram_tree, shape, dad, level, pos) - new item constructor
        """
        if level is not None and pos is not None:
            # Constructor with shape, dad, level, pos
            super().__init__(diagram_tree, item_or_dad, None)
            self._x_rectangle_shape = dad_or_shape
            self._rectangle_name = (
                self._diagram_tree.get_org_chart().get_shape_name(dad_or_shape)
                if dad_or_shape
                else ""
            )
            self.set_level(level)
            self.set_pos(pos)
        else:
            # Copy constructor
            super().__init__(diagram_tree, dad_or_shape, item_or_dad)

        diagram_tree.register_item(self)

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
            self._first_child = OrgChartTreeItem(
                self.get_diagram_tree(), self, tree_item.get_first_child()
            )
            self._first_child.convert_tree_items(tree_item.get_first_child())

        if tree_item.is_first_sibling():
            self._first_sibling = OrgChartTreeItem(
                self.get_diagram_tree(), self.get_dad(), tree_item.get_first_sibling()
            )
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
        last_hor_level = self._diagram_tree.LAST_HOR_LEVEL

        x_first_child_shape = self.get_diagram_tree().get_first_child_shape(
            self._x_rectangle_shape
        )
        if x_first_child_shape is not None:
            first_child_level = self._level + 1
            first_child_pos = 0.0
            if first_child_level <= last_hor_level:
                first_child_pos = (
                    OrgChartTreeItem._max_positions[first_child_level] + 1.0
                )
            else:
                first_child_pos = self._pos + 0.5
            self._first_child = OrgChartTreeItem(
                self.get_diagram_tree(),
                x_first_child_shape,
                self,
                first_child_level,
                first_child_pos,
            )
            self._first_child.init_tree_items()

        if self._level == last_hor_level:
            deep = self.get_number_of_items_in_branch(self)
            if deep > 2:
                max_pos_in_level = OrgChartTreeItem._max_branch_positions[
                    self._level + deep - 1
                ]
                if self._pos < max_pos_in_level + 0.5:
                    if self.is_first_child():
                        self.get_first_child().increase_pos_in_branch(
                            max_pos_in_level + 0.5 - self._pos
                        )
                        self.set_pos(max_pos_in_level + 0.5)
            self.set_max_pos_of_branch()

        x_first_sibling_shape = self.get_diagram_tree().get_first_sibling_shape(
            self._x_rectangle_shape, self._dad
        )
        if x_first_sibling_shape is not None:
            first_sibling_level = self._level
            first_sibling_pos = self._pos + 1.0

            if first_sibling_level > last_hor_level:
                first_sibling_pos = self._pos
                first_sibling_level = self._level + self.get_number_of_items_in_branch(
                    self
                )

            self._first_sibling = OrgChartTreeItem(
                self.get_diagram_tree(),
                x_first_sibling_shape,
                self._dad,
                first_sibling_level,
                first_sibling_pos,
            )
            self._first_sibling.init_tree_items()

        if (
            self._level <= last_hor_level
            and self._dad is not None
            and self._dad.get_first_child() == self
        ):
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
        last_hor_level = self._diagram_tree.LAST_HOR_LEVEL

        if self._first_child is not None:
            first_child_level = self._level + 1
            first_child_pos = 0.0

            if first_child_level <= last_hor_level:
                first_child_pos = (
                    OrgChartTreeItem._max_positions[first_child_level] + 1.0
                )

            if first_child_level > last_hor_level:
                first_child_pos = self._pos + 0.5

            self._first_child.set_level(first_child_level)
            self._first_child.set_pos(first_child_pos)
            self._first_child.set_positions_of_items()

        # Handle branch positioning
        if self._level == last_hor_level:
            deep = self.get_number_of_items_in_branch(self)
            if deep > 2:
                max_pos_in_level = OrgChartTreeItem._max_branch_positions[
                    self._level + deep - 1
                ]
                if self._pos < max_pos_in_level + 0.5:
                    if self.is_first_child():
                        self.get_first_child().increase_pos_in_branch(
                            max_pos_in_level + 0.5 - self._pos
                        )
                        self.set_pos(max_pos_in_level + 0.5)
            self.set_max_pos_of_branch()

        if self._first_sibling is not None:
            first_sibling_level = self._level
            first_sibling_pos = self._pos + 1.0

            if first_sibling_level > last_hor_level:
                first_sibling_pos = self._pos
                first_sibling_level = self._level + self.get_number_of_items_in_branch(
                    self
                )

            self._first_sibling.set_level(first_sibling_level)
            self._first_sibling.set_pos(first_sibling_pos)
            self._first_sibling.set_positions_of_items()

        # Position adjustment
        if (
            self._level <= last_hor_level
            and self._dad is not None
            and self._dad.get_first_child() == self
        ):
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
        last_hor_level = self._diagram_tree.LAST_HOR_LEVEL

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
        if self.get_diagram_tree().get_org_chart().is_hidden_root_element_prop():
            pass

        # Use fixed dimensions instead of scaling to fit available space
        org_chart = self.get_diagram_tree().get_org_chart()
        configured_height = get_default_symbol_height_cm(org_chart._x_context)
        OrgChartTreeItem._configured_symbol_height = configured_height

        # Set fixed shape dimensions (convert to appropriate units)
        OrgChartTreeItem._shape_width = org_chart.get_shape_width() * 1000
        OrgChartTreeItem._shape_height = configured_height
        OrgChartTreeItem._hor_space = org_chart.get_hor_space() * 1000
        OrgChartTreeItem._ver_space = org_chart.get_ver_space() * 1000

        control_shape_pos = self.get_diagram_tree().get_control_shape_pos()
        OrgChartTreeItem._group_pos_x = control_shape_pos.X if control_shape_pos else 0
        OrgChartTreeItem._group_pos_y = control_shape_pos.Y if control_shape_pos else 0

    @staticmethod
    def horizontal_pos_unit():
        """The x distance that one unit of layout position stands for."""
        return (
            OrgChartTreeItem._shape_width + OrgChartTreeItem._hor_space
        ) * OrgChartTreeItem.HORIZONTAL_STEP_FACTOR

    @staticmethod
    def column_gap():
        """The empty x distance between the edge of one column and the start of the next."""
        return OrgChartTreeItem._hor_space * OrgChartTreeItem.HORIZONTAL_STEP_FACTOR

    @staticmethod
    def channel_gap():
        """The empty x distance between the connector channel that runs down from a
        column head's anchor and the nearest left edge of a shape stacked below it. A
        quarter of the configured symbol height, the same for every column.
        """
        if OrgChartTreeItem._configured_symbol_height is not None:
            return OrgChartTreeItem._configured_symbol_height // 4
        return OrgChartTreeItem._shape_height // 10

    def left_overhang(self):
        """The x distance from the left edge of this item's shape to the left edge of its
        frame octagon, in 1/100 mm. Zero for a picture without symbol geometry, whose
        whole extent then counts as the octagon.
        """
        geometry = self.get_symbol_geometry()
        if geometry is None:
            return 0
        width = self._calculate_size_for_aspect_ratio()[0]
        return int(geometry.octagon_left() * width)

    def anchor_x_offset(self):
        """The x distance from the left edge of this item's shape to its symbol anchor, in
        1/100 mm. The middle of the shape for a picture without symbol geometry.
        """
        geometry = self.get_symbol_geometry()
        width = self._calculate_size_for_aspect_ratio()[0]
        if geometry is None:
            return width // 2
        return int(geometry.anchor_x() * width)

    def measure_column(self):
        """Measure the column this item heads. Returns the width the column needs, the x
        offset of every shape stacked below the head and the y offset of each of them.

        The x offsets are measured from the column start. The connector channel runs down
        from the head's anchor, and the stacked shapes sit to the right of it with their
        frame octagons on one line, a shape further right the deeper it nests. The line
        keeps the widest left overhang in the column clear of the channel, so the labels
        left of a symbol never cross the connectors. A head without symbol geometry has
        its channel on its left edge, so that a column of pictures without geometry keeps
        the width it had when the shapes lined up on their left edges. The width reaches
        to the right edge of the widest shape, head or stacked.

        The y offsets are measured from the top of the head. The stacked shapes follow
        each other down the column, each starting below the one before it, so a taller
        symbol moves every shape below it further down.
        """
        unit = OrgChartTreeItem.horizontal_pos_unit()
        quarter_space = OrgChartTreeItem._ver_space // 4
        base_pos = self.get_pos()
        width, head_height = self._calculate_size_for_aspect_ratio()
        if self.get_symbol_geometry() is not None:
            channel_x = self.anchor_x_offset()
        else:
            channel_x = 0

        stacked = []
        pending = [self._first_child]
        while pending:
            item = pending.pop()
            if item is None:
                continue
            stacked.append(item)
            pending.append(item.get_first_child())
            pending.append(item.get_first_sibling())

        # The level of a stacked item counts the rows above it in its column, so sorting
        # by level walks the column from top to bottom
        stacked.sort(key=lambda item: item.get_level())

        measures = [
            (item.get_pos() - base_pos, item._calculate_size_for_aspect_ratio()[0],
             item.left_overhang())
            for item in stacked
        ]
        x_offsets, right_edge = stacked_column_offsets(
            channel_x, OrgChartTreeItem.channel_gap(), unit, measures
        )
        if right_edge > width:
            width = right_edge

        x_offset_by_item = {}
        y_offset_by_item = {}
        y_offset = head_height + quarter_space
        for item, x_offset in zip(stacked, x_offsets):
            x_offset_by_item[item] = int(x_offset)
            y_offset_by_item[item] = y_offset
            y_offset += item._calculate_size_for_aspect_ratio()[1] + quarter_space

        return width, x_offset_by_item, y_offset_by_item

    def set_pos_of_rect(self):
        """Set position of rectangle"""
        last_hor_level = self._diagram_tree.LAST_HOR_LEVEL
        x_coord = OrgChartTreeItem._group_pos_x + int(
            self._diagram_tree.horizontal_offset_of(self)
        )
        y_coord = OrgChartTreeItem._group_pos_y + int(
            self._diagram_tree.vertical_offset_of(self)
        )

        if self.get_diagram_tree().get_org_chart().is_hidden_root_element_prop():
            if self == self.get_diagram_tree().get_root_item():
                y_coord = OrgChartTreeItem._group_pos_y - 10
            else:
                # With the root hidden, the whole diagram moves up by the row the root
                # would have taken
                y_coord -= int(self._diagram_tree.hidden_root_row_shift())

        # Calculate size based on graphic aspect ratio while fitting within default bounds
        calculated_width, calculated_height = self._calculate_size_for_aspect_ratio()

        # A stacked shape's place to the right of the connector channel is part of the
        # horizontal offset already, so every shape is put where the offsets say
        self.set_position_if_changed(Point(X=x_coord, Y=y_coord))

        self.set_size_if_changed(Size(Width=calculated_width, Height=calculated_height))

        org_chart = self.get_diagram_tree().get_org_chart()
        geometry = self.get_symbol_geometry()
        if geometry is not None:
            # The connectors of this item leave from and arrive at points that follow from
            # where the frame octagon and the anchor sit on its picture
            org_chart.update_anchor_glue_points(self._x_rectangle_shape, geometry)
        elif self._level >= last_hor_level and self.is_first_child():
            # The children of this item are stacked, and the glue point their connectors
            # start on sits at a place on the bottom edge that follows from the width
            # just set
            org_chart.update_stacked_glue_point(self._x_rectangle_shape)

    def note_symbol_svg(self, svg_data):
        """Keep what the layout needs to know about the SVG this item's shape shows.

        The string carries the symbol geometry, the proportions of the picture and the
        height of the frame within it, so an item told about its SVG here is never asked
        to read the picture back from the office. Called with the drawing the shape was
        just given, or with the drawing read back from the office; either way the string
        is the one the picture was made from.
        """
        self._symbol_geometry_known = True
        geometry = parse_svg_symbol_geometry(svg_data)
        self._symbol_geometry = geometry
        if geometry is not None:
            size = parse_svg_dimensions(svg_data)
            if size.Width > 0 and size.Height > 0:
                self._graphic_aspect_ratio = size.Width / size.Height
            octagon_height = geometry.octagon[3]
            if octagon_height > 0:
                self._graphic_frame_factor = 1.0 / octagon_height

    def get_symbol_geometry(self):
        """Where the frame octagon and the anchor sit on the picture of this item, as a
        SymbolGeometry of fractions of the picture, or None when the picture carries no
        such information (a drawing made before it was recorded, or a plain picture).

        An item whose drawing was set through the extension in this session already knows
        its geometry. Only a picture the extension has not seen as a string yet, such as
        one from a loaded document, is read back from the office, once, and kept.
        """
        if self._symbol_geometry_known:
            return self._symbol_geometry

        context = self.get_diagram_tree().get_org_chart()._x_context
        self.note_symbol_svg(read_shape_svg(context, self._x_rectangle_shape))
        return self._symbol_geometry

    def get_graphic_aspect_ratio(self):
        """Width divided by height of the picture on this item, or None when it has none.

        The ratio is read from the office once and kept, because reading it back asks the
        office for the picture's pixel size, which for a drawing has to be worked out
        rather than looked up. A picture that carries the symbol geometry gives its
        proportions together with that geometry.
        """
        if self._graphic_aspect_ratio is not None:
            return self._graphic_aspect_ratio
        if self.get_symbol_geometry() is not None and self._graphic_aspect_ratio is not None:
            return self._graphic_aspect_ratio

        count("shape: read graphic aspect ratio")
        try:
            graphic = self._x_rectangle_shape.Graphic
            if graphic:
                graphic_size = graphic.SizePixel
                if graphic_size.Height > 0 and graphic_size.Width > 0:
                    self._graphic_aspect_ratio = (
                        graphic_size.Width / graphic_size.Height
                    )
        except Exception as ex:
            print(f"Could not get graphic aspect ratio: {ex}")

        return self._graphic_aspect_ratio

    def get_graphic_frame_factor(self):
        """Height of the whole picture divided by the height of its frame octagon, or
        None when the shape does not carry a milsymbol drawing.

        The milsymbol size argument recorded on the shape is the height of the frame
        octagon in pixels, and the pixel height of the picture also counts decorations
        such as echelon markers or text labels, so the quotient says how much taller the
        decorations make the symbol. A bare symbol gives 1.0.

        The factor is read from the office once and kept, like the aspect ratio. A picture
        that carries the symbol geometry gives the factor together with that geometry.
        """
        if self._graphic_frame_factor is not None:
            return self._graphic_frame_factor
        if self.get_symbol_geometry() is not None and self._graphic_frame_factor is not None:
            return self._graphic_frame_factor

        count("shape: read graphic frame factor")
        try:
            attribute_hash = self._x_rectangle_shape.UserDefinedAttributes
            if attribute_hash is None or not attribute_hash.hasByName("MilSymSize"):
                return None
            frame_height_px = float(attribute_hash.getByName("MilSymSize").Value)
            if frame_height_px <= 0:
                return None
            graphic = self._x_rectangle_shape.Graphic
            if graphic:
                graphic_size = graphic.SizePixel
                if graphic_size.Height > 0:
                    self._graphic_frame_factor = graphic_size.Height / frame_height_px
        except Exception as ex:
            print(f"Could not get graphic frame factor: {ex}")

        return self._graphic_frame_factor

    def _calculate_size_for_aspect_ratio(self):
        """Calculate size that puts the frame octagon at the configured height"""
        default_width = OrgChartTreeItem._shape_width
        fixed_height = OrgChartTreeItem._configured_symbol_height
        if fixed_height is None:
            context = self.get_diagram_tree().get_org_chart()._x_context
            fixed_height = get_default_symbol_height_cm(context)

        aspect_ratio = self.get_graphic_aspect_ratio()
        if aspect_ratio is None:
            # Without a picture to measure, the shape keeps the width every item gets
            return default_width, fixed_height

        frame_factor = self.get_graphic_frame_factor()
        if frame_factor is not None:
            # The frame octagon keeps the configured height, and decorations such as
            # echelon markers or text labels make the whole symbol larger than that.
            calculated_height = int(fixed_height * frame_factor)
            calculated_width = int(calculated_height * aspect_ratio)
            return calculated_width, calculated_height

        # Without a recorded frame size, the whole picture gets the configured height,
        # and a wide picture shrinks to the width every item gets.
        calculated_width = int(fixed_height * aspect_ratio)
        if calculated_width > OrgChartTreeItem._shape_width:
            scale_factor = OrgChartTreeItem._shape_width / calculated_width
            calculated_width = OrgChartTreeItem._shape_width
            calculated_height = int(fixed_height * scale_factor)
        else:
            calculated_height = fixed_height

        return calculated_width, calculated_height
