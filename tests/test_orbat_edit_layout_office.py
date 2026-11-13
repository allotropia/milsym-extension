# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

"""An orbat edit leaves every member where a fresh layout of the diagram puts it.

In Writer the orbat group is anchored, and the office places the group frame again when
a shape inside it grows or arrives, so an edit can move the whole group on the page
while it runs. The diagram tree keeps a record of where it last read or wrote each
member and skips writes that match the record, so after such a move a single layout
pass can leave some members measured from the old group position and others from the
new one, one move-distance apart, with the connectors routed around the misplaced ones.

The test adds a child through the control dialog handler, the same path the dialog's
toolbar uses, and then compares every member's position, taken relative to the group's
control shape, against the positions after the manual refresh action, which lays the
diagram out from a fresh read of the settled document. A consistent edit leaves nothing
for the manual refresh to move.
"""

from conftest import run_as_office_event


def _shape_position(shape):
    """A shape's drawing-layer position on the page, in 1/100 mm.

    The Transformation property is the office's own coordinate, the same for every
    shape on the page, where getPosition on a Writer group member is measured from the
    group box and moves with it.
    """
    matrix = shape.getPropertyValue("Transformation")
    return matrix.Line1.Column3, matrix.Line2.Column3


def _member_positions(group):
    """Positions of the symbol members of an orbat group, relative to its control shape.

    The group holds one RectangleShape, the invisible control shape the layout is
    measured from, GraphicObjectShape members, and the connectors between them.
    Positions relative to the control shape compare the inner layout of two states of
    the diagram even when the whole group sits elsewhere on the page.
    """
    control = None
    members = {}
    for index in range(group.getCount()):
        shape = group.getByIndex(index)
        if shape.supportsService("com.sun.star.drawing.RectangleShape"):
            control = _shape_position(shape)
        elif shape.supportsService("com.sun.star.drawing.GraphicObjectShape"):
            members[shape.getName()] = _shape_position(shape)
    assert control is not None, "the group carries no control shape"
    return {
        name: (x - control[0], y - control[1]) for name, (x, y) in members.items()
    }


def _settled_member_positions(group):
    """Member positions once two consecutive reads agree.

    Reading a position can itself make Writer lay the anchored group out and move it, so
    a single pass over the members can mix positions from before and after such a move.
    Two identical passes in a row mean the document held still for a whole pass.
    """
    previous = _member_positions(group)
    for _ in range(20):
        current = _member_positions(group)
        if current == previous:
            return current
        previous = current
    raise AssertionError("the member positions did not settle over twenty reads")


def test_members_stay_put_when_the_diagram_is_refreshed_after_an_edit(office):
    from smart.controller import Controller
    from smart.diagram.data_of_diagram import DataOfDiagram
    from smart.gui import Gui

    document = office.desktop.loadComponentFromURL(
        "private:factory/swriter", "_blank", 0, ()
    )
    try:
        frame = document.getCurrentController().getFrame()
        controller = Controller(None, office.context, frame)

        data = DataOfDiagram()
        data.add(0, "")
        data.add(1, "")
        data.add(1, "")
        data.add(1, "")
        Gui._user_closed_dialog = False
        run_as_office_event(
            office.context, lambda: controller.create_diagram(data)
        )

        diagram = controller.get_diagram()
        group = diagram.get_group_shape()
        root_shape = diagram.get_diagram_tree().get_root_item().get_rectangle_shape()

        _settled_member_positions(group)

        # The edit under test: add a child below the root, the way the control dialog's
        # add button does, with the group moving while the edit is under way. In the
        # office that move is Writer placing the anchored group frame again when the
        # arriving shape grows the group; here the move is made explicitly between the
        # addition and the closing layout, so it happens on every run. Writer keeps the
        # X of an anchored group at its anchor, so the move is in Y.
        def add_child_below_root_while_the_group_moves():
            controller.remove_selection_listener()
            diagram.add_shape(root_shape)
            position = group.getPosition()
            position.Y += 2000
            group.setPosition(position)
            # The closing layout of every dialog edit. Before the fix for the moved
            # group frame it was the plain refresh_diagram, and reverting that fix
            # removes refresh_after_edit, so the plain refresh stands in for it again.
            closing_refresh = getattr(
                diagram, "refresh_after_edit", diagram.refresh_diagram
            )
            closing_refresh()

        run_as_office_event(office.context, add_child_below_root_while_the_group_moves)

        after_edit = _settled_member_positions(group)
        assert len(after_edit) == 5, "the edit added a fifth symbol to the group"

        # The manual refresh action: initialize the diagram again from the settled
        # document and lay it out.
        def refresh_manually():
            diagram.init_diagram(group_shape=group)
            diagram.init_properties()
            diagram.refresh_diagram()

        run_as_office_event(office.context, refresh_manually)

        after_manual_refresh = _settled_member_positions(group)

        moved = {
            name: (position, after_manual_refresh.get(name))
            for name, position in after_edit.items()
            if name not in after_manual_refresh
            or max(
                abs(position[0] - after_manual_refresh[name][0]),
                abs(position[1] - after_manual_refresh[name][1]),
            )
            > 2
        }
        assert not moved, (
            "the manual refresh moved members the edit had misplaced "
            f"(position after edit, after refresh): {moved}"
        )
    finally:
        document.setModified(False)
        document.close(False)
