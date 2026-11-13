# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

"""A copied and pasted orbat is recognized shape by shape, despite the renaming.

Pasting an orbat group into the document it was copied from renames it, because the
office keeps shape names in a document unique: a space and a counter are appended to
whatever the name already ends in, so "OrbatDiagram12-RectangleShape" becomes
"OrbatDiagram12-RectangleShape 1" and the numbered "OrbatDiagram12-GraphicObjectShape3"
becomes "OrbatDiagram12-GraphicObjectShape3 2". Office versions differ in how far this
reaches: some rename every shape in the group, others only the group itself. The
extension classifies shapes by name, so the classification has to see through that
suffix; when it does not, the copy's control shape and members go unrecognized, the
diagram tree never accounts for the whole group, and editing the copy misbehaves, from
the dialog tree collapsing on every click to a delete that does nothing.
"""

from conftest import run_as_office_event
from smart.diagram.diagram import Diagram


def test_shape_names_of_a_pasted_copy_are_classified():
    """The name a shape was created with and its pasted form give the same role."""
    for suffix in ("", " 1", " 12"):
        assert (
            Diagram.shape_role("OrbatDiagram12-RectangleShape" + suffix)
            == Diagram.DIAGRAM_BASE_SHAPE_TYPE
        )
        assert (
            Diagram.shape_role("OrbatDiagram12-GraphicObjectShape3" + suffix)
            == Diagram.DIAGRAM_SHAPE_TYPE
        )
        assert (
            Diagram.shape_role("OrbatDiagram12-ConnectorShape10" + suffix)
            == Diagram.CONNECTOR_SHAPE
        )


def test_foreign_shape_names_are_not_classified():
    assert Diagram.shape_role("Rectangle 5") == ""
    assert Diagram.shape_role("GraphicObjectShape1") == ""
    assert Diagram.shape_role("OrbatDiagram12-TitleTextShape") == ""
    assert Diagram.shape_role("") == ""
    assert Diagram.shape_role(None) == ""


def test_a_pasted_orbat_copy_is_a_complete_diagram(office):
    from smart.controller import Controller
    from smart.diagram.data_of_diagram import DataOfDiagram
    from smart.gui import Gui

    document = office.desktop.loadComponentFromURL(
        "private:factory/simpress", "_blank", 0, ()
    )
    try:
        frame = document.getCurrentController().getFrame()
        controller = Controller(None, office.context, frame)

        data = DataOfDiagram()
        data.add(0, "")
        data.add(1, "")
        data.add(1, "")
        Gui._user_closed_dialog = False
        run_as_office_event(office.context, lambda: controller.create_diagram(data))

        group = controller.get_diagram().get_group_shape()
        page = document.getCurrentController().getCurrentPage()
        shapes_before = page.getCount()

        # Copy the whole group and paste it into the same document. Creating the
        # diagram leaves the view editing inside the group, so the view steps out of
        # it first, and the selection listener is taken off so the test's controller
        # does not react to the selection changing under the dispatches.
        dispatcher = office.context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.DispatchHelper", office.context
        )

        def copy_and_paste_the_group():
            controller.remove_selection_listener()
            dispatcher.executeDispatch(frame, ".uno:LeaveGroup", "", 0, ())
            frame.getController().select(group)
            dispatcher.executeDispatch(frame, ".uno:Copy", "", 0, ())
            dispatcher.executeDispatch(frame, ".uno:Paste", "", 0, ())

        run_as_office_event(office.context, copy_and_paste_the_group)

        assert page.getCount() == shapes_before + 1, "the paste added a group"
        pasted = page.getByIndex(page.getCount() - 1)
        assert pasted.supportsService("com.sun.star.drawing.GroupShape")
        assert pasted.getName().startswith("OrbatDiagram")
        assert pasted.getName() != group.getName(), "the office renamed the pasted group"

        # Office versions differ in how far the renaming reaches: some rename every
        # member of the copy, others only the group. Give any member the office left
        # alone its renamed form here, so the copy looks the same to the extension on
        # every office version.
        def rename_members_the_way_a_paste_does():
            for index in range(pasted.getCount()):
                shape = pasted.getByIndex(index)
                if " " not in shape.getName():
                    shape.setName(shape.getName() + " 2")

        run_as_office_event(office.context, rename_members_the_way_a_paste_does)
        member_names = [
            pasted.getByIndex(index).getName() for index in range(pasted.getCount())
        ]

        # Bind a diagram to the pasted group, the way the edit and refresh menu
        # actions do for whatever orbat group is selected.
        def bind_diagram_to_pasted_group():
            controller.set_group_type(controller.ORGANIGROUP)
            controller.set_diagram_type(controller.ORGANIGRAM)
            controller.instantiate_diagram()
            controller.get_diagram().init_diagram(group_shape=pasted)
            controller.get_diagram().init_properties()

        run_as_office_event(office.context, bind_diagram_to_pasted_group)

        tree = controller.get_diagram().get_diagram_tree()
        assert tree.get_root_item() is not None, "the copy's root symbol was not found"
        assert tree.knows_every_shape_in_the_group(), (
            "the renamed shapes of the pasted copy were not all recognized: "
            f"{member_names}"
        )
    finally:
        document.setModified(False)
        document.close(False)
