# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

"""Inserting a symbol is one undo step, checked against a running office.

The tests start a headless office with a scratch profile and drive insertSvgGraphic over
a UNO socket connection, the way the symbol dialog's Save does for a new symbol. Impress,
Draw and Calc record nothing on their own when a shape is added to a draw page through
the API, but they do record setting its name, so without the extension registering one
action for the whole insertion, undo would only take back the name and leave the shape
on the page.

The drawing is handed in as ready-made SVG, so no extension package has to be installed
in the office: only the GraphicProvider service and the undo manager take part.
"""

import pytest
from com.sun.star.beans import NamedValue
from conftest import run_as_office_event

SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40">'
    '<rect width="40" height="40" fill="#8cc"/></svg>'
)
SIDC = "30031000001211000000"

EDITED_SVG = (
    '<svg xmlns="http://www.w3.org/2000/svg" width="40" height="60">'
    '<rect width="40" height="60" fill="#c88"/></svg>'
)
EDITED_SIDC = "30031000001211000001"


def _insert_page_of(document):
    """The draw page insertSvgGraphic puts a new symbol onto."""
    controller = document.getCurrentController()
    if document.supportsService("com.sun.star.text.TextDocument"):
        return document.getDrawPage()
    if document.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
        return controller.getActiveSheet().getDrawPage()
    return controller.getCurrentPage()


@pytest.mark.parametrize("factory", ["simpress", "sdraw", "scalc"])
def test_undo_removes_an_inserted_symbol_and_redo_restores_it(office, factory):
    from utils import insertSvgGraphic

    document = office.desktop.loadComponentFromURL(
        f"private:factory/{factory}", "_blank", 0, ()
    )
    try:
        page = _insert_page_of(document)
        undo_manager = document.getUndoManager()
        undo_manager.clear()
        shapes_before = page.getCount()
        name = "Symbol (undo test)"

        insertSvgGraphic(
            office.context,
            document,
            SVG,
            [SIDC, NamedValue("size", 37.8)],
            None,
            name,
        )

        assert page.getCount() == shapes_before + 1, "the symbol was not inserted"
        assert len(undo_manager.getAllUndoActionTitles()) == 1, (
            "the insertion is one undo step, with no step of its own for naming the shape"
        )

        undo_manager.undo()
        assert page.getCount() == shapes_before, "undo takes the symbol off the page"

        undo_manager.redo()
        assert page.getCount() == shapes_before + 1, "redo puts the symbol back"
        redone_names = [
            page.getByIndex(index).Name for index in range(page.getCount())
        ]
        assert name in redone_names, "the redone symbol carries the recorded name"

        undo_manager.undo()
        assert page.getCount() == shapes_before, "the redone symbol can be undone again"
    finally:
        document.setModified(False)
        document.close(False)


@pytest.mark.parametrize("factory", ["swriter", "simpress", "scalc"])
def test_editing_a_selected_symbol_is_one_undo_step(office, factory):
    """Editing a symbol changes it in place, and one undo restores what it was.

    The context menu's edit action runs with the symbol selected, so in Writer the view
    holds an object selection and no text cursor while the edit is saved. The edit must
    leave the shape on its page where it is, with the new drawing, attributes and name,
    and the document marked as changed; undoing it must restore the old drawing,
    attributes, name and size without taking the shape off the page.
    """
    from utils import insertSvgGraphic

    document = office.desktop.loadComponentFromURL(
        f"private:factory/{factory}", "_blank", 0, ()
    )
    try:
        page = _insert_page_of(document)
        undo_manager = document.getUndoManager()
        undo_manager.clear()
        old_name = "Symbol (edit test old)"
        new_name = "Symbol (edit test new)"

        run_as_office_event(
            office.context,
            lambda: insertSvgGraphic(
                office.context,
                document,
                SVG,
                [SIDC, NamedValue("size", 37.8)],
                None,
                old_name,
            ),
        )
        shapes_after_insert = page.getCount()
        shape = page.getByIndex(shapes_after_insert - 1)
        assert shape.getName() == old_name
        old_size = (shape.getSize().Width, shape.getSize().Height)
        steps_after_insert = len(undo_manager.getAllUndoActionTitles())

        # Select the shape, the way the context menu edit leaves it selected, and
        # save an edit the way the symbol dialog does.
        run_as_office_event(
            office.context,
            lambda: document.getCurrentController().select(shape),
        )
        document.setModified(False)
        run_as_office_event(
            office.context,
            lambda: insertSvgGraphic(
                office.context,
                document,
                EDITED_SVG,
                [EDITED_SIDC, NamedValue("size", 37.8)],
                shape,
                new_name,
            ),
        )

        assert shape.getName() == new_name, "the edit renamed the shape"
        assert shape.UserDefinedAttributes.getByName("MilSymCode").Value == EDITED_SIDC
        assert document.isModified(), "the edit marked the document as changed"
        assert page.getCount() == shapes_after_insert, "the edit added no shape"
        assert len(undo_manager.getAllUndoActionTitles()) == steps_after_insert + 1, (
            "the edit is one undo step"
        )

        run_as_office_event(office.context, undo_manager.undo)
        assert page.getCount() == shapes_after_insert, "undo keeps the shape in place"
        assert shape.getName() == old_name, "undo restores the name"
        assert shape.UserDefinedAttributes.getByName("MilSymCode").Value == SIDC, (
            "undo restores the attributes"
        )
        assert (shape.getSize().Width, shape.getSize().Height) == old_size, (
            "undo restores the size"
        )

        run_as_office_event(office.context, undo_manager.redo)
        assert shape.getName() == new_name, "redo applies the edit again"
        assert shape.UserDefinedAttributes.getByName("MilSymCode").Value == EDITED_SIDC
    finally:
        document.setModified(False)
        document.close(False)
