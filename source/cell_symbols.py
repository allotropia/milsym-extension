# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

"""Generate one military symbol for each row of a selected cell range.

Every row of the selection holds the symbol identification code (SIDC)
of one symbol in one of its cells. When the first row of the selection
is a heading row, its headings say which column holds the codes and
which columns hold the text modifiers shown around the symbol, such as
the unique designation or the higher formation. See EXTRA_OPTION_NAMES
below for what is currently recognized.

Example:

1) simple collection of SIDCs, single column:

130310001812110000130100000000
130310001412140000000000000000
130310001712060000000000000000

2) header row with extra attributes:

|sidc                          | size | uniqueDesignation | country |
| :--------------------------- | ---: | :---------------- | ------: |
|130310001812110000130100000000|  80  | unit 1            |  at     |
|130310001412140000000000000000| 100  | unit 2            |  uk     |
|130310001712060000000000000000| 120  | unit 3            |  de     |

"""

import re

import unohelper
from com.sun.star.document import XUndoAction
from com.sun.star.sheet.CellFlags import FORMULA, STRING, VALUE

from symbol_dialog_handler import TEXTBOX_OPTION_NAMES
from translator import translate
from utils import (
    build_symbol_shape,
    createMilSymbolScriptInstance,
    generate_icon_svg,
    get_symbol_generation_size_px,
    locked_controllers,
    locked_undo_manager,
    mark_document_modified,
    symbol_script_args,
)

# The properties of a symbol shape that drawing it anew changes.
SYMBOL_SHAPE_PROPERTIES = ("Graphic", "Name", "UserDefinedAttributes")

# A symbol identification code is the 20 or 30 digit code of APP-6D and MIL-STD-2525D, or
# the 15 character letter code of MIL-STD-2525C, whose first letter names the coding scheme.
SIDC_PATTERN = re.compile(r"^(?:[0-9]{20}|[0-9]{30}|[SGWIOE][A-Z0-9-]{14})$")

# The heading of the column that holds the codes, in its normalized form.
SIDC_HEADING = "sidc"

# Options a column heading can name besides the text modifiers of the symbol dialog. The
# cell text is passed to milsymbol as it is.
EXTRA_OPTION_NAMES = (
    "stack",
    "size",
    "reinforced",
    "signature",
    "engagementType",
    "additionalInformation",
    "uniqueDesignation",
    "higherFormation",
    "country",
    "colorMode",
    "fillColor",
    "fill",
)


def normalized_heading(text):
    """The form of a column heading that is compared: lower case, without spaces,
    underscores and hyphens, so that "Unique Designation" and "uniqueDesignation" are the
    same heading.
    """
    return re.sub(r"[\s_-]", "", text).lower()


# The milsymbol option each recognized column heading stands for, keyed by the
# normalized heading.
OPTION_BY_HEADING = {
    normalized_heading(name): name
    for name in list(TEXTBOX_OPTION_NAMES.values()) + list(EXTRA_OPTION_NAMES)
}


def looks_like_sidc(text):
    return SIDC_PATTERN.match(text.strip()) is not None


def attribute_name(option):
    """The MilSym attribute a shape carries for a milsymbol option: the option name with
    its first letter raised, behind the MilSym prefix.
    """
    return "MilSym" + option[0].upper() + option[1:]


def selected_cell_ranges(selection):
    """The cell ranges a selection covers, when it spans more than one cell.

    Returns a list of cell range objects. The list is empty for a selection that is not
    a cell selection, and for one that covers a single cell.
    """
    if selection is None:
        return []
    try:
        if selection.supportsService("com.sun.star.sheet.SheetCellRanges"):
            ranges = [selection.getByIndex(i) for i in range(selection.getCount())]
        elif selection.supportsService("com.sun.star.sheet.SheetCellRange"):
            ranges = [selection]
        else:
            return []

        cell_count = 0
        for cell_range in ranges:
            address = cell_range.getRangeAddress()
            cell_count += (address.EndColumn - address.StartColumn + 1) * (
                address.EndRow - address.StartRow + 1
            )
    except Exception:
        return []
    return ranges if cell_count > 1 else []


def content_cells_by_row(cell_range):
    """The columns of the filled cells of a range, keyed by row and in row order.

    Only cells with content are looked at, so a selection of whole columns costs as
    much as the cells that are actually filled.
    """
    rows = {}
    content = cell_range.queryContentCells(STRING | VALUE | FORMULA)
    for address in content.getRangeAddresses():
        for row in range(address.StartRow, address.EndRow + 1):
            rows.setdefault(row, set()).update(
                range(address.StartColumn, address.EndColumn + 1)
            )
    return {row: sorted(columns) for row, columns in sorted(rows.items())}


def cell_text(sheet, column, row):
    return sheet.getCellByPosition(column, row).getString().strip()


def read_heading_row(sheet, row, columns):
    """The option each column of a heading row names, keyed by column index.

    The SIDC column is recorded under the name SIDC_HEADING. The map is empty when the
    row is not a heading row: when none of its cells is a known heading, or when one of
    them holds a code.
    """
    headings = {}
    for column in columns:
        text = cell_text(sheet, column, row)
        if looks_like_sidc(text):
            return {}
        key = normalized_heading(text)
        if key == SIDC_HEADING:
            headings[column] = SIDC_HEADING
        elif key in OPTION_BY_HEADING:
            headings[column] = OPTION_BY_HEADING[key]
    return headings


def symbol_attributes_for_row(sheet, row, columns, headings):
    """The MilSym attributes of the symbol one row describes, or None when the row holds
    no code.

    The code comes from the SIDC column when the headings name one, and otherwise from
    the first filled cell of the row that looks like a code. Every other heading whose
    cell is filled adds one attribute. A country adds the country flag as well, as the
    symbol dialog does.
    """
    sidc_columns = [
        column for column, name in headings.items() if name == SIDC_HEADING
    ]
    if sidc_columns:
        candidates = [cell_text(sheet, sidc_columns[0], row)]
    else:
        candidates = [cell_text(sheet, column, row) for column in columns]

    sidc = next((text for text in candidates if looks_like_sidc(text)), None)
    if sidc is None:
        return None

    attributes = {"MilSymCode": sidc}
    for column, name in sorted(headings.items()):
        if name == SIDC_HEADING:
            continue
        value = cell_text(sheet, column, row)
        if not value:
            continue
        attributes[attribute_name(name)] = value
        if name == "country":
            attributes[attribute_name("country_flag")] = "true"
    return attributes


def symbol_shapes_by_anchor_cell(draw_page):
    """The symbol shapes of a draw page that are anchored to a cell, keyed by the
    (column, row) of that cell.
    """
    shapes = {}
    for index in range(draw_page.getCount()):
        shape = draw_page.getByIndex(index)
        try:
            attributes = shape.UserDefinedAttributes
            if attributes is None or not attributes.hasByName("MilSymCode"):
                continue
            anchor = shape.getPropertyValue("Anchor")
            if anchor is None or not anchor.supportsService(
                "com.sun.star.sheet.SheetCell"
            ):
                continue
            address = anchor.getCellAddress()
        except Exception:
            continue
        shapes.setdefault((address.Column, address.Row), []).append(shape)
    return shapes


def symbol_name(attributes):
    return "Symbol (" + attributes["MilSymCode"] + ")"


def shape_state(shape):
    """The values of the properties drawing a symbol anew changes, and the size."""
    state = {name: shape.getPropertyValue(name) for name in SYMBOL_SHAPE_PROPERTIES}
    state["Size"] = shape.getSize()
    return state


def apply_shape_state(shape, state):
    for name in SYMBOL_SHAPE_PROPERTIES:
        shape.setPropertyValue(name, state[name])
    shape.setSize(state["Size"])


def insert_symbol_shape(ctx, model, draw_page, svg_data, attributes, size, cell):
    """Put a new symbol shape at the top left corner of a cell, anchored to it, and
    return the shape.
    """
    shape = build_symbol_shape(ctx, model, svg_data, symbol_script_args(attributes, size))
    shape.setPropertyValue("Name", symbol_name(attributes))
    draw_page.add(shape)
    shape.setPosition(cell.getPropertyValue("Position"))
    shape.setPropertyValue("Anchor", cell)
    return shape


def set_row_height_unrecorded(undo_manager, row_properties, height):
    """Set the height of a row under the lock of the undo manager, without leaving a
    record of it.

    Calc turns undo recording back on once it has changed a row height, whatever the lock
    said before. The lock is released and taken again afterwards, which turns recording
    off again for what follows. The manager reports itself as locked only while recording
    is off, so that report cannot tell whether the lock is still held, and the caller has
    to hold exactly one lock.
    """
    row_properties.Height = height
    undo_manager.unlock()
    undo_manager.lock()


class InsertedSymbol:
    """One symbol a run put onto a page: the shape while it is on the page, and what it
    takes to draw it there again.
    """

    def __init__(self, draw_page, cell, svg_data, attributes, size, shape):
        self.draw_page = draw_page
        self.cell = cell
        self.svg_data = svg_data
        self.attributes = attributes
        self.size = size
        self.shape = shape


class ReplacedSymbol:
    """One symbol shape a run drew anew in place, with its state from before and after."""

    def __init__(self, shape, before, after):
        self.shape = shape
        self.before = before
        self.after = after


class GenerateSymbolsUndoAction(unohelper.Base, XUndoAction):
    """The undo step of one run: the symbols it put onto the page, the symbols it drew
    anew in place, and the rows it made taller.

    Undo takes the new shapes off their pages and lets go of them, and redo draws them
    again from what was recorded, so the action holds only shapes that are on a page. A
    shape held while off its page outlives the document it belonged to, and its release
    after the document is gone crashes the office.
    """

    def __init__(self, ctx, model, title, undo_manager):
        self.ctx = ctx
        self.model = model
        self.Title = title
        self.undo_manager = undo_manager
        self.inserted = []
        self.replaced = []
        self.row_heights = []

    def record_inserted(self, inserted_symbol):
        self.inserted.append(inserted_symbol)

    def record_replaced(self, replaced_symbol):
        self.replaced.append(replaced_symbol)

    def record_row_height(self, row_properties, old_height, new_height):
        self.row_heights.append((row_properties, old_height, new_height))

    def undo(self):
        """Take the symbols of the run off their pages, and put back what it changed.

        Every symbol is answered for on its own, so that one that cannot be reached any
        more, because the user removed it after the run, does not keep the rest of the
        step from being undone.
        """
        try:
            with locked_undo_manager(self.undo_manager):
                for symbol in reversed(self.inserted):
                    if symbol.shape is None:
                        continue
                    try:
                        symbol.draw_page.remove(symbol.shape)
                    except Exception as e:
                        print(f"Error removing a generated symbol: {e}")
                    symbol.shape = None
                for symbol in self.replaced:
                    try:
                        apply_shape_state(symbol.shape, symbol.before)
                    except Exception as e:
                        print(f"Error putting back a redrawn symbol: {e}")
                for row_properties, old_height, new_height in self.row_heights:
                    try:
                        set_row_height_unrecorded(
                            self.undo_manager, row_properties, old_height
                        )
                    except Exception as e:
                        print(f"Error putting back a row height: {e}")
        except Exception as e:
            print(f"Error undoing the generated symbols: {e}")

    def redo(self):
        """Put the symbols of the run back, one by one as undo takes them off."""
        try:
            with locked_undo_manager(self.undo_manager):
                for symbol in self.replaced:
                    try:
                        apply_shape_state(symbol.shape, symbol.after)
                    except Exception as e:
                        print(f"Error drawing a symbol again: {e}")
                for symbol in self.inserted:
                    try:
                        symbol.shape = insert_symbol_shape(
                            self.ctx,
                            self.model,
                            symbol.draw_page,
                            symbol.svg_data,
                            symbol.attributes,
                            symbol.size,
                            symbol.cell,
                        )
                    except Exception as e:
                        print(f"Error inserting a generated symbol again: {e}")
                for row_properties, old_height, new_height in self.row_heights:
                    try:
                        set_row_height_unrecorded(
                            self.undo_manager, row_properties, new_height
                        )
                    except Exception as e:
                        print(f"Error setting a row height again: {e}")
        except Exception as e:
            print(f"Error redoing the generated symbols: {e}")


def place_symbol_at_cell(
    ctx, model, draw_page, svg_data, attributes, size, cell, existing_shape, undo_action
):
    """Put the symbol for one row at a cell, and record what was done on the undo action.

    A symbol shape that an earlier run anchored to the cell is drawn anew in place and
    keeps its size. Otherwise a new shape goes to the top left corner of the cell,
    anchored to it.
    """
    if existing_shape is not None:
        before = shape_state(existing_shape)
        shape = build_symbol_shape(
            ctx, model, svg_data, symbol_script_args(attributes, size), existing_shape
        )
        shape.setPropertyValue("Name", symbol_name(attributes))
        undo_action.record_replaced(ReplacedSymbol(shape, before, shape_state(shape)))
    else:
        shape = insert_symbol_shape(
            ctx, model, draw_page, svg_data, attributes, size, cell
        )
        undo_action.record_inserted(
            InsertedSymbol(draw_page, cell, svg_data, attributes, size, shape)
        )

    row_properties = cell.getSpreadsheet().getRows().getByIndex(
        cell.getCellAddress().Row
    )
    symbol_height = shape.getSize().Height
    if row_properties.Height < symbol_height:
        undo_action.record_row_height(
            row_properties, row_properties.Height, symbol_height
        )
        set_row_height_unrecorded(
            undo_action.undo_manager, row_properties, symbol_height
        )


def generate_symbols_for_range(ctx, model, script, size, cell_range, undo_action):
    """Insert the symbols the rows of one cell range describe, and return how many.

    The symbols go into the column right of the range, or into its last column when the
    range already reaches the last column of the sheet. A symbol that an earlier run
    anchored to the same cell is drawn anew in place.
    """
    sheet = cell_range.getSpreadsheet()
    address = cell_range.getRangeAddress()
    draw_page = sheet.getDrawPage()

    rows = content_cells_by_row(cell_range)
    if not rows:
        return 0

    target_column = address.EndColumn + 1
    if target_column >= sheet.getColumns().getCount():
        target_column = address.EndColumn

    headings = {}
    if address.StartRow in rows:
        headings = read_heading_row(sheet, address.StartRow, rows[address.StartRow])

    existing_symbols = symbol_shapes_by_anchor_cell(draw_page)

    inserted = 0
    for row, columns in rows.items():
        if headings and row == address.StartRow:
            continue
        attributes = symbol_attributes_for_row(sheet, row, columns, headings)
        if attributes is None:
            continue
        curr_size = size
        if attributes.get('MilSymSize'):
            curr_size = attributes['MilSymSize']
        svg_data = generate_icon_svg(script, attributes, curr_size)
        if not svg_data:
            print(f"No symbol drawing for row {row + 1}: {attributes['MilSymCode']}")
            continue

        existing = existing_symbols.get((target_column, row), [])
        target_cell = sheet.getCellByPosition(target_column, row)
        place_symbol_at_cell(
            ctx,
            model,
            draw_page,
            svg_data,
            attributes,
            curr_size,
            target_cell,
            existing[0] if existing else None,
            undo_action,
        )
        inserted += 1
    return inserted


def generate_symbols_from_rows(ctx, model, selection):
    """Insert one military symbol for each row of the selected cell ranges, and return
    how many were inserted.

    The whole run is one undo step and the views repaint once it is done.
    """
    ranges = selected_cell_ranges(selection)
    if not ranges:
        return 0

    script = createMilSymbolScriptInstance(ctx, model)
    size = get_symbol_generation_size_px(ctx)

    undo_manager = model.getUndoManager()
    undo_action = GenerateSymbolsUndoAction(
        ctx, model, translate(ctx, "ContextMenu.GenerateSymbolsFromRows"), undo_manager
    )

    inserted = 0
    with locked_controllers(model), locked_undo_manager(undo_manager):
        for cell_range in ranges:
            inserted += generate_symbols_for_range(
                ctx, model, script, size, cell_range, undo_action
            )

    if inserted:
        undo_manager.addUndoAction(undo_action)
        mark_document_modified(model)
    return inserted
