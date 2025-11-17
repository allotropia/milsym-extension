# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import sys
import os
import uno
import xml.etree.ElementTree as ET

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from com.sun.star.awt import Size, Point
from com.sun.star.beans import PropertyValue
from com.sun.star.text.TextContentAnchorType import AT_PARAGRAPH
from com.sun.star.xml import AttributeData


def parse_svg_dimensions(svg_data):
    """Parse SVG dimensions and return width and height in 1/100mm units.

    Args:
        svg_data: SVG content as string

    Returns:
        Size(width, height)
    """
    width = 4000  # Default width
    height = 930  # Default height
    factor = 26.46  # Conversion factor from pixels to 1/100mm (assuming 96 DPI)

    try:
        # Parse SVG using ElementTree
        root = ET.fromstring(svg_data)

        # Extract width and height attributes
        width_str = root.get('width')
        height_str = root.get('height')

        if width_str:
            # Remove units like 'px', 'pt', etc. and extract numeric value
            width_num = ''.join(c for c in width_str if c.isdigit() or c == '.')
            if width_num:
                width = int(float(width_num) * factor)

        if height_str:
            # Remove units like 'px', 'pt', etc. and extract numeric value
            height_num = ''.join(c for c in height_str if c.isdigit() or c == '.')
            if height_num:
                height = int(float(height_num) * factor)
    except Exception as e:
        print(f"Warning: Could not parse SVG dimensions, using defaults: {e}")

    shape_size = Size()
    shape_size.Height = height
    shape_size.Width = width
    return shape_size


def insertSvgGraphic(ctx, model, svg_data, params):
    is_writer = model.supportsService("com.sun.star.text.TextDocument")
    is_calc = model.supportsService("com.sun.star.sheet.SpreadsheetDocument")
    is_draw_impress = model.supportsService("com.sun.star.presentation.PresentationDocument") or \
            model.supportsService("com.sun.star.drawing.DrawingDocument")

    try:
        pipe = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.io.Pipe", ctx)
        pipe.writeBytes(uno.ByteSequence(svg_data.encode('utf-8')))
        pipe.flush()
        pipe.closeOutput()
        graphic_provider = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.graphic.GraphicProvider", ctx)
        media_properties = (PropertyValue("InputStream", 0, pipe, 0),)
        graphic = graphic_provider.queryGraphic(media_properties)

        # For Writer, create a TextGraphicObject which behaves better (keeps aspect ratio, etc.)
        if is_writer:
            shape = model.createInstance("com.sun.star.text.TextGraphicObject")
        else:
            shape = model.createInstance("com.sun.star.drawing.GraphicObjectShape")
        shape.setPropertyValue("Graphic", graphic)

        size = parse_svg_dimensions(svg_data)
        shape.setSize(size)

        # set MilSym-specific user defined attributes
        insertGraphicAttributes(shape, params)

        # Writer
        if is_writer:
            controller = model.getCurrentController()
            view_cursor_supplier = controller
            cursor = view_cursor_supplier.getViewCursor()
            text = cursor.getText()
            text.insertTextContent(cursor, shape, True)
        # Calc - for spreadsheets, we'll use the shape directly since frames are not well supported
        elif is_calc:
            controller = model.getCurrentController()
            active_sheet = controller.getActiveSheet()
            draw_page = active_sheet.getDrawPage()
            draw_page.add(shape)

            try:
                # Try to position at current selection
                current_selection = model.getCurrentSelection()
                cell_position = current_selection.getPropertyValue("Position")
                shape.setPosition(cell_position)
            except:
                # Default position if we can't get cell position
                default_pos = Point()
                default_pos.X = 1000
                default_pos.Y = 1000
                shape.setPosition(default_pos)
        # Impress/Draw - for presentations and drawings, we'll use the shape directly
        elif is_draw_impress:
            controller = model.getCurrentController()
            current_page = controller.getCurrentPage()
            current_page.add(shape)
        else:
            print("Unsupported document type for graphic insertion")
    except Exception as e:
        print(f"Error inserting SVG graphic: {e}")


def insertGraphicAttributes(shape, params):
    attributeHash = shape.UserDefinedAttributes
    userAttrs = AttributeData()

    # first tuple is unnamed 'milsym code' entry. special handling.
    userAttrs.Type = "CDATA"
    userAttrs.Value = params[0]
    attributeHash["MilSymCode"] = userAttrs

    for entry in params[1:]:
        userAttrs.Type = "CDATA"
        userAttrs.Value = entry.Value
        attributeHash["MilSym" + entry.Name[0].upper() + entry.Name[1:]] = userAttrs

    # seems we're getting a copy above; set it explicitely
    shape.setPropertyValue("UserDefinedAttributes", attributeHash)


def getExtensionBasePath(ctx, extensionName="com.collabora.milsymbol"):
    srv = ctx.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
    return os.path.basename(srv.getPackageLocation(extensionName))
