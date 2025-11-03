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

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from com.sun.star.awt import Size, Point
from com.sun.star.beans import NamedValue, PropertyValue
from com.sun.star.xml import AttributeData


def insertSvgGraphic(ctx, model, svg_data, params):
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
        shape = model.createInstance("com.sun.star.drawing.GraphicObjectShape")
        shape.setPropertyValue("Graphic", graphic)

        # TODO: Make this dependent on actual SVG dimensions
        # (parse width and height properties of the svg element)
        shape_size = Size()
        shape_size.Height = 930
        shape_size.Width = 4000
        shape.setSize(shape_size)

        # set MilSym-specific user defined attributes
        insertGraphicAttributes(shape, params)

        # TODO: Set default anchoring for text documents
        #try:
        #    shape.setPropertyValue("AnchorType", TextContentAnchorType.AT_PARAGRAPH)
        #except:
        #    pass  # Not all document types support anchoring

        # Writer
        if model.supportsService("com.sun.star.text.TextDocument"):
            controller = model.getCurrentController()
            view_cursor_supplier = controller
            cursor = view_cursor_supplier.getViewCursor()
            text = cursor.getText()
            text.insertTextContent(cursor, shape, True)
            # Calc
        elif model.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            current_selection = model.getCurrentSelection()
            try:
                cell_position = current_selection.getPropertyValue("Position")
                shape.setPosition(cell_position)
            except:
                # Default position if we can't get cell position
                default_pos = Point()
                default_pos.X = 1000
                default_pos.Y = 1000
                shape.setPosition(default_pos)
                controller = model.getCurrentController()
                active_sheet = controller.getActiveSheet()
                draw_page_supplier = active_sheet
                draw_page = draw_page_supplier.getDrawPage()
                draw_page.add(shape)
                # Impress/Draw
        elif (model.supportsService("com.sun.star.presentation.PresentationDocument") or
              model.supportsService("com.sun.star.drawing.DrawingDocument")):
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
