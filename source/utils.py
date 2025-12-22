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
from unohelper import fileUrlToSystemPath

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from com.sun.star.awt import Size, Point
from com.sun.star.beans import PropertyValue
from com.sun.star.text.TextContentAnchorType import AT_PARAGRAPH
from com.sun.star.xml import AttributeData
from com.sun.star.beans import NamedValue


def parse_svg_dimensions(svg_data, scale_factor=1):
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
    shape_size.Height = height * scale_factor
    shape_size.Width = width * scale_factor
    return shape_size


def extractGraphicAttributes(shape):
    """Extract symbol attributes from shape's UserDefinedAttributes

    Args:
        shape: The shape object to extract attributes from

    Returns:
        Dictionary of attribute name to value mappings
    """
    attributeHash = shape.UserDefinedAttributes

    attributes = {}
    for name in attributeHash.getElementNames():
        attr_data = attributeHash.getByName(name)
        attributes[name] = attr_data.Value
    return attributes


def insertSvgGraphic(ctx, model, svg_data, params, selected_shape, smybol_name, scale_factor=1):
    is_writer = model.supportsService("com.sun.star.text.TextDocument")
    is_calc = model.supportsService("com.sun.star.sheet.SpreadsheetDocument")
    is_draw_impress = model.supportsService("com.sun.star.presentation.PresentationDocument") or \
            model.supportsService("com.sun.star.drawing.DrawingDocument")

    try:
        graphic = create_graphic_from_svg(ctx, svg_data)

        # For Writer, create a TextGraphicObject which behaves better (keeps aspect ratio, etc.)
        if selected_shape is None:
            if is_writer:
                shape = model.createInstance("com.sun.star.text.TextGraphicObject")
            else:
                shape = model.createInstance("com.sun.star.drawing.GraphicObjectShape")
        else:
            shape = selected_shape

        shape.setPropertyValue("Graphic", graphic)

        size = parse_svg_dimensions(svg_data, scale_factor)
        shape.setSize(size)

        # set MilSym-specific user defined attributes
        insertGraphicAttributes(shape, params)

        # Writer
        if is_writer:
            controller = model.getCurrentController()
            view_cursor_supplier = controller
            cursor = view_cursor_supplier.getViewCursor()
            text = cursor.getText()
            shape.setName(smybol_name)
            text.insertTextContent(cursor, shape, True)
        # Calc - for spreadsheets, we'll use the shape directly since frames are not well supported
        elif is_calc:
            controller = model.getCurrentController()
            active_sheet = controller.getActiveSheet()
            draw_page = active_sheet.getDrawPage()
            shape.setPropertyValue("Name", smybol_name)
            pos = shape.getPosition()
            draw_page.add(shape)
            shape.setPosition(pos)

            if selected_shape is None:
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
            shape.setPropertyValue("Name", smybol_name)
            pos = shape.getPosition()
            current_page.add(shape)
            shape.setPosition(pos)
        else:
            print("Unsupported document type for graphic insertion")
    except Exception as e:
        print(f"Error inserting SVG graphic: {e}")


def insertGraphicAttributes(shape, params):
    attributeHash = shape.UserDefinedAttributes
    userAttrs = AttributeData()

    for name in list(attributeHash.getElementNames()):
        attributeHash.removeByName(name)

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

def generate_icon_svg(script, attributes, size):
    """Generate SVG icon from symbol attributes

    Args:
        attributes: Dictionary of symbol attributes extracted from shape

    Returns:
        SVG string data or None if generation fails
    """
    try:
        sidc_code = attributes.get("MilSymCode")
        if not sidc_code:
            return None

        args = [sidc_code, NamedValue("size", size)]

        if "MilSymStack" in attributes:
            args.append(NamedValue("stack", attributes["MilSymStack"]))

        if "MilSymReinforced" in attributes:
            args.append(NamedValue("reinforced", attributes["MilSymReinforced"]))

        if "MilSymStaff" in attributes:
            args.append(NamedValue("staff", attributes["MilSymStaff"]))

        if "MilSymSpecialheadquarters" in attributes:
            args.append(
                NamedValue(
                    "specialheadquarters", attributes["MilSymSpecialheadquarters"]
                )
            )

        if "MilSymCountrycode" in attributes:
            args.append(NamedValue("countrycode", attributes["MilSymCountrycode"]))

        result = script.invoke(args, (), ())
        svg_data = str(result[0])
        return svg_data

    except Exception as e:
        print(f"Error generating icon SVG: {e}")
        return None

def create_graphic_from_svg(ctx, svg_data):
    """Create XGraphic from SVG data"""
    try:
        if not svg_data:
            return None

        # Create a pipe to stream the SVG data
        pipe = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.io.Pipe", ctx)
        pipe.writeBytes(uno.ByteSequence(svg_data.encode('utf-8')))
        pipe.flush()
        pipe.closeOutput()

        # Create graphic provider
        graphic_provider = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.graphic.GraphicProvider", ctx)

        # Create media properties for the SVG data
        media_properties = (PropertyValue("InputStream", 0, pipe, 0),)

        # Query the graphic from the provider
        graphic = graphic_provider.queryGraphic(media_properties)
        return graphic

    except Exception as e:
        print(f"Error creating graphic from SVG: {e}")
        return None


def get_package_location(ctx, extensionName="com.collabora.milsymbol"):
    """Get package location from package information provider"""
    srv = ctx.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider")
    return fileUrlToSystemPath(srv.getPackageLocation(extensionName))

def getExtensionBasePath(ctx, extensionName="com.collabora.milsymbol"):
    """Get the base path of the extension installation directory"""
    return os.path.basename(get_package_location(ctx, extensionName))
