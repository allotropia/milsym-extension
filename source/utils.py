# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import os
import tempfile
import xml.etree.ElementTree as ET
from contextlib import contextmanager

import uno

from com.sun.star.awt import Point, Rectangle, Size
from com.sun.star.beans import NamedValue, PropertyValue
from com.sun.star.xml import AttributeData

from perf import count

# Conversion factor from pixels to 1/100mm, assuming 96 DPI (2540 / 96)
PX_TO_MM100 = 26.46

# geometry attributes namespace, we've hacked into combine.sh.
#
# values are in pixel units of the width and height attributes, measured
# from the top left corner of the drawing:
#
#   octagon = "x y width height" of the frame octagon
#   anchor  = "x y" of the symbol anchor, the end of the staff for
#             a headquarters and the octagon centre for every other symbol
MILSYM_SVG_NAMESPACE = "urn:collabora:milsym"


@contextmanager
def locked_controllers(model):
    """Hold the model's controller lock for the body, and release it however the body ends.

    While the lock is held the views do not repaint, so a run of changes appears as one
    step instead of letting the reader watch the diagram being rebuilt. The model counts
    the locks it is given, so holding one inside another is safe.
    """
    if model is None:
        yield
        return

    model.lockControllers()
    try:
        yield
    finally:
        model.unlockControllers()


@contextmanager
def locked_undo_manager(undo_manager):
    """Hold the lock of an undo manager for the body, and release it however the body ends.

    While the lock is held, the changes made to the document are not recorded for undo.
    The manager counts the locks it is given, so holding one inside another is safe.
    """
    if undo_manager is None:
        yield
        return

    undo_manager.lock()
    try:
        yield
    finally:
        undo_manager.unlock()


# The settings branch that holds the extension's own configuration
SETTINGS_NODEPATH = "/com.collabora.milsymbol.Configuration/Settings"

# The configuration provider of this process, kept because building one is far more
# expensive than reading a value through it. It is None until the first read.
_config_provider = None


def get_settings_access(ctx):
    """Open the extension's settings branch for reading.

    The provider is built once per process and then reused. A fresh access to the branch
    is opened on every call, so a value changed while the office is running is seen.
    """
    global _config_provider

    if _config_provider is None:
        _config_provider = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", ctx
        )

    prop = PropertyValue()
    prop.Name = "nodepath"
    prop.Value = SETTINGS_NODEPATH

    return _config_provider.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationAccess", (prop,)
    )


def get_default_symbol_height_cm(ctx):
    """Get the default height of the symbol frame (octagon) from the configuration.

    Returns height in 1/100mm units (1cm = 1000 units), hidden config item name is: DefaultSymbolHeightCm
    """
    default_height = 1000  # 1cm in 1/100mm units
    count("config: read DefaultSymbolHeightCm")

    try:
        config_access = get_settings_access(ctx)

        # Get the DefaultSymbolHeightCm setting
        if config_access.hasByName("DefaultSymbolHeightCm"):
            height_cm = config_access.getByName("DefaultSymbolHeightCm")
            # Convert cm to 1/100mm units (1cm = 1000 units in 1/100mm)
            default_height = int(float(height_cm) * 1000)

    except Exception as e:
        print(
            f"Warning: Could not read symbol height configuration, using default: {e}"
        )

    return default_height


def get_symbol_generation_size_px(ctx):
    """Get the milsymbol size argument that makes the symbol frame (octagon) come out at
    the configured default height.

    The milsymbol size argument is the height of the frame octagon in pixels. An SVG
    generated with this size and inserted at its intrinsic pixel dimensions (at 96 DPI)
    has a frame of exactly the configured height, independent of decorations like
    echelon markers or text labels that enlarge the whole symbol.

    Returns the size in pixels as a float.
    """
    return get_default_symbol_height_cm(ctx) / PX_TO_MM100


def get_recorded_symbol_size_px(attributes, ctx):
    """Get the milsymbol size argument that a symbol's drawing was generated with.

    The size is kept on the shape as the MilSymSize attribute. A symbol that does not
    carry one gets the size that makes the frame octagon come out at the configured
    default height.

    Args:
        attributes: Dictionary of symbol attributes extracted from a shape

    Returns the size in pixels as a float.
    """
    try:
        return float(attributes["MilSymSize"])
    except (KeyError, TypeError, ValueError):
        return get_symbol_generation_size_px(ctx)


def containing_orbat_group(shape):
    """The ORBAT group shape a shape belongs to, or None when it belongs to none.

    A group shape named for an ORBAT answers for itself. A shape inside such a group,
    at any depth, answers with that group.
    """
    current = shape
    while current is not None:
        try:
            if current.supportsService(
                "com.sun.star.drawing.GroupShape"
            ) and current.getName().startswith("OrbatDiagram"):
                return current
            parent = current.getParent()
        except Exception:
            return None
        if parent is None or not hasattr(parent, "supportsService"):
            return None
        try:
            if not parent.supportsService("com.sun.star.drawing.Shape"):
                return None
        except Exception:
            return None
        current = parent
    return None


def is_orbat_feature_enabled(ctx):
    """Check if hidden feature flag for Orbat handling is enabled.

    Returns True or False
    """
    default_state = True

    try:
        config_access = get_settings_access(ctx)

        # Get the OrbatFeatureFlag setting
        if config_access.hasByName("OrbatFeatureFlag"):
            status = config_access.getByName("OrbatFeatureFlag")
            default_state = bool(status)

    except Exception as e:
        print(
            f"Warning: Could not read Orbat feature flag, using default: {e}"
        )

    return default_state


def parse_svg_dimensions(svg_data):
    """Parse SVG dimensions and return width and height in 1/100mm units.

    Args:
        svg_data: SVG content as string

    Returns:
        Size(width, height)
    """
    width = 4000  # Default width
    height = 930  # Default height

    try:
        width_px, height_px = svg_size_px(ET.fromstring(svg_data))
        if width_px:
            width = width_px * PX_TO_MM100
        if height_px:
            height = height_px * PX_TO_MM100
    except Exception as e:
        print(f"Warning: Could not parse SVG dimensions, using defaults: {e}")

    shape_size = Size()
    shape_size.Height = height
    shape_size.Width = width
    return shape_size


def svg_length_px(text):
    """The number in an SVG length such as "158" or "158px", or None when there is none."""
    if not text:
        return None
    number = "".join(c for c in text if c.isdigit() or c == ".")
    return float(number) if number else None


def svg_size_px(root):
    """The width and height attributes of an SVG root element as a pair of pixel numbers.

    Either value is None when the attribute is missing or carries no number.
    """
    return svg_length_px(root.get("width")), svg_length_px(root.get("height"))


def parse_svg_geometry(svg_data, name, count):
    """Read one milsym geometry attribute off the root element of a generated symbol SVG.

    The attribute holds count numbers in pixels. They come back as fractions of the
    drawing's width (even positions) and height (odd positions), so the result stays
    right at whatever size the drawing is shown at later.

    Returns a tuple of count floats, or None when the SVG carries no such attribute or
    its size is unknown.
    """
    try:
        root = ET.fromstring(svg_data)
        value = root.get("{%s}%s" % (MILSYM_SVG_NAMESPACE, name))
        if value is None:
            return None
        numbers = [float(part) for part in value.split()]
        if len(numbers) != count:
            return None
        width_px, height_px = svg_size_px(root)
        if not width_px or not height_px:
            return None
        return tuple(
            number / (width_px if index % 2 == 0 else height_px)
            for index, number in enumerate(numbers)
        )
    except Exception as e:
        print(f"Warning: Could not parse SVG {name} geometry: {e}")
        return None


def parse_svg_octagon(svg_data):
    """The frame octagon of a generated symbol SVG, as fractions of the drawing size.

    Returns (x, y, width, height), each a fraction of the drawing's width or height with
    the top left corner of the drawing at (0, 0). Returns None for an SVG without octagon
    information, such as one generated before the attribute was added.
    """
    return parse_svg_geometry(svg_data, "octagon", 4)


def parse_svg_anchor(svg_data):
    """The anchor point of a generated symbol SVG, as fractions of the drawing size.

    Returns (x, y), each a fraction of the drawing's width or height with the top left
    corner of the drawing at (0, 0). Returns None for an SVG without anchor information.
    """
    return parse_svg_geometry(svg_data, "anchor", 2)


def parse_svg_frame(svg_data):
    """The bounding box of the visible frame outline of a generated symbol SVG, as
    fractions of the drawing size.

    Returns (x, y, width, height), each a fraction of the drawing's width or height with
    the top left corner of the drawing at (0, 0). Returns None for an SVG without frame
    information, such as one generated before the attribute was added.
    """
    return parse_svg_geometry(svg_data, "frame", 4)


class SymbolGeometry:
    """Where the frame octagon, the frame outline and the anchor of a milsymbol drawing sit.

    Every value is a fraction of the drawing's width or height, with the top left corner
    of the drawing at (0, 0), so the geometry holds at whatever size the drawing is shown.
    octagon is (x, y, width, height) of the frame octagon: the square the symbol icon fits
    in, which the visible frame surrounds and can extend beyond. frame is (x, y, width,
    height) of the bounding box of the visible frame outline; for a drawing recorded
    before that value was added it falls back to the octagon. anchor is (x, y) of the
    symbol anchor: the end of the staff for a headquarters and the octagon centre for
    every other symbol.
    """

    def __init__(self, octagon, anchor, frame=None):
        self.octagon = octagon
        self.anchor = anchor
        self.frame = frame if frame is not None else octagon

    def octagon_left(self):
        return self.octagon[0]

    def octagon_top(self):
        return self.octagon[1]

    def octagon_bottom(self):
        return self.octagon[1] + self.octagon[3]

    def octagon_centre_x(self):
        return self.octagon[0] + self.octagon[2] / 2

    def octagon_centre_y(self):
        return self.octagon[1] + self.octagon[3] / 2

    def frame_left(self):
        return self.frame[0]

    def frame_top(self):
        return self.frame[1]

    def frame_bottom(self):
        return self.frame[1] + self.frame[3]

    def frame_centre_y(self):
        return self.frame[1] + self.frame[3] / 2

    def anchor_x(self):
        return self.anchor[0]

    def anchor_y(self):
        return self.anchor[1]


def parse_svg_symbol_geometry(svg_data):
    """The octagon, frame and anchor of a generated symbol SVG as a SymbolGeometry, or None
    for an SVG that carries no octagon or anchor information, such as one generated before
    the attributes were added or a picture that is not a milsymbol drawing."""
    if not svg_data:
        return None
    octagon = parse_svg_octagon(svg_data)
    anchor = parse_svg_anchor(svg_data)
    if octagon is None or anchor is None:
        return None
    return SymbolGeometry(octagon, anchor, parse_svg_frame(svg_data))


def read_shape_svg(ctx, shape):
    """The SVG source of the picture a shape shows, or None when it has none.

    The office keeps the bytes a vector picture was loaded from and writes them back
    unchanged when the picture is exported as SVG, so the string returned is the one the
    picture was made from, with every attribute it carried. The export goes through a
    temporary file, which is removed again before this returns.
    """
    try:
        graphic = shape.getPropertyValue("Graphic")
        if graphic is None:
            return None
        count("shape: read graphic svg")
        provider = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.graphic.GraphicProvider", ctx
        )
        handle, path = tempfile.mkstemp(suffix=".svg")
        os.close(handle)
        try:
            provider.storeGraphic(
                graphic,
                (
                    PropertyValue("URL", 0, uno.systemPathToFileUrl(path), 0),
                    PropertyValue("MimeType", 0, "image/svg+xml", 0),
                ),
            )
            with open(path, "rb") as exported:
                data = exported.read()
        finally:
            if os.path.exists(path):
                os.remove(path)
        if not data:
            return None
        return data.decode("utf-8")
    except Exception as e:
        print(f"Warning: Could not read the SVG of a shape: {e}")
        return None


def octagon_rect_in_shape(shape, svg_data):
    """The frame octagon of a symbol shape, in the shape's coordinate system.

    The shape shows svg_data at some size that keeps the drawing's aspect ratio, so the
    octagon fractions of the drawing map straight onto the shape's position and size.

    Returns a Rectangle in 1/100 mm, or None when the SVG carries no octagon information.
    """
    octagon = parse_svg_octagon(svg_data)
    if octagon is None:
        return None
    position = shape.getPosition()
    size = shape.getSize()
    rect = Rectangle()
    rect.X = int(round(position.X + octagon[0] * size.Width))
    rect.Y = int(round(position.Y + octagon[1] * size.Height))
    rect.Width = int(round(octagon[2] * size.Width))
    rect.Height = int(round(octagon[3] * size.Height))
    return rect


def anchor_point_in_shape(shape, svg_data):
    """The anchor point of a symbol shape, in the shape's coordinate system.

    Returns a Point in 1/100 mm, or None when the SVG carries no anchor information.
    """
    anchor = parse_svg_anchor(svg_data)
    if anchor is None:
        return None
    position = shape.getPosition()
    size = shape.getSize()
    point = Point()
    point.X = int(round(position.X + anchor[0] * size.Width))
    point.Y = int(round(position.Y + anchor[1] * size.Height))
    return point


def fit_size_to_aspect_ratio(bounding_size, intrinsic_size):
    """Fit a size to the aspect ratio of another size by shrinking one dimension.

    Returns a Size with the aspect ratio of intrinsic_size that fits inside
    bounding_size. One dimension of bounding_size is kept, the other is reduced.
    If either size has a non-positive dimension, bounding_size is returned
    unchanged.
    """
    if (
        bounding_size.Width <= 0
        or bounding_size.Height <= 0
        or intrinsic_size.Width <= 0
        or intrinsic_size.Height <= 0
    ):
        return bounding_size

    fitted = Size()
    fitted.Width = bounding_size.Width
    fitted.Height = bounding_size.Height
    if (
        bounding_size.Width * intrinsic_size.Height
        > bounding_size.Height * intrinsic_size.Width
    ):
        # The box is proportionally wider than the content, so the width shrinks
        fitted.Width = int(
            bounding_size.Height * intrinsic_size.Width / intrinsic_size.Height
        )
    else:
        # The box is proportionally taller than the content, so the height shrinks
        fitted.Height = int(
            bounding_size.Width * intrinsic_size.Height / intrinsic_size.Width
        )
    return fitted


def extractGraphicAttributes(shape):
    """Extract symbol attributes from shape's UserDefinedAttributes

    Args:
        shape: The shape object to extract attributes from

    Returns:
        Dictionary of attribute name to value mappings
    """
    count("shape: extractGraphicAttributes")
    attributeHash = shape.UserDefinedAttributes

    attributes = {}
    for name in attributeHash.getElementNames():
        attr_data = attributeHash.getByName(name)
        attributes[name] = attr_data.Value
    return attributes


def build_symbol_shape(ctx, model, svg_data, params, selected_shape=None):
    """Put the drawing of a symbol and its MilSym attributes onto a shape.

    With selected_shape given, its picture is replaced and its size is kept, adjusted to
    the aspect ratio of the new drawing so the content is not distorted. Otherwise a new
    graphic shape is created, sized to the drawing: the SVG is generated so that its
    intrinsic size gives the frame octagon the configured height, and decorations enlarge
    the shape beyond that. The returned shape still has to be put onto a page.
    """
    graphic = create_graphic_from_svg(ctx, svg_data)

    if selected_shape is None:
        shape = model.createInstance("com.sun.star.drawing.GraphicObjectShape")
    else:
        shape = selected_shape

    existing_size = selected_shape.getSize() if selected_shape is not None else None

    shape.setPropertyValue("Graphic", graphic)

    if existing_size is not None and existing_size.Width > 0 and existing_size.Height > 0:
        shape.setSize(
            fit_size_to_aspect_ratio(existing_size, parse_svg_dimensions(svg_data))
        )
    else:
        shape.setSize(parse_svg_dimensions(svg_data))

    insertGraphicAttributes(shape, params)
    return shape


def insertSvgGraphic(
    ctx, model, svg_data, params, selected_shape, smybol_name
):
    is_writer = model.supportsService("com.sun.star.text.TextDocument")
    is_calc = model.supportsService("com.sun.star.sheet.SpreadsheetDocument")
    is_draw_impress = model.supportsService(
        "com.sun.star.presentation.PresentationDocument"
    ) or model.supportsService("com.sun.star.drawing.DrawingDocument")

    try:
        shape = build_symbol_shape(ctx, model, svg_data, params, selected_shape)

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

        mark_document_modified(model)
    except Exception as e:
        print(f"Error inserting SVG graphic: {e}")


def mark_document_modified(model):
    """Record on the document that it has unsaved changes.

    Replacing the picture of a drawing shape, or changing its user defined attributes,
    changes only the drawing layer. In Writer and Calc the drawing layer keeps its own
    changed flag and does not pass it on to the document, so the save prompt on close
    and the modified indicator stay silent unless the document is told explicitly.
    """
    try:
        model.setModified(True)
    except Exception as e:
        print(f"Error marking the document as modified: {e}")


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


def symbol_script_args(attributes, size):
    """Build the argument list a milsymbol script invocation takes, from the MilSym
    attributes recorded on a shape.

    The first argument is the SIDC code and the second is the drawing size. Every other
    attribute whose name starts with MilSym becomes a named option: the prefix is
    stripped and the first letter of the rest is lowered, so MilSymFillColor turns into
    the option fillColor. The recorded MilSymSize is left out, because the size argument
    already says how large the drawing is made.
    """
    args = [attributes.get("MilSymCode", ""), NamedValue("size", size)]
    for name, value in attributes.items():
        if name in ("MilSymCode", "MilSymSize") or not name.startswith("MilSym"):
            continue
        option = name[6:]
        args.append(NamedValue(option[0].lower() + option[1:], value))
    return args


# Drawings already made in this process, keyed by the attributes and size they were made
# from. Each miss runs the whole milsymbol script, which is over a megabyte of
# JavaScript that the office reads, compiles and evaluates from the beginning every
# time, so a symbol that appears many times in one order of battle is worth keeping.
#
# A symbol whose attributes change gets a different key, so a drawing never goes stale
# and nothing has to be told to drop it.
_icon_svg_cache = {}

# How many drawings to keep before starting again, so that a long editing session does
# not hold on to every symbol that was ever shown.
ICON_CACHE_LIMIT = 500


def icon_cache_key(attributes, size):
    """The attributes and size that together decide the drawing of a symbol.

    Every MilSym attribute except the recorded MilSymSize takes part, since each one
    can change the drawing. The pairs are sorted, so two shapes carrying the same
    attributes in a different order share one key.
    """
    return (size,) + tuple(
        sorted(
            (name, value)
            for name, value in attributes.items()
            if name.startswith("MilSym") and name != "MilSymSize"
        )
    )


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

        cache_key = icon_cache_key(attributes, size)
        if cache_key in _icon_svg_cache:
            count("javascript: milsymbol drawing reused")
            return _icon_svg_cache[cache_key]

        args = symbol_script_args(attributes, size)

        count("javascript: milsymbol invoke")
        result = script.invoke(args, (), ())
        svg_data = str(result[0])
        if len(_icon_svg_cache) >= ICON_CACHE_LIMIT:
            _icon_svg_cache.clear()
        _icon_svg_cache[cache_key] = svg_data
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
        pipe = ctx.ServiceManager.createInstanceWithContext("com.sun.star.io.Pipe", ctx)
        pipe.writeBytes(uno.ByteSequence(svg_data.encode("utf-8")))
        pipe.flush()
        pipe.closeOutput()

        # Create graphic provider
        graphic_provider = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.graphic.GraphicProvider", ctx
        )

        # Create media properties for the SVG data
        media_properties = (PropertyValue("InputStream", 0, pipe, 0),)

        # Query the graphic from the provider
        graphic = graphic_provider.queryGraphic(media_properties)
        return graphic

    except Exception as e:
        print(f"Error creating graphic from SVG: {e}")
        return None
    
def extract_symbol_params_from_shape(ctx, model, shape):
    """Extract symbol parameters from a selected shape for sidebar insertion.
    
    Converts a shape's MilSym attributes into the format required by
    sidebar_panel.insert_symbol_node(). The category name is derived from
    the SIDC code by looking up the symbol set in the symbols_data.
    
    Args:
        ctx: LibreOffice component context
        model: Current document model
        shape: The shape object to extract parameters from
    
    Returns:
        Tuple of (category_name, svg_data, svg_args, is_editing) or (None, None, None, None) on error
    """
    try:        
        from data import symbols_data
        from translator import Translator
    
        # Extract attributes from shape
        attributes = extractGraphicAttributes(shape)
        
        if not attributes or "MilSymCode" not in attributes:
            print("Shape has no MilSym attributes")
            return None, None, None, None
        
        # Get script instance
        script = createMilSymbolScriptInstance(ctx, model)

        sidc_code = attributes.get("MilSymCode", "")

        # Build svg_args parameter list with all attributes, at the sidebar preview size
        svg_args = symbol_script_args(attributes, 20.0)

        # Generate SVG with ALL parameters
        try:
            svg_data = str(script.invoke(svg_args, (), ())[0])
        except Exception as e:
            print(f"Failed to generate SVG data: {e}")
            return None, None, None, None
        
        if not svg_data:
            print("Generated SVG is empty")
            return None, None, None, None
        
        # Determine category name from SIDC code (symbol set)
        # SIDC format: positions 4-5 contain the symbol set
        symbol_set = sidc_code[4:6] if len(sidc_code) >= 6 else ""
        
        # Look up the symbol set in symbols_data to get the translated category name
        category_name = None
        
        if symbol_set:
            translator = Translator(ctx)
            symbol_meta = next(
                (item for item in symbols_data.SYMBOLS if item["value"] == symbol_set),
                None
            )
            if symbol_meta:
                # Get translated category name (like in symbol_dialog_handler.py)
                category_name = translator.translate(symbol_meta["label"])
        
        # Return parameters
        return category_name, svg_data, svg_args, False
        
    except Exception as e:
        print(f"Error extracting symbol parameters from shape: {e}")
        return None, None, None, None


def get_package_location(ctx, extensionName="com.collabora.milsymbol"):
    """Get package location from package information provider"""
    srv = ctx.getByName(
        "/singletons/com.sun.star.deployment.PackageInformationProvider"
    )
    return srv.getPackageLocation(extensionName)


def getExtensionBasePath(ctx, extensionName="com.collabora.milsymbol"):
    """Get the base path of the extension installation directory"""
    return os.path.basename(get_package_location(ctx, extensionName))


def createMilSymbolScriptInstance(ctx, model):
    """Create an instance of the MilSymbol.js scripting library"""
    factory = ctx.getServiceManager().createInstanceWithContext(
        "com.sun.star.script.provider.MasterScriptProviderFactory", ctx
    )
    provider = factory.createScriptProvider(model)

    package_path = get_package_location(ctx, "com.collabora.milsymbol")
    location = ""
    if "share/uno_packages" in package_path:
        location = "share:uno_packages/"
    else:
        location = "user:uno_packages/"

    return provider.getScript(
        "vnd.sun.star.script:milsymbol.milsymbol.js?language=JavaScript&location="
        + location
        + os.path.basename(package_path)
    )
