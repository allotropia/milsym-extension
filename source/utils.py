# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import os
import xml.etree.ElementTree as ET
from contextlib import contextmanager

import uno

from com.sun.star.awt import Point, Size
from com.sun.star.beans import NamedValue, PropertyValue
from com.sun.star.xml import AttributeData

from perf import count

# Conversion factor from pixels to 1/100mm, assuming 96 DPI (2540 / 96)
PX_TO_MM100 = 26.46


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
    factor = PX_TO_MM100

    try:
        # Parse SVG using ElementTree
        root = ET.fromstring(svg_data)

        # Extract width and height attributes
        width_str = root.get("width")
        height_str = root.get("height")

        if width_str:
            # Remove units like 'px', 'pt', etc. and extract numeric value
            width_num = "".join(c for c in width_str if c.isdigit() or c == ".")
            if width_num:
                width = float(width_num) * factor

        if height_str:
            # Remove units like 'px', 'pt', etc. and extract numeric value
            height_num = "".join(c for c in height_str if c.isdigit() or c == ".")
            if height_num:
                height = float(height_num) * factor
    except Exception as e:
        print(f"Warning: Could not parse SVG dimensions, using defaults: {e}")

    shape_size = Size()
    shape_size.Height = height
    shape_size.Width = width
    return shape_size


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


def insertSvgGraphic(
    ctx, model, svg_data, params, selected_shape, smybol_name
):
    is_writer = model.supportsService("com.sun.star.text.TextDocument")
    is_calc = model.supportsService("com.sun.star.sheet.SpreadsheetDocument")
    is_draw_impress = model.supportsService(
        "com.sun.star.presentation.PresentationDocument"
    ) or model.supportsService("com.sun.star.drawing.DrawingDocument")

    try:
        graphic = create_graphic_from_svg(ctx, svg_data)

        # For Writer, create a TextGraphicObject which behaves better (keeps aspect ratio, etc.)
        if selected_shape is None:
            shape = model.createInstance("com.sun.star.drawing.GraphicObjectShape")
        else:
            shape = selected_shape

        # Preserve user's custom size when editing an existing shape
        existing_size = selected_shape.getSize() if selected_shape is not None else None

        shape.setPropertyValue("Graphic", graphic)

        if existing_size is not None and existing_size.Width > 0 and existing_size.Height > 0:
            # Keep the user's size, adjusted to the aspect ratio of the new graphic so
            # the content is not distorted
            shape.setSize(
                fit_size_to_aspect_ratio(existing_size, parse_svg_dimensions(svg_data))
            )
        else:
            # The SVG is generated so that its intrinsic size gives the frame octagon the
            # configured height. Decorations enlarge the shape beyond that.
            shape.setSize(parse_svg_dimensions(svg_data))

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


# The attributes that decide what a symbol looks like. Two symbols that agree on all
# of them, drawn at the same size, produce the same drawing.
ICON_ATTRIBUTES = (
    "MilSymCode",
    "MilSymStack",
    "MilSymReinforced",
    "MilSymStaff",
    "MilSymSpecialheadquarters",
    "MilSymCountrycode",
)

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
    """The attributes and size that together decide the drawing of a symbol."""
    return (size,) + tuple(attributes.get(name) for name in ICON_ATTRIBUTES)


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
        
        # Build svg_args parameter list with ALL attributes
        svg_args = []
        
        # First element: SIDC code
        sidc_code = attributes.get("MilSymCode", "")
        svg_args.append(sidc_code)
        
        # Add size for sidebar preview
        svg_args.append(NamedValue("size", 20.0))
        
        # Add all other MilSym* attributes as NamedValue objects
        for key, value in attributes.items():
            if key in ("MilSymCode", "MilSymSize"):
                continue  # Already handled
            if key.startswith("MilSym"):
                # e.g. convert "MilSymStack" -> "stack"
                attr_name = key[6:]  # Remove "MilSym" prefix
                attr_name = attr_name[0].lower() + attr_name[1:]  # Lowercase first letter
                
                nv = NamedValue()
                nv.Name = attr_name
                nv.Value = value
                svg_args.append(nv)
        
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
