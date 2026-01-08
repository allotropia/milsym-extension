# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import os
import uno
import json
import base64
import unohelper

from sidebar_tree import SidebarTree, TreeKeyListener, TreeMouseListener, TreeSelectionChangeListener
from symbol_dialog import open_symbol_dialog
from utils import get_package_location, parse_svg_dimensions
from sidebar_rename_dialog import RenameDialog

from unohelper import systemPathToFileUrl, fileUrlToSystemPath
from com.sun.star.ui.dialogs.TemplateDescription import FILESAVE_AUTOEXTENSION, FILEOPEN_SIMPLE
from com.sun.star.ui import XUIElement, XUIElementFactory, XToolPanel, XSidebarPanel, LayoutSize
from com.sun.star.awt import XWindowListener, XActionListener
from com.sun.star.awt import XFocusListener, XKeyListener
from com.sun.star.view.SelectionType import SINGLE
from com.sun.star.datatransfer.dnd import XDragGestureListener, XDragSourceListener
from com.sun.star.datatransfer import XTransferable, DataFlavor
from com.sun.star.datatransfer.dnd.DNDConstants import ACTION_COPY

class SidebarFactory(unohelper.Base, XUIElementFactory):
    def __init__(self, ctx):
        self.ctx = ctx

    def createUIElement(self, url, properties):
        xParentWindow = None

        for prop in properties:
            if prop.Name == "ParentWindow":
                xParentWindow = prop.Value

        try:
            xUIElement = SidebarPanel(self.ctx, xParentWindow, url)
            xUIElement.getRealInterface()
            panelWin = xUIElement.Window
            panelWin.Visible = True
            return xUIElement
        except Exception as e:
            print("Sidebar factory error:", e)

class SidebarPanel(unohelper.Base, XSidebarPanel, XUIElement, XToolPanel):

    # POS_X + POS_Y + POS_WIDTH + POS_HEIGHT = 1 + 2 + 4 + 8 = 15
    POS_ALL = 15

    MIN_WIDTH = 250
    NEW_BUTTON_WIDTH = 60
    BUTTON_WIDTH = 35
    BUTTON_HEIGHT = 30
    TEXTBOX_HEIGHT = 28
    VERTICAL_SPACING = 6
    LEFT_MARGIN = TOP_MARGIN = RIGHT_MARGIN = BOTTOM_MARGIN = 6

    def __init__(self, ctx, xParentWindow, url):
        self.ctx = ctx
        self.xParentWindow = xParentWindow
        self.ResourceURL = url
        self.toolpanel = None
        self.root_node = None
        self.tree_control = None
        self.mutable_tree_data_model = None
        self.selected_node = None
        self.selected_node_name = None
        self.removed_nodes = []

        self.sidebar_tree = SidebarTree(ctx, self)

        self.favorites_dir_path = self.get_favorites_dir_path(ctx)

        self.desktop = self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)

        self._resizeListener = WindowResizeListener(self.onResize)
        self.xParentWindow.addWindowListener(self._resizeListener)

    # XUIElement
    def getRealInterface(self):
        if self.toolpanel is None:
            self.toolpanel = self.getOrCreatePanelRootWindow()
            self.init_favorites_sidebar()
            self.update_export_button_state()
        return self

    # XToolPanel
    def createAccessible(self, parent):
        return self

    @property
    def Window(self):
        return self.toolpanel

    # XSidebarPanel
    def getHeightForWidth(self, width):
        return LayoutSize(0, -1, 0)

    def getMinimalWidth(self):
        return self.MIN_WIDTH

    def getOrCreatePanelRootWindow(self):
        try:
            sm = self.ctx.ServiceManager
            toolkit = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)

            container = sm.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", self.ctx)
            container_model = sm.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", self.ctx)
            container.setModel(container_model)
            container.createPeer(toolkit, self.xParentWindow)

            # New button
            x = self.LEFT_MARGIN
            y = self.TOP_MARGIN
            width = self.NEW_BUTTON_WIDTH
            height = self.BUTTON_HEIGHT
            names = ("Name", "Label")
            values = ("btNew", "New")
            btNew = self.createControl(self.ctx, "com.sun.star.awt.UnoControlButton", "com.sun.star.awt.UnoControlButtonModel", x, y, width, height, names, values)
            listener = NewButtonListener(self.ctx, self)
            btNew.addActionListener(listener)

            # Import button
            x = 0  # X position will be set later in onResize()
            y = self.TOP_MARGIN
            width = self.BUTTON_WIDTH
            height = self.BUTTON_HEIGHT
            names = ("Name", "Label")
            values = ("btImport", "\u21E5")
            btImport = self.createControl(self.ctx, "com.sun.star.awt.UnoControlButton", "com.sun.star.awt.UnoControlButtonModel", x, y, width, height, names, values)
            btImport.getModel().setPropertyValue("HelpText", "Import symbols")

            btImport_listener = ImportButtonListener(self.ctx, self)
            btImport.addActionListener(btImport_listener)

            # Export button
            x = 0  # X position will be set later in onResize()
            y = self.TOP_MARGIN
            width = self.BUTTON_WIDTH
            height = self.BUTTON_HEIGHT
            names = ("Name", "Label")
            values = ("btExport", "\u21E4")
            btExport = self.createControl(self.ctx, "com.sun.star.awt.UnoControlButton", "com.sun.star.awt.UnoControlButtonModel", x, y, width, height, names, values)
            btExport.getModel().setPropertyValue("HelpText", "Export symbols")

            btExport_listener = ExportButtonListener(self.ctx, self)
            btExport.addActionListener(btExport_listener)

            # Filter textbox
            x = self.LEFT_MARGIN
            y = self.TOP_MARGIN + self.BUTTON_HEIGHT + self.VERTICAL_SPACING
            width = 0 # width will be set later in onResize()
            height = self.TEXTBOX_HEIGHT
            names = ("Name", "Text",)
            values = ("tbFilter", "Filtering...",)
            tbFilter = self.createControl(self.ctx, "com.sun.star.awt.UnoControlEdit", "com.sun.star.awt.UnoControlEditModel", x, y, width, height, names, values)

            tb_focus_listener = TextboxFocusListener(tbFilter)
            tbFilter.addFocusListener(tb_focus_listener)

            tb_key_listener = TextboxKeyListener(self, tbFilter)
            tbFilter.addKeyListener(tb_key_listener)

            # Tree control
            x = self.LEFT_MARGIN
            y = self.TOP_MARGIN + self.BUTTON_HEIGHT + self.VERTICAL_SPACING + self.TEXTBOX_HEIGHT + self.BOTTOM_MARGIN
            width = 0   # width will be set later in onResize()
            height = 0  # height will be set later in onResize()
            names = ("Name",)
            values = ("myTree",)
            treeCtrl = self.createControl(self.ctx, "com.sun.star.awt.tree.TreeControl", "com.sun.star.awt.tree.TreeControlModel", x, y, width, height, names, values)
            self.tree_control = treeCtrl

            mouse_listener = TreeMouseListener(self.ctx, self, self.sidebar_tree)
            treeCtrl.addMouseListener(mouse_listener)
            treeCtrl.addMouseMotionListener(mouse_listener)

            key_listener = TreeKeyListener(self, self.sidebar_tree)
            treeCtrl.addKeyListener(key_listener)

            selection_listener = TreeSelectionChangeListener(self)
            treeCtrl.addSelectionChangeListener(selection_listener)

            container.addControl("btNew", btNew)
            container.addControl("btImport", btImport)
            container.addControl("btExport", btExport)
            container.addControl("tbFilter", tbFilter)
            container.addControl("treeCtrl", treeCtrl)

            # Setup native drag and drop after adding controls to container
            # This ensures the tree control's peer is properly initialized
            drag_handler = TreeDragDropHandler(self.ctx, treeCtrl, self)

            # Add drag gesture recognizer for native drag support
            try:
                # Try to set up drag source
                try:
                    drag_gesture_recognizer = toolkit.getDragGestureRecognizer(treeCtrl.getPeer())
                    drag_gesture_recognizer.addDragGestureListener(drag_handler)
                except Exception as e:
                    print(f"Could not set up drag source: {e}")

            except Exception as e:
                print(f"Could not setup native drag support, falling back to mouse listener: {e}")

            return container
        except Exception as e:
            print("Panel window error:", e)

    def createControl(self, ctx, ctrlType, ctrlTypeModel, x, y, width, height, names, values):
        try:
            sm = ctx.ServiceManager
            ctrl = sm.createInstanceWithContext(ctrlType, ctx)
            ctrl_model = sm.createInstanceWithContext(ctrlTypeModel, ctx)
            ctrl_model.setPropertyValues(names, values)
            ctrl.setModel(ctrl_model)
            ctrl.setPosSize(x, y, width, height, 15)
            return ctrl
        except Exception as e:
            print("Control error:", e)

    def insert_symbol_node(self, category_name, svg_data, svg_args, is_editing):
        self.sidebar_tree.create_node(self.root_node, self.mutable_tree_data_model,
                                      category_name, svg_data, svg_args, is_editing,
                                      self.selected_node)

        self.update_export_button_state()

    def get_favorites_dir_path(self, ctx):
        ps = ctx.getByName("/singletons/com.sun.star.util.thePathSettings")
        user_config = ps.UserConfig
        user_profile_path = os.path.dirname(user_config)

        favorites_path_URL = os.path.join(user_profile_path, "milsymbol_favorites")
        favorites_dir_path = fileUrlToSystemPath(favorites_path_URL)

        os.makedirs(favorites_dir_path, exist_ok=True)

        return favorites_dir_path

    def import_json_data(self, file_name, category_path):
        order_indexes = []
        symbol_params = []
        symbol_name = os.path.splitext(file_name)[0]
        json_path = os.path.join(category_path, symbol_name + ".json")
        if os.path.exists(json_path):
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    json_data = json.load(f)

                if json_data:
                    sidc_value = json_data.get("sidc", "")
                    symbol_params.append(str(sidc_value))

                    for key, value in json_data.items():
                        if key in ("sidc", "order_index"):
                            if key == "order_index":
                                order_indexes.append((value, symbol_name + ".svg"))
                            continue

                        nv = uno.createUnoStruct("com.sun.star.beans.NamedValue")
                        nv.Name = key
                        nv.Value = value
                        symbol_params.append(nv)

            except Exception as e:
                print("JSON read error:", e)

        return order_indexes, symbol_params

    def get_ordered_symbols(self, category_path):
        ordered_symbols = []

        try:
            if not os.path.isdir(category_path):
                print(f"Category path is not a directory: {category_path}")
                return ordered_symbols

            file_list = sorted(os.listdir(category_path))
        except Exception as e:
            print(f"Error listing category path {category_path}: {e}")
            return ordered_symbols

        for file_name in file_list:
            if file_name.lower().endswith(".json"):
                try:
                    order_indexes, symbol_params = self.import_json_data(file_name, category_path)
                    if order_indexes:  # Check if list is not empty
                        order_index_value, svg_name = order_indexes[0]
                    else:
                        # Provide default values if no order_index found in JSON
                        order_index_value = 0  # Default order
                        svg_name = os.path.splitext(file_name)[0] + ".svg"  # Construct SVG name from JSON name
                    ordered_symbols.append((order_index_value, svg_name, symbol_params))
                except Exception as e:
                    print(f"Error processing JSON file {file_name}: {e}")
                    continue

        ordered_symbols.sort(key=lambda x: x[0])

        return ordered_symbols

    def init_favorites_sidebar(self):
        smgr = self.ctx.ServiceManager
        self.mutable_tree_data_model = smgr.createInstanceWithContext("com.sun.star.awt.tree.MutableTreeDataModel", self.ctx)

        tree_ctrl = self.tree_control
        tree_model = tree_ctrl.getModel()

        tree_model.setPropertyValue("SelectionType", SINGLE)
        tree_model.setPropertyValue("RootDisplayed", True)
        tree_model.setPropertyValue("ShowsHandles", True)
        tree_model.setPropertyValue("ShowsRootHandles", True)
        tree_model.setPropertyValue("Editable", False)

        self.root_node = self.mutable_tree_data_model.createNode("Favorites", True)
        self.mutable_tree_data_model.setRoot(self.root_node)

        for category_name in sorted(os.listdir(self.favorites_dir_path)):
            category_node = self.mutable_tree_data_model.createNode(category_name, True)
            self.root_node.appendChild(category_node)

            category_path = os.path.join(self.favorites_dir_path, category_name)
            ordered_symbols = self.get_ordered_symbols(category_path)
            for _, file_name, symbol_params in ordered_symbols:
                file_path = os.path.join(category_path, file_name)
                file_url = systemPathToFileUrl(file_path)

                symbol_node = self.mutable_tree_data_model.createNode(
                    os.path.splitext(file_name)[0],
                    False
                )

                symbol_node.DataValue = symbol_params
                symbol_node.setNodeGraphicURL(file_url)
                category_node.appendChild(symbol_node)

        tree_model.setPropertyValue("DataModel", self.mutable_tree_data_model)

        tree_ctrl.expandNode(self.root_node)
        for i in range(self.root_node.getChildCount()):
            category_node = self.root_node.getChildAt(i)
            tree_ctrl.expandNode(category_node)

    def rename_symbol(self):
        RenameDialog(self.ctx, self.selected_node, self.favorites_dir_path).run()

    def update_export_button_state(self):
        has_symbol = bool(os.listdir(self.favorites_dir_path))
        self.toolpanel.getControl("btExport").getModel().State = 0 if has_symbol else 1

    def get_symbol_path(self, symbol_name):
        category_name = self.selected_node.getParent().getDisplayValue()
        category_path = os.path.join(self.favorites_dir_path, category_name)
        path = os.path.join(category_path, symbol_name)
        return path

    def node_name_exists(self, parent_node, name: str, exclude_node=None) -> bool:
        for i in range(parent_node.getChildCount()):
            child = parent_node.getChildAt(i)
            if child == exclude_node:
                continue
            if child.getDisplayValue() == name:
                return True

        return False

    def rename_symbol_files(self):
        old_name = self.selected_node_name
        if not old_name:
            return

        node = self.selected_node
        new_name = node.getDisplayValue()
        if old_name == new_name:
            return

        exists = self.node_name_exists(node.getParent(), new_name, node)
        if exists:
            n = 1
            base_name = new_name
            new_name = f"{base_name} ({n})"

            while self.node_name_exists(node.getParent(), new_name, node):
                n += 1
                new_name = f"{base_name} ({n})"

            node.setDisplayValue(new_name)

        old_svg = self.get_symbol_path(old_name) + ".svg"
        old_json = self.get_symbol_path(old_name) + ".json"

        new_svg = self.get_symbol_path(new_name) + ".svg"
        new_json = self.get_symbol_path(new_name) + ".json"

        os.rename(old_svg, new_svg)
        os.rename(old_json, new_json)

        node.setNodeGraphicURL(systemPathToFileUrl(new_svg))
        self.selected_node_name = None

    def onResize(self, event):
        try:
            toolpanel_size = event.Source.getPosSize()
            toolpanel_width = toolpanel_size.Width
            toolpanel_height = toolpanel_size.Height

            treeCtrl = self.toolpanel.getControl("treeCtrl")
            if treeCtrl:
                rect = treeCtrl.getPosSize()
                new_treeCtrl_width = toolpanel_width - self.LEFT_MARGIN - self.RIGHT_MARGIN
                new_treeCtrl_height = toolpanel_height - self.TOP_MARGIN - self.BUTTON_HEIGHT - self.VERTICAL_SPACING \
                                      - self.TEXTBOX_HEIGHT - self.VERTICAL_SPACING - self.BOTTOM_MARGIN
                treeCtrl.setPosSize(rect.X, rect.Y , new_treeCtrl_width, new_treeCtrl_height, self.POS_ALL)

            tbFilter = self.toolpanel.getControl("tbFilter")
            if tbFilter:
                rect = tbFilter.getPosSize()
                new_tbFilter_width = toolpanel_width - self.LEFT_MARGIN - self.RIGHT_MARGIN
                tbFilter.setPosSize(rect.X, rect.Y , new_tbFilter_width, rect.Height, self.POS_ALL)

            btImport = self.toolpanel.getControl("btImport")
            if btImport:
                rect = btImport.getPosSize()
                new_btImport_x_pos = toolpanel_width - (self.BUTTON_WIDTH * 2) - self.LEFT_MARGIN
                btImport.setPosSize(new_btImport_x_pos, rect.Y , rect.Width, rect.Height, self.POS_ALL)

            btExport = self.toolpanel.getControl("btExport")
            if btExport:
                rect = btExport.getPosSize()
                new_btExport_x_pos = toolpanel_width - self.BUTTON_WIDTH - self.LEFT_MARGIN
                btExport.setPosSize(new_btExport_x_pos, rect.Y , rect.Width, rect.Height, self.POS_ALL)

        except Exception as e:
            print("Resize error:", e)

class WindowResizeListener(unohelper.Base, XWindowListener):
    def __init__(self, callback):
        self.callback = callback

    def windowResized(self, event):
        self.callback(event)

    def windowHidden(self, event): pass
    def windowMoved(self, event): pass
    def windowShown(self, event): pass
    def disposing(self, event): pass

class NewButtonListener(unohelper.Base, XActionListener):
    def __init__(self, ctx, sidebar_panel):
        self.ctx = ctx
        self.sidebar_panel = sidebar_panel

    def actionPerformed(self, event):
        model = self.sidebar_panel.desktop.getCurrentComponent()
        open_symbol_dialog(self.ctx, model, None, self.sidebar_panel, None, None)

    def disposing(self, event):
        pass

class TextboxFocusListener(unohelper.Base, XFocusListener):

    def __init__(self, textbox, placeholder="Filtering..."):
        self.textbox = textbox
        self.placeholder = placeholder

    def focusGained(self, event):
        if self.textbox.getText() == self.placeholder:
            self.textbox.setText("")

    def focusLost(self, event):
        if not self.textbox.getText():
            self.textbox.setText(self.placeholder)

class TextboxKeyListener(unohelper.Base, XKeyListener):
    def __init__(self, sidebar, textbox):
        self.textbox = textbox
        self.sidebar = sidebar
        self.tree_restored = True

    def keyPressed(self, event):
        pass

    def keyReleased(self, event):
        text = self.textbox.getText()
        self.filter_sidebar_tree(text)

    def filter_sidebar_tree(self, text: str):
        if not text:
            if self.tree_restored:
                self.sidebar.init_favorites_sidebar()
                self.sidebar.removed_nodes.clear()
                self.tree_restored = False
            return

        self.tree_restored = True
        search_text = text.lower()
        root_node = self.sidebar.root_node

        for i in  reversed(range(root_node.getChildCount())):
                parent_node = root_node.getChildAt(i)

                for j in reversed(range(parent_node.getChildCount())):
                    child_node = parent_node.getChildAt(j)
                    node_name = child_node.getDisplayValue().lower()
                    found = (node_name.startswith(search_text) or
                             any(word.startswith(search_text) for word in node_name.split()))

                    if not found:
                        self.sidebar.removed_nodes.append(child_node)
                        parent_node.removeChildByIndex(j)
                        if parent_node.getChildCount() == 0:
                            root_node.removeChildByIndex(i)

        for node in self.sidebar.removed_nodes:
            node_name = node.getDisplayValue().lower()
            matches_search = (node_name.startswith(search_text) or
                              any(word.startswith(search_text) for word in node_name.split()))

            if matches_search:
                self.sidebar.init_favorites_sidebar()
                self.sidebar.removed_nodes.clear()
                self.filter_sidebar_tree(search_text)
                break

class ImportButtonListener(unohelper.Base, XActionListener):
    def __init__(self, ctx, sidebar):
        self.ctx = ctx
        self.sidebar = sidebar

    def actionPerformed(self, event):
        try:
            file_picker = self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.ui.dialogs.FilePicker", self.ctx)
            file_picker.initialize((FILEOPEN_SIMPLE,))
            file_picker.appendFilter("JSON File", "*.json")

            if file_picker.execute() != 1:
                file_picker.dispose()
                return

            selected_file = file_picker.getFiles()[0]
            if selected_file.startswith("file:///"):
                path = uno.fileUrlToSystemPath(selected_file)
            else:
                path = selected_file

            file_picker.dispose()

            with open(path, "r", encoding="utf-8") as f:
                all_data = json.load(f)

            favorites_dir = self.sidebar.favorites_dir_path
            os.makedirs(favorites_dir, exist_ok=True)

            for category_name, symbols in all_data.items():
                category_path = os.path.join(favorites_dir, category_name)
                os.makedirs(category_path, exist_ok=True)

                for symbol_name, symbol_content in symbols.items():
                    json_path = os.path.join(category_path, f"{symbol_name}.json")
                    with open(json_path, "w", encoding="utf-8") as f:
                        json.dump(symbol_content["data"], f, indent=4, ensure_ascii=False)

                    svg_path = os.path.join(category_path, f"{symbol_name}.svg")
                    with open(svg_path, "w", encoding="utf-8") as f:
                        f.write(symbol_content["svg"])

            self.sidebar.init_favorites_sidebar()
            self.sidebar.update_export_button_state()

        except Exception as e:
            print("File opening error:", e)

    def disposing(self, event):
        pass


class TreeDragDropHandler(unohelper.Base, XDragGestureListener, XDragSourceListener):
    """Native drag and drop handler for tree control"""

    def __init__(self, ctx, tree_control, sidebar_panel):
        self.ctx = ctx
        self.tree_control = tree_control
        self.sidebar_panel = sidebar_panel

    def dragGestureRecognized(self, event):
        """Handle drag gesture recognition"""
        try:
            # Get the node at the drag start location
            node = self.tree_control.getNodeForLocation(event.DragOriginX, event.DragOriginY)
            if node and node.getChildCount() == 0:  # Only leaf nodes are draggable
                # Create transferable data
                transferable = SymbolTransferable(self.ctx, node, self.sidebar_panel.favorites_dir_path)

                # Start the drag operation
                drag_source = event.DragSource
                drag_source.startDrag(event, ACTION_COPY, 0, 0, transferable, self)
        except Exception as e:
            print(f"Error in dragGestureRecognized: {e}")

    # XDragSourceListener methods
    def dragEnter(self, event):
        """Drag entered a drop target"""
        pass

    def dragExit(self, event):
        """Drag exited a drop target"""
        pass

    def dragOver(self, event):
        """Drag is over a drop target"""
        pass

    def dragDropEnd(self, event):
        """Drag operation ended"""
        # Force exiting the drag state by clearing and restoring the selection:
        selection = self.tree_control.getSelection()
        self.tree_control.clearSelection()
        if selection:
            self.tree_control.select(selection)

        # Focus the document window and select the dropped shape
        try:
            model = self.sidebar_panel.desktop.getCurrentComponent()
            if model:
                controller = model.getCurrentController()
                if controller:
                    # For Writer documents, select the last shape
                    if model.supportsService("com.sun.star.text.TextDocument"):
                        draw_page = model.getDrawPage()
                        if draw_page and draw_page.getCount() > 0:
                            last_shape = draw_page.getByIndex(draw_page.getCount() - 1)
                            controller.select(last_shape)

                    # Focus the document window
                    frame = controller.getFrame()
                    if frame:
                        component_window = frame.getComponentWindow()
                        if component_window:
                            component_window.setFocus()
        except Exception as e:
            print(f"Error focusing document after drop: {e}")


    def dropActionChanged(self, event):
        pass

    def disposing(self, event):
        pass


class SymbolTransferable(unohelper.Base, XTransferable):
    """Transferable data for symbol drag and drop"""

    def __init__(self, ctx, node, favorites_dir_path):
        self.ctx = ctx
        self.node = node
        self.favorites_dir_path = favorites_dir_path

        # Setup data flavors
        self.data_flavor = DataFlavor()
        self.data_flavor.MimeType = "application/x-openoffice-drawing;windows_formatname=\"Drawing Format\""

    def getTransferData(self, flavor):
        """Get the transferable data"""
        if flavor.MimeType == self.data_flavor.MimeType:

            # LibreOffice expects a byte sequence containing a document with the graphic
            # Read the template document from the data folder
            package_path = fileUrlToSystemPath(get_package_location(self.ctx))
            template_path = os.path.join(package_path, "source", "data", "dragdropgraphic.fodg")
            with open(template_path, 'r', encoding='utf-8') as f:
                data_string = f.read()

            # Get SVG content from the dragged node
            svg_string = self._get_svg_content_from_node()
            # Base64 encode the SVG content
            svg_base64 = base64.b64encode(svg_string.encode('utf-8')).decode('utf-8')
            svg_size = parse_svg_dimensions(svg_string)
            factor = 5
            width_cm = svg_size.Width / 1000.0 * factor  # 1/100mm to cm
            height_cm = svg_size.Height / 1000.0 * factor # 1/100mm to cm
            data_string = data_string.replace('SVG_BASE_64_ENCODED', svg_base64)
            data_string = data_string.replace('SVG_WIDTH_CM', str(width_cm)+'cm')
            data_string = data_string.replace('SVG_HEIGHT_CM', str(height_cm)+'cm')
            data_string = data_string.replace('SYMBOL_NAME', self.node.getDisplayValue())

            style_start = data_string.find('<style:style style:name="gr1"')
            style_end = data_string.find('</style:style>', style_start)

            props_start = data_string.find('<style:graphic-properties', style_start, style_end)
            props_end = data_string.find('/>', props_start)
            props_tag = data_string[props_start:props_end]

            milsym_values = {}
            for item in self.node.DataValue[1:]:
                milsym_values[item.Name] = str(item.Value)

            attrs = []
            attrs.append(f'MilSymCode="{self.node.DataValue[0]}"')
            for name, value in milsym_values.items():
                attrs.append(f'MilSym{name[0].upper() + name[1:]}="{value}"')

            new_props_tag = props_tag + ' ' + ' '.join(attrs) + '/>'

            data_string = (
                data_string[:props_start] +
                new_props_tag +
                data_string[props_end+2:]
            )

            data_bytes = data_string.encode('utf-8')
            return uno.ByteSequence(data_bytes)
        else:
            raise uno.RuntimeException("Unsupported data flavor", self)

    def getTransferDataFlavors(self):
        """Get available data flavors"""
        return (self.data_flavor,)

    def isDataFlavorSupported(self, flavor):
        """Check if data flavor is supported"""
        return flavor.MimeType == self.data_flavor.MimeType

    def _get_svg_content_from_node(self):
        """Get SVG content from the dragged node"""
        try:
            # Get the symbol name from the node
            symbol_name = self.node.getDisplayValue()

            # Get the category name from the parent node
            parent_node = self.node.getParent()
            category_name = parent_node.getDisplayValue()

            # Construct the path to the SVG file
            category_path = os.path.join(self.favorites_dir_path, category_name)
            svg_file_path = os.path.join(category_path, f"{symbol_name}.svg")

            # Read and return the SVG content
            if os.path.exists(svg_file_path):
                with open(svg_file_path, 'r', encoding='utf-8') as f:
                    return f.read()
            else:
                print(f"SVG file not found: {svg_file_path}")
                return ""

        except Exception as e:
            print(f"Error reading SVG content: {e}")
            return ""

class ExportButtonListener(unohelper.Base, XActionListener):
    def __init__(self, ctx, sidebar):
        self.ctx = ctx
        self.sidebar = sidebar

    def actionPerformed(self, event):
        try:
            state = event.Source.getModel().State
            if state == 1:
                return

            file_picker = self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.ui.dialogs.FilePicker", self.ctx)
            file_picker.initialize((FILESAVE_AUTOEXTENSION,))
            file_picker.setDefaultName("sidebar_data.json")
            file_picker.appendFilter("JOSN File", "*.json")

            if file_picker.execute() != 1:
                file_picker.dispose()
                return

            selected_file = file_picker.getFiles()[0]
            if selected_file.startswith("file:///"):
                path = uno.fileUrlToSystemPath(selected_file)
            else:
                path = selected_file

            file_picker.dispose()

            all_data = {}
            favorites_dir = self.sidebar.favorites_dir_path
            for category_name in sorted(os.listdir(favorites_dir)):
                category_path = os.path.join(favorites_dir, category_name)
                if not os.path.isdir(category_path):
                    continue

                all_data[category_name] = {}
                for file_name in sorted(os.listdir(category_path)):
                    if file_name.endswith(".json"):
                        file_base = os.path.splitext(file_name)[0]

                        json_path = os.path.join(category_path, f"{file_base}.json")
                        with open(json_path, "r", encoding="utf-8") as f:
                            data = json.load(f)

                        svg_path = os.path.join(category_path, f"{file_base}.svg")
                        with open(svg_path, "r", encoding="utf-8") as f:
                            svg_content = f.read()

                        all_data[category_name][file_base] = {
                            "data": data,
                            "svg": svg_content
                        }

            with open(path, "w", encoding="utf-8") as f:
                json.dump(all_data, f, indent=4, ensure_ascii=False)

        except Exception as e:
            print("Save file error:", e)

    def disposing(self, event):
        pass
