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
import unohelper

from sidebar_tree import SidebarTree, TreeKeyListener, TreeMouseListener, TreeSelectionChangeListener
from symbol_dialog import open_symbol_dialog
from utils import insertSvgGraphic
from sidebar_rename_dialog import RenameDialog

from unohelper import systemPathToFileUrl, fileUrlToSystemPath
from com.sun.star.ui.dialogs.TemplateDescription import FILESAVE_AUTOEXTENSION, FILEOPEN_SIMPLE
from com.sun.star.ui import XUIElement, XUIElementFactory, XToolPanel, XSidebarPanel, LayoutSize
from com.sun.star.awt import XWindowListener, XActionListener
from com.sun.star.awt import XFocusListener, XKeyListener
from com.sun.star.view.SelectionType import SINGLE
from com.sun.star.datatransfer.dnd import XDragGestureListener, XDragSourceListener, XDropTargetListener
from com.sun.star.datatransfer import XTransferable, DataFlavor
from com.sun.star.datatransfer.dnd.DNDConstants import ACTION_COPY
from com.sun.star.frame import XFrameActionListener

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

        self.sidebar_tree = SidebarTree(ctx)

        self.favorites_dir_path = self.get_favorites_dir_path(ctx)

        self.desktop = self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)

        self._resizeListener = WindowResizeListener(self.onResize)
        self.xParentWindow.addWindowListener(self._resizeListener)

        # Setup drop targets on document windows
        self._setup_document_drop_targets()

    # XUIElement
    def getRealInterface(self):
        if self.toolpanel is None:
            self.toolpanel = self.getOrCreatePanelRootWindow()
            self.init_favorites_sidebar()
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
            self.sidebar_tree.set_tree_control(treeCtrl)

            key_listener =  TreeKeyListener(self, self.favorites_dir_path)
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
            drag_handler = TreeDragDropHandler(self.ctx, treeCtrl, self, self.favorites_dir_path)

            # Add drag gesture recognizer for native drag support
            try:
                # Try to set up drag source
                try:
                    drag_gesture_recognizer = toolkit.getDragGestureRecognizer(treeCtrl.getPeer())
                    drag_gesture_recognizer.addDragGestureListener(drag_handler)
                except Exception as e:
                    print(f"Could not set up drag source: {e}")

                # Try to set up drop target
                """ try:
                    drop_target = toolkit.getDropTarget(peer)
                    drop_target.addDropTargetListener(drop_handler)
                    drop_target.setActive(True)
                except Exception as e:
                    print(f"Could not set up drop target: {e}") """


            except Exception as e:
                print(f"Could not setup native drag support, falling back to mouse listener: {e}")
                # Fallback to mouse listener approach
                #mouse_handler = TreeMouseListener(self.ctx, treeCtrl, self, self.favorites_dir_path)
                #treeCtrl.addMouseListener(mouse_handler)
                #treeCtrl.addMouseMotionListener(mouse_handler)

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

    def get_favorites_dir_path(self, ctx):
        ps = ctx.getByName("/singletons/com.sun.star.util.thePathSettings")
        user_config = ps.UserConfig
        user_profile_path = os.path.dirname(user_config)

        favorites_path_URL = os.path.join(user_profile_path, "milsymbol_favorites")
        favorites_dir_path = fileUrlToSystemPath(favorites_path_URL)

        os.makedirs(favorites_dir_path, exist_ok=True)

        self.sidebar_tree.set_favorites_dir_path(favorites_dir_path)

        return favorites_dir_path

    def _setup_document_drop_targets(self):
        """Setup drop targets on document windows"""
        try:
            print("setting up document drop targets")
            # Listen for frame events to setup drop targets on new documents
            frame_listener = DocumentFrameListener(self.ctx, self)
            self.desktop.addFrameActionListener(frame_listener)

            # Setup on all existing frames
            self._setup_on_existing_frames()
        except Exception as e:
            print(f"Could not setup document drop targets: {e}")

    def _setup_on_existing_frames(self):
        """Setup drop targets on all existing frames"""
        try:
            # First try the current frame
            current_frame = self.desktop.getCurrentFrame()
            print("current frame: ", current_frame)
            if current_frame:
                self._setup_drop_target_on_frame(current_frame)
                return

            # If no current frame, iterate through all frames
            frames = self.desktop.getFrames()
            print(f"Found {frames.getCount()} frames")

            for i in range(frames.getCount()):
                frame = frames.getByIndex(i)
                if frame and self._is_document_frame(frame):
                    self._setup_drop_target_on_frame(frame)

        except Exception as e:
            print(f"Could not setup on existing frames: {e}")

    def _is_document_frame(self, frame):
        """Check if frame is a document frame (not toolbar, etc.)"""
        try:
            # Check if frame has a component (document)
            component = frame.getController()
            if component:
                model = component.getModel()
                return model is not None
            return False
        except:
            return False

    def _setup_drop_target_on_frame(self, frame):
        """Setup drop target on a specific frame"""
        print("Setting up drop target on frame:", frame)
        try:
            toolkit = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)

            # Create drop target listener for the document window
            drop_listener = DocumentDropTargetListener(self.ctx, self)
            drop_target = toolkit.getDropTarget(frame.getComponentWindow())
            drop_target.addDropTargetListener(drop_listener)
            drop_target.setActive(True)
        except Exception as e:
            print(f"Could not setup drop target on frame: {e}")

    def import_json_data(self, file_name, category_path):
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
                        if key == "sidc":
                            continue

                        nv = uno.createUnoStruct("com.sun.star.beans.NamedValue")
                        nv.Name = key
                        nv.Value = value
                        symbol_params.append(nv)

            except Exception as e:
                print("JSON read error:", e)

        return symbol_params

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

        for category_name in os.listdir(self.favorites_dir_path):
            category_path = os.path.join(self.favorites_dir_path, category_name)

            category_node = self.mutable_tree_data_model.createNode(category_name, True)
            self.root_node.appendChild(category_node)

            for file_name in os.listdir(category_path):
                if file_name.lower().endswith(".svg"):
                    file_path = os.path.join(category_path, file_name)
                    file_url = systemPathToFileUrl(file_path)

                    symbol_node = self.mutable_tree_data_model.createNode(
                        os.path.splitext(file_name)[0],
                        False
                    )
                    symbol_params = self.import_json_data(file_name, category_path)
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

        except Exception as e:
            print("File opening error:", e)

    def disposing(self, event):
        pass


class TreeDragDropHandler(unohelper.Base, XDragGestureListener, XDragSourceListener):
    """Native drag and drop handler for tree control"""

    def __init__(self, ctx, tree_control, sidebar_panel, favorites_dir_path):
        self.ctx = ctx
        self.tree_control = tree_control
        self.sidebar_panel = sidebar_panel
        self.favorites_dir_path = favorites_dir_path

    def dragGestureRecognized(self, event):
        """Handle drag gesture recognition"""
        try:
            print("Drag gesture recognized")
            # Get the node at the drag start location
            node = self.tree_control.getNodeForLocation(event.DragOriginX, event.DragOriginY)
            if node and node.getChildCount() == 0:  # Only leaf nodes are draggable
                # Create transferable data
                transferable = SymbolTransferable(self.ctx, node, self.favorites_dir_path)

                # Start the drag operation
                drag_source = event.DragSource
                drag_source.startDrag(event, ACTION_COPY, 0, 0, transferable, self)  # 1 = MOVE action

                print(f"Started drag for node: {node.getDisplayValue()}")
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
        if event.DropSuccess:
            print("Drag drop completed successfully")
        else:
            print("Drag drop was cancelled")

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
        self.data_flavor.MimeType = "application/x-milsymbol-node"
        self.data_flavor.HumanPresentableName = "Military Symbol Node"
        self.data_flavor.DataType = uno.getTypeByName("string")

    def getTransferData(self, flavor):
        """Get the transferable data"""
        if flavor.MimeType == self.data_flavor.MimeType:
            # Return node data as JSON string
            node_data = {
                "displayValue": self.node.getDisplayValue(),
                "graphicURL": self.node.getNodeGraphicURL() if hasattr(self.node, 'getNodeGraphicURL') else None,
                "dataValue": self.node.DataValue if hasattr(self.node, 'DataValue') else None
            }
            return json.dumps(node_data)
        else:
            raise uno.RuntimeException("Unsupported data flavor", self)

    def getTransferDataFlavors(self):
        """Get available data flavors"""
        return (self.data_flavor,)

    def isDataFlavorSupported(self, flavor):
        """Check if data flavor is supported"""
        return flavor.MimeType == self.data_flavor.MimeType


class DocumentFrameListener(unohelper.Base, XFrameActionListener):
    """Listen for frame events to setup drop targets on new documents"""

    def __init__(self, ctx, sidebar_panel):
        print("frame listener created")
        self.ctx = ctx
        self.sidebar_panel = sidebar_panel
        self.processed_frames = set()  # Track which frames we've already processed

    def frameAction(self, event):
        """Handle frame action events"""
        try:
            from com.sun.star.frame.FrameAction import FRAME_ACTIVATED, COMPONENT_ATTACHED

            # Setup drop target when frame is activated or component is attached
            if event.Action in (FRAME_ACTIVATED, COMPONENT_ATTACHED):
                self.sidebar_panel._setup_drop_target_on_frame(event.Frame)
        except Exception as e:
            print(f"Error in frame action listener: {e}")

    def disposing(self, event):
        pass


class DocumentDropTargetListener(unohelper.Base, XDropTargetListener):
    """Handle drops on document windows"""

    def __init__(self, ctx, sidebar_panel):
        self.ctx = ctx
        self.sidebar_panel = sidebar_panel

    def drop(self, event):
        """Handle drop event with shape hit testing"""
        try:
            # Check if we can handle this drop
            transferable = event.Transferable
            data_flavors = transferable.getTransferDataFlavors()

            for flavor in data_flavors:
                if flavor.MimeType == "application/x-milsymbol-node":
                    # Get the dropped data
                    node_data_json = transferable.getTransferData(flavor)
                    node_data = json.loads(node_data_json)

                    # Insert the symbol at the drop location
                    model = self.sidebar_panel.desktop.getCurrentComponent()
                    if model:
                        svg_url = node_data.get("graphicURL")
                        if svg_url:
                            svg_path = fileUrlToSystemPath(svg_url)
                            with open(svg_path, 'r', encoding='utf-8') as f:
                                svg_data = f.read()

                            # Convert drop coordinates and check for shapes at that location
                            drop_position, target_shape = self._get_drop_position_and_target(event, model)

                            if target_shape:
                                # Update the existing target shape with the new SVG
                                self._update_shape_with_svg(target_shape, svg_data, node_data.get("dataValue", []))
                                target_name = target_shape.getName() if hasattr(target_shape, 'getName') else 'unnamed shape'
                                print(f"Updated existing shape '{target_name}' with symbol '{node_data.get('displayValue')}'")
                            else:
                                # Create new shape at drop coordinates if no target shape found
                                params = node_data.get("dataValue", [])
                                insertSvgGraphic(self.ctx, model, svg_data, params, drop_position, 3)
                                print(f"Created new symbol '{node_data.get('displayValue')}' at coordinates ({drop_position.X}, {drop_position.Y})")

                            event.acceptDrop(1)  # Accept the drop
                            return

            # Reject the drop if we can't handle it
            event.rejectDrop()

        except Exception as e:
            print(f"Error in drop handler: {e}")
            event.rejectDrop()

    def _get_drop_position_and_target(self, event, model):
        """Convert drop coordinates and detect target shape using hit testing"""
        try:
            # Create drop position point
            drop_position = uno.createUnoStruct('com.sun.star.awt.Point')
            drop_position.X = event.LocationInTarget.X * 100  # Convert to document units (1/100mm)
            drop_position.Y = event.LocationInTarget.Y * 100

            # Try to find a shape at the drop location using hit testing
            target_shape = self._find_shape_at_position(model, drop_position)

            return drop_position, target_shape

        except Exception as e:
            print(f"Error calculating drop position: {e}")
            # Fallback to basic position
            drop_position = uno.createUnoStruct('com.sun.star.awt.Point')
            drop_position.X = event.LocationInTarget.X * 100
            drop_position.Y = event.LocationInTarget.Y * 100
            return drop_position, None

    def _find_shape_at_position(self, model, position):
        """Find shape at given position using hit testing"""
        try:
            controller = model.getCurrentController()
            if not controller:
                return None

            # Get the current page/slide
            current_page = None
            if hasattr(controller, 'getCurrentPage'):
                current_page = controller.getCurrentPage()
            elif hasattr(model, 'getCurrentPage'):
                current_page = model.getCurrentPage()
            elif hasattr(model, 'getDrawPage'):
                current_page = model.getDrawPage()

            if not current_page:
                return None

            # Check each shape to see if the drop position intersects with it
            for i in range(current_page.getCount()):
                shape = current_page.getByIndex(i)
                if self._point_in_shape(position, shape):
                    return shape

            return None

        except Exception as e:
            print(f"Error in shape hit testing: {e}")
            return None

    def _point_in_shape(self, point, shape):
        """Check if a point is inside a shape's bounding rectangle"""
        try:
            shape_pos = shape.getPosition()
            shape_size = shape.getSize()

            # Check if point is within shape bounds
            if (point.X >= shape_pos.X and
                point.X <= shape_pos.X + shape_size.Width and
                point.Y >= shape_pos.Y and
                point.Y <= shape_pos.Y + shape_size.Height):
                return True

            return False

        except Exception as e:
            print(f"Error checking point in shape: {e}")
            return False

    def _update_shape_with_svg(self, target_shape, svg_data, params):
        """Update an existing shape with new SVG content"""
        try:
            # Store the original position and size
            original_pos = target_shape.getPosition()
            original_size = target_shape.getSize()

            # Get the page that contains the target shape
            parent_page = None
            if hasattr(target_shape, 'getParent'):
                parent_page = target_shape.getParent()

            if not parent_page:
                print("Could not find parent page for target shape")
                return False

            # Create a temporary SVG file for the new content
            import tempfile
            with tempfile.NamedTemporaryFile(mode='w', suffix='.svg', delete=False, encoding='utf-8') as temp_file:
                temp_file.write(svg_data)
                temp_svg_path = temp_file.name

            try:
                # Import the new SVG
                from com.sun.star.beans import PropertyValue

                # Create import properties
                import_props = []

                # Import filter for SVG
                filter_prop = PropertyValue()
                filter_prop.Name = "FilterName"
                filter_prop.Value = "draw_svg_Export"  # For import, use the export filter name
                import_props.append(filter_prop)

                # Create URL from temp file
                from unohelper import systemPathToFileUrl
                svg_url = systemPathToFileUrl(temp_svg_path)

                # Try to replace the graphic content
                if hasattr(target_shape, 'setPropertyValue'):
                    try:
                        # For GraphicObjectShape, update the GraphicURL property
                        target_shape.setPropertyValue("GraphicURL", svg_url)

                        # Try to maintain the original position and size
                        target_shape.setPosition(original_pos)
                        # Optionally maintain size or let it adapt to new content
                        # target_shape.setSize(original_size)

                        print(f"Successfully updated shape with new SVG content")
                        return True

                    except Exception as prop_e:
                        print(f"Could not update GraphicURL property: {prop_e}")

                # Alternative approach: Replace the shape entirely but preserve position
                model = self.sidebar_panel.desktop.getCurrentComponent()
                if model and parent_page:
                    # Remove the old shape
                    parent_page.remove(target_shape)

                    # Insert new shape at the same position
                    insertSvgGraphic(self.ctx, model, svg_data, params, original_pos, 3)

                    print(f"Replaced shape with new SVG content at original position")
                    return True

            finally:
                # Clean up temporary file
                import os
                try:
                    os.unlink(temp_svg_path)
                except:
                    pass

        except Exception as e:
            print(f"Error updating shape with SVG: {e}")
            return False

    def dragEnter(self, event):
        """Drag entered the drop target"""
        # Check if we can accept this drag
        print("Drag entered document drop target")
        for flavor in event.SupportedDataFlavors:
            if flavor.MimeType == "application/x-milsymbol-node":
                event.Context.acceptDrag(1)  # Accept drag with MOVE action
                return
        event.Context.rejectDrag()

    def dragExit(self, event):
        """Drag exited the drop target"""
        pass

    def dragOver(self, event):
        """Drag is over the drop target"""
        print("Drag is over the drop target")
        pass

    def dropActionChanged(self, event):
        """Drop action changed"""
        pass

    def disposing(self, event):
        pass



class ExportButtonListener(unohelper.Base, XActionListener):
    def __init__(self, ctx, sidebar):
        self.ctx = ctx
        self.sidebar = sidebar

    def actionPerformed(self, event):
        try:
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
            for category_name in os.listdir(favorites_dir):
                category_path = os.path.join(favorites_dir, category_name)
                if not os.path.isdir(category_path):
                    continue

                all_data[category_name] = {}
                for file_name in os.listdir(category_path):
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