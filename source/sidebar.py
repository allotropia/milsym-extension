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

from sidebar_tree import SidebarTree, TreeKeyListener, TreeMouseListener
from symbol_dialog import open_symbol_dialog
from utils import insertSvgGraphic

from unohelper import systemPathToFileUrl, fileUrlToSystemPath
from com.sun.star.beans import PropertyValue
from com.sun.star.ui import XUIElement, XUIElementFactory, XToolPanel, XSidebarPanel, LayoutSize
from com.sun.star.awt import XWindowPeer, XWindowListener, XActionListener, Size
from com.sun.star.view.SelectionType import SINGLE

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
    BUTTON_WIDTH = 60
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

        self.sidebar_tree = SidebarTree(ctx)

        self.favorites_dir_path = self.get_favorites_dir_path(ctx)

        self.desktop = self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)

        self._resizeListener = WindowResizeListener(self.onResize)
        self.xParentWindow.addWindowListener(self._resizeListener)

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
            width = self.BUTTON_WIDTH
            height = self.BUTTON_HEIGHT
            names = ("Name", "Label")
            values = ("btNew", "New")
            btNew = self.createControl(self.ctx, "com.sun.star.awt.UnoControlButton", "com.sun.star.awt.UnoControlButtonModel", x, y, width, height, names, values)
            listener = NewButtonListener(self.ctx, self)
            btNew.addActionListener(listener)

            # Import/Export button
            x = 0  # X position will be set later in onResize()
            y = self.TOP_MARGIN
            width = self.BUTTON_WIDTH
            height = self.BUTTON_HEIGHT
            names = ("Name", "Label")
            values = ("btImport", "...")
            btImport = self.createControl(self.ctx, "com.sun.star.awt.UnoControlButton", "com.sun.star.awt.UnoControlButtonModel", x, y, width, height, names, values)

            # Filter textbox
            x = self.LEFT_MARGIN
            y = self.TOP_MARGIN + self.BUTTON_HEIGHT + self.VERTICAL_SPACING
            width = 0 # width will be set later in onResize()
            height = self.TEXTBOX_HEIGHT
            names = ("Name", "Text",)
            values = ("tbFilter", "Filtering...",)
            tbFilter = self.createControl(self.ctx, "com.sun.star.awt.UnoControlEdit", "com.sun.star.awt.UnoControlEditModel", x, y, width, height, names, values)

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

            drag_handler = TreeMouseListener(self.ctx, treeCtrl, self, self.favorites_dir_path)
            treeCtrl.addMouseListener(drag_handler)
            treeCtrl.addMouseMotionListener(drag_handler)

            key_listener =  TreeKeyListener(self, self.favorites_dir_path)
            treeCtrl.addKeyListener(key_listener)

            container.addControl("btNew", btNew)
            container.addControl("btImport", btImport)
            container.addControl("tbFilter", tbFilter)
            container.addControl("treeCtrl", treeCtrl)

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

    def insert_symbol_node(self, category_name, svg_data, svg_args):
        self.sidebar_tree.create_node(self.root_node, self.mutable_tree_data_model,
                                      category_name, svg_data, svg_args)

    def get_favorites_dir_path(self, ctx):
        ps = ctx.getByName("/singletons/com.sun.star.util.thePathSettings")
        user_config = ps.UserConfig
        user_profile_path = os.path.dirname(user_config)

        favorites_path_URL = os.path.join(user_profile_path, "milsymbol_favorites")
        favorites_dir_path = fileUrlToSystemPath(favorites_path_URL)

        os.makedirs(favorites_dir_path, exist_ok=True)

        self.sidebar_tree.set_favorites_dir_path(favorites_dir_path)

        return favorites_dir_path

    def import_json_data(self, file_name, category_path):
        symbol_params = []
        symbol_name = os.path.splitext(file_name)[0]
        json_path = os.path.join(category_path, symbol_name + ".json")
        if os.path.exists(json_path):
            try:
                with open(json_path, "r", encoding="utf-8") as f:
                    json_data = json.load(f)

                if json_data:
                    sidc_value = json_data[0]
                    symbol_params.append(str(sidc_value))

                    for item in json_data[1:]:
                        if isinstance(item, dict):
                            for key, value in item.items():
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
                new_btImport_x_pos = toolpanel_width - self.BUTTON_WIDTH - self.LEFT_MARGIN
                btImport.setPosSize(new_btImport_x_pos, rect.Y , rect.Width, rect.Height, self.POS_ALL)

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
        open_symbol_dialog(self.ctx, model, None, self.sidebar_panel)

    def disposing(self, event):
        pass
