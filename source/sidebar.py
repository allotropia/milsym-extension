# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import unohelper
from com.sun.star.ui import XUIElement, XUIElementFactory
from com.sun.star.ui import XToolPanel, XSidebarPanel, LayoutSize
from com.sun.star.awt import XWindowPeer, XWindowListener

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

        self._resizeListener = WindowResizeListener(self.onResize)
        self.xParentWindow.addWindowListener(self._resizeListener)

    # XUIElement
    def getRealInterface(self):
        if self.toolpanel is None:
            self.toolpanel = self.getOrCreatePanelRootWindow()
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
