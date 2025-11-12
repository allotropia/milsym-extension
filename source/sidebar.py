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
        return 300

    def getOrCreatePanelRootWindow(self):
        try:
            sm = self.ctx.ServiceManager
            toolkit = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctx)
            
            container = sm.createInstanceWithContext("com.sun.star.awt.UnoControlContainer", self.ctx)
            container_model = sm.createInstanceWithContext("com.sun.star.awt.UnoControlContainerModel", self.ctx)
            container.setModel(container_model)
            container.createPeer(toolkit, self.xParentWindow)

            names = ("Name", "Label")
            values = ("btNew", "New")
            btNew = self.createControl(self.ctx, "com.sun.star.awt.UnoControlButton", "com.sun.star.awt.UnoControlButtonModel",
                                                    6, 6, 75, 35, names, values)
            names = ("Name", "Label")
            values = ("btImportExport", "...")
            btImportExport = self.createControl(self.ctx, "com.sun.star.awt.UnoControlButton", "com.sun.star.awt.UnoControlButtonModel",
                                                219, 6, 75, 35, names, values)
            names = ("Name", "Text",)
            values = ("tbFilter", "",)
            tbFilter = self.createControl(self.ctx, "com.sun.star.awt.UnoControlEdit", "com.sun.star.awt.UnoControlEditModel",
                                          6, 47, 0, 35, names, values)
            names = ("Name",)
            values = ("myTree",)
            treeCtrl = self.createControl(self.ctx, "com.sun.star.awt.tree.TreeControl", "com.sun.star.awt.tree.TreeControlModel",
                                          6, 88, 0, 0, names, values)

            container.addControl("btNew", btNew)
            container.addControl("btImportExport", btImportExport)
            container.addControl("tbFilter", tbFilter)
            container.addControl("treeCtrl", treeCtrl)

            print("Coainter:", container)
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
            toolpanel_size = self.toolpanel.getPosSize()
            toolpanel_width = toolpanel_size.Width
            toolpanel_height = toolpanel_size.Height
            
            treeCtrl = self.toolpanel.getControl("treeCtrl")
            if treeCtrl:
                rect = treeCtrl.getPosSize()
                treeCtrl.setPosSize(rect.X, rect.Y , toolpanel_width - 12, toolpanel_height - 94, 15)
                
            tbFilter = self.toolpanel.getControl("tbFilter")
            if tbFilter:
                rect = tbFilter.getPosSize()
                tbFilter.setPosSize(rect.X, rect.Y , toolpanel_width - 12, rect.Height, 15)
                
            btImportExport = self.toolpanel.getControl("btImportExport")
            if btImportExport:
                rect = btImportExport.getPosSize()
                x = toolpanel_width - 75 - 6
                btImportExport.setPosSize(x, rect.Y , rect.Width, rect.Height, 15)

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
