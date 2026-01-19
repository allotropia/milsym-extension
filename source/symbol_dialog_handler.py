# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import os
import re
import sys
import uno
import unohelper
from com.sun.star.awt import XDialogEventHandler
from com.sun.star.awt.ImageScaleMode import ISOTROPIC
from com.sun.star.beans import NamedValue

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from data import symbols_data
from data import country_data
from utils import insertSvgGraphic, insertGraphicAttributes, getExtensionBasePath, create_graphic_from_svg
from translator import Translator
from com.sun.star.view.SelectionType import SINGLE
from com.sun.star.awt import XMouseListener, XFocusListener, XKeyListener
from com.sun.star.awt.Key import UP, DOWN, LEFT, RIGHT, RETURN
from collections import defaultdict

class SymbolDialogHandler(unohelper.Base, XDialogEventHandler):
    TREES_CACHE = {}

    def __init__(self, ctx, model, controller, dialog, sidebar_panel, selected_shape, selected_node_value):
        self.ctx = ctx
        self.model = model
        self.controller = controller
        self.dialog = dialog
        self.sidebar_panel = sidebar_panel
        self.sidc_options = {}
        self.tree_values = {}
        self.ui_indexes = {}
        self.ignore_event = False
        self.is_editing = False
        self.color = None
        self.hex_color = None
        self.final_svg_data = None
        self.final_svg_args = None
        self.sidebar_symbol_svg_data = None
        self.tree_category_name = None
        self.current_symbolSet_index = 4
        self.active_tree_ctrl = None
        self.symbol_id = None
        self.search_index = None
        self.selected_node_value = selected_node_value
        self.selected_shape = selected_shape
        self.translator = Translator(self.ctx)
        self.factory = self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", self.ctx)
        self.provider = self.factory.createScriptProvider(model)
        self.script = self.provider.getScript(
            "vnd.sun.star.script:milsymbol.milsymbol.js?language=JavaScript&location=user:uno_packages/" +
            getExtensionBasePath(self.ctx))

    def init_dialog_controls(self):
        self.init_textboxes()
        self.dialog.Model.Step = 1
        self.init_tree_controls()

        has_selected_item = False
        if self.selected_shape or self.selected_node_value:
            has_selected_item = self.init_selected_shape_params(self.dialog, self.selected_shape, self.selected_node_value)

        if not has_selected_item:
            self.init_buttons()
            self.init_default_values()
            self.init_base_preview()

    def init_textboxes(self):
        # Mapping of dialog textbox control names to their corresponding option names
        self.textbox_map = {
            "tbSpecialHeadquart":   "specialHeadquarters",
            "tbUnitNameUniqDesign": "uniqueDesignation",
            "tbHigherFormation":    "higherFormation",
            "tbAdditionalInfo":     "additionalInformation",
            "tbAltitudeDepth":      "altitudeDepth",
            "tbCombatEffect":       "combatEffectiveness",
            "tbCommonIdentifier":   "commonIdentifier",
            "tbDateTimeGroup":      "dtg",
            "tbEngageBarText":      "engagementBar",
            "tbEquipTeardownTime":  "equipmentTeardownTime",
            "tbEvaluatRating":      "evaluationRating",
            "tbGuardedUnit":        "guardedUnit",
            "tbIFF_SIF_AIS":        "iffSif",
            "tbLocation":           "location",
            "tbPlatformType":       "platformType",
            "tbQuantity":           "quantity",
            "tbSpecialDesign":      "specialDesignator",
            "tbSpeed":              "speed",
            "tbStaffComments":      "staffComments",
            "tbType":               "type",
            "tbDirection":          "direction"
        }
        self.reverse_textbox_map = {value: key for key, value in self.textbox_map.items()}

        tbSearch_ctrl = self.dialog.getControl("tbSearch")
        self.search_placeholder = tbSearch_ctrl.Text

        search_textbox_focus_listener = SearchTextBoxFocusListener(self)
        tbSearch_ctrl.addFocusListener(search_textbox_focus_listener)

        tbSearch_key_listener = SearchTextboxKeyListener(self, self.ctx)
        tbSearch_ctrl.addKeyListener(tbSearch_key_listener)

    def init_buttons(self, is_reset=False):
        if not is_reset:
            self.dialog.getControl("btReality").getModel().State = 1
            self.dialog.getControl("btFriend").getModel().State = 1
            self.context = symbols_data.BUTTONS["CONTEXT"]["btReality"]
            self.affiliation= symbols_data.BUTTONS["AFFILIATION"]["btFriend"]

        self.dialog.getControl("btPresent").getModel().State = 1
        self.dialog.getControl("btNotApplicableReinReduc").getModel().State = 1
        self.dialog.getControl("btStack1").getModel().State = 1
        self.dialog.getControl("btLight").getModel().State = 1
        self.dialog.getControl("btTarget").getModel().State = 1
        self.dialog.getControl("btNotApplicableSignature").getModel().State = 1

        self.version = symbols_data.VERSION
        self.status = symbols_data.BUTTONS["STATUS"]["btPresent"]
        self.reinforced = symbols_data.BUTTONS["REINFORCED_REDUCED"]["btNotApplicableReinReduc"]
        self.stack = symbols_data.BUTTONS["STACK"]["btStack1"]
        self.color = symbols_data.BUTTONS["COLOR"]["btLight"]
        self.signature = symbols_data.BUTTONS["SIGNATURE"]["btNotApplicableSignature"]
        self.engagement = symbols_data.BUTTONS["ENGAGEMENT"]["btTarget"]

    def tree_mapping(self):
        tree_names = (
            "treeSymbolSet",
            "treeMainIcon",
            "treeFirstIcon",
            "treeSecondIcon",
            "treeHeadTaskDummy",
            "treeEchelonMobility",
            "treeCountry",
            "treeSearch"
        )

        self.tree_map = {
            name: name[4].lower() + name[5:]
            for name in tree_names
        }

        self.tree_ctrls = {
            name: self.dialog.getControl(name)
            for name in tree_names
        }

    def init_tree_controls(self):
        self.tree_mapping()

        treeSearch_ctrl = self.tree_ctrls["treeSearch"]
        treeSearch_ctrl.setVisible(False)

        treeSearch_mosuse_listener = SearchTreeMouseListener(self)
        treeSearch_ctrl.addMouseListener(treeSearch_mosuse_listener)

        for name, tree_ctrl in self.tree_ctrls.items():
            if name == "treeSearch":
                continue

            tree_ctrl.setVisible(False)

            suffix = name.removeprefix("tree")
            listbox_ctrl = self.dialog.getControl(f"ltb{suffix}")
            ltb_mouse_listener = ListboxMouseListener(self, tree_ctrl)
            listbox_ctrl.addMouseListener(ltb_mouse_listener)

            tree_mosuse_listener = TreeMouseListener(self, listbox_ctrl)
            tree_ctrl.addMouseListener(tree_mosuse_listener)

            tree_key_listener = TreeKeyListener(self, listbox_ctrl)
            tree_ctrl.addKeyListener(tree_key_listener)

    def init_default_values(self, selected_index = 4, update_country = True):
        self.init_default_tree(update_country)

        label = self.translator.translate(symbols_data.SYMBOLS[selected_index]["label"])
        listbox_control = self.dialog.getControl("ltbSymbolSet")
        listbox_control.addItems([label], 0)
        listbox_control.selectItemPos(0, True)
        self.symbolSet = symbols_data.SYMBOLS[selected_index]["value"]

        current_symbol = self.get_current_symbol(selected_index)

        label = self.translator.translate(current_symbol["MainIcon"][0]["label"])
        listbox_control = self.dialog.getControl("ltbMainIcon")
        listbox_control.addItems([label], 0)
        listbox_control.selectItemPos(0, True)
        self.mainIcon = current_symbol["MainIcon"][0]["value"]

        label = self.translator.translate(current_symbol["FirstIconModifier"][0]["label"])
        listbox_control = self.dialog.getControl("ltbFirstIcon")
        listbox_control.addItems([label], 0)
        listbox_control.selectItemPos(0, True)
        self.firstIcon = current_symbol["FirstIconModifier"][0]["value"]

        label = self.translator.translate(current_symbol["SecondIconModifier"][0]["label"])
        listbox_control = self.dialog.getControl("ltbSecondIcon")
        listbox_control.addItems([label], 0)
        listbox_control.selectItemPos(0, True)
        self.secondIcon = current_symbol["SecondIconModifier"][0]["value"]

        label = self.translator.translate(current_symbol["EchelonMobility"][0]["label"])
        listbox_control = self.dialog.getControl("ltbEchelonMobility")
        listbox_control.addItems([label], 0)
        listbox_control.selectItemPos(0, True)
        self.echelonMobility = current_symbol["EchelonMobility"][0]["value"]

        label = self.translator.translate(current_symbol["HeadquartersTaskforceDummy"][0]["label"])
        listbox_control = self.dialog.getControl("ltbHeadTaskDummy")
        listbox_control.addItems([label], 0)
        listbox_control.selectItemPos(0, True)
        self.headTaskDummy = current_symbol["HeadquartersTaskforceDummy"][0]["value"]

        if update_country:
            label = self.translator.translate(country_data.COUNTRY_CODES[0]["label"])
            listbox_control = self.dialog.getControl("ltbCountry")
            listbox_control.addItems([label], 0)
            listbox_control.selectItemPos(0, True)
            self.country = country_data.COUNTRY_CODES[0]["value"]

    def init_default_tree(self, update_country):
        for tree_ctrl in self.tree_ctrls.values():
            data_model = tree_ctrl.getModel().DataModel
            if data_model:
                if (tree_ctrl.getModel().Name == "treeCountry"
                    and not update_country
                ):
                    continue

                root_node = tree_ctrl.getModel().DataModel.getRoot()
                selected_node = root_node.getChildAt(0)
                tree_ctrl.select(selected_node)

    def populate_symbolSet(self, selected_index):
        self.fill_tree_control(
            "treeSymbolSet", "ltbSymbolSet", symbols_data.SYMBOLS, selected_index
        )

    def populate_mainIcon(self, current_symbol, selected_index):
        self.fill_tree_control(
            "treeMainIcon", "ltbMainIcon", current_symbol["MainIcon"], selected_index
        )

    def populate_firstIcon(self, current_symbol, selected_index):
        self.fill_tree_control(
            "treeFirstIcon", "ltbFirstIcon", current_symbol["FirstIconModifier"], selected_index
        )

    def populate_secondIcon(self, current_symbol, selected_index):
        self.fill_tree_control(
            "treeSecondIcon", "ltbSecondIcon", current_symbol["SecondIconModifier"], selected_index
        )

    def populate_echelonMobility(self, current_symbol, selected_index):
        self.fill_tree_control(
            "treeEchelonMobility", "ltbEchelonMobility", current_symbol["EchelonMobility"], selected_index
        )

    def populate_headTaskDummy(self, current_symbol, selected_index):
        self.fill_tree_control(
            "treeHeadTaskDummy", "ltbHeadTaskDummy", current_symbol["HeadquartersTaskforceDummy"], selected_index
        )

    def populate_country(self, selected_index):
        self.fill_tree_control(
            "treeCountry", "ltbCountry", country_data.COUNTRY_CODES, selected_index
        )

    def fill_tree_control(self, tree_name, listbox_name, items, selected_index):
        root_node = None
        tree_control = self.tree_ctrls[tree_name]
        tree_model = tree_control.getModel()
        if listbox_name != "ltbCountry":
            tree_model.setPropertyValue("RowHeight", 30)

        symbolset_index = self.current_symbolSet_index
        if symbolset_index not in SymbolDialogHandler.TREES_CACHE:
            SymbolDialogHandler.TREES_CACHE[symbolset_index] = {}

        cached_model = SymbolDialogHandler.TREES_CACHE[symbolset_index].get(tree_name)
        if cached_model:
            tree_model.setPropertyValue("DataModel", cached_model)
        else:
            smgr = self.ctx.ServiceManager
            mutable_tree_data_model = smgr.createInstanceWithContext(
                "com.sun.star.awt.tree.MutableTreeDataModel", self.ctx
            )

            tree_model.setPropertyValue("SelectionType", SINGLE)
            tree_model.setPropertyValue("RootDisplayed", False)
            tree_model.setPropertyValue("ShowsHandles", False)
            tree_model.setPropertyValue("ShowsRootHandles", False)
            tree_model.setPropertyValue("Editable", False)

            root_node = mutable_tree_data_model.createNode("root_node", False)
            mutable_tree_data_model.setRoot(root_node)

            if listbox_name == "ltbCountry":
                BASE_ICON_URL = "vnd.sun.star.extension://com.collabora.milsymbol/img/preview/countries"
            elif listbox_name == "ltbSymbolSet":
                BASE_ICON_URL = "vnd.sun.star.extension://com.collabora.milsymbol/img/preview/symbol_set"
            else:
                category = self.symbol_id.lower()
                sub_category = re.sub(r'(?<!^)(?=[A-Z])', '_', tree_name[4:]).lower()
                BASE_ICON_URL = f"vnd.sun.star.extension://com.collabora.milsymbol/img/preview/"f"{category}/{sub_category}"

            for idx, item in enumerate(items):
                img_file = item.get("img")
                if listbox_name == "ltbCountry":
                    img_file = item.get("value") + ".png"
                icon_url = f"{BASE_ICON_URL}/{img_file}"

                label = self.translator.translate(item["label"])
                node = mutable_tree_data_model.createNode(label, False)
                node.DataValue = idx
                node.setCollapsedGraphicURL(icon_url)
                root_node.appendChild(node)

            tree_model.setPropertyValue("DataModel", mutable_tree_data_model)
            SymbolDialogHandler.TREES_CACHE[symbolset_index][tree_name] = mutable_tree_data_model

        if root_node is None:
            root_node = tree_control.getModel().DataModel.getRoot()

        selected_node = root_node.getChildAt(selected_index)
        node_name = selected_node.getDisplayValue()
        tree_control.select(selected_node)

        listbox_control = self.dialog.getControl(listbox_name)
        listbox_control.addItems([node_name], 0)
        listbox_control.selectItemPos(0, True)

        self.tree_values[tree_name] = [item["value"] for item in items]

    def handle_search_tree_node_click(self, node_name, category):
        tbSearch_ctrl = self.dialog.getControl("tbSearch")
        tbSearch_ctrl.Text = self.translator.translate(node_name)

        groups = symbols_data.SYMBOL_DETAILS.get(category, {})
        labels = groups.get("MainIcon", [])

        self.search_index = next(
            (
                index
                for index, item in enumerate(labels)
                if isinstance(item, dict)
                and self.translator.translate(item.get("label")) == node_name
            ),
            None
        )

        if self.search_index is not None:
            keys = list(symbols_data.SYMBOL_DETAILS.keys())
            symbolSet_index = keys.index(category)
            self.current_symbolSet_index = symbolSet_index
            index = self.search_index

            self.reset_symbol(self.dialog, symbolSet_index)

            current_symbol = self.get_current_symbol(symbolSet_index)
            label = self.translator.translate(current_symbol["MainIcon"][index]["label"])
            listbox_control = self.dialog.getControl("ltbMainIcon")
            listbox_control.addItems([label], 0)
            listbox_control.selectItemPos(0, True)
            self.mainIcon = current_symbol["MainIcon"][index]["value"]

            self.updatePreview()

    def apply_tree_selection(self, node, tree_ctrl, listbox_ctrl):
        tree_ctrl.select(node)
        tree_ctrl.setVisible(False)
        listbox_ctrl.getPeer().setFocus()

        node_name = node.getDisplayValue()
        label = self.translator.translate(node_name)

        count = listbox_ctrl.ItemCount
        if count > 0:
            listbox_ctrl.removeItems(0, count)

        listbox_ctrl.addItems([label], 0)
        listbox_ctrl.selectItemPos(0, True)

        selected_index = node.DataValue
        control_name = tree_ctrl.getModel().Name
        values = self.tree_values[control_name]
        value = values[selected_index]

        if control_name == "treeSymbolSet":
            self.symbolSet = value
            self.init_default_values(selected_index, update_country=False)
            self.current_symbolSet_index = selected_index
            self.ui_indexes.clear()
        elif control_name == "treeMainIcon":
            self.mainIcon = value
        elif control_name == "treeFirstIcon":
            self.firstIcon = value
        elif control_name == "treeSecondIcon":
            self.secondIcon = value
        elif control_name == "treeEchelonMobility":
            self.echelonMobility = value
        elif control_name == "treeHeadTaskDummy":
            self.headTaskDummy = value
        elif control_name == "treeCountry":
            self.country = value

        self.updatePreview()

    def get_current_symbol(self, selected_index):
        symbol_meta = symbols_data.SYMBOLS[selected_index]
        self.symbol_id = symbol_meta["id"]
        current_symbol = symbols_data.SYMBOL_DETAILS[self.symbol_id]
        self.tree_category_name=self.translator.translate(symbol_meta["label"])
        return current_symbol

    def update_ui_state(self):
        if self.active_tree_ctrl:
            self.active_tree_ctrl.setVisible(False)

        tbSearch_ctrl = self.dialog.getControl("tbSearch")
        if tbSearch_ctrl.getPeer().hasFocus():
            tbSearch_ctrl.Text = self.search_placeholder
            treeSearch_ctrl = self.tree_ctrls["treeSearch"]
            treeSearch_ctrl.getPeer().setFocus()

    def callHandlerMethod(self, dialog, eventObject, methodName):
        # Generic help button handler - handles any button with id "btHelp<FieldName>"
        if methodName.startswith("btHelp"):
            field_name = methodName[6:]
            self.show_help_dialog(field_name)
            return True
        elif methodName == "btCustom":
            hex_color = self.pick_custom_color()
            if hex_color:
                self.hex_color = hex_color
                self.button_handler(dialog, methodName)
            return True
        elif methodName.startswith("focus"): # textboxes
            if self.active_tree_ctrl:
                self.active_tree_ctrl.setVisible(False)
            return True
        elif methodName.startswith("tb"): # textboxes
            self.textbox_handler(dialog, methodName)
            return True
        elif methodName.startswith("tabbed"):
            self.update_ui_state()
            self.tabbed_button_switch_handler(dialog, methodName)
            return True
        elif methodName.startswith("bt"):
            self.update_ui_state()
            self.button_handler(dialog, methodName)
            return True
        elif methodName == "click_dialog":
            self.update_ui_state()
            return True
        elif methodName == "dialog_btSave":
            if not self.final_svg_args:
                self.sidc = "130310000000000000000000000000"
                self.updatePreview()

            if self.controller is not None:
                shape = self.controller.get_diagram().set_svg_data(
                    self.final_svg_data)
                insertGraphicAttributes(shape,
                                        self.final_svg_args)
            elif self.sidebar_panel is not None:
                self.sidebar_panel.insert_symbol_node(self.tree_category_name, self.sidebar_symbol_svg_data,
                                                      self.final_svg_args, self.is_editing)
            else: # document
                insertSvgGraphic(
                    self.ctx, self.model,
                    self.final_svg_data,
                    self.final_svg_args,
                    self.selected_shape,
                    "Symbol " + "(" + self.sidc + ")")
            dialog.endExecute()
            return True
        elif methodName == "dialog_btCancel":
            dialog.endExecute()
            return True
        elif methodName == "dialog_btReset":
            self.reset_symbol(dialog)
            return True
        else:
            return False

    def getSupportedMethodNames(self):
        return self.buttons

    def disposing(self, event):
        pass

    def reset_symbol(self, dialog, selected_index = None):
        self.ignore_event = True
        for textbox, option_name in self.textbox_map.items():
            dialog.getControl(textbox).Text = ""
            self.sidc_options[option_name] = ""
        self.ignore_event = False

        if selected_index is not None:
            index = selected_index
        else:
            symbolSet_item = next((item for item in symbols_data.SYMBOLS
                                   if item["value"] == self.symbolSet),None)
            index = symbols_data.SYMBOLS.index(symbolSet_item)
            self.search_index = None

        self.init_default_values(index)
        self.init_buttons(True)
        self.update_buttons_state(dialog)
        self.ui_indexes.clear()

        self.updatePreview()

    def pick_custom_color(self, init_color=2938211):
        smgr = self.ctx.getServiceManager()
        color_picker = smgr.createInstanceWithContext(
            "com.sun.star.ui.dialogs.ColorPicker", self.ctx
        )
        prop = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        prop.Name = "Color"
        prop.Value = int(init_color)
        color_picker.initialize((prop,))
        color_picker.setPropertyValues((prop,))

        result = color_picker.execute()
        if result == 1:
            props = color_picker.getPropertyValues()
            color = props[0].Value
            return self.color_to_hex(color)
        return None

    def color_to_hex(self, color_val):
        red   = (color_val >> 16) & 0xFF
        green = (color_val >> 8) & 0xFF
        blue  = color_val & 0xFF
        return f"#{red:02X}{green:02X}{blue:02X}"

    def show_help_dialog(self, field_name):
        """Display a message box with help text for the specified field.

        Args:
            field_name: The field name (without "btHelp" prefix), used to construct
                        translation keys Help.<field_name>.Title and Help.<field_name>.Message
        """
        from com.sun.star.awt.MessageBoxType import INFOBOX
        from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK

        title_key = f"Help.{field_name}.Title"
        message_key = f"Help.{field_name}.Message"

        toolkit = self.ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.awt.Toolkit", self.ctx
        )
        parent_window = self.dialog.getPeer()
        msgbox = toolkit.createMessageBox(
            parent_window,
            INFOBOX,
            BUTTONS_OK,
            self.translator.translate(title_key),
            self.translator.translate(message_key),
        )
        msgbox.execute()

    def textbox_handler(self, dialog, methodName):
        if getattr(self, "ignore_event", False):
            return

        options_name = self.textbox_map.get(methodName)
        text = dialog.getControl(methodName).Text

        if text:
            self.sidc_options[options_name] = text
        else:
            self.sidc_options.pop(options_name, None) # remove

        self.updatePreview()

    def tabbed_button_switch_handler(self, dialog, methodName):
        # Handling tabbed page buttons
        if methodName == "tabbed_btBasic":
            dialog.Model.Step = 1
            for tree_ctrl in self.tree_ctrls.values():
                tree_ctrl.setVisible(False)

        elif methodName == "tabbed_btAdvance":
            dialog.Model.Step = 2
            self.tree_ctrls["treeCountry"].setVisible(False)

    def button_handler(self, dialog, active_button_id, updatePreview = True):
        group_name = None
        group_buttons = None

        for name, buttons in symbols_data.BUTTONS.items():
            if active_button_id in buttons:
                group_name = name
                group_buttons = buttons
                break

        if group_buttons is None:
            return False

        for button_id in group_buttons:
            self.update_button(dialog, button_id, active_button_id, group_name, updatePreview)
        return True

    def update_button(self, dialog, button_id, active_button_id, group_name, updatePreview):
        state = 0
        if button_id == active_button_id:
            state = 1
            group_buttons= symbols_data.BUTTONS.get(group_name)
            value = group_buttons.get(button_id)

            if group_name == "CONTEXT":
                self.context = value
            elif group_name == "AFFILIATION":
                self.affiliation = value
            elif group_name == "STATUS":
                self.status = value
            elif group_name == "REINFORCED_REDUCED":
                self.reinforced = value
            elif group_name == "STACK":
                self.stack = value
            elif group_name == "COLOR":
                self.color = value
            elif group_name == "SIGNATURE":
                self.signature = value
            elif group_name == "ENGAGEMENT":
                self.engagement = value

            if updatePreview:
                self.updatePreview()

        dialog.getControl(button_id).getModel().State = state

    def init_base_preview(self):
        imgPreview = self.dialog.getModel().getByName("imgPreview")
        imgPreview.ScaleImage = True
        imgPreview.ScaleMode = ISOTROPIC
        imgPreview.ImageURL =  "vnd.sun.star.extension://com.collabora.milsymbol/img/base.svg"

    def updatePreview(self):
        svg_data = None
        if self.selected_node_value is not None:
            svg_data = self.get_tree_node_svg_data()
        else:
            svg_data = self.insertSymbolToPreview()

        if svg_data:
            graphic = create_graphic_from_svg(self.ctx, svg_data)

            imgPreview = self.dialog.getModel().getByName("imgPreview")
            imgPreview.ScaleImage = True
            imgPreview.ScaleMode = ISOTROPIC
            imgPreview.Graphic = graphic

    def create_sidc(self):
        sidc = [
            self.version,
            self.context,
            self.affiliation,
            self.symbolSet,
            self.status,
            self.headTaskDummy,
            self.echelonMobility,
            self.mainIcon,
            self.firstIcon[-2:],
            self.secondIcon[-2:],
            self.firstIcon[0],
            self.secondIcon[0],
            "00000",
            "000" # Country flag
        ]
        self.sidc = ''.join(sidc)

        return self.sidc

    def get_tree_node_svg_data(self):
        args = list(self.selected_node_value)
        args[1] = NamedValue("size", 150.0)
        result = self.script.invoke(args, (), ())
        svg_data = str(result[0]) if result else None
        self.selected_node_value = None
        return svg_data

    def insertSymbolToPreview(self):
        sidc_code = self.create_sidc()

        args = [
            sidc_code,
            NamedValue("size", 150.0),
            NamedValue("stack", self.stack),
            NamedValue("reinforced", self.reinforced),
            NamedValue("signature", self.signature),
            NamedValue("engagementType", self.engagement)
        ]

        if self.country:
            args.extend([
                NamedValue("country", self.country),
                NamedValue("country_flag", "true")
            ])

        if self.color == "Custom":
            args.append(NamedValue("fillColor", self.hex_color))
        elif self.color== "NoFill":
            args.append(NamedValue("fill", "false"))
        else:
            args.append(NamedValue("colorMode", self.color))

        for key, value in self.sidc_options.items():
            if value:
                args.append(NamedValue(key, value))

        try:
            result = self.script.invoke(args, (), ())

            if self.sidebar_panel is not None:
                args[1] = NamedValue("size", 20.0)
                self.sidebar_symbol_svg_data  = str(self.script.invoke(args, (), ())[0])

            # Assuming the result contains SVG data
            if result and len(result) > 0:
                svg_data = str(result[0])
                self.final_svg_data = svg_data
                self.final_svg_args = args
                return svg_data
        except Exception as e:
            print(f"Error executing script: {e}")
            return

    def get_textbox_name(self, name):
        return name[6:][0].lower() + name[6:][1:]

    def init_selected_shape_params(self, dialog, shape, tree_node_value):
        attrs = self.get_attrs(shape, tree_node_value)

        if not attrs:
            return False

        self.is_editing = True

        for element, value in attrs.items():
            name = self.get_textbox_name(element)
            textbox = self.reverse_textbox_map.get(name)
            if textbox:
                self.ignore_event = True
                self.sidc_options[name] = value
                self.dialog.getControl(textbox).Text = value
                self.ignore_event = False
            else:
                if element   == "MilSymStack":          self.stack      = value
                elif element == "MilSymReinforced":     self.reinforced = value
                elif element == "MilSymColorMode":      self.color      = value
                elif element == "MilSymSignature":      self.signature  = value
                elif element == "MilSymEngagementType": self.engagement = value
                elif element == "MilSymFillColor":      self.hex_color  = value

        if not self.color:
            if self.hex_color:
                self.color = "Custom"
            else:
                self.color = "NoFill"

        self.update_tree_controls()
        self.update_buttons_state(dialog)

        self.updatePreview()

        return True

    def update_tree_controls(self):
        symbolSet_index, symbolSet_label = self.find_index_and_label(
            symbols_data.SYMBOLS, self.symbolSet
        )
        label = self.translator.translate(symbolSet_label)
        listbox_control = self.dialog.getControl("ltbSymbolSet")
        listbox_control.addItems([label], 0)
        listbox_control.selectItemPos(0, True)

        self.current_symbolSet_index = symbolSet_index
        current_symbol = self.get_current_symbol(symbolSet_index)

        mainIcon_index, mainIcon_label = self.find_index_and_label(
            current_symbol["MainIcon"], self.mainIcon
        )
        self.ui_indexes["treeMainIcon"] = mainIcon_index
        label = self.translator.translate(mainIcon_label)
        listbox_control = self.dialog.getControl("ltbMainIcon")
        listbox_control.addItems([label], 0)
        listbox_control.selectItemPos(0, True)

        firstIcon_index, firstIcon_label = self.find_index_and_label(
            current_symbol["FirstIconModifier"], self.firstIcon
        )
        self.ui_indexes["treeFirstIcon"] = firstIcon_index
        label = self.translator.translate(firstIcon_label)
        listbox_control = self.dialog.getControl("ltbFirstIcon")
        listbox_control.addItems([label], 0)
        listbox_control.selectItemPos(0, True)

        secondIcon_index, secondIcon_label = self.find_index_and_label(
            current_symbol["SecondIconModifier"], self.secondIcon
        )
        self.ui_indexes["treeSecondIcon"] = secondIcon_index
        label = self.translator.translate(secondIcon_label)
        listbox_control = self.dialog.getControl("ltbSecondIcon")
        listbox_control.addItems([label], 0)
        listbox_control.selectItemPos(0, True)

        echelonMobility_index, echelonMobility_label = self.find_index_and_label(
            current_symbol["EchelonMobility"], self.echelonMobility
        )
        self.ui_indexes["treeEchelonMobility"] = echelonMobility_index
        label = self.translator.translate(echelonMobility_label)
        listbox_control = self.dialog.getControl("ltbEchelonMobility")
        listbox_control.addItems([label], 0)
        listbox_control.selectItemPos(0, True)

        headTaskDummy_index, headTaskDummy_label = self.find_index_and_label(
            current_symbol["HeadquartersTaskforceDummy"], self.headTaskDummy
        )
        self.ui_indexes["treeHeadTaskDummy"] = headTaskDummy_index
        label = self.translator.translate(headTaskDummy_label)
        listbox_control = self.dialog.getControl("ltbHeadTaskDummy")
        listbox_control.addItems([label], 0)
        listbox_control.selectItemPos(0, True)

        country_index, country_label = self.find_index_and_label(
            country_data.COUNTRY_CODES, self.country
        )
        self.ui_indexes["treeCountry"] = country_index
        label = self.translator.translate(country_label)
        listbox_control = self.dialog.getControl("ltbCountry")
        listbox_control.addItems([label], 0)
        listbox_control.selectItemPos(0, True)

    def find_index_and_label(self, items, value):
        return next(
            ((i, item["label"]) for i, item in enumerate(items) if item["value"] == value),
            (0, None)
        )

    def update_buttons_state(self, dialog):
        self.set_button_state(dialog, self.stack,       "STACK")
        self.set_button_state(dialog, self.reinforced,  "REINFORCED_REDUCED")
        self.set_button_state(dialog, self.signature,   "SIGNATURE")
        self.set_button_state(dialog, self.engagement,  "ENGAGEMENT")
        self.set_button_state(dialog, self.color,       "COLOR")
        self.set_button_state(dialog, self.context,     "CONTEXT")
        self.set_button_state(dialog, self.affiliation, "AFFILIATION")
        self.set_button_state(dialog, self.status,      "STATUS")

    def set_button_state(self, dialog, option, group_button):
        group_buttons = symbols_data.BUTTONS.get(group_button, {})
        for key, val in group_buttons.items():
            if val == option:
                self.button_handler(dialog, key, False)
                break

    def get_attrs(self, shape, tree_node_value):
        attrs = {}
        other_attrs = {}
        self.country = ""

        if shape is not None:
            user_attrs = shape.getPropertyValue("UserDefinedAttributes")
            if hasattr(user_attrs, "getElementNames"):
                for element in user_attrs.getElementNames():
                    attr = user_attrs.getByName(element)
                    other_attrs[element] = str(attr.Value)
        else:
            other_attrs["MilSymCode"] = str(tree_node_value[0])
            for entry in tree_node_value[1:]:
                element = "MilSym" + entry.Name[0].upper() + entry.Name[1:]
                other_attrs[element] = str(entry.Value)

        for element, value in other_attrs.items():
            if element == "MilSymCode":
                self.version = value[0:2]
                self.context= value[2]
                self.affiliation = value[3]
                self.symbolSet = value[4:6]
                self.status = value[6]
                self.headTaskDummy = value[7]
                self.echelonMobility = value[8:10]
                self.mainIcon = value[10:16]
                self.firstIcon = value[20] + value[16:18]
                self.secondIcon = value[21] + value[18:20]
                # others value[23:]
            elif element == "MilSymCountry":
                self.country = value
            else:
                attrs[element] = value

        return attrs

class SearchTreeMouseListener(unohelper.Base, XMouseListener):
    def __init__(self, dialog_handler):
        self.dialog_handler = dialog_handler
        self.tbSearch_ctrl = dialog_handler.dialog.getControl("tbSearch")
        self.pressed_node = None

    def mousePressed(self, event):
        x, y = event.X, event.Y
        node = event.Source.getNodeForLocation(x, y)
        if not node:
            return

        self.pressed_node = node
        event.Source.select(node)
        event.Source.makeNodeVisible(node)

    def mouseReleased(self, event):
        if not self.pressed_node:
            return

        node_name = self.pressed_node.getDisplayValue()
        category = self.pressed_node.DataValue[1]
        self.dialog_handler.handle_search_tree_node_click(node_name, category)
        self.tbSearch_ctrl.Text = node_name

        event.Source.setVisible(False)

    def mouseEntered(self, event):
        pass

    def mouseExited(self, event):
        pass

    def disposing(self, event):
        pass


class SearchTextboxKeyListener(unohelper.Base, XKeyListener):
    cached_prefix_index = None

    def __init__(self, dialog_handler, ctx):
        self.ctx = ctx
        self.dialog_handler = dialog_handler
        self.treeSearch_ctrl = dialog_handler.dialog.getControl("treeSearch")
        self.index = None

        self.tree_model = self.treeSearch_ctrl.getModel()
        self.tree_model.setPropertyValue("SelectionType", SINGLE)
        self.tree_model.setPropertyValue("RootDisplayed", False)
        self.tree_model.setPropertyValue("ShowsHandles", False)
        self.tree_model.setPropertyValue("ShowsRootHandles", False)
        self.tree_model.setPropertyValue("Editable", False)
        self.tree_model.setPropertyValue("RowHeight", 30)

    def ensure_search_index(self):
        if SearchTextboxKeyListener.cached_prefix_index is None:
            SearchTextboxKeyListener.cached_prefix_index = self.build_token_index()

        self.index = SearchTextboxKeyListener.cached_prefix_index

    def build_token_index(self):
        index = defaultdict(set)
        TOKEN_SPLIT = re.compile(r'[ /-]+').split
        translate = self.dialog_handler.translator.translate

        for category_name, data in symbols_data.SYMBOL_DETAILS.items():
            for icon in data.get("MainIcon", []):
                raw = icon.get("label", "")
                label = translate(raw)
                img = icon.get("img", "")
                main_part = label.split(" - ", 1)[0].lower()
                for token in TOKEN_SPLIT(main_part):
                    if token:
                        index[token].add((label, img, category_name))

        return index

    def keyPressed(self, event):
        if event.KeyCode in (UP, DOWN, LEFT, RIGHT):
            self.handle_tree_navigation(event)
            return

        if event.KeyCode == RETURN:
            node = self.treeSearch_ctrl.getSelection()
            if not node:
                return
            self.treeSearch_ctrl.getPeer().setFocus()
            node_name = node.getDisplayValue()
            category = node.DataValue[1]
            self.dialog_handler.handle_search_tree_node_click(node_name, category)

            self.treeSearch_ctrl.setVisible(False)

    def keyReleased(self, event):
        if event.KeyCode in (UP, DOWN, LEFT, RIGHT, RETURN):
            return

        text = event.Source.getText().strip().lower()
        if not text:
            self.treeSearch_ctrl.setVisible(False)
            return

        matches = self.run_search(text)
        self.rebuild_tree(matches)
        self.treeSearch_ctrl.setVisible(bool(matches))

    def handle_tree_navigation(self, event):
        root = self.tree_model.DataModel.getRoot()
        selection = self.treeSearch_ctrl.getSelection()

        if not selection and root.getChildCount() > 0:
            selection = root.getChildAt(0)
            self.treeSearch_ctrl.select(selection)
            return

        idx = selection.DataValue[0]

        if event.KeyCode == DOWN and idx + 1 < root.getChildCount():
            node = root.getChildAt(idx + 1)
            self.treeSearch_ctrl.select(node)
            self.treeSearch_ctrl.makeNodeVisible(node)

        elif event.KeyCode == UP and idx > 0:
            node = root.getChildAt(idx - 1)
            self.treeSearch_ctrl.select(node)
            self.treeSearch_ctrl.makeNodeVisible(node)

    def run_search(self, text):
        self.ensure_search_index()

        words = text.lower().split()
        if not words:
            return []

        matches = set()
        first_word = words[0]

        for token, symbol_tuples in self.index.items():
            if token.startswith(first_word):
                matches |= symbol_tuples

        if not matches:
            return []

        for i, word in enumerate(words[1:], start=1):
            tmp_matches = set()
            for label, img, category in matches:
                parts = label.lower().split()
                for start in range(len(parts)):
                    if parts[start].startswith(first_word):
                        pos = start + i
                        if pos < len(parts) and parts[pos].startswith(word):
                            tmp_matches.add((label, img, category))
                        break
            matches = tmp_matches
            if not matches:
                return []

        return list(matches)

    def rebuild_tree(self, items):
        mutable_tree_data_model = self.ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.awt.tree.MutableTreeDataModel", self.ctx)

        root_node = mutable_tree_data_model.createNode("root_node", False)
        mutable_tree_data_model.setRoot(root_node)

        BASE_ICON_URL = "vnd.sun.star.extension://com.collabora.milsymbol/img/preview"

        for idx, (label, img, category) in enumerate(items):
            node = mutable_tree_data_model.createNode(label, False)
            node.DataValue = (idx, category)
            category = category.replace(" - ", "_").replace(" ", "_").lower()
            icon_url = f"{BASE_ICON_URL}/{category}/main_icon/{img}"
            node.setCollapsedGraphicURL(icon_url)
            root_node.appendChild(node)

        self.tree_model.setPropertyValue("DataModel", mutable_tree_data_model)

        self.tree_model.Height = min(root_node.getChildCount() * 14, 120)

        self.dialog_handler.active_tree_ctrl = self.treeSearch_ctrl

    def disposing(self, event):
        pass


class SearchTextBoxFocusListener(unohelper.Base, XFocusListener):
    def __init__(self, dialog_handler):
        self.dialog_handler = dialog_handler
        self.tree_ctrl = dialog_handler.dialog.getControl("treeSearch")

    def focusGained(self, event):
        event.Source.Text = ""

        if self.dialog_handler.active_tree_ctrl:
            self.dialog_handler.active_tree_ctrl.setVisible(False)

    def focusLost(self, event):
        if not self.tree_ctrl.getPeer().hasFocus():
            event.Source.Text = self.dialog_handler.search_placeholder

    def disposing(self, event):
        pass


class ListboxMouseListener(unohelper.Base, XMouseListener):
    def __init__(self, dialog_handler, tree_ctrl):
        self.dialog_handler = dialog_handler
        self.tree_ctrl = tree_ctrl

    def mousePressed(self, event):
        control_name = self.tree_ctrl.getModel().Name
        symbolSet_index = self.dialog_handler.current_symbolSet_index

        if control_name == "treeSymbolSet":
            self.dialog_handler.populate_symbolSet(symbolSet_index)
        else:
            selected_index = self.get_selected_index(control_name)

            current_symbol = self.dialog_handler.get_current_symbol(symbolSet_index)
            if control_name == "treeMainIcon":
                self.dialog_handler.populate_mainIcon(current_symbol, selected_index)
            elif control_name == "treeFirstIcon":
                self.dialog_handler.populate_firstIcon(current_symbol, selected_index)
            elif control_name == "treeSecondIcon":
                self.dialog_handler.populate_secondIcon(current_symbol, selected_index)
            elif control_name == "treeEchelonMobility":
                self.dialog_handler.populate_echelonMobility(current_symbol, selected_index)
            elif control_name == "treeHeadTaskDummy":
                self.dialog_handler.populate_headTaskDummy(current_symbol, selected_index)
            elif control_name == "treeCountry":
                self.dialog_handler.populate_country(selected_index)

        self.tree_ctrl.getPeer().setFocus()
        self.tree_ctrl.setVisible(not self.tree_ctrl.isVisible())

        if (self.dialog_handler.active_tree_ctrl
            and self.dialog_handler.active_tree_ctrl != self.tree_ctrl
        ):
            self.dialog_handler.active_tree_ctrl.setVisible(False)

    def get_selected_index(self, control_name):
        if (control_name == "treeMainIcon"
            and self.dialog_handler.search_index is not None
        ):
            idx = self.dialog_handler.search_index
            self.dialog_handler.search_index = None
            return idx

        node = self.tree_ctrl.getSelection()
        if node:
            return node.DataValue

        return self.dialog_handler.ui_indexes.get(control_name, 0)

    def mouseReleased(self, event):
        node = self.tree_ctrl.getSelection()
        self.tree_ctrl.makeNodeVisible(node)
        self.dialog_handler.active_tree_ctrl = self.tree_ctrl

    def mouseEntered(self, event):
        pass

    def mouseExited(self, event):
        pass

    def disposing(self, event):
        pass


class TreeKeyListener(unohelper.Base, XKeyListener):
    def __init__(self, dialog_handler, current_listbox):
        self.dialog_handler = dialog_handler
        self.current_listbox = current_listbox

    def keyPressed(self, event):
        if event.KeyCode == RETURN:
            node = event.Source.getSelection()
            if not node:
                return

            self.dialog_handler.apply_tree_selection(
                node, event.Source, self.current_listbox
            )

    def keyReleased(self, event): pass

    def disposing(self, event):
        pass


class TreeMouseListener(unohelper.Base, XMouseListener):
    def __init__(self, dialog_handler, current_listbox):
        self.dialog_handler = dialog_handler
        self.current_listbox = current_listbox
        self.pressed_node = None

    def mousePressed(self, event):
        x, y = event.X, event.Y
        node = event.Source.getNodeForLocation(x, y)
        if not node:
            return

        event.Source.makeNodeVisible(node)
        self.pressed_node = node

    def mouseReleased(self, event):
        if not self.pressed_node:
            return

        self.dialog_handler.apply_tree_selection(
            self.pressed_node, event.Source, self.current_listbox
        )

    def mouseEntered(self, event):
        pass

    def mouseExited(self, event):
        pass

    def disposing(self, event):
        pass
