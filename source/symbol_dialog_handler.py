# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import os
import sys
import uno
import unohelper
from com.sun.star.awt import XDialogEventHandler
from com.sun.star.awt.ImageScaleMode import ISOTROPIC
from com.sun.star.beans import NamedValue
from xml.etree.ElementPath import prepare_self

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from data import symbols_data
from data import country_data
from data.help_content import HELP_CONTENT
from utils import insertSvgGraphic, insertGraphicAttributes, getExtensionBasePath, create_graphic_from_svg
from translator import Translator

class SymbolDialogHandler(unohelper.Base, XDialogEventHandler):

    def __init__(self, ctx, model, controller, dialog, sidebar_panel, selected_shape, selected_node_value):
        self.ctx = ctx
        self.model = model
        self.controller = controller
        self.dialog = dialog
        self.sidebar_panel = sidebar_panel
        self.sidc_options = {}
        self.listbox_values = {}
        self.disable_callHandler = False
        self.is_editing = False
        self.color = None
        self.hex_color = None
        self.final_svg_data = None
        self.final_svg_args = None
        self.sidebar_symbol_svg_data = None
        self.tree_category_name = None
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

        has_selected_item = False
        if self.selected_shape or self.selected_node_value:
            has_selected_item = self.init_selected_shape_params(self.dialog, self.selected_shape, self.selected_node_value)

        if not has_selected_item:
            self.init_listbox(self.dialog)
            self.init_buttons()
            self.updatePreview()

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

    def init_listbox(self, dialog, selected_index = 4):
        self.populate_symbol_listboxes(dialog, selected_index)

        self.listbox_map = {
            "ltbMainIcon":        "mainIcon",
            "ltbFirstIcon":       "firstIcon",
            "ltbSecondIcon":      "secondIcon",
            "ltbHeadTaskDummy":   "headquartersTaskforceDummy",
            "ltbEchelonMobility": "echelonMobility",
            "ltbCountry":         "country"
        }

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

    def get_current_symbol(self, selected_index):
        symbol_meta = symbols_data.SYMBOLS[selected_index]
        symbol_id = symbol_meta["id"]
        current_symbol = symbols_data.SYMBOL_DETAILS[symbol_id]
        self.tree_category_name=self.translator.translate(symbol_meta["label"])
        return current_symbol

    def fill_listbox(self, dialog, control_name, items, selected_index):
        labels = [self.translator.translate(item["label"]) for item in items]

        listbox = dialog.getControl(control_name)
        listbox.removeItems(0, listbox.getItemCount())
        listbox.addItems(labels, 0)
        self.disable_callHandler = True
        listbox.selectItemPos(selected_index, True)
        self.disable_callHandler = False
        listbox.getModel().LineCount = 12

        self.listbox_values[control_name] = [item["value"] for item in items]
        return items[selected_index]["value"]

    def callHandlerMethod(self, dialog, eventObject, methodName):
        if getattr(self, "disable_callHandler", False):
            return

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
        elif methodName.startswith("ltb"): # listboxes
            self.listbox_handler(dialog, eventObject, methodName)
            return True
        elif methodName.startswith("tb"): # textboxes
            self.textbox_handler(dialog, methodName)
            return True
        elif methodName.startswith(("click_tb", "click_ltb")):
            selected_value = dialog.getControl("tbSearch").Text
            self.update_listbox_value(dialog, selected_value)
        elif methodName.startswith("search"):
            self.search_box_handler(dialog, methodName)
            return True
        elif methodName.startswith("tabbed"):
            self.tabbed_button_switch_handler(dialog, methodName)
            return True
        elif methodName == "action_ltbSearch":
            selected_index = eventObject.Source.getSelectedItemPos()
            selected_value = dialog.getControl("ltbSearch").getItem(selected_index)
            self.update_listbox_value(dialog, selected_value)
            return True
        elif methodName == "click_Search":
            self.apply_search_selection(dialog, eventObject)
            return True
        elif methodName == "click_dialog":
            selected_value = dialog.getControl("tbSearch").Text
            self.update_listbox_value(dialog, selected_value)
            return True
        elif self.button_handler(dialog, methodName):
            return True
        elif methodName == "dialog_btSave":
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

    def reset_symbol(self, dialog):
        self.disable_callHandler = True
        for textbox, option_name in self.textbox_map.items():
                dialog.getControl(textbox).Text = ""
                self.sidc_options[option_name] = ""
        self.disable_callHandler = False

        symbolSet_item = next((item for item in symbols_data.SYMBOLS
                               if item["value"] == self.symbolSet),None)
        index = symbols_data.SYMBOLS.index(symbolSet_item)
        self.init_listbox(self.dialog, index)

        self.init_buttons(True)
        self.update_buttons_state(dialog)

        self.updatePreview()

    def listbox_handler(self, dialog, eventObject, methodName):
        if methodName == "ltbSymbolSet":
            self.update_symbolSet_listbox(dialog, eventObject)
        else:
            target_attr = self.listbox_map.get(methodName)
            selected_index = eventObject.Source.getSelectedItemPos()
            self.update_listbox_value_by_index(selected_index, methodName, target_attr)

    def update_listbox_value_by_index(self, selected_index, control_name, attr_name):
        values = self.listbox_values[control_name]
        value = values[selected_index]
        setattr(self, attr_name, value)
        self.updatePreview()

    def update_symbolSet_listbox(self, dialog, eventObject):
        selected_index = eventObject.Source.getSelectedItemPos()
        self.populate_symbol_listboxes(dialog, selected_index)
        self.updatePreview()

    def populate_symbol_listboxes(self, dialog, selected_index):
        current_symbol = self.get_current_symbol(selected_index)

        self.symbolSet =                  self.fill_listbox(dialog, "ltbSymbolSet",       symbols_data.SYMBOLS, selected_index)
        self.mainIcon =                   self.fill_listbox(dialog, "ltbMainIcon",        current_symbol["MainIcon"], 0)
        self.firstIcon =                  self.fill_listbox(dialog, "ltbFirstIcon",       current_symbol["FirstIconModifier"], 0)
        self.secondIcon =                 self.fill_listbox(dialog, "ltbSecondIcon",      current_symbol["SecondIconModifier"], 0)
        self.echelonMobility =            self.fill_listbox(dialog, "ltbEchelonMobility", current_symbol["EchelonMobility"], 0)
        self.headquartersTaskforceDummy = self.fill_listbox(dialog, "ltbHeadTaskDummy",   current_symbol["HeadquartersTaskforceDummy"], 0)
        self.country =                    self.fill_listbox(dialog, "ltbCountry",         country_data.COUNTRY_CODES, 0)

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
            field_name: The field name (without "btHelp" prefix) to look up in HELP_CONTENT
        """
        from com.sun.star.awt.MessageBoxType import INFOBOX
        from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK

        help_info = HELP_CONTENT.get(field_name)
        if not help_info:
            return

        toolkit = self.ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.awt.Toolkit", self.ctx)
        parent_window = self.dialog.getPeer()
        msgbox = toolkit.createMessageBox(
            parent_window,
            INFOBOX,
            BUTTONS_OK,
            self.translator.translate(help_info["title"]),
            self.translator.translate(help_info["message"])
        )
        msgbox.execute()

    def textbox_handler(self, dialog, methodName):
        options_name = self.textbox_map.get(methodName)
        text = dialog.getControl(methodName).Text

        if text:
            self.sidc_options[options_name] = " " + text
        else:
            self.sidc_options.pop(options_name, None) # remove

        self.updatePreview()

    # Update listbox value after search
    def update_listbox_value(self, dialog, selected_value):
        valid_value = False
        for symbolSet_index, groups in enumerate(symbols_data.SYMBOL_DETAILS.values()):
            labels = groups.get("MainIcon", [])

            mainIcon_index = next(
               (i for i, item in enumerate(labels)
                if isinstance(item, dict) and self.translator.translate(item.get("label")) == selected_value),
               None
            )

            if mainIcon_index is not None:
                dialog.getControl("ltbSymbolSet").selectItemPos(symbolSet_index, True)
                dialog.getControl("ltbMainIcon").selectItemPos(mainIcon_index, True)
                valid_value = True
                break

        if not valid_value:
            self.disable_callHandler = True
            dialog.getControl("tbSearch").Text = " Type to search..."
            self.disable_callHandler = False

        dialog.getControl("ltbSearch").setVisible(False)

    def search_box_handler(self, dialog, methodName):
        if methodName == "search_click":
            self.disable_callHandler = True
            dialog.getControl("tbSearch").Text = ""
            self.disable_callHandler = False
        elif methodName == "search_change":
            search_text = dialog.getControl("tbSearch").Text
            search_listbox = dialog.getControl("ltbSearch")
            if search_text:
                matches = []
                search = search_text.lower()

                for data in symbols_data.SYMBOL_DETAILS.values():
                    for icon in data.get("MainIcon", []):
                        label = self.translator.translate(icon.get("label", ""))
                        label_to_search = label.split("-")[0].strip()
                        label_lower = label_to_search.lower()
                        if search in label_lower:
                            matches.append(label)

                search_listbox.removeItems(0, search_listbox.getItemCount())
                for item in matches:
                    search_listbox.addItem(item, search_listbox.getItemCount())

                item_height = 10
                max_visible_items = 5
                visible_rows = min(len(matches), max_visible_items)
                search_listbox.getModel().Height = visible_rows * item_height

            if search_text and search_text.strip() != "Type to search...":
                search_listbox.setVisible(True)
            else:
                search_listbox.setVisible(False)

            search_listbox.selectItemPos(0, True)

    def search_box_listbox_handler(self, dialog, eventObject):
        selected_index = eventObject.Source.getSelectedItemPos()
        selected_value = dialog.getControl("ltbSearch").getItem(selected_index)
        symbolSet_name = None
        for symbolSet, groups in symbols_data.SYMBOL_DETAILS.items():
            labels = groups.get("MainIcon", [])

            self.mainIcon = next(
                (item.get("value") for item in labels
                 if isinstance(item, dict) and self.translator.translate(item.get("label")) == selected_value),
                None
            )

            if self.mainIcon is not None:
                symbolSet_name = symbolSet
                break

        self.symbolSet = next(
            (item.get("value") for item in symbols_data.SYMBOLS
             if item.get("id") == symbolSet_name),
            None
        )

        self.disable_callHandler = True
        self.updatePreview()
        dialog.getControl("tbSearch").Text = selected_value
        self.disable_callHandler = False

    def apply_search_selection(self, dialog, eventObject):
        search_listbox = dialog.getControl("ltbSearch")
        pos = search_listbox.getSelectedItemPos()
        if pos >= 0:
            self.search_box_listbox_handler(dialog, eventObject)

    def tabbed_button_switch_handler(self, dialog, methodName):
        # Handling tabbed page buttons
        if methodName == "tabbed_btBasic":
            dialog.Model.Step = 1
        elif methodName == "tabbed_btAdvance":
            dialog.Model.Step = 2

        selected_value = dialog.getControl("tbSearch").Text
        self.update_listbox_value(dialog, selected_value)
        dialog.getControl("ltbSearch").setVisible(False)

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
            self.headquartersTaskforceDummy,
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
        attrs, listbox_attrs = self.get_attrs(shape, tree_node_value)

        if not attrs:
            return False

        self.is_editing = True

        for element, value in attrs.items():
            name = self.get_textbox_name(element)
            textbox = self.reverse_textbox_map.get(name)
            if textbox:
                self.sidc_options[name] = " " + value
                self.disable_callHandler = True
                self.dialog.getControl(textbox).Text = value
                self.disable_callHandler = False
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

        self.update_listboxes(dialog, listbox_attrs)
        self.update_buttons_state(dialog)

        self.updatePreview()

        return True

    def update_listboxes(self, dialog, listbox_attrs):
        symbolSet_item = next(
            (item for item in symbols_data.SYMBOLS
             if item["value"] == self.symbolSet),None)

        index = symbols_data.SYMBOLS.index(symbolSet_item)
        self.init_listbox(self.dialog, index)

        for attr, value in listbox_attrs.items():
            setattr(self, attr, value)

        symbol_id = self.translator.translate(symbolSet_item["id"])
        current_symbol = symbols_data.SYMBOL_DETAILS[symbol_id]

        mainIcon_item = next(
            (item for item in current_symbol["MainIcon"]
             if item["value"] == self.mainIcon), None)
        self.set_selected_shapes_listbox_item(dialog, mainIcon_item, "ltbMainIcon")

        firstIcon_item = next(
            (item for item in current_symbol["FirstIconModifier"]
             if item["value"] == self.firstIcon), None)
        self.set_selected_shapes_listbox_item(dialog, firstIcon_item, "ltbFirstIcon")

        secondIcon_item = next(
            (item for item in current_symbol["SecondIconModifier"]
             if item["value"] == self.secondIcon), None)
        self.set_selected_shapes_listbox_item(dialog, secondIcon_item, "ltbSecondIcon")

        echelonMobility_item = next(
            (item for item in current_symbol["EchelonMobility"]
             if item["value"] == self.echelonMobility), None)
        self.set_selected_shapes_listbox_item(dialog, echelonMobility_item, "ltbEchelonMobility")

        headquartersTaskforceDummy_item = next(
            (item for item in current_symbol["HeadquartersTaskforceDummy"]
             if item["value"] == self.headquartersTaskforceDummy), None)
        self.set_selected_shapes_listbox_item(dialog, headquartersTaskforceDummy_item, "ltbHeadTaskDummy")

    def set_selected_shapes_listbox_item(self, dialog, item, listbox_control):
        if item:
            target_label = self.translator.translate(item["label"])
            listbox = dialog.getControl(listbox_control)
            labels = listbox.getItems()
            if target_label in labels:
                index = labels.index(target_label)
                self.disable_callHandler = True
                listbox.selectItemPos(index, True)
                self.disable_callHandler = False

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
        listbox_attrs = {}
        other_attrs = {}

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
                listbox_attrs["headquartersTaskforceDummy"] = value[7]
                listbox_attrs["echelonMobility"] = value[8:10]
                listbox_attrs["mainIcon"] = value[10:16]
                listbox_attrs["firstIcon"] = value[16:18]
                listbox_attrs["secondIcon"] = value[18:20]
                # others value[20:]
            else:
                attrs[element] = value

        return attrs, listbox_attrs
