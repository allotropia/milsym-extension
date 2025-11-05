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
import tempfile
from unohelper import systemPathToFileUrl
from com.sun.star.awt import XDialogEventHandler
from com.sun.star.awt import Size, Point
from com.sun.star.beans import NamedValue, PropertyValue

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from data import symbols_data
from data import country_data
from utils import insertSvgGraphic, insertGraphicAttributes, getExtensionBasePath

class SymbolDialogHandler(unohelper.Base, XDialogEventHandler):

    def __init__(self, ctx, model, controller, dialog):
        self.ctx = ctx
        self.model = model
        self.controller = controller
        self.dialog = dialog
        self.sidc_options = {}
        self.listbox_values = {}
        self.disable_callHandler = False
        self.hex_color_value = None
        self.final_svg_data = None
        self.final_svg_args = None

        self.factory = self.ctx.getServiceManager().createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", self.ctx)
        self.provider = self.factory.createScriptProvider(model)
        self.script = self.provider.getScript(
            "vnd.sun.star.script:milsymbol.milsymbol.js?language=JavaScript&location=user:uno_packages/" +
            getExtensionBasePath(self.ctx))

    def init_dialog_controls(self):
        self.init_listbox(self.dialog)
        self.init_textboxes()

        self.dialog.Model.Step = 1

        self.dialog.getControl("btReality").getModel().State = 1
        self.dialog.getControl("btFriend").getModel().State = 1
        self.dialog.getControl("btPresent").getModel().State = 1
        self.dialog.getControl("btNotApplicableReinReduc").getModel().State = 1
        self.dialog.getControl("btStack1").getModel().State = 1
        self.dialog.getControl("btLight").getModel().State = 1
        self.dialog.getControl("btTarget").getModel().State = 1
        self.dialog.getControl("btNotApplicableSignature").getModel().State = 1

        self.version_value = symbols_data.VERSION
        self.context_value = symbols_data.BUTTONS["CONTEXT"]["btReality"]
        self.affiliation_value = symbols_data.BUTTONS["AFFILIATION"]["btFriend"]
        self.status_value = symbols_data.BUTTONS["STATUS"]["btPresent"]
        self.reinforced_reduced_option = symbols_data.BUTTONS["REINFORCED_REDUCED"]["btNotApplicableReinReduc"]
        self.stack_option = symbols_data.BUTTONS["STACK"]["btStack1"]
        self.color_mode_option = symbols_data.BUTTONS["COLOR"]["btLight"]
        self.signature_option = symbols_data.BUTTONS["SIGNATURE"]["btNotApplicableSignature"]
        self.engagement_option = symbols_data.BUTTONS["ENGAGEMENT"]["btTarget"]

        self.updatePreview()

    def init_listbox(self, dialog):
        selected_index = 4 # Land unit
        self.populate_symbol_listboxes(dialog, selected_index)

        self.listbox_map = {
            "ltbMainIcon":        "mainIcon_value",
            "ltbFirstIcon":       "firstIcon_value",
            "ltbSecondIcon":      "secondIcon_value",
            "ltbHeadTaskDummy":   "headquartersTaskforceDummy_value",
            "ltbEchelonMobility": "echelonMobility_value",
            "ltbCountry":         "country_code_value"
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

    def get_current_symbol(self, selected_index):
        symbol_meta = symbols_data.SYMBOLS[selected_index]
        symbol_id = symbol_meta["id"]
        current_symbol = symbols_data.SYMBOL_DETAILS[symbol_id]
        return current_symbol

    def fill_listbox(self, dialog, control_name, items, selected_index):
        labels = [item["label"] for item in items]

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

        if methodName.startswith("ltb"): # listboxes
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
            dialog.getControl("tlbPreview").setFocus()
            return True
        elif methodName == "click_Search":
            self.apply_search_selection(dialog, eventObject)
            return True
        elif methodName == "click_dialog":
            selected_value = dialog.getControl("tbSearch").Text
            self.update_listbox_value(dialog, selected_value)
            dialog.getControl("tlbPreview").setFocus()
            return True
        elif self.button_handler(dialog, methodName):
            return True
        elif methodName == "dialog_btSave":
            if self.controller is not None:
                shape = self.controller.get_diagram().set_svg_data(
                    self.final_svg_data)
                insertGraphicAttributes(shape,
                                        self.final_svg_args)
            else:
                insertSvgGraphic(
                    self.ctx, self.model,
                    self.final_svg_data,
                    self.final_svg_args)
            dialog.endExecute()
            return True
        elif methodName == "dialog_btCancel":
            dialog.endExecute()
            return True
        else:
            return False

    def getSupportedMethodNames(self):
        return self.buttons

    def disposing(self, event):
        pass

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

        self.symbolSet_value =                  self.fill_listbox(dialog, "ltbSymbolSet",       symbols_data.SYMBOLS, selected_index)
        self.mainIcon_value =                   self.fill_listbox(dialog, "ltbMainIcon",        current_symbol["MainIcon"], 1)
        self.firstIcon_value =                  self.fill_listbox(dialog, "ltbFirstIcon",       current_symbol["FirstIconModifier"], 0)
        self.secondIcon_value =                 self.fill_listbox(dialog, "ltbSecondIcon",      current_symbol["SecondIconModifier"], 0)
        self.echelonMobility_value =            self.fill_listbox(dialog, "ltbEchelonMobility", current_symbol["EchelonMobility"], 0)
        self.headquartersTaskforceDummy_value = self.fill_listbox(dialog, "ltbHeadTaskDummy",   current_symbol["HeadquartersTaskforceDummy"], 0)
        self.country_code_value =               self.fill_listbox(dialog, "ltbCountry",         country_data.COUNTRY_CODES, 0)

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
            color_value = props[0].Value
            return self.color_to_hex(color_value)
        return None

    def color_to_hex(self, color_val):
        red   = (color_val >> 16) & 0xFF
        green = (color_val >> 8) & 0xFF
        blue  = color_val & 0xFF
        return f"#{red:02X}{green:02X}{blue:02X}"

    def textbox_handler(self, dialog, methodName):
        options_name = self.textbox_map.get(methodName)
        text_value = dialog.getControl(methodName).Text

        if text_value:
            self.sidc_options[options_name] = " " + text_value
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
                if isinstance(item, dict) and item.get("label") == selected_value),
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
            if search_text:
                matches = []
                search = search_text.lower()

                for data in symbols_data.SYMBOL_DETAILS.values():
                    for icon in data.get("MainIcon", []):
                        label = icon.get("label", "")
                        label_to_search = label.split("–")[0].strip()
                        word_match = any(w.lower().startswith(search) for w in label_to_search.split())
                        prefix_match = label_to_search.lower().startswith(search)
                        if word_match or prefix_match:
                            matches.append(label)

                search_listbox = dialog.getControl("ltbSearch")
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

            self.mainIcon_value = next(
                (item.get("value") for item in labels
                 if isinstance(item, dict) and item.get("label") == selected_value),
                None
            )

            if self.mainIcon_value is not None:
                symbolSet_name = symbolSet
                break

        self.symbolSet_value = next(
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
        dialog.getControl("tlbPreview").setFocus()

    def button_handler(self, dialog, active_button_id):
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
            self.update_button(dialog, button_id, active_button_id, group_name)
        return True

    def update_button(self, dialog, button_id, active_button_id, group_name):
        state = 0
        if button_id == active_button_id:
            state = 1

            group_buttons= symbols_data.BUTTONS.get(group_name)
            value = group_buttons.get(button_id)

            if group_name == "CONTEXT":
                self.context_value = value
            elif group_name == "AFFILIATION":
                self.affiliation_value = value
            elif group_name == "STATUS":
                self.status_value = value
            elif group_name == "REINFORCED_REDUCED":
                self.reinforced_reduced_option = value
            elif group_name == "STACK":
                self.stack_option = value
            elif group_name == "COLOR":
                self.color_mode_option = value
                if value == "Custom":
                    hex_color = self.pick_custom_color()
                    if hex_color:
                        self.hex_color_value = hex_color
            elif group_name == "SIGNATURE":
                self.signature_option = value
            elif group_name == "ENGAGEMENT":
                self.engagement_option = value

            self.updatePreview()

        dialog.getControl(button_id).getModel().State = state

    def updatePreview(self):
        file_url = self.insertSymbolToPreview()
        imgPreview = self.dialog.getModel().getByName("imgPreview")
        imgPreview.ImageURL = ""
        imgPreview.ImageURL = file_url

    def create_sidc(self):
        sidc = [
            self.version_value,
            self.context_value,
            self.affiliation_value,
            self.symbolSet_value,
            self.status_value,
            self.headquartersTaskforceDummy_value,
            self.echelonMobility_value,
            self.mainIcon_value,
            self.firstIcon_value[-2:],
            self.secondIcon_value[-2:],
            self.firstIcon_value[0],
            self.secondIcon_value[0],
            "00000",
            "000" # Country flag
        ]
        self.sidc = ''.join(sidc)
        return self.sidc

    def insertSymbolToPreview(self):
        sidc_code = self.create_sidc()

        args = [
            sidc_code,
            NamedValue("size", 60.0),
            NamedValue("stack", self.stack_option),
            NamedValue("reinforced", self.reinforced_reduced_option),
            NamedValue("signature", self.signature_option),
            NamedValue("engagementType", self.engagement_option)
        ]

        if self.country_code_value:
            args.extend([
                NamedValue("country", self.country_code_value),
                NamedValue("country_flag", "true")
            ])

        if self.color_mode_option and self.color_mode_option != "false":
            if self.color_mode_option == "Custom":
                args.append(NamedValue("fillColor", self.hex_color_value))
            else:
                args.append(NamedValue("colorMode", self.color_mode_option))
        else:
            args.append(NamedValue("fill", "false"))

        for key, value in self.sidc_options.items():
            if value:
                args.append(NamedValue(key, value))

        temp_svg_path = os.path.join(tempfile.gettempdir(), "preview.svg")

        try:
            result = self.script.invoke(args, (), ())
            # Assuming the result contains SVG data
            if result and len(result) > 0:
                svg_data = str(result[0])
                self.final_svg_data = svg_data
                self.final_svg_args = args
                with open(temp_svg_path, 'w', encoding='utf-8') as preview_file:
                    preview_file.write(svg_data)
                svg_url = systemPathToFileUrl(temp_svg_path)
                return svg_url
        except Exception as e:
            print(f"Error executing script: {e}")
            return

    def remove_temp_preview_svg(self):
        temp_svg_path = os.path.join(tempfile.gettempdir(), "preview.svg")

        try:
            if os.path.exists(temp_svg_path):
                os.remove(temp_svg_path)
                print(f"Temporary SVG deleted: {temp_svg_path}")
            else:
                print("No temporary SVG file found to delete.")
        except Exception as e:
            print(f"Error deleting temporary SVG: {e}")
