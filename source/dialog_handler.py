# SPDX-License-Identifier: MPL-2.0

import os
import sys
import unohelper
from com.sun.star.awt import XDialogEventHandler
from com.sun.star.beans import NamedValue

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

import symbols_data

class DialogHandler(unohelper.Base, XDialogEventHandler):

    def __init__(self, ctx, dialog=None):
        self.ctx = ctx
        self.dialog = dialog
        self.sidc_options = {}
        self.listbox_values = {}
        self.disable_callHandler = False

    def init_dialog_controls(self):
        self.init_listbox(self.dialog)

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

    def init_listbox(self, dialog):
        selected_index = 4 # Land unit
        self.populate_symbol_listboxes(dialog, selected_index)

        self.listbox_map = {
            "ltbMainIcon":        "mainIcon_value",
            "ltbFirstIcon":       "firstIcon_value",
            "ltbSecondIcon":      "secondIcon_value",
            "ltbHeadTaskDummy":   "headquartersTaskforceDummy_value",
            "ltbEchelonMobility": "echelonMobility_value"
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

        if methodName.startswith("ltb"):
            self.listbox_handler(dialog, eventObject, methodName)
            return True
        elif methodName.startswith("tabbed"):
            self.tabbed_button_switch_handler(dialog, methodName)
            return True
        elif self.button_handler(dialog, methodName):
            return True
        elif methodName == "dialog_btCancel":
            dialog.endExecute()
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

    def update_symbolSet_listbox(self, dialog, eventObject):
        selected_index = eventObject.Source.getSelectedItemPos()
        self.populate_symbol_listboxes(dialog, selected_index)

    def populate_symbol_listboxes(self, dialog, selected_index):
        current_symbol = self.get_current_symbol(selected_index)

        self.symbolSet_value =                  self.fill_listbox(dialog, "ltbSymbolSet",       symbols_data.SYMBOLS, selected_index)
        self.mainIcon_value =                   self.fill_listbox(dialog, "ltbMainIcon",        current_symbol["MainIcon"], 0)
        self.firstIcon_value =                  self.fill_listbox(dialog, "ltbFirstIcon",       current_symbol["FirstIconModifier"], 0)
        self.secondIcon_value =                 self.fill_listbox(dialog, "ltbSecondIcon",      current_symbol["SecondIconModifier"], 0)
        self.echelonMobility_value =            self.fill_listbox(dialog, "ltbEchelonMobility", current_symbol["EchelonMobility"], 0)
        self.headquartersTaskforceDummy_value = self.fill_listbox(dialog, "ltbHeadTaskDummy",   current_symbol["HeadquartersTaskforceDummy"], 0)

    def tabbed_button_switch_handler(self, dialog, methodName):
        # Handling tabbed page buttons
        if methodName == "tabbed_btBasic":
            dialog.Model.Step = 1
        elif methodName == "tabbed_btAdvance":
            dialog.Model.Step = 2
            dialog.getControl("ltbSearch").setVisible(False)

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
            elif group_name == "SIGNATURE":
                self.signature_option = value
            elif group_name == "ENGAGEMENT":
                self.engagement_option = value

        dialog.getControl(button_id).getModel().State = state