# SPDX-License-Identifier: MPL-2.0

import os
import sys
import uno
import unohelper
from com.sun.star.awt import XDialogEventHandler

base_dir = os.path.dirname(__file__)
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

import listbox_data

class DialogHandler(unohelper.Base, XDialogEventHandler):
    # "group name": ["button1", "button2",...]
    buttonGroups = {
        "mode":         ["mode_btReality", "mode_btExercise", "mode_btSimulation"],
        "stack":        ["stack_bt1", "stack_bt2", "stack_bt3", "stack_bt4", "stack_bt5", "stack_bt6"],
        "color":        ["color_btNoFill", "color_btLight", "color_btMedium", "color_btDark", "color_btCustom"],
        "status":       ["status_btPresent", "status_btPlanned", "status_btFullyCapable",
                        "status_btDamaged", "status_btDestroyed", "status_btFullToCapacity"],
        "reinforced":   ["reinforced_btNotApplicable", "reinforced_btReinforced",
                        "reinforced_btReduced", "reinforced_btReinforcedReduced"],
        "affiliation":  ["affiliation_btPending", "affiliation_btNeutral", "affiliation_btUnknown", "affiliation_btFriend",
                        "affiliation_btAssumedFriend", "affiliation_btSuspectJoker", "affiliation_btHostileFaker"],
        "engagement":   ["engagement_btTarget", "engagement_btNonTarget", "engagement_btExpired"],
        "signature":    ["signature_btNotApplicable", "signature_btElectSign"]
    }

    def __init__(self, dialog):
        self.dialog = dialog

    def initialize_dialog_controls(self, dialog):
        self.update_listboxes(dialog, 4)
        self.dialog.getControl("mode_btReality").getModel().State = 1
        self.dialog.getControl("affiliation_btFriend").getModel().State = 1
        self.dialog.getControl("status_btPresent").getModel().State = 1
        self.dialog.getControl("reinforced_btNotApplicable").getModel().State = 1
        self.dialog.getControl("stack_bt1").getModel().State = 1
        self.dialog.getControl("color_btLight").getModel().State = 1
        self.dialog.getControl("engagement_btTarget").getModel().State = 1
        self.dialog.getControl("signature_btNotApplicable").getModel().State = 1

    def callHandlerMethod(self, dialog, eventObject, methodName):
        if methodName == "action_ltbSymbolSet":
            selected_index = eventObject.Source.getSelectedItemPos()
            self.update_listboxes(dialog, selected_index)
            return True
        elif methodName == "dialog_btCancel":
            self.dialog.dispose()
        elif self.buttonStateHandler(methodName):
            return True
        elif self.tabbedButtonSwitchHandler(methodName):
            return True
        else:
            return False

    def getSupportedMethodNames(self):
        return self.buttons

    def disposing(self, event):
        pass

    def tabbedButtonSwitchHandler(self, methodName):
        # Handling tabbed page buttons
        if methodName.startswith("tabbed"):
            if methodName == "tabbed_btBasic":
                self.dialog.Model.Step = 1
                return True
            elif methodName == "tabbed_btAdvance":
                self.dialog.Model.Step = 2
                return True

    def buttonStateHandler(self, methodName):
        # Check if the clicked button belongs to the current group and
        # set STATE=1 for the clicked button, and STATE=0 for all others in the same group
        # This ensures that only one button can be selected at a time within a group, 
        # and it does not affect the selection of buttons in other groups.
        for group_name, buttons in self.buttonGroups.items():
            if methodName.startswith(group_name):
                for btn_name in buttons:
                    self.dialog.getControl(btn_name).getModel().State = 1 if (btn_name == methodName) else 0
                return True

    def update_listboxes(self, dialog, index):
        data_dict = listbox_data.LISTBOX_LIST[index]

        main_icons = data_dict["MainIcon"]
        first_icon = data_dict["FirstIconModifier"]
        second_icon = data_dict["SecondIconModifier"]
        echelon_mobility = data_dict["EchelonMobility"]
        head_task_dummy = data_dict["HeadquartersTaskforceDummy"]

        # Main Icon ListBox
        ltbMainIcon = dialog.getControl("ltbMainIcon")
        ltbMainIcon.removeItems(0, ltbMainIcon.getItemCount())
        ltbMainIcon.addItems(tuple(main_icons), 0)
        ltbMainIcon.selectItemPos(0, True)
        # First Icon Modifier ListBox
        ltbFirstIcon = dialog.getControl("ltbFirstIcon")
        ltbFirstIcon.removeItems(0, ltbFirstIcon.getItemCount())
        ltbFirstIcon.addItems(tuple(first_icon), 0)
        ltbFirstIcon.selectItemPos(0, True)
        # Second Icon Modifier ListBox
        ltbSecondIcon = dialog.getControl("ltbSecondIcon")
        ltbSecondIcon.removeItems(0, ltbSecondIcon.getItemCount())
        ltbSecondIcon.addItems(tuple(second_icon), 0)
        ltbSecondIcon.selectItemPos(0, True)
        # Echelon/Mobility List Box
        ltbEchelonMobility = dialog.getControl("ltbEchelonMobility")
        ltbEchelonMobility.removeItems(0, ltbEchelonMobility.getItemCount())
        ltbEchelonMobility.addItems(tuple(echelon_mobility), 0)
        ltbEchelonMobility.selectItemPos(0, True)
        # Headquarters/Taskforce/Dummy List Box
        ltbHeadTaskDummy = dialog.getControl("ltbHeadTaskDummy")
        ltbHeadTaskDummy.removeItems(0, ltbHeadTaskDummy.getItemCount())
        ltbHeadTaskDummy.addItems(tuple(head_task_dummy), 0)
        ltbHeadTaskDummy.selectItemPos(0, True)
