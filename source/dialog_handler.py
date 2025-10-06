# SPDX-License-Identifier: MPL-2.0

import uno
import unohelper
from com.sun.star.awt import XDialogEventHandler

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

    def callHandlerMethod(self, dialog, eventObject, methodName):
        if methodName == "dialog_btCancel":
            self.dialog.dispose()
            return True

        if self.handleButtonState(methodName):
            return True
        if self.handleTabbedButtonSwitch(methodName):
            return True   
        return False

    def getSupportedMethodNames(self):
        return self.buttons

    def disposing(self, event):
        pass

    def handleTabbedButtonSwitch(self, methodName):
        # Handling tabbed page buttons
        if methodName.startswith("tabbed"):
            if methodName == "tabbed_btBasic":
                self.dialog.Model.Step = 1
                return True
            elif methodName == "tabbed_btAdvance":
                self.dialog.Model.Step = 2
                return True
            
    def handleButtonState(self, methodName):
        # Check if the clicked button belongs to the current group and
        # set STATE=1 for the clicked button, and STATE=0 for all others in the same group
        # This ensures that only one button can be selected at a time within a group, 
        # and it does not affect the selection of buttons in other groups.
        for group_name, buttons in self.buttonGroups.items():
            if methodName.startswith(group_name):
                for btn_name in buttons:
                    self.dialog.getControl(btn_name).getModel().State = 1 if (btn_name == methodName) else 0
                return True