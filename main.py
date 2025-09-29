# SPDX-License-Identifier: MPL-2.0

import sys
import unohelper
import officehelper

from com.sun.star.task import XJobExecutor


# The MainJob is a UNO component derived from unohelper.Base class
# and also the XJobExecutor, the implemented interface
class MainJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx
        # handling different situations (inside LibreOffice or other process)
        try:
            self.sm = ctx.getServiceManager()
            self.desktop = XSCRIPTCONTEXT.getDesktop()
        except NameError:
            self.sm = ctx.ServiceManager
            self.desktop = self.ctx.getServiceManager().createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.ctx)

    def trigger(self, args):
        # ServiceManager
        smgr = self.ctx.getServiceManager()
        # DialogProvider
        dialog_provider = smgr.createInstanceWithContext("com.sun.star.awt.DialogProvider2", self.ctx)
        dialog_url = "vnd.sun.star.extension://com.collabora.milsymbol/dialog/TacticalSymbolDlg.xdl"
        try:
            dialog = dialog_provider.createDialog(dialog_url)
            dialog.execute()
        except Exception as e:
            pass

# Starting from Python IDE
def main():
    try:
        ctx = XSCRIPTCONTEXT
    except NameError:
        ctx = officehelper.bootstrap()
        if ctx is None:
            print("ERROR: Could not bootstrap default Office.")
            sys.exit(1)
    job = MainJob(ctx)
    job.trigger("hello")


# Starting from command line
if __name__ == "__main__":
    main()


# pythonloader loads a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    MainJob,  # UNO object class
    "com.collabora.milsymbol.do",  # implementation name (customize for yourself)
    ("com.sun.star.task.Job",), )  # implemented services (only 1)
