# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

# This file incorporates work covered by the following license notice:
#   SPDX-License-Identifier: LGPL-3.0-only

"""
Controller class for LibreOffice extension
"""

import unohelper

from perf import timed
from utils import containing_orbat_group, locked_controllers

from .gui import Gui

from com.sun.star.document import XUndoManagerListener
from com.sun.star.view import XSelectionChangeListener

from smart.diagram.organizationcharts.orgchart.orgchart import OrgChart


class DocumentUndoWatch(unohelper.Base, XUndoManagerListener):
    """Tells the diagram to forget what it keeps about its shapes when the document is undone.

    Laying a diagram out asks the office as little as it can. It decides whether a shape
    needs moving or resizing by comparing against the position and size it last wrote, and
    it keeps the proportions of each symbol's picture. Undoing or redoing puts an earlier
    state of the document back without telling the extension, so from then on those answers
    describe a document that is no longer there, and every write they would save is a write
    the diagram needs.

    What is kept is dropped rather than read again. Asking a shape inside a group where it
    is makes the office work out the bounding box of the group and route every connector in
    it again, so the next layout writes each shape once instead, which is what it did before
    any of this was kept.
    """

    def __init__(self, controller):
        self._controller = controller

    def _forget_kept_shape_state(self):
        """Drop what the diagram keeps, if there is a diagram to tell."""
        try:
            diagram = self._controller.get_diagram()
            if diagram is not None:
                diagram.forget_kept_shape_state()
        except Exception as ex:
            print(f"Error clearing what the diagram keeps: {ex}")

    # The two events that put an earlier state of the document back
    def actionUndone(self, event):
        self._forget_kept_shape_state()

    def actionRedone(self, event):
        self._forget_kept_shape_state()

    # The rest leave the document as it is, and only say what the undo stacks now hold
    def undoActionAdded(self, event):
        pass

    def allActionsCleared(self, event):
        pass

    def redoActionsCleared(self, event):
        pass

    def resetAll(self, event):
        pass

    def enteredContext(self, event):
        pass

    def enteredHiddenContext(self, event):
        pass

    def leftContext(self, event):
        pass

    def leftHiddenContext(self, event):
        pass

    def cancelledContext(self, event):
        pass

    def disposing(self, event):
        pass


class Controller(unohelper.Base, XSelectionChangeListener):
    """Controller class for LibreOffice extension"""

    # Group types
    ORGANIGROUP = 0
    RELATIONGROUP = 1
    PROCESSGROUP = 2
    LISTGROUP = 3
    MATRIXGROUP = 4

    # Diagram types
    NOTDIAGRAM = -1
    SIMPLEORGANIGRAM = 0
    HORIZONTALORGANIGRAM = 1
    TABLEHIERARCHYDIAGRAM = 2
    ORGANIGRAM = 3
    VENNDIAGRAM = 10
    CYCLEDIAGRAM = 11
    PYRAMIDDIAGRAM = 12
    TARGETDIAGRAM = 13
    CONTINUOUSBLOCKPROCESS = 20
    STAGGEREDPROCESS = 21
    BENDINGPROCESS = 22
    UPWARDARROWPROCESS = 23

    def __init__(self, smart_ph, x_context, x_frame):
        """Initialize Controller"""
        self._smart_ph = smart_ph
        self._x_context = x_context
        self._x_frame = x_frame
        self._x_controller = x_frame.getController()
        self._gui = None
        self._x_selection_supplier = None

        self._diagram = None
        self._diagram_type = None
        self._group_type = None
        self._last_diagram_type = -1
        # The group shape of the diagram the control dialog was last built for. The
        # group shape is what stands for a diagram: two diagrams in one document can
        # carry the same names, for example after copying one, so a name or an id
        # parsed from a name cannot tell them apart.
        self._last_diagram_group_shape = None
        # The group shape this controller entered on the user's behalf. None if
        # nothing is selected.
        self._entered_group = None

        self._undo_watch = None

        self._gui = Gui(self, self._x_context, self._x_frame)
        self.add_selection_listener()
        self.add_undo_listener()

    def dispose(self):
        """Dispose controller and all associated resources"""
        try:
            self.remove_selection_listener()
            self.remove_undo_listener()

            if self._gui is not None:
                self._gui.close_and_dispose_control_dialog()
                self._gui = None

            self._diagram = None
            self._last_diagram_group_shape = None
            self._entered_group = None

            self._x_controller = None
            self._x_frame = None
            self._x_selection_supplier = None

        except Exception as e:
            print(f"Error disposing controller: {e}")

    def is_smart_diagram_shape(self, shape_name):
        """Check if shape is a smart diagram shape"""
        return shape_name.startswith("OrbatDiagram")

    def get_containing_diagram_group(self, shape):
        """The diagram group shape a shape belongs to, or None when it belongs to none.

        A diagram group shape answers for itself. A shape inside a diagram group answers
        with that group.
        """
        return containing_orbat_group(shape)

    def _shape_is_inside(self, shape, group):
        """Whether a shape sits inside a group shape, at any depth."""
        current = shape
        while current is not None:
            try:
                parent = current.getParent()
            except Exception:
                return False
            if parent is None or not hasattr(parent, "supportsService"):
                return False
            try:
                if not parent.supportsService("com.sun.star.drawing.Shape"):
                    return False
            except Exception:
                return False
            if parent == group:
                return True
            current = parent
        return False

    def set_new_size(self):
        """Set new diagram size"""
        self.get_diagram().increase_size_prop()
        width = 0
        height = 0
        x_pos = 0
        y_pos = 0

        if self.get_diagram().get_size_prop() == self.get_diagram().UD_SIZE:
            width = self.get_diagram().get_ud_width_prop()
            height = self.get_diagram().get_ud_height_prop()
            x_pos = self.get_diagram().get_ud_x_pos_prop()
            y_pos = self.get_diagram().get_ud_y_pos_prop()

        if self.get_diagram().get_size_prop() == self.get_diagram().FULL_SIZE:
            s = self.get_diagram().get_group_shape_size()
            self.get_diagram().set_ud_width_prop(s.Width)
            self.get_diagram().set_ud_height_prop(s.Height)
            p = self.get_diagram().get_group_shape_pos()
            self.get_diagram().set_ud_x_pos_prop(p.X)
            self.get_diagram().set_ud_y_pos_prop(p.Y)
            width = (
                self.get_diagram().page_props.Width
                - self.get_diagram().page_props.BorderLeft
                - self.get_diagram().page_props.BorderRight
            )
            height = (
                self.get_diagram().page_props.Height
                - self.get_diagram().page_props.BorderTop
                - self.get_diagram().page_props.BorderBottom
            )
            x_pos = self.get_diagram().page_props.BorderLeft
            y_pos = self.get_diagram().page_props.BorderTop

        self.get_diagram().set_group_shape_size_and_pos(width, height, x_pos, y_pos)
        self.get_diagram().refresh_diagram()

    def get_number_of_pages(self):
        """Get number of pages in document"""
        o_document = self._x_frame.getController().getModel()
        pages_supplier = o_document
        pages = pages_supplier.getDrawPages()
        return pages.getCount()

    def get_smart_ph(self):
        """Get SmartProtocolHandler"""
        return self._smart_ph

    def get_diagram(self):
        """Get current diagram"""
        return self._diagram

    def set_null_diagram(self):
        """Set diagram to None"""
        if self._diagram is not None:
            self._diagram = None

    def dispose_diagram(self):
        """Dispose diagram and clear all shape references before document closes.

        This method should be called from queryClosing() BEFORE the document
        actually closes to ensure all Python references to UNO shapes are
        released before LibreOffice destroys the underlying C++ objects.
        """
        try:
            # Clear all undo action references from the dialog handler
            if self._gui is not None:
                dialog_handler = Gui._global_control_dlg_listener
                if dialog_handler is not None:
                    if hasattr(dialog_handler, "clear_all_undo_action_references"):
                        dialog_handler.clear_all_undo_action_references()

            self._diagram = None
            self._last_diagram_group_shape = None
        except Exception as e:
            print(f"Error in dispose_diagram: {e}")

    def set_group_type(self, d_type):
        """Set group type"""
        self._group_type = d_type

    def get_group_type(self):
        """Get group type"""
        return self._group_type

    def set_diagram_type(self, d_type):
        """Set diagram type"""
        self._diagram_type = d_type

    def get_diagram_type(self):
        """Get diagram type"""
        return self._diagram_type

    def set_last_diagram_type(self, d_type):
        """Set last diagram type"""
        self._last_diagram_type = d_type

    def get_last_diagram_type(self):
        """Get last diagram type"""
        return self._last_diagram_type

    def set_last_diagram_group_shape(self, group_shape):
        """Set the group shape of the diagram the control dialog was last built for"""
        self._last_diagram_group_shape = group_shape

    def get_last_diagram_group_shape(self):
        """Get the group shape of the diagram the control dialog was last built for"""
        return self._last_diagram_group_shape

    def add_selection_listener(self):
        """Add selection change listener"""
        if self._x_selection_supplier is None:
            self._x_selection_supplier = self._x_controller
        if self._x_selection_supplier is not None:
            self._x_selection_supplier.addSelectionChangeListener(self)

    def remove_selection_listener(self):
        """Remove selection change listener"""
        if self._x_selection_supplier is not None:
            self._x_selection_supplier.removeSelectionChangeListener(self)

    def _get_undo_manager(self):
        """The undo manager of the document, or None where it has none."""
        try:
            model = self._x_controller.getModel() if self._x_controller else None
            if model is None:
                return None
            if hasattr(model, "getUndoManager"):
                return model.getUndoManager()
            return getattr(model, "UndoManager", None)
        except Exception as ex:
            print(f"Could not get undo manager: {ex}")
        return None

    def add_undo_listener(self):
        """Listen for undo and redo, so that a diagram can be told the document changed."""
        if self._undo_watch is not None:
            return

        undo_manager = self._get_undo_manager()
        if undo_manager is None:
            return

        try:
            self._undo_watch = DocumentUndoWatch(self)
            undo_manager.addUndoManagerListener(self._undo_watch)
        except Exception as ex:
            self._undo_watch = None
            print(f"Error listening for undo: {ex}")

    def remove_undo_listener(self):
        """Stop listening for undo and redo"""
        if self._undo_watch is None:
            return

        try:
            undo_manager = self._get_undo_manager()
            if undo_manager is not None:
                undo_manager.removeUndoManagerListener(self._undo_watch)
        except Exception as ex:
            print(f"Error giving up the undo listener: {ex}")

        self._undo_watch = None

    def execute_gallery_dialog(self):
        """Execute gallery dialog"""
        if self._gui is not None:
            return self._gui.execute_gallery_dialog()
        return 0

    def get_current_page(self):
        model = self._x_frame.getController().getModel()
        if model.supportsService("com.sun.star.text.TextDocument"):
            return model.getDrawPage()
        elif model.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            return self._x_controller.getActiveSheet().getDrawPage()  # Calc
        else:
            return self._x_controller.getCurrentPage()  # Impress/Draw

    def get_location(self):
        """Get current locale"""
        locale = None
        try:
            x_mcf = self._x_context.getServiceManager()
            o_configuration_provider = x_mcf.createInstanceWithContext(
                "com.sun.star.configuration.ConfigurationProvider", self._x_context
            )
            x_localizable = o_configuration_provider
            locale = x_localizable.getLocale()
        except Exception as ex:
            print(f"Error getting locale: {ex}")
        return locale

    def get_current_diagram_id(self):
        """Get current diagram ID from selected shape"""
        name = self.get_diagram().get_shape_name(self.get_selected_shape())
        s = ""
        i = 0

        # Skip non-digit characters
        while i < len(name) and not name[i].isdigit():
            i += 1

        # Collect digits until dash
        while i < len(name) and name[i] != "-":
            s += name[i]
            i += 1

        return int(s) if s else 0

    def get_shape_id(self, name):
        """Get shape ID from name"""
        s = ""
        i = 0

        # Skip to dash
        while i < len(name) and name[i] != "-":
            i += 1

        # Skip non-digit characters after dash
        while i < len(name) and not name[i].isdigit():
            i += 1

        # Collect digits
        while i < len(name) and name[i].isdigit():
            s += name[i]
            i += 1

        return int(s) if s else 0

    def get_selected_shape(self):
        """Get currently selected shape"""
        try:
            selection = self._x_selection_supplier.getSelection()
            if selection and selection.getCount() > 0:
                return selection.getByIndex(0)
        except Exception:
            pass
        return None

    def create_diagram(self, data=None):
        """Create diagram with optional data"""
        self.remove_selection_listener()
        self.instantiate_diagram()

        model = self._x_frame.getController().getModel()
        with locked_controllers(model):
            if self.get_diagram() is not None:
                if data is not None:
                    self.get_diagram().create_diagram(data)
                else:
                    self.get_diagram().create_diagram()

                # Initialize object tree in organigrams. The geometry of the shapes
                # was just written, so it is already known and must not be read back
                # from a menu command.
                if self.get_group_type() == self.ORGANIGROUP:
                    self.get_diagram().init_diagram(read_geometry=False)

        if self.get_diagram() is not None:
            # Showing the dialog is left outside the lock, because it takes the focus
            self._gui.set_visible_control_dialog(True)
        self.add_selection_listener()

    def is_shape_service(self, obj):
        """Check if object is a shape service"""
        is_shape = False
        if obj is not None:
            try:
                # In Python UNO, check for supported services
                if hasattr(obj, "supportsService"):
                    if obj.supportsService("com.sun.star.drawing.Shape"):
                        is_shape = True
                    if obj.supportsService("com.sun.star.drawing.GroupShape"):
                        return False
            except Exception:
                pass
        return is_shape

    def create_diagram_from_list(self):
        """Create diagram from selected text list"""
        self.remove_selection_listener()

        if (
            self.get_selected_shape() is not None
            and self.get_selected_shapes().getCount() == 1
            and self.is_shape_service(self.get_selected_shape())
        ):
            try:
                x_text = self.get_selected_shape()
                text_content = (
                    x_text.getString() if hasattr(x_text, "getString") else ""
                )

                if not text_content or text_content.strip() == "":
                    title = self._gui.get_dialog_property_value(
                        "GenericStrings", "CouldnotCreateDiagram.Title"
                    )
                    message = self._gui.get_dialog_property_value(
                        "GenericStrings", "CouldnotCreateDiagram.Message"
                    )
                    self._gui.show_message_box(title, message)
                else:
                    # TODO: Implement DataOfDiagram creation from text
                    # This would parse the text content and create diagram data
                    pass

            except Exception as ex:
                print(f"Error creating diagram from list: {ex}")

        self.add_selection_listener()

    def instantiate_diagram(self):
        """Instantiate diagram based on type"""
        self._diagram = OrgChart(self, self._gui, self._x_frame, self._x_context)

    def get_selected_shapes(self):
        """Get selected shapes collection"""
        try:
            return self._x_selection_supplier.getSelection()
        except Exception:
            return None

    def set_selected_shape(self, obj):
        """Set the selected shape"""
        try:
            self._x_selection_supplier.select(obj)
        except Exception as ex:
            print(f"Error setting selected shape: {ex}")

    def disposing(self, event):
        """Handle disposing event from XEventListener"""
        pass

    @timed("selectionChanged")
    def selectionChanged(self, event):
        """Handle selection change events - XSelectionChangeListener implementation"""
        selected_shape = self.get_selected_shape()

        if selected_shape:
            # Get shape name
            selected_shape_name = ""
            try:
                if hasattr(selected_shape, "getName"):
                    selected_shape_name = selected_shape.getName()
            except Exception:
                selected_shape_name = ""

            if selected_shape.supportsService("com.sun.star.drawing.GroupShape"):
                if (
                    self._entered_group is not None
                    and selected_shape == self._entered_group
                ):
                    # The office selects the group it has just
                    # left. The user wants out, don't touch things &
                    # keep stuff un-entered.
                    self._entered_group = None
                else:
                    try:
                        # Enter group only if drag orbat checkbox is
                        # disabled **and** dialog is visible
                        dialog_handler = self._gui._global_control_dlg_listener
                        drag_orbat = dialog_handler.is_drag_orbat_enabled()
                        if not drag_orbat and self._gui.is_visible_control_dialog():
                            selected_shape.enterGroup()
                            self._entered_group = selected_shape
                    except Exception:
                        print("Error entering group shape")
            elif self._entered_group is not None and not self._shape_is_inside(
                selected_shape, self._entered_group
            ):
                # The selection moved to a shape outside the entered
                # group, so lets have the next click on that group be
                # a fresh one
                self._entered_group = None

            # Listen for clicks on diagrams
            if self.is_smart_diagram_shape(selected_shape_name):
                x_group_shape = self.get_containing_diagram_group(selected_shape)

                current_group_shape = None
                if self._diagram is not None:
                    try:
                        current_group_shape = self._diagram.get_group_shape()
                    except Exception:
                        current_group_shape = None

                current_group_name = ""
                if current_group_shape is not None:
                    try:
                        current_group_name = current_group_shape.getName()
                    except Exception:
                        current_group_name = ""

                # A shape removed or added by something other than the diagram's own
                # bookkeeping, such as the office's own Delete or an undo it did not
                # tell the diagram about, leaves the tree accounting for fewer or more
                # shapes than the group holds. Rebuild the diagram rather than reuse a
                # tree that no longer matches the group.
                diagram_tree_is_stale = False
                if current_group_shape is not None and self._diagram is not None:
                    try:
                        diagram_tree = self._diagram.get_diagram_tree()
                        diagram_tree_is_stale = (
                            diagram_tree is not None
                            and not diagram_tree.knows_every_shape_in_the_group()
                        )
                    except Exception:
                        diagram_tree_is_stale = False

                # The control dialog can be showing a diagram of another document's
                # controller, and then this document's diagram is set up afresh even
                # when the click stayed in the same group.
                dialog_belongs_elsewhere = False
                if Gui._global_control_dlg_listener is not None:
                    try:
                        dialog_belongs_elsewhere = (
                            Gui._global_control_dlg_listener.get_controller() != self
                        )
                    except Exception:
                        dialog_belongs_elsewhere = True

                # The group shapes stand for the diagrams. Two diagrams in one
                # document can carry the same names, for example after copying one, so
                # comparing the shapes themselves is what tells whether the click
                # landed in another diagram. When the clicked shape's group cannot be
                # resolved to a shape, fall back to comparing the diagram name prefix,
                # so a click on a diagram whose group is unreachable is still told
                # apart from the diagram that is currently open.
                needs_new_diagram = (
                    dialog_belongs_elsewhere
                    or current_group_shape is None
                    or diagram_tree_is_stale
                    or (
                        x_group_shape is not None
                        and x_group_shape != current_group_shape
                    )
                    or (
                        x_group_shape is None
                        and selected_shape_name.split("-", 1)[0]
                        != current_group_name.split("-", 1)[0]
                    )
                )

                if needs_new_diagram:
                    # The id parsed from the name only matters when no group shape was
                    # found and init_diagram falls back to searching the page by name.
                    group_name = selected_shape_name
                    if x_group_shape is not None:
                        try:
                            group_name = x_group_shape.getName()
                        except Exception:
                            pass
                    diagram_id = int(
                        "".join(c for c in group_name.split("-", 1)[0] if c.isdigit())
                        or "0"
                    )

                    # Set diagram types based on shape name
                    if selected_shape_name.startswith("OrbatDiagram"):
                        self.set_group_type(self.ORGANIGROUP)
                        self.set_diagram_type(self.ORGANIGRAM)

                    self.instantiate_diagram()

                    self.get_diagram().init_diagram(
                        diagram_id, group_shape=x_group_shape
                    )
                    self.get_diagram().init_properties()

                    # control dialog is open? just follow the
                    # selection. but keep it closed, if user did so
                    # earlier.
                    if (
                        self._gui.is_visible_control_dialog()
                        or not Gui._user_closed_dialog
                    ):
                        self._gui.set_visible_control_dialog(True)

                # Handle organization chart shape selection
                if selected_shape_name.startswith(
                    "OrbatDiagram"
                ) and selected_shape_name.endswith("RectangleShape0"):
                    if self.get_diagram() is not None:
                        self.get_diagram().select_shapes()

                # Focus the dialog if it's visible (don't auto-open)
                if self._gui is not None and self._gui.is_visible_control_dialog():
                    self._gui.set_focus_control_dialog()

    def disappear_control_dialog(self):
        """Hide control dialog"""
        if self._gui is not None:
            self._gui.set_visible_control_dialog(False)

    def is_only_simple_item_selected(self):
        """Check if only simple item is selected"""
        if self.get_selected_shapes().getCount() == 1:
            selected_shape = self.get_selected_shape()
            selected_shape_name = selected_shape.getName()

            if (
                selected_shape_name.startswith("OrbatDiagram")
                and "RectangleShape" in selected_shape_name
                and not selected_shape_name.endswith("RectangleShape0")
            ):
                return True
        return False


# pythonloader loads a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    Controller,  # UNO object class
    "com.collabora.milsymbol.Controller",  # implementation name
    ("com.sun.star.view.XSelectionChangeListener",),
)  # implemented services (only 1)
