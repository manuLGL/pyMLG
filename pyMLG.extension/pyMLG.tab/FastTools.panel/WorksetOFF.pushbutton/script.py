# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import WorksetVisibility, Transaction, WorksetId
from pyrevit import revit

__title__ = "WorksetOFF"

doc = revit.doc
uidoc = revit.uidoc

if doc.IsWorkshared:
    selection = list(uidoc.Selection.GetElementIds())

    if selection:
        element = doc.GetElement(selection[0])
        workset_id = element.WorksetId

        if workset_id != WorksetId.InvalidWorksetId:
            t = Transaction(doc, "Workset ausblenden")
            t.Start()

            try:
                uidoc.ActiveView.SetWorksetVisibility(workset_id, WorksetVisibility.Hidden)
                t.Commit()
            except:
                t.RollBack()