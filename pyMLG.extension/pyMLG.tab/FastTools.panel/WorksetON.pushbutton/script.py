# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import WorksetVisibility, Transaction, FilteredWorksetCollector, WorksetKind
from pyrevit import revit

__title__ = "WorksetsON"

doc = revit.doc
uidoc = revit.uidoc

if doc.IsWorkshared:
    active_view = uidoc.ActiveView

    # Alle User-Worksets sammeln
    worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()

    t = Transaction(doc, "Alle Worksets einblenden")
    t.Start()

    try:
        for workset in worksets:
            active_view.SetWorksetVisibility(workset.Id, WorksetVisibility.Visible)

        t.Commit()
    except:
        t.RollBack()