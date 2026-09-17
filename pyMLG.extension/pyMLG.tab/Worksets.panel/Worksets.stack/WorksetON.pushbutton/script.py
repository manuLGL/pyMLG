# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import WorksetVisibility, Transaction, FilteredWorksetCollector, WorksetKind
from pyrevit import revit
from mlg_sprache import t

__title__ = "WorksetsON"

doc = revit.doc
uidoc = revit.uidoc

if doc.IsWorkshared:
    active_view = uidoc.ActiveView

    # Alle User-Worksets sammeln
    worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()

    transaktion = Transaction(doc, t("Alle Worksets einblenden", u"Show all worksets", u"Mostrar todos los subproyectos"))
    transaktion.Start()

    try:
        for workset in worksets:
            active_view.SetWorksetVisibility(workset.Id, WorksetVisibility.Visible)

        transaktion.Commit()
    except:
        transaktion.RollBack()