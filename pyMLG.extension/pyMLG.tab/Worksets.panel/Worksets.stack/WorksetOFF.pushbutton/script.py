# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import WorksetVisibility, Transaction, WorksetId
from pyrevit import revit
from mlg_sprache import t

__title__ = "WorksetOFF"

doc = revit.doc
uidoc = revit.uidoc

if doc.IsWorkshared:
    selection = list(uidoc.Selection.GetElementIds())

    if selection:
        element = doc.GetElement(selection[0])
        workset_id = element.WorksetId

        if workset_id != WorksetId.InvalidWorksetId:
            transaktion = Transaction(doc, t("Workset ausblenden", u"Hide workset", u"Ocultar subproyecto"))
            transaktion.Start()

            try:
                uidoc.ActiveView.SetWorksetVisibility(workset_id, WorksetVisibility.Hidden)
                transaktion.Commit()
            except:
                transaktion.RollBack()