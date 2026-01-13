# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import WorksetVisibility, Transaction, WorksetId, FilteredWorksetCollector, WorksetKind
from pyrevit import revit

__title__ = "WorksetREVERSE"

doc = revit.doc
uidoc = revit.uidoc

if doc.IsWorkshared:
    selection = list(uidoc.Selection.GetElementIds())

    if selection:
        element = doc.GetElement(selection[0])
        selected_workset_id = element.WorksetId

        if selected_workset_id != WorksetId.InvalidWorksetId:
            active_view = uidoc.ActiveView

            # Alle User-Worksets sammeln
            all_worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()

            t = Transaction(doc, "Nur ausgewaehltes Workset anzeigen")
            t.Start()

            try:
                for workset in all_worksets:
                    if workset.Id == selected_workset_id:
                        # Ausgewähltes Workset einblenden
                        active_view.SetWorksetVisibility(workset.Id, WorksetVisibility.Visible)
                    else:
                        # Alle anderen ausblenden
                        active_view.SetWorksetVisibility(workset.Id, WorksetVisibility.Hidden)

                t.Commit()
            except:
                t.RollBack()