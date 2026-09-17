# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import WorksetVisibility, Transaction, WorksetId, FilteredWorksetCollector, WorksetKind
from pyrevit import revit
from mlg_sprache import t

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

            transaktion = Transaction(doc, t("Nur ausgewaehltes Workset anzeigen", u"Show selected workset only", u"Mostrar solo el subproyecto seleccionado"))
            transaktion.Start()

            try:
                for workset in all_worksets:
                    if workset.Id == selected_workset_id:
                        # Ausgewähltes Workset einblenden
                        active_view.SetWorksetVisibility(workset.Id, WorksetVisibility.Visible)
                    else:
                        # Alle anderen ausblenden
                        active_view.SetWorksetVisibility(workset.Id, WorksetVisibility.Hidden)

                transaktion.Commit()
            except:
                transaktion.RollBack()