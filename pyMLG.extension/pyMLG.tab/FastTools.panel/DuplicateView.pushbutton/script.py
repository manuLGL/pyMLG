# -*- coding: utf-8 -*-
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *

# Aktuelles Dokument
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# TaskDialog zur Info
result = TaskDialog.Show(
    "Views kopieren",
    "Wähle im Project Browser mehrere Views aus (Strg + Klick),\n" +
    "dann klicke OK um sie zu kopieren.",
    TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel
)

if result == TaskDialogResult.Ok:
    # Hole alle im Project Browser ausgewählten Views
    selected_ids = uidoc.Selection.GetElementIds()

    if selected_ids.Count == 0:
        TaskDialog.Show("Fehler", "Keine Views ausgewählt!\n\nBitte Views im Project Browser markieren.")
    else:
        # Filtere nur Views aus der Auswahl
        selected_views = []
        for elem_id in selected_ids:
            elem = doc.GetElement(elem_id)
            if isinstance(elem, View):
                # Prüfe ob duplizierbar
                if elem.CanViewBeDuplicated(ViewDuplicateOption.Duplicate):
                    selected_views.append(elem)

        if len(selected_views) == 0:
            TaskDialog.Show(
                "Fehler",
                "Keine duplizierbaren Views ausgewählt!\n\n" +
                "Hinweis: Schedules, Sheets und Legends können nicht kopiert werden."
            )
        else:
            # Transaction starten
            t = Transaction(doc, "Views kopieren")
            t.Start()

            try:
                created_views = []
                failed_views = []

                # Alle existierenden View-Namen sammeln
                all_view_names = [v.Name for v in FilteredElementCollector(doc).OfClass(View).ToElements()]

                for view in selected_views:
                    try:
                        # View duplizieren
                        new_view_id = view.Duplicate(ViewDuplicateOption.Duplicate)
                        new_view = doc.GetElement(new_view_id)

                        # Namen generieren mit automatischer Nummerierung
                        base_name = view.Name
                        counter = 1
                        new_name = "{} - Kopie {}".format(base_name, counter)

                        while new_name in all_view_names:
                            counter += 1
                            new_name = "{} - Kopie {}".format(base_name, counter)

                        # Namen setzen
                        new_view.Name = new_name
                        created_views.append(new_name)
                        all_view_names.append(new_name)

                    except Exception as ex:
                        failed_views.append(view.Name)

                # Transaction abschließen
                t.Commit()

                # Erfolgsmeldung
                message = ""

                if len(created_views) > 0:
                    message += "✓ {} View(s) erfolgreich kopiert:\n\n".format(len(created_views))
                    for name in created_views:
                        message += "  • {}\n".format(name)

                if len(failed_views) > 0:
                    message += "\n✗ {} View(s) konnten nicht kopiert werden:\n\n".format(len(failed_views))
                    for name in failed_views:
                        message += "  • {}\n".format(name)

                TaskDialog.Show("Ergebnis", message)

            except Exception as e:
                t.RollBack()
                TaskDialog.Show("Fehler", "Fehler beim Kopieren:\n{}".format(str(e)))