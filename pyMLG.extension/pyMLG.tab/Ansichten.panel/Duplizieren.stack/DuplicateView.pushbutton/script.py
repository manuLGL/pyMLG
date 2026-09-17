# -*- coding: utf-8 -*-
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from mlg_sprache import t

# Aktuelles Dokument
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# TaskDialog zur Info
result = TaskDialog.Show(
    t("Views kopieren", u"Copy views", u"Copiar vistas"),
    t("Wähle im Project Browser mehrere Views aus (Strg + Klick),\n", u"Select several views in the Project Browser (Ctrl + click),\n", u"Seleccione varias vistas en el navegador de proyectos (Ctrl + clic),\n") +
    t("dann klicke OK um sie zu kopieren.", u"then click OK to copy them.", u"y pulse Aceptar para copiarlas."),
    TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel
)

if result == TaskDialogResult.Ok:
    # Hole alle im Project Browser ausgewählten Views
    selected_ids = uidoc.Selection.GetElementIds()

    if selected_ids.Count == 0:
        TaskDialog.Show(t("Fehler", u"Error", u"Error"), t("Keine Views ausgewählt!\n\nBitte Views im Project Browser markieren.", u"No views selected!\n\nPlease select views in the Project Browser.", u"¡No hay vistas seleccionadas!\n\nSeleccione vistas en el navegador de proyectos."))
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
                t("Fehler", u"Error", u"Error"),
                t("Keine duplizierbaren Views ausgewählt!\n\n", u"No duplicable views selected!\n\n", u"¡No hay vistas duplicables seleccionadas!\n\n") +
                t("Hinweis: Schedules, Sheets und Legends können nicht kopiert werden.", u"Note: schedules, sheets and legends cannot be copied.", u"Nota: no se pueden copiar tablas, planos ni leyendas.")
            )
        else:
            # Transaction starten
            transaktion = Transaction(doc, t("Views kopieren", u"Copy views", u"Copiar vistas"))
            transaktion.Start()

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
                        new_name = t("{} - Kopie {}", u"{} - Copy {}", u"{} - Copia {}").format(base_name, counter)

                        while new_name in all_view_names:
                            counter += 1
                            new_name = t("{} - Kopie {}", u"{} - Copy {}", u"{} - Copia {}").format(base_name, counter)

                        # Namen setzen
                        new_view.Name = new_name
                        created_views.append(new_name)
                        all_view_names.append(new_name)

                    except Exception as ex:
                        failed_views.append(view.Name)

                # Transaction abschließen
                transaktion.Commit()

                # Erfolgsmeldung
                message = ""

                if len(created_views) > 0:
                    message += t("✓ {} View(s) erfolgreich kopiert:\n\n", u"✓ {} view(s) copied successfully:\n\n", u"✓ {} vista(s) copiadas correctamente:\n\n").format(len(created_views))
                    for name in created_views:
                        message += "  • {}\n".format(name)

                if len(failed_views) > 0:
                    message += t("\n✗ {} View(s) konnten nicht kopiert werden:\n\n", u"\n✗ {} view(s) could not be copied:\n\n", u"\n✗ No se pudieron copiar {} vista(s):\n\n").format(len(failed_views))
                    for name in failed_views:
                        message += "  • {}\n".format(name)

                TaskDialog.Show(t("Ergebnis", u"Result", u"Resultado"), message)

            except Exception as e:
                transaktion.RollBack()
                TaskDialog.Show(t("Fehler", u"Error", u"Error"), t("Fehler beim Kopieren:\n{}", u"Error while copying:\n{}", u"Error al copiar:\n{}").format(str(e)))