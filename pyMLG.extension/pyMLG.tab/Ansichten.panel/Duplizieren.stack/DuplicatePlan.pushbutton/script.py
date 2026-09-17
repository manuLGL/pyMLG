# -*- coding: utf-8 -*-
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from mlg_sprache import t

# Aktuelles Dokument
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

try:
    # Hole ausgewählte Sheets
    selected_ids = uidoc.Selection.GetElementIds()

    if selected_ids.Count == 0:
        # Nichts ausgewählt -> Still beenden (kein Fehler)
        import sys

        sys.exit()

    selected_sheets = []
    for elem_id in selected_ids:
        try:
            elem = doc.GetElement(elem_id)
            if isinstance(elem, ViewSheet):
                selected_sheets.append(elem)
        except:
            pass  # Fehlerhafte Elements überspringen

    # Keine gültigen Sheets gefunden
    if len(selected_sheets) == 0:
        import sys

        sys.exit()

    # Transaction starten (mit Fehlerbehandlung)
    transaktion = Transaction(doc, t("Sheets kopieren", u"Copy sheets", u"Copiar planos"))
    transaktion.Start()

    try:
        # Titleblock-Typen holen
        titleblock_types = FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_TitleBlocks) \
            .WhereElementIsElementType() \
            .ToElements()

        if len(titleblock_types) == 0:
            # Keine Titleblocks im Projekt
            transaktion.RollBack()
            TaskDialog.Show(t("Fehler", u"Error", u"Error"), t("Keine Titleblock-Typen im Projekt gefunden!", u"No title block types found in the project!", u"¡No se encontraron tipos de cajetín en el proyecto!"))
            import sys

            sys.exit()

        default_tb = titleblock_types[0].Id

        # Alle existierenden Sheet-Nummern (cached für Performance)
        all_numbers = set()
        try:
            all_sheets = FilteredElementCollector(doc).OfClass(ViewSheet).ToElements()
            all_numbers = set([s.SheetNumber for s in all_sheets])
        except:
            # Fallback: Leere Liste
            all_numbers = set()

        # Zähler für Erfolg/Fehler
        success_count = 0
        failed_count = 0

        for sheet in selected_sheets:
            try:
                # Titleblock vom Original holen
                tbs = FilteredElementCollector(doc, sheet.Id) \
                    .OfCategory(BuiltInCategory.OST_TitleBlocks) \
                    .WhereElementIsNotElementType() \
                    .ToElements()

                tb_type = tbs[0].GetTypeId() if len(tbs) > 0 else default_tb

                # Neues Sheet erstellen
                new_sheet = ViewSheet.Create(doc, tb_type)

                # Nummer generieren (mit Limit gegen Endlosschleife)
                base_num = sheet.SheetNumber
                counter = 1
                max_attempts = 1000  # Sicherheit gegen Endlosschleife
                new_num = t("{} - Kopie {}", u"{} - Copy {}", u"{} - Copia {}").format(base_num, counter)

                while new_num in all_numbers and counter < max_attempts:
                    counter += 1
                    new_num = t("{} - Kopie {}", u"{} - Copy {}", u"{} - Copia {}").format(base_num, counter)

                # Wenn 1000 Kopien erreicht -> Eindeutige ID anhängen
                if counter >= max_attempts:
                    import random

                    new_num = "{}-K{}".format(base_num, random.randint(1000, 9999))

                # Sheet-Nummer und Name setzen
                new_sheet.SheetNumber = new_num
                new_sheet.Name = t("{} - Kopie {}", u"{} - Copy {}", u"{} - Copia {}").format(sheet.Name, counter)
                all_numbers.add(new_num)

                success_count += 1

            except Exception as ex:
                # Einzelnes Sheet fehlgeschlagen -> weiter mit nächstem
                failed_count += 1
                # Bei kritischem Fehler: Sheet wieder löschen falls erstellt
                try:
                    if 'new_sheet' in locals() and new_sheet:
                        doc.Delete(new_sheet.Id)
                except:
                    pass

        # Transaction abschließen
        if success_count > 0:
            transaktion.Commit()
            # Optional: Stille Erfolgsmeldung (auskommentiert für "still mode")
            # TaskDialog.Show("Erfolg", "{} Sheet(s) kopiert".format(success_count))
        else:
            # Nichts erfolgreich -> Rollback
            transaktion.RollBack()
            TaskDialog.Show(t("Fehler", u"Error", u"Error"), t("Keine Sheets konnten kopiert werden.", u"No sheets could be copied.", u"No se pudo copiar ningún plano."))

    except Exception as e:
        # Transaction fehlgeschlagen -> Rollback
        transaktion.RollBack()
        TaskDialog.Show(t("Fehler", u"Error", u"Error"), t("Fehler beim Kopieren:\n{}", u"Error while copying:\n{}", u"Error al copiar:\n{}").format(str(e)))

except Exception as e:
    # Kritischer Fehler außerhalb der Transaction
    TaskDialog.Show(t("Kritischer Fehler", u"Critical error", u"Error crítico"), t("Unerwarteter Fehler:\n{}", u"Unexpected error:\n{}", u"Error inesperado:\n{}").format(str(e)))