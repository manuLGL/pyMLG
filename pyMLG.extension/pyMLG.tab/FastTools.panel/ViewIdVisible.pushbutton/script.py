# -*- coding: utf-8 -*-
__doc__ = "ExcelExport Exportiert ausgewählte Listen nach Excel"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import revit, forms
import sys
import subprocess
import clr
import System

clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop import Excel

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Wir filtern alle Schedules aus dem Proyect
view_Schedule = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
scheduleList = []
schedule_dict = {}

for sched in view_Schedule:
    if not sched.IsInternalKeynoteSchedule and not sched.IsTitleblockRevisionSchedule:
        scheduleName = sched.Name
        scheduleList.append(scheduleName)
        schedule_dict[scheduleName] = sched

# 2.Erstellen einer UI zur Auswahl der Schedules
selected_names = forms.SelectFromList.show(
    scheduleList,
    title="Select Schedules",
    multiselect=True,
)

if not selected_names:
    sys.exit()

# 4. Nach Excel exportieren
datei_pfad = forms.save_file(
    file_ext="xlsx",
    default_name="Schedule Export"
)

if not datei_pfad:
    sys.exit()

if selected_names and datei_pfad:
    print("User selected:" + datei_pfad)

    # Excel starten
    excel_app = Excel.ApplicationClass()
    excel_app.Visible = False
    workbook = excel_app.Workbooks.Add()

    # Datenschleife um Excel zu füllen:
    for name in selected_names:
        schedule_obj = schedule_dict[name]

        # TableData holen
        tableData = schedule_obj.GetTableData()
        sectionData = tableData.GetSectionData(SectionType.Body)
        headerData = tableData.GetSectionData(SectionType.Header)

        # Worksheet erstellen
        worksheet = workbook.Worksheets.Add()
        worksheet.Name = name

        # Header-Zeilen kopieren (falls vorhanden)
        if headerData:
            header_rows = headerData.NumberOfRows
            header_cols = headerData.NumberOfColumns
            for row in range(header_rows):
                for col in range(header_cols):
                    wert = headerData.GetCellText(row, col)
                    worksheet.Cells[row + 1, col + 1] = wert
            start_row = header_rows + 1
        else:
            start_row = 1

        # Body-Zeilen kopieren (HYBRID: GetCellText + direkter Parameter-Zugriff)
        body_rows = sectionData.NumberOfRows
        body_cols = sectionData.NumberOfColumns

        # Elemente aus Schedule holen
        collector = FilteredElementCollector(doc, schedule_obj.Id)
        element_list = list(collector.ToElements())

        schedule_definition = schedule_obj.Definition
        field_count = schedule_definition.GetFieldCount()

        element_index = 0

        for row in range(body_rows):
            # Prüfe ob es eine Element-Zeile ist (nicht Gruppierung/Summe)
            first_cell = sectionData.GetCellText(row, 0)

            # Wenn erste Spalte nicht leer → könnte Element-Zeile sein
            is_element_row = first_cell != "" and element_index < len(element_list)

            print("Zeile " + str(row) + ": first_cell='" + first_cell + "', is_element_row=" + str(
                is_element_row) + ", element_index=" + str(element_index))  # DEBUG

            for col in range(body_cols):
                wert = sectionData.GetCellText(row, col)

                print("  Spalte " + str(col) + ": GetCellText='" + wert + "'")  # DEBUG

                # Wenn Wert leer UND es ist eine Element-Zeile → vom Element holen
                if wert == "" and is_element_row:
                    element = element_list[element_index]
                    field = schedule_definition.GetField(col)
                    param_id = field.ParameterId

                    print("    Versuche Parameter zu holen: param_id=" + str(param_id.IntegerValue))  # DEBUG

                    # Versuche Parameter zu holen
                    try:
                        built_in = System.Enum.ToObject(BuiltInParameter, param_id.IntegerValue)
                        param = element.get_Parameter(built_in)

                        print("    param gefunden: " + str(param is not None))  # DEBUG

                        if param:
                            if param.StorageType == StorageType.String:
                                wert = param.AsString() or ""
                            elif param.StorageType == StorageType.Double:
                                wert = param.AsValueString() or ""
                            elif param.StorageType == StorageType.Integer:
                                wert = str(param.AsInteger())
                            print("    Wert aus Parameter: '" + wert + "'")  # DEBUG
                    except Exception as e:
                        print("    Fehler: " + str(e))  # DEBUG
                        wert = ""

                worksheet.Cells[row + start_row, col + 1] = wert

            # Erhöhe element_index nur bei Element-Zeilen
            if is_element_row:
                element_index += 1

    # Speichern
    workbook.SaveAs(datei_pfad)
    workbook.Close()
    excel_app.Quit()
    print("Excel wurde erstellt!")
else:
    print("User didn't selected anything")