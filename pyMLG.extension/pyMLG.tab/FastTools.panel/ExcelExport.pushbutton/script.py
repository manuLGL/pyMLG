# -*- coding: utf-8 -*-
__doc__ = "ExcelExport Exportiert ausgewählte Listen nach Excel"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import revit, forms
import sys
import subprocess
import clr

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
    import sys

    sys.exit()

# 4. Nach Excel exportieren

# Speicherort wählen
datei_pfad = forms.save_file(
    file_ext="xlsx",
    default_name="Schedule Export"
)

if selected_names and datei_pfad:
    print("User selected:" + datei_pfad)

    # Excel starten
    excel_app = Excel.ApplicationClass()
    excel_app.Visible = False
    workbook = excel_app.Workbooks.Add()

    #Datenschlefe um Excel zu füllen:
    for name in selected_names:
        schedule_obj = schedule_dict[name]
        tableData = schedule_obj.GetTableData()
        body_Headers = tableData.GetSectionData(SectionType.Body)

        # Worksheet erstellen
        worksheet = workbook.Worksheets.Add()
        worksheet.Name = name

        #Ids holen von Elementen aus Liste:
        collector = FilteredElementCollector(doc, schedule_obj.Id)
        element_List = collector.ToElements()


        zeile = 3
        for element in element_List:
            elem_id = element.Id.IntegerValue
            elem_type_id = element.GetTypeId()

            if elem_type_id != ElementId.InvalidElementId:
                elem_type_obj = doc.GetElement(elem_type_id)
                if elem_type_obj:
                    elem_type_name = Element.Name.GetValue(elem_type_obj)
                else:
                    elem_type_name = "Kein Typ (None)"
            else:
                elem_type_name = "Invalid ID"

            worksheet.Cells[zeile, 1] = elem_id
            worksheet.Cells[zeile, 2] = elem_type_name

            zeile = zeile + 1

        anzahl_Spalten = body_Headers.NumberOfColumns
        anzahl_Reihen = body_Headers.NumberOfRows

        # Daten schreiben
        for reihe in range(anzahl_Reihen):
            for spalte in range(anzahl_Spalten):
                zellenText = body_Headers.GetCellText(reihe, spalte)
                worksheet.Cells[reihe + 1, spalte + 3] = zellenText

    # Speichern
    workbook.SaveAs(datei_pfad)
    workbook.Close()
    excel_app.Quit()
    print("Excel wurde erstellt!")
else:
    print("User didn't selected anything")