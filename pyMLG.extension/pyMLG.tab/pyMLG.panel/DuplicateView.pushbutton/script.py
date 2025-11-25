# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


#_________________________________________________________________________
#_________________________________________________________________________
# Nombre de la plantilla que quieres aplicar
template_name = "WIP_Wall_Control"

# Buscar la plantilla
templates = FilteredElementCollector(doc).OfClass(View).ToElements()
template = None
for t in templates:
    if t.IsTemplate and t.Name == template_name:
        template = t
        break

if not template:
    raise Exception("No se encontró la plantilla con el nombre especificado.")

# Vista seleccionada (ejemplo: primera vista seleccionada)
selected_ids = uidoc.Selection.GetElementIds()
if not selected_ids:
    raise Exception("Selecciona una vista primero.")

view = doc.GetElement(selected_ids[0])
if not isinstance(view, View):
    raise Exception("El elemento seleccionado no es una vista.")

#_________________________________________________________________________
#_________________________________________________________________________

# Funktion für eindeutige Blattnummer
def get_unique_sheet_number(prefix, start):
    existing_numbers = {s.SheetNumber for s in FilteredElementCollector(doc).OfClass(ViewSheet)}
    num = start
    while "{}-{:03d}".format(prefix, num) in existing_numbers:
        num += 1
    return "{}-{:03d}".format(prefix, num)

#_________________________________________________________________________
#_________________________________________________________________________

# Alle Plankopf-Typen sammeln
titleblock_types = FilteredElementCollector(doc)\
    .OfCategory(BuiltInCategory.OST_TitleBlocks)\
    .WhereElementIsElementType()\
    .ToElements()

if not titleblock_types:
    TaskDialog.Show("Fehler", "Keine Planvorlage (Titleblock) im Projekt gefunden.")
    raise SystemExit

# Plankopf mit Name "A0" suchen

titleblock_type = None
for tb in titleblock_types:
    tb_name = tb.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
    if tb_name:
        parts = [p.strip() for p in tb_name.split(":")]
        if len(parts) == 2 and parts[0] == "B+K Plankopf BA A3" and parts[1] == "B+K Plankopf BA A3":
            titleblock_type = tb
            break

if not titleblock_type:
    TaskDialog.Show("Fehler", "Kein Titleblock mit diesem Namen gefunden.")
    raise SystemExit

# Aktuelle Auswahl filtern
selected_ids = uidoc.Selection.GetElementIds()
views = []
for id in selected_ids:
    el = doc.GetElement(id)
    if isinstance(el, View) and not el.IsTemplate and el.CanBePrinted:
        views.append(el)

if not views:
    TaskDialog.Show("Fehler", "Bitte wähle gültige Ansichten aus.")
    raise SystemExit


#_________________________________________________________________________________________
#AUSFÜHRUNG TRANSACTION
#_________________________________________________________________________________________

t = Transaction(doc, "Create Sheet View")
t.Start()

for i, view in enumerate(views):
    try:
        # Aplicar la plantilla
        view.ViewTemplateId = template.Id
# ________________________________________________________________________________________
#________________________________________________________________________________________
        new_sheet = ViewSheet.Create(doc, titleblock_type.Id)
        new_sheet.Name = view.Name
        new_sheet.SheetNumber = get_unique_sheet_number("AP", i + 1)

        x_cm = -57.0
        y_cm = 40.0

        # Umrechnung in Fuß
        x_ft = x_cm / 30.48
        y_ft = y_cm / 30.48

        # Punkt erstellen
        point = XYZ(x_ft, y_ft, 0)

        # Ansicht platzieren
        vp = Viewport.Create(doc, new_sheet.Id, view.Id, point)
        if vp is None:
            print("Viewport konnte nicht erstellt werden für:", view.Name)

        print("Plan erstellt für Ansicht:", view.Name)
    except Exception as e:
        print("Fehler bei {}: {}".format(view.Name, e))

t.Commit()

TaskDialog.Show("Olé", "You created {} sheets.".format(len(views)))