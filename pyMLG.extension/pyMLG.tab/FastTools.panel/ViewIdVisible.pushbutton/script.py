# -*- coding: utf-8 -*-
__doc__ = "Exportiert die ViewId in den Parameter ViewId und macht die Id somit sichtbar"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import revit, forms


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

views = FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType()

t = Transaction(doc, 'Set ViewId Parameter')
t.Start()

try:
    for view in views:
        param = view.LookupParameter('ViewId')

        if param and not param.IsReadOnly:
            if param.StorageType == StorageType.String:
                param.Set(str(view.Id.IntegerValue))
            elif param.StorageType == StorageType.Integer:
                param.Set(view.Id.IntegerValue)
            elif param.StorageType == StorageType.ElementId:
                param.Set(view.Id)

except Exception as e:
    print("Fehler:", e)

t.Commit()

