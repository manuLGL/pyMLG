# -*- coding: utf-8 -*-
__doc__ = "Exportiert die ViewId in den Parameter ViewId und macht die Id somit sichtbar"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import revit, forms
from mlg_sprache import t


def id_wert(element_id):
    """Zahlenwert einer ElementId (Revit 2024+: .Value, IntegerValue entfällt ab 2026)."""
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

views = FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType()

transaktion = Transaction(doc, t('Set ViewId Parameter', u"Set ViewId parameter", u"Definir parámetro ViewId"))
transaktion.Start()

try:
    for view in views:
        param = view.LookupParameter('ViewId')

        if param and not param.IsReadOnly:
            if param.StorageType == StorageType.String:
                param.Set(str(id_wert(view.Id)))
            elif param.StorageType == StorageType.Integer:
                param.Set(id_wert(view.Id))
            elif param.StorageType == StorageType.ElementId:
                param.Set(view.Id)

except Exception as e:
    print(t(u"Fehler:", u"Error:", u"Error:"), e)

transaktion.Commit()

