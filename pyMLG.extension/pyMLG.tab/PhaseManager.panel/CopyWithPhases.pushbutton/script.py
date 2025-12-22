# -*- coding: utf-8 -*-
"""Copy With Phases"""

__title__ = "Copy With\nPhases"
__author__ = "Manuel"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import revit
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc

selected_ids = uidoc.Selection.GetElementIds()

if not selected_ids or selected_ids.Count == 0:
    print("Keine Elemente ausgewählt!")
else:
    original_phases = {}
    for elem_id in selected_ids:
        elem = doc.GetElement(elem_id)
        if elem:
            phase_created = elem.get_Parameter(BuiltInParameter.PHASE_CREATED)
            phase_demolished = elem.get_Parameter(BuiltInParameter.PHASE_DEMOLISHED)
            original_phases[elem_id.IntegerValue] = {
                'created': phase_created.AsElementId() if phase_created else None,
                'demolished': phase_demolished.AsElementId() if phase_demolished else None
            }

    t = None
    try:
        print("Wähle Basispunkt...")
        base_point = uidoc.Selection.PickPoint("Basispunkt wählen")
        print("Wähle Zielpunkt...")
        target_point = uidoc.Selection.PickPoint("Zielpunkt wählen")
        translation = target_point - base_point

        t = Transaction(doc, "Copy with Phases")
        t.Start()

        element_ids_list = List[ElementId](selected_ids)
        copied_ids = ElementTransformUtils.CopyElements(doc, element_ids_list, translation)

        for i, copied_id in enumerate(copied_ids):
            original_id = list(selected_ids)[i]
            copied_elem = doc.GetElement(copied_id)
            if original_id.IntegerValue in original_phases:
                phase_info = original_phases[original_id.IntegerValue]
                if phase_info['created']:
                    param = copied_elem.get_Parameter(BuiltInParameter.PHASE_CREATED)
                    if param and not param.IsReadOnly:
                        param.Set(phase_info['created'])
                if phase_info['demolished']:
                    param = copied_elem.get_Parameter(BuiltInParameter.PHASE_DEMOLISHED)
                    if param and not param.IsReadOnly:
                        param.Set(phase_info['demolished'])

        uidoc.Selection.SetElementIds(copied_ids)
        t.Commit()

    except:
        if t and t.HasStarted():
            t.RollBack()
        print("Abgebrochen")