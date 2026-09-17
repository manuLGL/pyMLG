# -*- coding: utf-8 -*-
"""Paste Aligned to Current View With Phases
Fügt Elemente ausgerichtet zur aktuellen Ansicht ein mit Phase-Beibehaltung
"""

__title__ = "Paste Aligned\nView"
__author__ = "Manuel"

from Autodesk.Revit.DB import *
from pyrevit import revit
from System.Collections.Generic import List

doc = revit.doc
uidoc = revit.uidoc

selected_ids = uidoc.Selection.GetElementIds()

if not selected_ids or selected_ids.Count == 0:
    print("Keine Elemente ausgewählt!")
else:
    active_view = doc.ActiveView

    if not hasattr(active_view, 'GenLevel') or not active_view.GenLevel:
        print("Aktive Ansicht hat kein zugeordnetes Level!")
    else:
        target_level = active_view.GenLevel

        # Speichere Original-Phasen und ermittle Basis-Level
        original_phases = {}
        base_level = None

        for elem_id in selected_ids:
            elem = doc.GetElement(elem_id)
            if elem:
                phase_created = elem.get_Parameter(BuiltInParameter.PHASE_CREATED)
                phase_demolished = elem.get_Parameter(BuiltInParameter.PHASE_DEMOLISHED)
                original_phases[elem_id.IntegerValue] = {
                    'created': phase_created.AsElementId() if phase_created else None,
                    'demolished': phase_demolished.AsElementId() if phase_demolished else None
                }

                if not base_level:
                    level_param = elem.get_Parameter(BuiltInParameter.LEVEL_PARAM)
                    if level_param:
                        base_level = doc.GetElement(level_param.AsElementId())

        # Berechne Z-Offset
        if base_level:
            z_offset = target_level.Elevation - base_level.Elevation
        else:
            z_offset = 0

        translation = XYZ(0, 0, z_offset)

        t = None
        try:
            t = Transaction(doc, "Paste Aligned to View")
            t.Start()

            element_ids_list = List[ElementId](selected_ids)
            copied_ids = ElementTransformUtils.CopyElements(doc, element_ids_list, translation)

            # WICHTIG: Erst Level setzen, dann Phasen
            for i, copied_id in enumerate(copied_ids):
                original_id = list(selected_ids)[i]
                copied_elem = doc.GetElement(copied_id)

                # 1. LEVEL SETZEN (wichtigster Schritt!)
                level_param = copied_elem.get_Parameter(BuiltInParameter.LEVEL_PARAM)
                if level_param and not level_param.IsReadOnly:
                    level_param.Set(target_level.Id)

                # Für Wände: Base Constraint und Top Constraint
                base_constraint = copied_elem.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)
                if base_constraint and not base_constraint.IsReadOnly:
                    base_constraint.Set(target_level.Id)

                # Base Offset auf 0 setzen damit es wirklich auf dem Level sitzt
                base_offset = copied_elem.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
                if base_offset and not base_offset.IsReadOnly:
                    original_elem = doc.GetElement(original_id)
                    original_offset = original_elem.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
                    if original_offset:
                        base_offset.Set(original_offset.AsDouble())

                # 2. PHASEN SETZEN
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
            print("{} Element(e) auf Level '{}' eingefuegt".format(copied_ids.Count, target_level.Name))

        except Exception as e:
            if t and t.HasStarted():
                t.RollBack()
            print("Fehler: {}".format(str(e)))