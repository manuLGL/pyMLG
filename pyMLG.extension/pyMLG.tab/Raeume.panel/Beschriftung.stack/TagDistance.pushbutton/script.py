# -*- coding: utf-8 -*-
__title__ = "WallTagsDistance"
__doc__ = "Setzt alle Wall Tags auf den gleichen Abstand zur Wand"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import revit, DB, forms

doc = revit.doc
uidoc = revit.uidoc


def move_tag_to_offset(tag, wall, offset):
    """Verschiebt das Tag auf den gewünschten Abstand zur Wand"""
    # Hole die aktuelle Position des Tags
    tag_head = tag.TagHeadPosition

    # Hole die Wandlinie
    location_curve = wall.Location
    if not location_curve:
        return False

    curve = location_curve.Curve

    # Finde den nächsten Punkt auf der Wandlinie zum Tag
    result = curve.Project(tag_head)
    if result:
        closest_point = result.XYZPoint

        # Berechne die Richtung vom Tag zur Wand
        direction_to_wall = (closest_point - tag_head).Normalize()

        # Berechne die neue Position mit dem gewünschten Abstand
        new_position = closest_point - direction_to_wall * offset

        # Setze die neue Position
        tag.TagHeadPosition = new_position

        # KORREKTUR: Überprüfe ob Tag einen Leader hat UND ob LeaderElbow existiert
        if tag.HasLeader:
            try:
                tag.LeaderEndCondition = LeaderEndCondition.Free
                # Versuche LeaderElbow zu setzen (funktioniert nicht bei allen Tag-Typen)
                if hasattr(tag, 'LeaderElbow'):
                    tag.LeaderElbow = closest_point - direction_to_wall * (offset * 0.7)
            except:
                # Wenn LeaderElbow nicht funktioniert, ignoriere es
                pass

        return True
    return False


# Sammle alle Wall Tags
collector = FilteredElementCollector(doc, doc.ActiveView.Id) \
    .OfCategory(BuiltInCategory.OST_WallTags) \
    .WhereElementIsNotElementType()

wall_tags = list(collector)

if not wall_tags:
    forms.alert("Keine Wall Tags in der aktuellen Ansicht gefunden.", exitscript=True)

# Frage den Benutzer nach dem gewünschten Abstand
distance_input = forms.ask_for_string(
    default="500",
    prompt="Gib den gewünschten Abstand in mm ein:",
    title="Abstand für Wall Tags"
)

if distance_input:
    try:
        # Konvertiere mm zu Fuß (Revit interne Einheit)
        distance_mm = float(distance_input)
        DESIRED_OFFSET = distance_mm / 304.8  # 1 Fuß = 304.8 mm
    except:
        forms.alert("Ungültige Eingabe. Verwende Standard-Abstand von 500mm.")
        DESIRED_OFFSET = 500 / 304.8
else:
    DESIRED_OFFSET = 500 / 304.8

# Starte eine Transaction
t = Transaction(doc, "Wall Tags ausrichten")
t.Start()

success_count = 0
failed_count = 0

try:
    for tag in wall_tags:
        # Hole die verknüpfte Wand
        tagged_element_ids = tag.GetTaggedLocalElementIds()

        # Konvertiere HashSet zu Liste
        if tagged_element_ids and tagged_element_ids.Count > 0:
            # Hole das erste Element aus dem HashSet
            element_id = list(tagged_element_ids)[0]
            wall = doc.GetElement(element_id)

            if wall and isinstance(wall, Wall):
                if move_tag_to_offset(tag, wall, DESIRED_OFFSET):
                    success_count += 1
                else:
                    failed_count += 1
            else:
                failed_count += 1
        else:
            failed_count += 1

    t.Commit()

    # Zeige Ergebnis
    message = "Fertig!\n\n"
    message += "{} Wall Tags erfolgreich ausgerichtet\n".format(success_count)
    if failed_count > 0:
        message += "{} Wall Tags konnten nicht ausgerichtet werden".format(failed_count)

    forms.alert(message, title="Ergebnis")

except Exception as e:
    t.RollBack()
    forms.alert("Fehler: {}".format(str(e)), title="Fehler")