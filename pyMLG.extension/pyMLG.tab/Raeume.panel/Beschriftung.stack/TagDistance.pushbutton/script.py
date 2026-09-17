# -*- coding: utf-8 -*-
__title__ = "WallTagsDistance"
__doc__ = "Setzt alle Wall Tags auf den gleichen Abstand zur Wand"

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from pyrevit import revit, DB, forms
from mlg_sprache import t

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
    forms.alert(t("Keine Wall Tags in der aktuellen Ansicht gefunden.", u"No wall tags found in the active view.", u"No se encontraron etiquetas de muro en la vista activa."), exitscript=True)

# Frage den Benutzer nach dem gewünschten Abstand
distance_input = forms.ask_for_string(
    default="500",
    prompt=t("Gib den gewünschten Abstand in mm ein:", u"Enter the desired distance in mm:", u"Introduzca la distancia deseada en mm:"),
    title=t("Abstand für Wall Tags", u"Wall tag distance", u"Distancia de etiquetas de muro")
)

if distance_input:
    try:
        # Konvertiere mm zu Fuß (Revit interne Einheit)
        distance_mm = float(distance_input)
        DESIRED_OFFSET = distance_mm / 304.8  # 1 Fuß = 304.8 mm
    except:
        forms.alert(t("Ungültige Eingabe. Verwende Standard-Abstand von 500mm.", u"Invalid input. Using the default distance of 500 mm.", u"Entrada no válida. Se usa la distancia por defecto de 500 mm."))
        DESIRED_OFFSET = 500 / 304.8
else:
    DESIRED_OFFSET = 500 / 304.8

# Starte eine Transaction
transaktion = Transaction(doc, t("Wall Tags ausrichten", u"Align wall tags", u"Alinear etiquetas de muro"))
transaktion.Start()

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

    transaktion.Commit()

    # Zeige Ergebnis
    message = t("Fertig!\n\n", u"Done!\n\n", u"¡Listo!\n\n")
    message += t("{} Wall Tags erfolgreich ausgerichtet\n", u"{} wall tags aligned successfully\n", u"{} etiquetas de muro alineadas correctamente\n").format(success_count)
    if failed_count > 0:
        message += t("{} Wall Tags konnten nicht ausgerichtet werden", u"{} wall tags could not be aligned", u"No se pudieron alinear {} etiquetas de muro").format(failed_count)

    forms.alert(message, title=t("Ergebnis", u"Result", u"Resultado"))

except Exception as e:
    transaktion.RollBack()
    forms.alert(t("Fehler: {}", u"Error: {}", u"Error: {}").format(str(e)), title=t("Fehler", u"Error", u"Error"))