# -*- coding: utf-8 -*-
"""
Fenster-Daten Export aus Fassadenmodell
========================================
Dieses Skript extrahiert alle Fenster-Informationen aus dem aktiven Modell
und speichert sie als JSON für die spätere Verarbeitung in den Wohnungsmodellen.

Verwendung: Als pyRevit Button im Fassadenmodell ausführen
"""

__title__ = "Export\nFenster"
__author__ = "Manuel"

import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog

import json
import os
import sys
import codecs  # Für Python 2 Kompatibilität

# .NET System Bibliothek
import System
from System.Collections.Generic import List

# Revit Dokument
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


def get_window_data(window):
    """
    Extrahiert alle relevanten Daten eines Fensters
    """
    try:
        # Location Point des Fensters (in Projektkoordinaten)
        location = window.Location
        if not isinstance(location, LocationPoint):
            return None

        point = location.Point

        # Fenster-Parameter auslesen
        width = window.LookupParameter("Breite") or window.LookupParameter("Width")
        height = window.LookupParameter("Höhe") or window.LookupParameter("Height")
        sill_height = window.LookupParameter("Brüstungshöhe") or window.LookupParameter("Sill Height")

        # Host-Wand Information
        host_wall = window.Host
        wall_id = host_wall.Id.IntegerValue if host_wall else -1

        # Orientierung des Fensters (wichtig für Platzierung)
        facing_orientation = window.FacingOrientation
        hand_orientation = window.HandOrientation

        window_data = {
            "id": window.Id.IntegerValue,
            "mark": window.get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString() or "",
            "type_name": window.Name,
            "family_name": window.Symbol.Family.Name,

            # Geometrische Daten
            "location": {
                "x": point.X,
                "y": point.Y,
                "z": point.Z
            },

            # Dimensionen (in Fuß - Revit intern)
            "width": width.AsDouble() if width else 0,
            "height": height.AsDouble() if height else 0,
            "sill_height": sill_height.AsDouble() if sill_height else 0,

            # Orientierung
            "facing_orientation": {
                "x": facing_orientation.X,
                "y": facing_orientation.Y,
                "z": facing_orientation.Z
            },
            "hand_orientation": {
                "x": hand_orientation.X,
                "y": hand_orientation.Y,
                "z": hand_orientation.Z
            },

            # Host Wand
            "host_wall_id": wall_id,

            # Level Information
            "level_id": window.LevelId.IntegerValue,
            "level_name": doc.GetElement(window.LevelId).Name if window.LevelId != ElementId.InvalidElementId else ""
        }

        return window_data

    except Exception as e:
        print("Fehler bei Fenster ID {}: {}".format(window.Id, str(e)))
        return None


def get_linked_models():
    """
    Findet alle Revit-Verknüpfungen im aktuellen Modell
    """
    collector = FilteredElementCollector(doc)
    links = collector.OfClass(RevitLinkInstance).ToElements()

    link_info = []
    for link in links:
        try:
            link_type = doc.GetElement(link.GetTypeId())

            if not link_type:
                print("Warnung: Link-Typ nicht gefunden für Link ID {}".format(link.Id))
                continue

            # Name sicher auslesen
            try:
                link_name = link_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
            except:
                try:
                    link_name = link_type.Name
                except:
                    link_name = "Unknown_Link_{}".format(link.Id.IntegerValue)

            # Pfad zur verknüpften Datei
            try:
                external_file_ref = link_type.GetExternalFileReference()
                link_path = ModelPathUtils.ConvertModelPathToUserVisiblePath(
                    external_file_ref.GetPath()
                )
            except:
                print("Warnung: Pfad nicht gefunden für Link: {}".format(link_name))
                continue

            # Transform der Verknüpfung (für Origin to Origin ist das meist Identity)
            try:
                transform = link.GetTotalTransform()
                transform_data = {
                    "origin": {
                        "x": transform.Origin.X,
                        "y": transform.Origin.Y,
                        "z": transform.Origin.Z
                    },
                    "basis_x": {
                        "x": transform.BasisX.X,
                        "y": transform.BasisX.Y,
                        "z": transform.BasisX.Z
                    },
                    "basis_y": {
                        "x": transform.BasisY.X,
                        "y": transform.BasisY.Y,
                        "z": transform.BasisY.Z
                    },
                    "basis_z": {
                        "x": transform.BasisZ.X,
                        "y": transform.BasisZ.Y,
                        "z": transform.BasisZ.Z
                    }
                }
            except:
                # Fallback: Identity Transform
                transform_data = {
                    "origin": {"x": 0.0, "y": 0.0, "z": 0.0},
                    "basis_x": {"x": 1.0, "y": 0.0, "z": 0.0},
                    "basis_y": {"x": 0.0, "y": 1.0, "z": 0.0},
                    "basis_z": {"x": 0.0, "y": 0.0, "z": 1.0}
                }

            link_data = {
                "name": link_name,
                "path": link_path,
                "id": link.Id.IntegerValue,
                "transform": transform_data
            }
            link_info.append(link_data)

        except Exception as e:
            print("Fehler bei Link ID {}: {}".format(link.Id, str(e)))
            continue

    return link_info


def main():
    """
    Hauptfunktion: Sammelt alle Daten und exportiert als JSON
    """

    try:
        # Alle Fenster im Modell sammeln
        collector = FilteredElementCollector(doc)
        windows = collector.OfCategory(BuiltInCategory.OST_Windows) \
            .WhereElementIsNotElementType() \
            .ToElements()

        print("Gefundene Fenster: {}".format(len(windows)))

        # Fenster-Daten extrahieren
        windows_data = []
        for window in windows:
            data = get_window_data(window)
            if data:
                windows_data.append(data)

        print("Erfolgreich verarbeitet: {}".format(len(windows_data)))

        # Verknüpfungen finden
        print("\nSuche nach Verknüpfungen...")
        linked_models = get_linked_models()
        print("Gefundene Verknüpfungen: {}".format(len(linked_models)))

        # Details der Verknüpfungen ausgeben
        if linked_models:
            print("\nVerknüpfte Modelle:")
            for i, link in enumerate(linked_models, 1):
                print("  {}. {}".format(i, link['name']))
                print("     Pfad: {}".format(link['path']))
        else:
            print("\n⚠️ WARNUNG: Keine Verknüpfungen gefunden!")
            print("   Sind die Verknüpfungen geladen?")

        # Gesamtdaten-Struktur
        export_data = {
            "project_name": doc.Title,
            "project_path": doc.PathName,
            "windows": windows_data,
            "linked_models": linked_models,
            "export_date": str(System.DateTime.Now)
        }

        # Export-Pfad wählen (gleicher Ordner wie Revit-Datei)
        project_dir = os.path.dirname(doc.PathName)
        export_path = os.path.join(project_dir, "windows_export_data.json")

        # Als JSON speichern (Python 2 kompatibel)
        with codecs.open(export_path, 'w', encoding='utf-8') as f:
            json.dump(export_data, f, indent=2, ensure_ascii=False)

        # Erfolgsmeldung
        message = "Fenster-Daten exportiert:\n\n{} Fenster\n{} Verknüpfungen\n\nGespeichert unter:\n{}".format(
            len(windows_data),
            len(linked_models),
            export_path
        )

        if len(linked_models) == 0:
            message += "\n\n⚠️ WARNUNG: Keine Verknüpfungen gefunden!\nBitte Verknüpfungen laden und erneut exportieren."

        TaskDialog.Show("Export erfolgreich", message)

    except Exception as e:
        error_message = "FEHLER beim Export:\n\n{}".format(str(e))
        print(error_message)
        import traceback
        traceback.print_exc()
        TaskDialog.Show("Export fehlgeschlagen", error_message)


# Skript ausführen
if __name__ == '__main__':
    main()