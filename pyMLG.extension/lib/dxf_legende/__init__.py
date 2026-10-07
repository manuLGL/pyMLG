# -*- coding: utf-8 -*-
"""DxfLegend: Legende aus einer DXF-Datei als Revit-Legende nachbauen.

geometrie.py  Bögen als Polylinien-Ausbuchtung (bulge), Abtasten, Matrizen
dxf.py        eigener DXF-Leser (ASCII) - Linien, Texte, Schraffuren, Blöcke
              in Weltkoordinaten mit aufgelöster Farbe/Linientyp/Stärke
logik.py      Revit-freie Umsetzung: Massstab, Farben, Linienmuster, Namen
              -> Legendenplan in Papier-Millimetern (offline testbar)
revit.py      Legende/Zeichnungsansicht, Linienstile, Texttypen, Füllbereiche
fenster.py    WPF-Dialog mit Vorschau
"""
