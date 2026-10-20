# -*- coding: utf-8 -*-
"""Wohnflaeche - Flächentabelle nach WoFlV je Wohnungsplan.

Module:
    logik - Raumarten, Anrechnung nach WoFlV, Tabellenzeilen (ohne Revit-API,
            offline testbar: tools/test_wohnflaeche.py)
    revit - Räume der Grundrisse auf dem Plan finden, Dachschrägen messen,
            Tabelle als Gruppe auf den Plan zeichnen
"""

VERSION = u"1.0.0"
