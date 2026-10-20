# -*- coding: utf-8 -*-
"""Raum-Automatik - Räume aus der Einrichtung erkennen und anlegen.

Module:
    logik - Regeln (Bett, Sanitär, Küche + Sofa, Türen), Zimmernummern je
            Wohnung (ohne Revit-API, offline testbar:
            tools/test_raum_automatik.py)
    revit - freie Bereiche des Geschosses finden, Einrichtung und Türen
            zuordnen, Räume anlegen und Parameter füllen
"""

VERSION = u"1.0.0"
