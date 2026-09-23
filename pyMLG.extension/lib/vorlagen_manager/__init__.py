# -*- coding: utf-8 -*-
"""Gemeinsame Logik des Ansichtsvorlagen-Managers (View Templates).

    logik   - Suche, Zusammenfassen von Werten mehrerer Vorlagen,
              Reihenfolge und Namensprüfung (ohne Revit-Importe, offline
              testbar)
    revit   - Lesen und Schreiben der Vorlagen über die Revit-API
    bloecke - Editoren für die Einstellungen, die kein einfacher Wert sind
              (Kategorien, Filter, Bearbeitungsbereiche, Ansichtsbereich)
    uebertragen - einzelne Einstellungen von einer Vorlage auf mehrere
              kopieren
    fenster - WPF-Hauptfenster
"""
