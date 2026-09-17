# -*- coding: utf-8 -*-
"""Revit-freier Auswahlbaum von FilterMore.

    Alle
    └─ Kategorie
       └─ Familie
          └─ Typ
             └─ Element (Blatt, trägt die Element-Id)

Die Markierung liegt nicht in den Knoten, sondern als Menge von Element-Ids
beim Aufrufer. Jeder Knoten kennt die Ids aller Elemente darunter; daraus
ergibt sich sein Zustand (alle / keine / teilweise markiert).
"""

import re

WURZEL = 0
KATEGORIE = 1
FAMILIE = 2
TYP = 3
ELEMENT = 4


class Knoten(object):
    def __init__(self, name, ebene):
        self.name = name
        self.ebene = ebene
        self.kinder = []
        self.ids = []              # Element-Ids aller Blätter darunter
        self.element_id = None     # nur Blätter
        self._index = {}
        # Oberfläche (fenster.py)
        self.item = None
        self.box = None
        self.zaehler = None
        self.geladen = False

    def kind(self, name, ebene):
        knoten = self._index.get(name)
        if knoten is None:
            knoten = self._index[name] = Knoten(name, ebene)
            self.kinder.append(knoten)
        return knoten

    def __repr__(self):
        return "<Knoten %s '%s' %d>" % (self.ebene, self.name, len(self.ids))


def natuerlich(text):
    """Sortierschlüssel: 'Wand 2' vor 'Wand 10', ohne Groß-/Kleinschreibung.
    re.split mit Gruppe liefert immer abwechselnd Text und Zahl."""
    teile = re.split(r"(\d+)", text or u"")
    return [int(teil) if i % 2 else teil.casefold() for i, teil in enumerate(teile)]


def baue(datensaetze, wurzelname=u"Alle"):
    """[(Kategorie, Familie, Typ, Bezeichnung, Id)] -> Wurzelknoten."""
    wurzel = Knoten(wurzelname, WURZEL)
    for kategorie, familie, typ, bezeichnung, element_id in datensaetze:
        pfad = [wurzel]
        pfad.append(pfad[-1].kind(kategorie, KATEGORIE))
        pfad.append(pfad[-1].kind(familie, FAMILIE))
        pfad.append(pfad[-1].kind(typ, TYP))
        blatt = Knoten(bezeichnung, ELEMENT)
        blatt.element_id = element_id
        blatt.ids.append(element_id)
        pfad[-1].kinder.append(blatt)
        for knoten in pfad:
            knoten.ids.append(element_id)
    _sortiere(wurzel)
    return wurzel


def _sortiere(knoten):
    knoten.kinder.sort(key=lambda k: (natuerlich(k.name), k.element_id or 0))
    for kind in knoten.kinder:
        _sortiere(kind)


def anzahl_markiert(knoten, markiert):
    return sum(1 for i in knoten.ids if i in markiert)


def zustand(knoten, markiert):
    """True (alle), False (keine) oder None (teilweise) - wie IsChecked
    einer Kontrollbox mit drei Zuständen."""
    n = anzahl_markiert(knoten, markiert)
    if n == 0:
        return False
    if n == len(knoten.ids):
        return True
    return None


def zaehler_text(knoten, markiert):
    n = anzahl_markiert(knoten, markiert)
    if n in (0, len(knoten.ids)):
        return u"%d" % len(knoten.ids)
    return u"%d / %d" % (n, len(knoten.ids))


def anzeige(knoten, markiert):
    """(Zustand, Zählertext) mit nur einer Zählung - wird bei jeder
    Markierungsänderung für alle sichtbaren Knoten aufgerufen."""
    n = anzahl_markiert(knoten, markiert)
    gesamt = len(knoten.ids)
    if n == 0:
        return False, u"%d" % gesamt
    if n == gesamt:
        return True, u"%d" % gesamt
    return None, u"%d / %d" % (n, gesamt)


def knoten_liste(wurzel):
    """Alle Knoten in Baumreihenfolge (Tiefensuche)."""
    ergebnis = [wurzel]
    for kind in wurzel.kinder:
        ergebnis.extend(knoten_liste(kind))
    return ergebnis
