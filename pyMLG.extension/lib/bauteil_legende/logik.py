# -*- coding: utf-8 -*-
"""Revit-freie Logik von ComponentLegend.

Aus den Ansichten kommen Funde (kat_id, kategorie, typ_schluessel, typ_ref,
familie, typname, element_schluessel). Daraus entstehen Typen (je
Familientyp einer), die Kategorienliste für den Dialog und die sortierte
Reihenfolge der Legende (gestapelt wird in revit.py, weil die Höhe eines
Bauteils erst nach dem Platzieren bekannt ist).
"""

import re

# Werte des Parameters LEGEND_COMPONENT_VIEW ("Ansichtsrichtung")
RICHTUNG_GRUNDRISS = -8

BESCHRIFTUNG_KEINE = 0
BESCHRIFTUNG_TYP = 1
BESCHRIFTUNG_FAMILIE_TYP = 2


class Typ(object):
    """Ein Familientyp mit allen Elementen, die in den Ansichten vorkommen."""

    def __init__(self, kat_id, kategorie, schluessel, ref, familie, name):
        self.kat_id = kat_id
        self.kategorie = kategorie
        self.schluessel = schluessel    # Zahl der Typ-Id
        self.ref = ref                  # Typ-Id für Revit (ElementId)
        self.familie = familie or u""
        self.name = name or u""
        self.elemente = set()

    @property
    def anzahl(self):
        return len(self.elemente)

    def __repr__(self):
        return u"Typ(%s: %s: %s)" % (self.kategorie, self.familie, self.name)


def sammle(funde, typen=None):
    """Funde in {typ_schluessel: Typ} einsammeln. Ein Element, das in
    mehreren Ansichten sichtbar ist, zählt nur einmal."""
    typen = {} if typen is None else typen
    for kat_id, kategorie, schluessel, ref, familie, name, element in funde:
        typ = typen.get(schluessel)
        if typ is None:
            typ = typen[schluessel] = Typ(kat_id, kategorie, schluessel, ref,
                                          familie, name)
        typ.elemente.add(element)
    return typen


def _natuerlich(text):
    """Sortierschlüssel, der Zahlen als Zahlen vergleicht (T2 < T10)."""
    return [(0, int(teil), u"") if teil.isdigit() else (1, 0, teil)
            for teil in re.split(r"(\d+)", (text or u"").lower()) if teil]


def kategorien(typen):
    """[(kat_id, kategorie, anzahl_typen)] nach Namen sortiert."""
    zaehler = {}
    namen = {}
    for typ in typen.values():
        zaehler[typ.kat_id] = zaehler.get(typ.kat_id, 0) + 1
        namen[typ.kat_id] = typ.kategorie
    return sorted(((k, namen[k], zaehler[k]) for k in zaehler),
                  key=lambda e: _natuerlich(e[1]))


def reihenfolge(typen, kat_ids):
    """Typen der gewählten Kategorien in Legendenreihenfolge:
    Kategorie, Familie, Typ - jeweils natürlich sortiert."""
    gewaehlt = set(kat_ids)
    return sorted((typ for typ in typen.values() if typ.kat_id in gewaehlt),
                  key=lambda typ: (_natuerlich(typ.kategorie), _natuerlich(typ.familie),
                                   _natuerlich(typ.name)))


def beschriftung(typ, art):
    """Text neben dem Bauteil - oder u"" ohne Beschriftung."""
    if art == BESCHRIFTUNG_TYP:
        return typ.name
    if art == BESCHRIFTUNG_FAMILIE_TYP:
        if typ.familie and typ.familie != typ.name:
            return u"%s: %s" % (typ.familie, typ.name)
        return typ.name
    return u""


def zahl(text):
    try:
        return float((text or u"").strip().replace(u",", u"."))
    except ValueError:
        return None


def pruefe_massstab(text):
    """Ganze Zahl 1..100000 oder None."""
    wert = zahl(text)
    if wert is None or not 1 <= wert <= 100000 or wert != int(wert):
        return None
    return int(wert)


def pruefe_abstand(text):
    """Abstand in Papier-mm (0..500) oder None."""
    wert = zahl(text)
    if wert is None or not 0 <= wert <= 500:
        return None
    return wert
