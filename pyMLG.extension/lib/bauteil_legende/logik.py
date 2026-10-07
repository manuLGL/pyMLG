# -*- coding: utf-8 -*-
"""Revit-freie Logik von ComponentLegend.

Aus den Ansichten kommen Funde (kat_id, kategorie, typ_schluessel, typ_ref,
familie, typname, element_schluessel). Daraus entstehen Typen (je
Familientyp einer), die Kategorienliste für den Dialog und die sortierte
Reihenfolge der Legende (gestapelt wird in revit.py, weil die Höhe eines
Bauteils erst nach dem Platzieren bekannt ist).
"""

import re

# Werte des Parameters LEGEND_COMPONENT_VIEW ("Ansichtsrichtung"), wie in
# pyRevits Werkzeugen Drawing Set > Legends. Welche Richtungen ein Typ
# anbietet, hängt von der Kategorie ab (Türen z.B. keinen Schnitt).
RICHTUNG_GRUNDRISS = -8
RICHTUNGEN = (
    (RICHTUNG_GRUNDRISS, (u"Grundriss", u"Floor plan", u"Planta")),
    (-4, (u"Deckenplan", u"Ceiling plan", u"Plano de techo")),
    (-5, (u"Schnitt", u"Section", u"Sección")),
    (-7, (u"Ansicht vorne", u"Elevation front", u"Alzado frontal")),
    (-6, (u"Ansicht hinten", u"Elevation back", u"Alzado posterior")),
    (-10, (u"Ansicht links", u"Elevation left", u"Alzado izquierdo")),
    (-9, (u"Ansicht rechts", u"Elevation right", u"Alzado derecho")),
    (-3, (u"3D", u"3D", u"3D")),
)

# Paso von Türen/Fenstern: als Text, als Maßkette oder beides
MASS_TEXT = 0
MASS_KETTE = 1
MASS_BEIDES = 2
MASS_ARTEN = (
    (MASS_KETTE, (u"Maßkette", u"Dimension string", u"Cota")),
    (MASS_TEXT, (u"Text", u"Text", u"Texto")),
    (MASS_BEIDES, (u"Maßkette + Text", u"Dimension + text", u"Cota + texto")),
)
# Ansichtsrichtungen, in denen die Höhe sichtbar ist (Ansichten)
RICHTUNGEN_MIT_HOEHE = (-7, -6, -10, -9)

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
        self.beispiel = None            # ein Exemplar (ElementId) für Exemplarparameter

    @property
    def anzahl(self):
        return len(self.elemente)

    def __repr__(self):
        return u"Typ(%s: %s: %s)" % (self.kategorie, self.familie, self.name)


def sammle(funde, typen=None):
    """Funde in {typ_schluessel: Typ} einsammeln. Ein Element, das in
    mehreren Ansichten sichtbar ist, zählt nur einmal."""
    typen = {} if typen is None else typen
    for fund in funde:
        kat_id, kategorie, schluessel, ref, familie, name, element = fund[:7]
        typ = typen.get(schluessel)
        if typ is None:
            typ = typen[schluessel] = Typ(kat_id, kategorie, schluessel, ref,
                                          familie, name)
        if typ.beispiel is None and len(fund) > 7:
            typ.beispiel = fund[7]
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


def masstext(breite_mm, hoehe_mm):
    """u"825 × 2030" - fehlt ein Wert, nur der andere; ohne beide u""."""
    werte = [u"%d" % int(round(w)) for w in (breite_mm, hoehe_mm) if w is not None]
    return u" × ".join(werte)


def mit_text(art):
    return art in (MASS_TEXT, MASS_BEIDES)


def mit_kette(art):
    return art in (MASS_KETTE, MASS_BEIDES)


def kettenlage(links, unten, breite_bauteil, breite_paso, hoehe_paso, abstand):
    """Lage der Maßketten eines Bauteils (Modelleinheiten).

    Breite: mittig unter dem Bauteil, Maßlinie um abstand tiefer.
    Höhe: links neben dem Bauteil ab der Unterkante, Maßlinie um abstand
    weiter links. Rückgabe {"breite": (x1, x2, y_linie),
    "hoehe": (y1, y2, x_linie)} - nur für vorhandene Werte."""
    lage = {}
    if breite_paso:
        mitte = links + breite_bauteil / 2.0
        lage[u"breite"] = (mitte - breite_paso / 2.0, mitte + breite_paso / 2.0,
                          unten - abstand)
    if hoehe_paso:
        lage[u"hoehe"] = (unten, unten + hoehe_paso, links - abstand)
    return lage


def zeilen(*texte):
    """Nicht leere Texte untereinander (Revit-Text: Zeilenumbruch \r)."""
    return u"\r".join(text for text in texte if text)


def ueberlappt(punkte, rahmen):
    """Berührt das Rechteck um die Punkte [(x, y)] den Rahmen
    (x0, y0, x1, y1)? Kante an Kante zählt als berührt."""
    if not punkte:
        return True
    xs = [p[0] for p in punkte]
    ys = [p[1] for p in punkte]
    x0, y0, x1, y1 = rahmen
    return max(xs) >= x0 and min(xs) <= x1 and max(ys) >= y0 and min(ys) <= y1


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
