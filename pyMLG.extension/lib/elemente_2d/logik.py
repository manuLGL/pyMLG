# -*- coding: utf-8 -*-
"""Revit-freie Logik von Elements2D.

Ein 2D-Element ist ein ansichtsspezifisches Element (Detaillinie, Füllbereich,
Detailbauteil, Text, Bemaßung, Beschriftung, Detailgruppe ...). revit.py
liest je Element einen Eintrag aus; hier wird eingegrenzt und in die
Datensätze für filter_more.baum.baue() übersetzt:

    nach Kategorie   Kategorie > Familie > Typ > Element · Ansicht
    nach Ansicht     Ansichtsart > Ansicht > Kategorie > Element
"""

NACH_KATEGORIE = "kategorie"
NACH_ANSICHT = "ansicht"

# Eingrenzung nach Platzierung der Ansicht
ALLE = "alle"
PLATZIERTE = "platzierte"
UNPLATZIERTE = "unplatzierte"


class Eintrag(object):
    """Was der Baum über ein 2D-Element wissen muss."""

    def __init__(self, element_id, kategorie, familie, typ, name, ansicht,
                 ansichtsart, platziert=False, auf_plan=False, plaene=()):
        self.element_id = element_id
        self.kategorie = kategorie
        self.familie = familie
        self.typ = typ
        self.name = name or typ
        self.ansicht = ansicht
        self.ansichtsart = ansichtsart
        self.plaene = frozenset(plaene)   # Plan-Ids, auf denen es erscheint
        self.platziert = platziert or bool(self.plaene)
        self.auf_plan = auf_plan      # Element liegt direkt auf einem Plan

    def __repr__(self):
        return "<Eintrag %d %s @ %s>" % (self.element_id, self.kategorie,
                                         self.ansicht)


def grenze_ein(eintraege, platzierung=ALLE, mit_plaenen=True, plaene=None):
    """platzierung  ALLE, PLATZIERTE (Ansicht liegt auf einem Plan) oder
                 UNPLATZIERTE (auf keinem Plan - Elemente direkt auf Plänen
                 fallen dann weg)
    mit_plaenen  Elemente, die direkt auf einem Plan liegen
    plaene       bei PLATZIERTE: nur diese Plan-Ids (None = alle Pläne)"""
    ergebnis = []
    for e in eintraege:
        if e.auf_plan and (platzierung == UNPLATZIERTE or not mit_plaenen):
            continue
        if platzierung == UNPLATZIERTE and e.platziert:
            continue
        if platzierung == PLATZIERTE and (
                not e.platziert
                or (plaene is not None and not e.plaene & plaene)):
            continue
        ergebnis.append(e)
    return ergebnis


def bezeichnung(eintrag, mit_ansicht):
    text = u"%s  [%d]" % (eintrag.name, eintrag.element_id)
    if mit_ansicht:
        text += u"  ·  %s" % eintrag.ansicht
    return text


def datensaetze(eintraege, gruppierung=NACH_KATEGORIE):
    """[Eintrag] -> [(Ebene 1, Ebene 2, Ebene 3, Bezeichnung, Id)]."""
    if gruppierung == NACH_ANSICHT:
        return [(e.ansichtsart, e.ansicht, e.kategorie,
                 bezeichnung(e, False), e.element_id) for e in eintraege]
    return [(e.kategorie, e.familie, e.typ, bezeichnung(e, True),
             e.element_id) for e in eintraege]


def anzahl_ansichten(eintraege):
    return len(set(e.ansicht for e in eintraege))


def kurz(text, laenge=60):
    """Auf eine Zeile gebracht und gekürzt - für Textinhalte im Baum."""
    zeile = u" ".join((text or u"").split())
    if len(zeile) > laenge:
        zeile = zeile[:laenge - 1].rstrip() + u"…"
    return zeile
