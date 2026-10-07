# -*- coding: utf-8 -*-
"""Revit-freie Logik von ColumnNumbering.

Bewusst ohne Revit-API, damit sie ausserhalb von Revit prüfbar ist
(siehe tools/test_stuetzen_nummer.py).

Punkte sind Tupel (x, y) im Grundriss (interne Fuss). Ein Linienstück ist
eine Punktliste in Zeichenrichtung (Revit: Curve.Tessellate()).
"""

import math
import re

from mlg_plaene.logik import nummer_plus

# Endpunkte, die näher beieinander liegen, gelten als verbunden (Fuss)
VERBINDUNG = 0.01

_ZAHL = re.compile(r"\d")


# ---------------------------------------------------------------- Nummern

def pruefe_start(start):
    """Fehlertext oder None. Die letzte Zahl der Startnummer wird
    hochgezählt - ohne Zahl gibt es nichts zu zählen."""
    if not (start or u"").strip():
        return u"leer"
    if not _ZAHL.search(start):
        return u"ohne Zahl"
    return None


class Zaehler(object):
    """S-01, S-02, ... - führende Nullen bleiben (nummer_plus)."""

    def __init__(self, start, schritt=1):
        if schritt == 0:
            raise ValueError(u"schritt")
        self.aktuell = start.strip()
        self.schritt = schritt

    def weiter(self):
        """Liefert die aktuelle Nummer und zählt weiter."""
        nummer = self.aktuell
        self.aktuell = nummer_plus(nummer, self.schritt)
        return nummer


def vorschau(start, schritt, anzahl=4):
    """'S-01, S-02, S-03, S-04 …' - wirft ValueError bei ungültiger Eingabe
    (oder wenn die Zahl negativ würde)."""
    zaehler = Zaehler(start, schritt)
    return u", ".join(zaehler.weiter() for _ in range(anzahl)) + u" …"


def zahl(text):
    """Dezimalzahl mit Komma oder Punkt -> float (ValueError sonst)."""
    return float((text or u"").strip().replace(u",", u"."))


# ---------------------------------------------------------------- Pfad

def _abstand(a, b):
    return math.hypot(a[0] - b[0], a[1] - b[1])


def _laenge(punkte):
    return sum(_abstand(a, b) for a, b in zip(punkte, punkte[1:]))


def ketten(stuecke, toleranz=VERBINDUNG):
    """Verbundene Linienstücke zu Ketten zusammenfassen.

    Jede Kette ist eine Punktliste. Durchlaufen wird sie in der Richtung, in
    der die meisten ihrer Stücke gezeichnet wurden (gleich lange: das
    erste Stück zählt). Die Ketten bleiben in der Reihenfolge ihres ersten
    Stücks - so bestimmt die Klickreihenfolge die Reihenfolge der Reihen.
    """
    stuecke = [list(s) for s in stuecke if len(s) >= 2]
    frei = list(range(len(stuecke)))
    ergebnis = []
    while frei:
        erstes = frei.pop(0)
        # (Index, vorwärts) in Durchlaufreihenfolge
        kette = [(erstes, True)]

        def anfang():
            i, vor = kette[0]
            return stuecke[i][0] if vor else stuecke[i][-1]

        def ende():
            i, vor = kette[-1]
            return stuecke[i][-1] if vor else stuecke[i][0]

        gefunden = True
        while gefunden:
            gefunden = False
            for i in list(frei):
                s = stuecke[i]
                if _abstand(ende(), s[0]) <= toleranz:
                    kette.append((i, True))
                elif _abstand(ende(), s[-1]) <= toleranz:
                    kette.append((i, False))
                elif _abstand(anfang(), s[-1]) <= toleranz:
                    kette.insert(0, (i, True))
                elif _abstand(anfang(), s[0]) <= toleranz:
                    kette.insert(0, (i, False))
                else:
                    continue
                frei.remove(i)
                gefunden = True

        punkte = []
        for i, vor in kette:
            teil = stuecke[i] if vor else stuecke[i][::-1]
            punkte.extend(teil[1:] if punkte else teil)
        # Das erste Stück läuft immer vorwärts -> bei Gleichstand bleibt es
        vorwaerts = sum(1 for _, vor in kette if vor)
        if len(kette) - vorwaerts > vorwaerts:
            punkte.reverse()
        ergebnis.append(punkte)
    return ergebnis


def segmente(ketten_liste):
    """Ketten -> [(a, b, Station bei a)]. Die Station läuft über alle
    Ketten weiter, der Sprung zwischen zwei Ketten zählt nicht."""
    ergebnis, station = [], 0.0
    for punkte in ketten_liste:
        for a, b in zip(punkte, punkte[1:]):
            laenge = _abstand(a, b)
            if laenge > 1e-12:
                ergebnis.append((a, b, station))
                station += laenge
    return ergebnis


def naechster(segmente_liste, p):
    """(Abstand, Station) des Pfadpunkts, der p am nächsten liegt."""
    beste = None
    for a, b, station in segmente_liste:
        dx, dy = b[0] - a[0], b[1] - a[1]
        laenge2 = dx * dx + dy * dy
        anteil = ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / laenge2
        anteil = min(1.0, max(0.0, anteil))
        fuss = (a[0] + anteil * dx, a[1] + anteil * dy)
        abstand = _abstand(p, fuss)
        if beste is None or abstand < beste[0] - 1e-9:
            beste = (abstand, station + anteil * math.sqrt(laenge2))
    return beste


def entlang(segmente_liste, punkte, max_abstand):
    """Indizes der Punkte höchstens max_abstand neben dem Pfad, sortiert
    nach ihrer Lage entlang des Pfads."""
    treffer = []
    for index, p in enumerate(punkte):
        abstand, station = naechster(segmente_liste, p)
        if abstand <= max_abstand:
            treffer.append((station, abstand, index))
    treffer.sort()
    return [index for _, _, index in treffer]


def uebereinander(p, punkte, toleranz):
    """Indizes der Punkte, die im Grundriss auf p liegen."""
    return [i for i, q in enumerate(punkte) if _abstand(p, q) <= toleranz]
