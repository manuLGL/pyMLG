# -*- coding: utf-8 -*-
"""Revit-freie Logik von ApartmentNumbering.

Bewusst ohne Revit-API, damit sie ausserhalb von Revit prüfbar ist
(siehe tools/test_wohnung_top.py).

Gruppentypname einer Wohnung:

    WHG_TYP A2a - TU_OG07-OG12
            ^^^   ^^ ^^^^^^^^^
            Typ   Bauteil  Geschoss

Punkte sind Tupel (x, y) im Grundriss (interne Fuss). Ein Ring ist eine
geschlossene Punktliste (Raumumgrenzung), der letzte Punkt darf fehlen.
Der Pfad kommt als Segmentliste aus stuetzen_nummer.logik.segmente.
"""

import math
import re

# Parameter der Räume (Namen wie im Projekt)
BAUTEIL = u"BAUTEIL"
STIEGE = u"STIEGE"
TOP = u"TOP"
GRUPPE_TYP = u"GRUPPE_TYP"
KATEGORIE = u"TOP_ID_Kategorie"
TYPNAME = u"TOP_ID_Typname"
GESCHOSS = u"TOP_ID_Geschoss"

# In dieser Reihenfolge werden sie geschrieben und gemeldet
ALLE = (BAUTEIL, STIEGE, TOP, GRUPPE_TYP, KATEGORIE, TYPNAME, GESCHOSS)

PREFIX = u"WHG_TYP"

_NAME = re.compile(r"^WHG_TYP\s+(\S+)\s+-\s+([^_\s]+)_(\S+)\s*$")


# ---------------------------------------------------------------- Namen

def ist_wohnung(name):
    """Gruppentyp einer Wohnung (auch wenn der Rest nicht passt)."""
    return (name or u"").strip().upper().startswith(PREFIX)


def zerlege(name):
    """'WHG_TYP A2a - TU_OG07-OG12' -> (Typ, Bauteil, Geschoss), sonst
    None. Der Typ muss mit einem Buchstaben beginnen (er gibt die
    Kategorie)."""
    treffer = _NAME.match((name or u"").strip())
    if not treffer or not treffer.group(1)[0].isalpha():
        return None
    return treffer.groups()


def kategorie(typ):
    """'A2a' -> 'A-TYP'."""
    return u"%s-TYP" % typ[0].upper()


def werte(name, bauteil=u"", stiege=u""):
    """{Parameter: Wert} aus dem Gruppentypnamen - ohne TOP. Leeres
    Bauteil oder leere Stiege wird nicht geschrieben. None, wenn der Name
    nicht zum Schema passt."""
    teile = zerlege(name)
    if teile is None:
        return None
    typ, _, geschoss = teile
    ergebnis = {
        GRUPPE_TYP: name.strip(),
        TYPNAME: typ,
        KATEGORIE: kategorie(typ),
        GESCHOSS: geschoss,
    }
    if (bauteil or u"").strip():
        ergebnis[BAUTEIL] = bauteil.strip()
    if (stiege or u"").strip():
        ergebnis[STIEGE] = stiege.strip()
    return ergebnis


# ---------------------------------------------------------------- Geometrie

def _kreuz(ax, ay, bx, by):
    return ax * by - ay * bx


def _kanten(ring):
    n = len(ring)
    for i in range(n):
        a, b = ring[i], ring[(i + 1) % n]
        if a != b:
            yield a, b


def innen(p, ringe):
    """Liegt p in der Fläche (gerade-ungerade über alle Ringe - Löcher
    zählen also nicht)?"""
    drin = False
    x, y = p
    for ring in ringe:
        for (x1, y1), (x2, y2) in _kanten(ring):
            if (y1 > y) != (y2 > y):
                xs = x1 + (y - y1) * (x2 - x1) / (y2 - y1)
                if x < xs:
                    drin = not drin
    return drin


def _schnitt(a, b, c, d):
    """Anteil t (0..1) auf a-b, an dem a-b die Kante c-d schneidet."""
    rx, ry = b[0] - a[0], b[1] - a[1]
    sx, sy = d[0] - c[0], d[1] - c[1]
    nenner = _kreuz(rx, ry, sx, sy)
    if abs(nenner) < 1e-12:
        return None
    qx, qy = c[0] - a[0], c[1] - a[1]
    t = _kreuz(qx, qy, sx, sy) / nenner
    u = _kreuz(qx, qy, rx, ry) / nenner
    if -1e-9 <= t <= 1 + 1e-9 and -1e-9 <= u <= 1 + 1e-9:
        return min(1.0, max(0.0, t))
    return None


def eintritt(segmente_liste, ringe):
    """Station, an der der Pfad zum ersten Mal in die Fläche kommt - None,
    wenn er sie nicht berührt. segmente_liste: [(a, b, Station bei a)] in
    Pfadrichtung."""
    for a, b, station in segmente_liste:
        if innen(a, ringe):
            return station
        beste = None
        for ring in ringe:
            for c, d in _kanten(ring):
                t = _schnitt(a, b, c, d)
                if t is not None and (beste is None or t < beste):
                    beste = t
        if beste is not None:
            return station + beste * math.hypot(b[0] - a[0], b[1] - a[1])
    return None


def eintritt_mehrere(segmente_liste, flaechen):
    """Frühester Eintritt in eine von mehreren Flächen (die Räume einer
    Wohnung) - None, wenn keine berührt wird."""
    stationen = [eintritt(segmente_liste, ringe) for ringe in flaechen]
    stationen = [s for s in stationen if s is not None]
    return min(stationen) if stationen else None


# ---------------------------------------------------------------- Reihenfolge

def reihenfolge(treffer):
    """treffer: [(Höhe des Geschosses, Station, Schlüssel)] -> Schlüssel
    von unten nach oben, im Geschoss entlang der Linie."""
    return [s for _, _, s in sorted(treffer, key=lambda e: (e[0], e[1]))]


def bereiche(nummerierte):
    """[(Geschossname, Nummer)] in Vergabereihenfolge -> [(Geschossname,
    erste, letzte)] für die Zusammenfassung."""
    ergebnis = []
    for geschoss, nummer in nummerierte:
        if ergebnis and ergebnis[-1][0] == geschoss:
            ergebnis[-1] = (geschoss, ergebnis[-1][1], nummer)
        else:
            ergebnis.append((geschoss, nummer, nummer))
    return ergebnis
