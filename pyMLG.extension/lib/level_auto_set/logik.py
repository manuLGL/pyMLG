# -*- coding: utf-8 -*-
"""Revit-freie Logik von LevelAutoSet.

Ein Element hängt mit Unterkante (Basis) und ggf. Oberkante an je einer
Ebene: absolute Höhe = Ebenenhöhe + Versatz. Beim Wechsel der Ebene bleibt
die absolute Höhe gleich, der Versatz wird umgerechnet:

    neuer Versatz = alte Ebenenhöhe + alter Versatz - neue Ebenenhöhe

Alle Höhen in Revit-internen Einheiten (Fuß).
"""

NAEHER = "naeher"
DARUEBER = "darueber"
DARUNTER = "darunter"
IGNORIEREN = "ignorieren"

# Eine Ebene auf genau der Höhe zählt als "darüber" und als "darunter"
# (ca. 0,03 mm)
TOLERANZ = 1e-4


def waehle_ebene(hoehe, ebenen, modus):
    """Ebene für eine absolute Höhe.

    ebenen  [(Schlüssel, Ebenenhöhe)]
    modus   NAEHER     nächstgelegene Ebene
            DARUEBER   nächste Ebene auf oder über der Höhe
            DARUNTER   nächste Ebene auf oder unter der Höhe
            IGNORIEREN nichts ändern
    Rückgabe: Schlüssel oder None (keine passende Ebene / ignorieren)
    """
    if modus == IGNORIEREN or not ebenen:
        return None
    if modus == DARUEBER:
        kandidaten = [(h - hoehe, k) for k, h in ebenen
                      if h >= hoehe - TOLERANZ]
    elif modus == DARUNTER:
        kandidaten = [(hoehe - h, k) for k, h in ebenen
                      if h <= hoehe + TOLERANZ]
    else:
        kandidaten = [(abs(h - hoehe), k) for k, h in ebenen]
    if not kandidaten:
        return None
    # Bei gleichem Abstand gewinnt die zuerst genannte Ebene
    return min(kandidaten, key=lambda e: e[0])[1]


def modus_fuer_basis(oberkante, modus_basis, modus_oben):
    """Modus für den einzigen Ebenenbezug eines Elements.

    Bei Geschossdecken bezieht sich die Ebene auf die Oberkante - dort gilt
    die Einstellung "Oben", solange sie nicht auf Ignorieren steht."""
    if oberkante and modus_oben != IGNORIEREN:
        return modus_oben
    return modus_basis


def neuer_versatz(alte_ebenenhoehe, versatz, neue_ebenenhoehe):
    return alte_ebenenhoehe + versatz - neue_ebenenhoehe


def bezugshoehe(ebenenhoehe, versaetze):
    """Absolute Höhe für die Ebenenwahl. Mehrere Versätze (Träger: Anfang
    und Ende) -> der niedrigere Punkt."""
    if not versaetze:
        return ebenenhoehe
    return ebenenhoehe + min(versaetze)
