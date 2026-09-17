# -*- coding: utf-8 -*-
"""Revit-freie Logik von JoinMultiple.

Priorität: Das Element mit der höheren Zahl schneidet das mit der
niedrigeren. Bei gleicher Priorität bleibt die Reihenfolge, die Revit wählt.
"""

import re

# Vorgaben je BuiltInCategory-Name. Nur diese Kategorien lassen sich
# zuverlässig mit "Geometrie verbinden" bearbeiten.
STANDARD_PRIORITAETEN = (
    ("OST_StructuralColumns", 500),
    ("OST_Columns", 450),
    ("OST_StructuralFraming", 400),
    ("OST_StructuralFoundation", 350),
    ("OST_Walls", 300),
    ("OST_Floors", 200),
    ("OST_Roofs", 150),
    ("OST_GenericModel", 100),
    ("OST_Ceilings", 50),
    ("OST_Toposolid", 0),
    ("OST_Mass", 0),
)

# Toleranz der Bounding-Box-Suche in Fuß (ca. 3 mm): Elemente, die sich nur
# berühren, gelten noch als Kandidaten
TOLERANZ = 0.01

# Zähler im Ergebnis
NEU = "neu"                          # neu verbunden
BEREITS = "bereits"                  # war schon verbunden
UEBERSPRUNGEN = "uebersprungen"      # schneiden sich nicht
UMGEDREHT = "umgedreht"              # Schnittreihenfolge getauscht
RICHTIG = "richtig"                  # Reihenfolge stimmte schon
GLEICH = "gleich"                    # gleiche Priorität - nichts zu tun
NICHT_GEPRUEFT = "nicht_geprueft"    # bereits verbunden, Anpassen aus
ABGELEHNT = "abgelehnt"              # Revit behält die falsche Reihenfolge


def standard_prioritaet(bic_name):
    return dict(STANDARD_PRIORITAETEN).get(bic_name, 0)


def lies_prioritaet(text):
    """Ganzzahl aus einem Eingabefeld. Leer -> 0, sonst ValueError."""
    text = (text or u"").strip()
    if not text:
        return 0
    if not re.match(r"^[+-]?\d+$", text):
        raise ValueError(u"'%s' ist keine ganze Zahl." % text)
    return int(text)


def zahl_aus_wert(wert):
    """Priorität aus einem Parameterwert (Zahl oder Text wie '200' / '1,5').
    None, wenn kein Zahlenwert erkennbar ist."""
    if wert is None or isinstance(wert, bool):
        return None
    if isinstance(wert, (int, float)):
        return float(wert)
    treffer = re.search(r"[+-]?\d+(?:[.,]\d+)?", u"%s" % wert)
    if not treffer:
        return None
    return float(treffer.group(0).replace(u",", u"."))


def kandidaten_paare(boxen, toleranz=TOLERANZ):
    """Indexpaare (i, j) mit i < j, deren Boxen sich (mit Toleranz)
    überlappen.

    boxen: [(min_xyz, max_xyz)] oder None für Elemente ohne Box.
    Sweep-and-Prune entlang X: sortieren nach min.x, dann nur Boxen
    vergleichen, deren X-Bereich noch offen ist.
    """
    reihenfolge = sorted((i for i, b in enumerate(boxen) if b is not None),
                         key=lambda i: boxen[i][0][0])
    aktiv = []
    paare = []
    for i in reihenfolge:
        min_i, max_i = boxen[i]
        aktiv = [j for j in aktiv if boxen[j][1][0] + toleranz >= min_i[0]]
        for j in aktiv:
            min_j, max_j = boxen[j]
            if all(min_i[k] <= max_j[k] + toleranz
                   and min_j[k] <= max_i[k] + toleranz for k in (1, 2)):
                paare.append((min(i, j), max(i, j)))
        aktiv.append(i)
    return sorted(paare)


def paar_erlaubt(kategorie_a, kategorie_b, gleiche_kategorie):
    return gleiche_kategorie or kategorie_a != kategorie_b


def schneidender(prio_a, prio_b):
    """0: a soll schneiden, 1: b soll schneiden, None: gleich."""
    if prio_a is None or prio_b is None or prio_a == prio_b:
        return None
    return 0 if prio_a > prio_b else 1


def verbinden_aktion(verbunden, schneiden_sich, nur_bei_schnitt):
    """Durchgang 1: NEU (jetzt verbinden), BEREITS oder UEBERSPRUNGEN."""
    if verbunden:
        return BEREITS
    if nur_bei_schnitt and not schneiden_sich:
        return UEBERSPRUNGEN
    return NEU


def reihenfolge_pruefen(aktion, bestehende_anpassen):
    """Durchgang 2 für dieses Paar? Neu verbundene immer, bereits verbundene
    nur mit der Option 'Bereits verbundene anpassen'."""
    return aktion == NEU or (aktion == BEREITS and bestehende_anpassen)


def oben_unten(prio_a, prio_b):
    """(0, 1), wenn a schneiden soll, (1, 0) für b, None bei Gleichstand."""
    soll = schneidender(prio_a, prio_b)
    if soll is None:
        return None
    return (0, 1) if soll == 0 else (1, 0)
