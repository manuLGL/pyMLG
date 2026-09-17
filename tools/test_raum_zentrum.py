# -*- coding: utf-8 -*-
"""Offline-Test der Schwerpunkt-Mathematik von RaumZentrum (ohne Revit).

    python tools\\test_raum_zentrum.py

Prüft typische Raumformen. Room.IsPointInRoom() selbst lässt sich nur in
Revit testen - dafür hat das Werkzeug den Modus "Probelauf".
"""
import math
import os
import sys

# Texte werden auf Deutsch verglichen
os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from raum_zentrum import geometrie as g  # noqa: E402

# Weit vom Ursprung entfernt, wie bei Projekten mit Landeskoordinaten
OX, OY = 150000.0, 250000.0


def verschoben(punkte):
    return [(x + OX, y + OY) for x, y in punkte]


def nahe(a, b, tol=1e-6):
    return abs(a - b) < tol


def pruefe(name, bedingung):
    print(("  OK    " if bedingung else "  FEHLER") + "  " + name)
    return bedingung


def main():
    ok = True

    # 1. Rechteck 10 x 4, im Uhrzeigersinn (Revit-Richtung egal)
    rechteck = verschoben([(0, 0), (0, 4), (10, 4), (10, 0)])
    e = g.schwerpunkt_mit_loechern([rechteck])
    print("Rechteck:", e["cx"] - OX, e["cy"] - OY, e["flaeche"])
    ok &= pruefe("Schwerpunkt (5, 2)", nahe(e["cx"] - OX, 5) and nahe(e["cy"] - OY, 2))
    ok &= pruefe("Fläche 40", nahe(e["flaeche"], 40))

    # 2. L-Form: 10x10 minus 6x6 oben rechts. Eckpunkt-Mittel wäre falsch.
    l_form = verschoben([(0, 0), (10, 0), (10, 4), (4, 4), (4, 10), (0, 10)])
    e = g.schwerpunkt_mit_loechern([l_form])
    soll = (100 * 5 - 36 * 7) / 64.0
    print("L-Form:", e["cx"] - OX, e["cy"] - OY)
    ok &= pruefe("Schwerpunkt (%.4f, %.4f)" % (soll, soll),
                 nahe(e["cx"] - OX, soll) and nahe(e["cy"] - OY, soll))
    ok &= pruefe("L-Schwerpunkt liegt im Umriss",
                 g.punkt_in_umriss(e["cx"], e["cy"], [l_form]))

    # 3. Raum mit Aussparung (Stütze 2x2 bei (6..8, 1..3)), gleiche Richtung
    aussen = verschoben([(0, 0), (10, 0), (10, 4), (0, 4)])
    stuetze = verschoben([(6, 1), (8, 1), (8, 3), (6, 3)])
    e = g.schwerpunkt_mit_loechern([stuetze, aussen])  # Reihenfolge egal
    soll_x = (40 * 5 - 4 * 7) / 36.0
    print("Aussparung:", e["cx"] - OX, e["cy"] - OY, e["flaeche"])
    ok &= pruefe("1 Loch erkannt", e["loecher"] == 1)
    ok &= pruefe("Fläche 36", nahe(e["flaeche"], 36))
    ok &= pruefe("Schwerpunkt (%.4f, 2)" % soll_x,
                 nahe(e["cx"] - OX, soll_x) and nahe(e["cy"] - OY, 2))
    ok &= pruefe("Punkt in Stütze gilt als aussen",
                 not g.punkt_in_umriss(OX + 7, OY + 2, [aussen, stuetze]))

    # 4. U-Form: Schwerpunkt liegt in der Öffnung -> Ersatzpunkt nötig
    u_form = verschoben([(0, 0), (10, 0), (10, 10), (8, 10), (8, 2),
                         (2, 2), (2, 10), (0, 10)])
    e = g.schwerpunkt_mit_loechern([u_form])
    print("U-Form:", e["cx"] - OX, e["cy"] - OY)
    ok &= pruefe("U-Schwerpunkt liegt ausserhalb",
                 not g.punkt_in_umriss(e["cx"], e["cy"], [u_form]))
    kandidaten = g.ersatzpunkte([u_form], e["cx"], e["cy"])
    bester = kandidaten[0]
    print("  Ersatzpunkt:", bester[0] - OX, bester[1] - OY,
          "(%d Kandidaten)" % len(kandidaten))
    ok &= pruefe("Ersatzpunkt = Mitte des U-Stegs (5, 1)",
                 nahe(bester[0] - OX, 5, 1e-6) and nahe(bester[1] - OY, 1, 1e-6))
    ok &= pruefe("alle Kandidaten im Umriss",
                 all(g.punkt_in_umriss(x, y, [u_form]) for x, y in kandidaten))

    # 5. Halbkreis aus tessellierten Bögen, Kurven teils umgedreht
    r = 5.0
    bogen = [(r * math.cos(math.pi * k / 200), r * math.sin(math.pi * k / 200))
             for k in range(201)]
    sehne = [(-r, 0.0), (r, 0.0)]
    kurven = [verschoben(bogen), list(reversed(verschoben(sehne)))]
    punkte = g.verkette_kurvenpunkte(kurven)
    e = g.schwerpunkt_mit_loechern([punkte])
    soll_y = 4 * r / (3 * math.pi)
    print("Halbkreis:", e["cx"] - OX, e["cy"] - OY, "Punkte:", len(punkte))
    ok &= pruefe("keine doppelten Stosspunkte", len(punkte) == 201)
    ok &= pruefe("Schwerpunkt y = 4r/(3 pi), Toleranz 1 mm",
                 nahe(e["cx"] - OX, 0, 1e-3) and nahe(e["cy"] - OY, soll_y, 3e-3))

    # 6. Entartet: Fläche 0
    try:
        g.schwerpunkt_mit_loechern([verschoben([(0, 0), (5, 0), (10, 0)])])
        ok &= pruefe("ValueError bei Fläche 0", False)
    except ValueError:
        ok &= pruefe("ValueError bei Fläche 0", True)

    print("\nALLE TESTS BESTANDEN" if ok else "\nTESTS FEHLGESCHLAGEN")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
