# -*- coding: utf-8 -*-
"""Offline-Test der Logik von JoinMultiple (ohne Revit).

    python tools\\test_join_multiple.py
"""
import os
import random
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from join_multiple import logik as lg  # noqa: E402


def _box(x0, y0, z0, x1, y1, z1):
    return ((x0, y0, z0), (x1, y1, z1))


def test_kandidaten_einfach():
    boxen = [
        _box(0, 0, 0, 10, 1, 3),      # Wand
        _box(-1, -5, 3, 11, 5, 3.5),  # Decke liegt auf der Wand (berührt)
        _box(20, 0, 0, 30, 1, 3),     # weit weg
        None,                         # ohne Geometrie
        _box(9.995, 0, 0, 12, 1, 3),  # berührt Wand innerhalb der Toleranz
    ]
    assert lg.kandidaten_paare(boxen) == [(0, 1), (0, 4), (1, 4)]


def test_kandidaten_wie_brute_force():
    zufall = random.Random(7)
    boxen = []
    for _ in range(300):
        x, y, z = (zufall.uniform(0, 50) for _ in range(3))
        boxen.append(_box(x, y, z, x + zufall.uniform(0, 4),
                          y + zufall.uniform(0, 4), z + zufall.uniform(0, 4)))
    t = lg.TOLERANZ
    erwartet = sorted(
        (i, j) for i in range(len(boxen)) for j in range(i + 1, len(boxen))
        if all(boxen[i][0][k] <= boxen[j][1][k] + t
               and boxen[j][0][k] <= boxen[i][1][k] + t for k in range(3)))
    assert lg.kandidaten_paare(boxen) == erwartet


def test_prioritaeten():
    assert lg.lies_prioritaet(u" 200 ") == 200
    assert lg.lies_prioritaet(u"-5") == -5
    assert lg.lies_prioritaet(u"") == 0
    for falsch in (u"2.5", u"abc", u"1 2"):
        try:
            lg.lies_prioritaet(falsch)
        except ValueError:
            continue
        raise AssertionError(falsch)
    assert lg.zahl_aus_wert(3) == 3.0
    assert lg.zahl_aus_wert(u"Prio 1,5") == 1.5
    assert lg.zahl_aus_wert(u"") is None
    assert lg.zahl_aus_wert(None) is None
    assert lg.standard_prioritaet("OST_Walls") == 300
    assert lg.standard_prioritaet("OST_Unbekannt") == 0


def test_entscheidungen():
    A = lg.verbinden_aktion
    assert A(True, None, True) == lg.BEREITS
    assert A(False, False, True) == lg.UEBERSPRUNGEN
    assert A(False, None, False) == lg.NEU
    assert A(False, True, True) == lg.NEU
    P = lg.reihenfolge_pruefen
    assert P(lg.NEU, False) and P(lg.BEREITS, True)
    assert not P(lg.BEREITS, False) and not P(lg.UEBERSPRUNGEN, True)
    assert lg.oben_unten(300, 200) == (0, 1)
    assert lg.oben_unten(200, 500) == (1, 0)
    assert lg.oben_unten(100, 100) is None
    assert lg.paar_erlaubt(1, 1, False) is False
    assert lg.paar_erlaubt(1, 2, False) is True
    assert lg.schneidender(1, 2) == 1 and lg.schneidender(2, 1) == 0


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items())
             if n.startswith("test_") and callable(f)]
    for name, funktion in tests:
        funktion()
        print("  ok:", name)
    print("%d Tests bestanden" % len(tests))
