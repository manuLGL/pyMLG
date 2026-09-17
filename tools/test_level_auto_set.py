# -*- coding: utf-8 -*-
"""Offline-Test der Logik von LevelAutoSet (ohne Revit).

    python tools\\test_level_auto_set.py
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from level_auto_set import logik as lg  # noqa: E402

EBENEN = [("DG", 26.0), ("OG1", 3.0), ("EG", 0.0), ("KG", -2.0)]


def test_waehle_ebene():
    W = lg.waehle_ebene
    assert W(1.0, EBENEN, lg.NAEHER) == "EG"
    assert W(2.0, EBENEN, lg.NAEHER) == "OG1"
    assert W(1.0, EBENEN, lg.DARUEBER) == "OG1"
    assert W(1.0, EBENEN, lg.DARUNTER) == "EG"
    # genau auf der Ebene: darüber und darunter liefern dieselbe
    assert W(3.0, EBENEN, lg.DARUEBER) == "OG1"
    assert W(3.0 + 1e-6, EBENEN, lg.DARUEBER) == "OG1"
    assert W(3.0 - 1e-6, EBENEN, lg.DARUNTER) == "OG1"
    # außerhalb aller Ebenen
    assert W(30.0, EBENEN, lg.DARUEBER) is None
    assert W(-5.0, EBENEN, lg.DARUNTER) is None
    assert W(-5.0, EBENEN, lg.NAEHER) == "KG"
    assert W(1.0, EBENEN, lg.IGNORIEREN) is None
    assert W(1.0, [], lg.NAEHER) is None
    # Gleichstand: zuerst genannte Ebene
    assert W(1.5, EBENEN, lg.NAEHER) == "OG1"


def test_modus_geschossdecke():
    M = lg.modus_fuer_basis
    # Geschossdecke: "Oben" gewinnt, außer es steht auf Ignorieren
    assert M(True, lg.IGNORIEREN, lg.DARUEBER) == lg.DARUEBER
    assert M(True, lg.NAEHER, lg.IGNORIEREN) == lg.NAEHER
    assert M(True, lg.IGNORIEREN, lg.IGNORIEREN) == lg.IGNORIEREN
    # alle anderen: immer "Unten"
    assert M(False, lg.IGNORIEREN, lg.DARUEBER) == lg.IGNORIEREN
    assert M(False, lg.DARUNTER, lg.DARUEBER) == lg.DARUNTER


def test_versatz():
    # Wand auf EG mit +3,5 -> auf OG1 (3,0): Versatz +0,5, Höhe bleibt 3,5
    v = lg.neuer_versatz(0.0, 3.5, 3.0)
    assert abs(v - 0.5) < 1e-12 and abs(3.0 + v - 3.5) < 1e-12
    assert lg.neuer_versatz(3.0, -1.0, 0.0) == 2.0
    assert lg.bezugshoehe(3.0, []) == 3.0
    assert lg.bezugshoehe(3.0, [0.2, -0.1]) == 2.9


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items())
             if n.startswith("test_") and callable(f)]
    for name, funktion in tests:
        funktion()
        print("  ok:", name)
    print("%d Tests bestanden" % len(tests))
