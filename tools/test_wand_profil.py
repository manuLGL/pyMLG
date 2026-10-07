# -*- coding: utf-8 -*-
"""Offline-Test der Logik von WallProfileCopy (ohne Revit).

    python tools\test_wand_profil.py
"""
import math
import os
import sys

os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from wand_profil import logik as lg  # noqa: E402


def _nah(a, b, tol=1e-9):
    return all(abs(x - y) < tol for x, y in zip(a, b))


def test_parallel_daneben():
    quelle = ((0.0, 0.0, 0.0), (10.0, 0.0, 0.0))
    ziel = ((0.0, 5.0, 0.0), (10.0, 5.0, 0.0))
    winkel, versch = lg.lage(quelle, 0.0, ziel, 0.0)
    assert abs(winkel) < 1e-12
    # Giebelspitze des Profils wandert mit
    assert _nah(lg.wende_an(winkel, versch, (5.0, 0.0, 12.0)),
                (5.0, 5.0, 12.0))


def test_gedreht_und_hoeher():
    quelle = ((1.0, 1.0, 0.0), (5.0, 1.0, 0.0))
    ziel = ((20.0, 0.0, 3.0), (20.0, 4.0, 3.0))       # 90° gedreht
    winkel, versch = lg.lage(quelle, 0.5, ziel, 3.5)
    assert abs(winkel - math.pi / 2) < 1e-12
    an = lambda p: lg.wende_an(winkel, versch, p)  # noqa: E731
    assert _nah(an((1.0, 1.0, 0.5)), (20.0, 0.0, 3.5))   # Anfang unten
    assert _nah(an((5.0, 1.0, 0.5)), (20.0, 4.0, 3.5))   # Ende unten
    assert _nah(an((3.0, 1.0, 4.0)), (20.0, 2.0, 7.0))   # Mitte oben


def test_umgekehrt_gezeichnet():
    # Zielwand läuft andersherum: Anfang landet auf Anfang
    quelle = ((0.0, 0.0, 0.0), (10.0, 0.0, 0.0))
    ziel = ((10.0, 3.0, 0.0), (0.0, 3.0, 0.0))
    winkel, versch = lg.lage(quelle, 0.0, ziel, 0.0)
    assert _nah(lg.wende_an(winkel, versch, (2.0, 0.0, 1.0)),
                (8.0, 3.0, 1.0))


def test_gleich_lang():
    a = ((0.0, 0.0, 0.0), (10.0, 0.0, 0.0))
    assert lg.gleich_lang(a, ((0.0, 1.0, 0.0), (10.002, 1.0, 0.0)))
    assert not lg.gleich_lang(a, ((0.0, 1.0, 0.0), (10.01, 1.0, 0.0)))
    assert lg.gleich_lang(a, ((0.0, 0.0, 0.0), (0.0, 10.0, 0.0)))


if __name__ == "__main__":
    for name, funktion in sorted(globals().items()):
        if name.startswith("test_") and callable(funktion):
            funktion()
            print("ok ", name)
