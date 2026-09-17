# -*- coding: utf-8 -*-
"""Offline-Test des Auswahlbaums von FilterMore (ohne Revit).

    python tools\\test_filter_more.py
"""
import os
import sys

# Texte werden auf Deutsch verglichen
os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from filter_more import baum as bm  # noqa: E402

DATEN = [
    (u"Wände", u"Basiswand", u"AW_36.5", u"AW_36.5  [10]", 10),
    (u"Wände", u"Basiswand", u"AW_36.5", u"AW_36.5  [11]", 11),
    (u"Wände", u"Basiswand", u"IW_11.5", u"IW_11.5  [12]", 12),
    (u"Geschossdecken", u"Geschossdecke", u"DA_25cm", u"DA_25cm  [20]", 20),
    (u"Türen", u"Tür 10", u"T1", u"T1  [30]", 30),
    (u"Türen", u"Tür 2", u"T1", u"T1  [31]", 31),
]


def test_aufbau():
    w = bm.baue(DATEN)
    assert w.ids and sorted(w.ids) == [10, 11, 12, 20, 30, 31]
    assert [k.name for k in w.kinder] == [u"Geschossdecken", u"Türen",
                                          u"Wände"]
    waende = w.kinder[2]
    assert len(waende.ids) == 3
    basis = waende.kinder[0]
    assert [t.name for t in basis.kinder] == [u"AW_36.5", u"IW_11.5"]
    aw = basis.kinder[0]
    assert [b.element_id for b in aw.kinder] == [10, 11]
    assert aw.kinder[0].ebene == bm.ELEMENT
    # natürliche Sortierung: "Tür 2" vor "Tür 10"
    assert [f.name for f in w.kinder[1].kinder] == [u"Tür 2", u"Tür 10"]
    assert len(bm.knoten_liste(w)) == 1 + 3 + 4 + 5 + 6


def test_zustand():
    w = bm.baue(DATEN)
    waende = w.kinder[2]
    assert bm.zustand(waende, set()) is False
    assert bm.zustand(waende, {10, 11, 12}) is True
    assert bm.zustand(waende, {10}) is None
    assert bm.zaehler_text(waende, {10}) == u"1 / 3"
    assert bm.zaehler_text(waende, {10, 11, 12, 20}) == u"3"
    assert bm.zaehler_text(waende, set()) == u"3"
    for markiert in (set(), {10}, {10, 11, 12}, {20}):
        assert bm.anzeige(waende, markiert) == (
            bm.zustand(waende, markiert), bm.zaehler_text(waende, markiert))


def test_natuerlich():
    namen = [u"W 10", u"w 2", u"W 1a", u""]
    assert sorted(namen, key=bm.natuerlich) == [u"", u"W 1a", u"w 2",
                                                u"W 10"]


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items())
             if n.startswith("test_") and callable(f)]
    for name, funktion in tests:
        funktion()
        print("  ok:", name)
    print("%d Tests bestanden" % len(tests))
