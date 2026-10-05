# -*- coding: utf-8 -*-
"""Offline-Test der Logik von LinkedViews (ohne Revit).

    python tools\test_link_ansichten.py
"""
import os
import sys

os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from link_ansichten import logik as lg  # noqa: E402


class _Ansicht(object):
    def __init__(self, wert, vorlage=None):
        self.wert = wert
        self.vorlage = vorlage
        self.Name = u"A%d" % wert


def test_gruppen():
    # Im Grundriss sind auch Flächen- und Deckenpläne wählbar
    assert lg.passt("FloorPlan", "AreaPlan")
    assert lg.passt("CeilingPlan", "FloorPlan")
    assert lg.passt("Section", "Detail")
    assert not lg.passt("FloorPlan", "Section")
    assert not lg.passt("Elevation", "Section")
    # Bauteillisten, Pläne, Legenden: keine verknüpfte Ansicht
    assert lg.gruppe("Schedule") is None
    assert not lg.passt("DrawingSheet", "DrawingSheet")


def test_suche():
    text = u"grundriss: ug02 - plan ebene ug02 ar.rvt"
    assert lg.trifft(text, lg.suchwoerter(u"UG02 ar"))
    assert lg.trifft(text, lg.suchwoerter(u""))
    assert not lg.trifft(text, lg.suchwoerter(u"ug02 tw"))


def test_doppelte():
    paare = [(1, u"AR.rvt"), (2, u"TW.rvt"), (1, u"AR.rvt"), (1, u"AR.rvt")]
    assert lg.doppelte(paare) == [u"AR.rvt"]
    assert lg.doppelte([(1, u"AR"), (2, u"TW")]) == []


def test_ziele_mit_vorlage():
    vorlage = _Ansicht(100)
    a1 = _Ansicht(1, vorlage)
    a2 = _Ansicht(2)
    a3 = _Ansicht(3, vorlage)
    ziele, umgeleitet = lg.ziele([a1, a2, a3, vorlage],
                                 lambda a: a.vorlage, lambda e: e.wert)
    # Vorlage nur einmal, freie Ansicht direkt
    assert [z.wert for z in ziele] == [100, 2]
    assert list(umgeleitet) == [100]
    assert [a.wert for a in umgeleitet[100][1]] == [1, 3]


def test_kurzliste():
    assert lg.kurzliste([u"a", u"b"]) == u"a, b"
    assert lg.kurzliste([u"x"] * 10, 3) == u"x, x, x … und 7 weitere"


if __name__ == "__main__":
    for name, funktion in sorted(globals().items()):
        if name.startswith("test_"):
            funktion()
            print("ok  ", name)
