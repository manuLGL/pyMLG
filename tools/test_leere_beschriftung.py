# -*- coding: utf-8 -*-
"""Offline-Test der Logik von EmptyTags (ohne Revit).

    python tools\\test_leere_beschriftung.py
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from leere_beschriftung import logik as lg  # noqa: E402

TYP_A = (u"T", 1)
TYP_B = (u"T", 2)
LEER = u"AS OK= üFOK\n"
VOLL = u"AS OK=112,00 üFOK\nVSG"


def test_normiere():
    assert lg.normiere(LEER) == u"AS OK= üFOK"
    assert lg.normiere(u"a  \r\nb\r\n\r\n") == u"a\nb"
    assert lg.normiere(None) is None


def test_treffer_nur_gleicher_text_und_typ():
    vorlagen = lg.muster([(TYP_A, LEER)])
    kandidaten = [(10, TYP_A, LEER), (11, TYP_A, VOLL),
                  (12, TYP_B, LEER), (13, TYP_A, u"AS OK= üFOK"),
                  (10, TYP_A, LEER)]
    assert lg.treffer(kandidaten, vorlagen) == [10, 13]


def test_mehrere_muster_je_typ():
    vorlagen = lg.muster([(TYP_A, LEER), (TYP_B, u"?"), (TYP_A, None)])
    kandidaten = [(1, TYP_A, LEER), (2, TYP_B, u"?"), (3, TYP_B, LEER),
                  (4, TYP_A, None)]
    assert lg.treffer(kandidaten, vorlagen) == [1, 2]


if __name__ == "__main__":
    for name, funktion in sorted(globals().items()):
        if name.startswith("test_") and callable(funktion):
            funktion()
            print("ok  ", name)
