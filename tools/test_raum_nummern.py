# -*- coding: utf-8 -*-
"""Offline-Test der festen Raumnummern (AutoRooms, ApartmentNumbering).

    python tools\test_raum_nummern.py
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

import raum_nummern as rn  # noqa: E402


def test_erkennen():
    assert rn.erkenne(u"WKÜ") == (rn.WOHNKUECHE, None)
    assert rn.erkenne(u"Wohnküche") == (rn.WOHNKUECHE, None)
    assert rn.erkenne(u"ZI 2") == (rn.ZIMMER, 2)
    assert rn.erkenne(u"ZIMMER 3") == (rn.ZIMMER, 3)
    assert rn.erkenne(u"Zimmer1") == (rn.ZIMMER, 1)
    assert rn.erkenne(u"B") == (rn.BAD, None)
    assert rn.erkenne(u"Bad 2") == (rn.BAD, 2)
    assert rn.erkenne(u"VR") == (rn.VORRAUM, None)
    assert rn.erkenne(u"v") == (rn.VORRAUM, None)
    assert rn.erkenne(u"Gang") == (rn.GANG, None)
    assert rn.erkenne(u"Schrankraum") == (rn.SR, None)
    assert rn.erkenne(u"Abstellraum") == (rn.AR, None)
    for name in (u"AR 2", u"Küche", u"Treppenhaus", u"", None):
        assert rn.erkenne(name) is None, name


def test_tabelle():
    raeume = [("v", u"V", 6.0), ("g", u"G", 4.0), ("b", u"B", 5.0),
              ("wk", u"WKÜ", 30.0), ("wc", u"WC", 2.0),
              ("z1", u"ZI 1", 14.0), ("z2", u"ZI 2", 12.0),
              ("z3", u"ZI 3", 10.0), ("ar", u"AR", 2.0),
              ("sr", u"SR", 3.0), ("th", u"Treppenhaus", 20.0)]
    assert rn.nummern(raeume) == {
        "v": u"001", "g": u"002", "b": u"003", "wk": u"005", "wc": u"011",
        "z1": u"006", "z2": u"007", "z3": u"008", "ar": u"009",
        "sr": u"010"}


def test_gang_ohne_vorraum():
    assert rn.nummern([("g", u"GANG", 8.0), ("b", u"BAD", 5.0)]) == {
        "g": u"001", "b": u"003"}


def test_reihen_ohne_index():
    # zwei Bäder ohne Nummer: das grössere ist Bad 1
    assert rn.nummern([("klein", u"BAD", 4.0), ("gross", u"BAD", 7.0)]) == {
        "gross": u"003", "klein": u"004"}
    # "BAD 1" ist belegt, das andere bekommt 004
    assert rn.nummern([("a", u"BAD", 9.0), ("b", u"Bad 1", 4.0)]) == {
        "a": u"004", "b": u"003"}
    # ZI 4 und ein drittes Bad gibt es in der Tabelle nicht
    nummern = rn.nummern([("z", u"ZI 4", 9.0), ("b1", u"B", 5.0),
                          ("b2", u"B", 4.0), ("b3", u"B", 3.0)])
    assert nummern == {"b1": u"003", "b2": u"004"}


if __name__ == "__main__":
    for name, funktion in sorted(globals().items()):
        if name.startswith("test_") and callable(funktion):
            funktion()
            print("ok ", name)
