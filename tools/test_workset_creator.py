# -*- coding: utf-8 -*-
"""Offline-Test der Logik des Workset-Creators (ohne Revit).

    python tools\\test_workset_creator.py
"""
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from workset_creator import logik as lg  # noqa: E402


def test_namen_aus_text():
    text = (u"﻿01_Raster\r\n\r\n  02_Wände  \n"
            u"03_Decken\t04_Dach\n\"05_Treppen\"\n02_WÄNDE\n")
    assert lg.namen_aus_text(text) == [
        u"01_Raster", u"02_Wände", u"03_Decken", u"04_Dach", u"05_Treppen"]
    assert lg.namen_aus_text(u"") == []
    assert lg.namen_aus_text(None) == []


def test_zusammenfuehren():
    liste, hinzu = lg.zusammenfuehren([u"A", u"B"], [u"b", u"C", u"C"])
    assert liste == [u"A", u"B", u"C"] and hinzu == 1
    liste, hinzu = lg.zusammenfuehren([u"A"], [u"X", u"a"], anhaengen=False)
    assert liste == [u"X", u"a"] and hinzu == 2


def test_filter():
    f = lg.filterfunktion(u"links haus8")
    assert f(u"01.3_LINKS_PW_Haus8")
    assert not f(u"01.4_LINKS_PW_Haus10")
    assert lg.filterfunktion(u"")(u"egal")
    r = lg.filterfunktion(r"^0\d_", regex=True)
    assert r(u"04_LINKS") and not r(u"50_H30")
    assert lg.filterfunktion(u"", regex=True)(u"egal")
    try:
        lg.filterfunktion(u"([", regex=True)
    except re.error:
        pass
    else:
        raise AssertionError(u"ungültiger Regex muss re.error werfen")


def test_gueltig():
    assert lg.ist_gueltig(u"01_Raster Ebenen")
    for name in (u"", u"  ", u"A:B", u"A[1]", u"A;B", u"A<B"):
        assert not lg.ist_gueltig(name), name


def test_plane_ueberspringen():
    plan = lg.plane([u"Neu", u"wände", u"Fal|sch"], [u"Wände"],
                    umbenennen=False)
    assert plan == [(u"Neu", u"Neu", lg.NEU),
                    (u"wände", None, lg.UEBERSPRUNGEN),
                    (u"Fal|sch", None, lg.UNGUELTIG)]


def test_plane_umbenennen():
    plan = lg.plane([u"Wände", u"Wände (2)", u"Dach"],
                    [u"Wände", u"Wände (3)"], umbenennen=True)
    # "Wände" -> (2) ist frei; danach ist "Wände (2)" belegt -> (4),
    # weil (3) im Projekt existiert und (2) gerade vergeben wurde
    assert plan == [(u"Wände", u"Wände (2)", lg.UMBENANNT),
                    (u"Wände (2)", u"Wände (2) (2)", lg.UMBENANNT),
                    (u"Dach", u"Dach", lg.NEU)]
    assert lg.eindeutiger_name(u"A", {u"a (2)", u"a (3)"}) == u"A (4)"
    assert lg.eindeutiger_name(u"A", set(), u"{name}_{n}") == u"A_2"


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items())
             if n.startswith("test_") and callable(f)]
    for name, funktion in tests:
        funktion()
        print("  ok:", name)
    print("%d Tests bestanden" % len(tests))
