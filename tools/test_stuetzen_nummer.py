# -*- coding: utf-8 -*-
"""Offline-Test der Logik von ColumnNumbering (ohne Revit).

    python tools\test_stuetzen_nummer.py
"""
import os
import sys

os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from stuetzen_nummer import logik as lg  # noqa: E402


def test_zaehler():
    z = lg.Zaehler(u"S-08", 1)
    assert [z.weiter() for _ in range(3)] == [u"S-08", u"S-09", u"S-10"]
    assert z.aktuell == u"S-11"
    z = lg.Zaehler(u"1", 2)
    assert [z.weiter() for _ in range(3)] == [u"1", u"3", u"5"]
    z = lg.Zaehler(u"ST1.01a", 1)
    assert [z.weiter() for _ in range(2)] == [u"ST1.01a", u"ST1.02a"]


def test_vorschau_und_pruefung():
    assert lg.vorschau(u"P01", 1, 3) == u"P01, P02, P03 …"
    assert lg.pruefe_start(u"S-1") is None
    assert lg.pruefe_start(u"S-") is not None
    assert lg.pruefe_start(u"  ") is not None
    for schritt in (0,):
        try:
            lg.Zaehler(u"1", schritt)
        except ValueError:
            pass
        else:
            raise AssertionError(u"Schritt 0 angenommen")
    try:
        lg.vorschau(u"2", -1, 4)            # 2, 1, 0, -1
    except ValueError:
        pass
    else:
        raise AssertionError(u"negative Nummer angenommen")
    assert lg.zahl(u"0,5") == 0.5 and lg.zahl(u" 1.25 ") == 1.25


def test_kette_in_zeichenrichtung():
    # Polylinie in drei Stücken, in beliebiger Reihenfolge gewählt
    a = [(0.0, 0.0), (10.0, 0.0)]
    b = [(10.0, 0.0), (10.0, 5.0)]
    c = [(10.0, 5.0), (0.0, 5.0)]
    ketten = lg.ketten([c, a, b])
    assert len(ketten) == 1
    assert ketten[0] == [(0.0, 0.0), (10.0, 0.0), (10.0, 5.0), (0.0, 5.0)]


def test_kette_mehrheit_bestimmt_richtung():
    # Ein Stück verkehrt herum gezeichnet: die Mehrheit gewinnt
    a = [(0.0, 0.0), (10.0, 0.0)]
    b = [(20.0, 0.0), (10.0, 0.0)]
    c = [(20.0, 0.0), (30.0, 0.0)]
    ketten = lg.ketten([b, a, c])
    assert ketten[0][0] == (0.0, 0.0) and ketten[0][-1] == (30.0, 0.0)


def test_getrennte_reihen_in_klickreihenfolge():
    oben = [(0.0, 10.0), (30.0, 10.0)]
    unten = [(0.0, 0.0), (30.0, 0.0)]
    seg = lg.segmente(lg.ketten([oben, unten]))
    punkte = [(20.0, 0.1), (10.0, 9.8), (0.0, 0.0), (29.0, 10.0),
              (15.0, 5.0)]                    # letzter: zu weit weg
    assert lg.entlang(seg, punkte, 0.5) == [1, 3, 2, 0]


def test_schlangenlinie():
    # Zeile 1 nach rechts, Zeile 2 zurück nach links
    pfad = [(0.0, 0.0), (30.0, 0.0), (30.0, 10.0), (0.0, 10.0)]
    seg = lg.segmente(lg.ketten([pfad]))
    punkte = [(0.0, 10.0), (10.0, 0.0), (20.0, 10.0), (0.0, 0.0),
              (20.0, 0.0)]
    assert lg.entlang(seg, punkte, 0.5) == [3, 1, 4, 2, 0]


def test_naechster_und_uebereinander():
    seg = lg.segmente([[(0.0, 0.0), (10.0, 0.0)]])
    abstand, station = lg.naechster(seg, (4.0, 3.0))
    assert abs(abstand - 3.0) < 1e-9 and abs(station - 4.0) < 1e-9
    punkte = [(1.0, 1.0), (1.01, 1.0), (5.0, 1.0)]
    assert lg.uebereinander((1.0, 1.0), punkte, 0.05) == [0, 1]


if __name__ == "__main__":
    for name, funktion in sorted(globals().items()):
        if name.startswith("test_") and callable(funktion):
            funktion()
            print("ok ", name)
