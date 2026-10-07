# -*- coding: utf-8 -*-
"""Offline-Test von ComponentLegend: Typen sammeln, sortieren, beschriften
(ohne Revit).

    python tools\test_bauteil_legende.py
"""
import os
import sys

os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from bauteil_legende import logik as lg  # noqa: E402

TUEREN, FENSTER = -2000023, -2000014

# (kat_id, kategorie, typ, ref, familie, typname, element)
ANSICHT_EG = [
    (TUEREN, u"Türen", 10, u"ref10", u"Tür 1-flg", u"T10", 1),
    (TUEREN, u"Türen", 11, u"ref11", u"Tür 1-flg", u"T2", 2),
    (TUEREN, u"Türen", 11, u"ref11", u"Tür 1-flg", u"T2", 3),
    (FENSTER, u"Fenster", 20, u"ref20", u"Fenster", u"F1", 4),
]
ANSICHT_OG = [
    (TUEREN, u"Türen", 11, u"ref11", u"Tür 1-flg", u"T2", 3),    # auch im EG sichtbar
    (TUEREN, u"Türen", 12, u"ref12", u"Schiebetür", u"S1", 5),
]


def test_sammeln_zaehlt_elemente_einmal():
    typen = lg.sammle(ANSICHT_EG)
    lg.sammle(ANSICHT_OG, typen)
    assert sorted(typen) == [10, 11, 12, 20]
    assert typen[11].anzahl == 2, typen[11].elemente
    assert typen[11].ref == u"ref11"


def test_kategorien_mit_typzahl():
    typen = lg.sammle(ANSICHT_EG + ANSICHT_OG)
    assert lg.kategorien(typen) == [(FENSTER, u"Fenster", 1), (TUEREN, u"Türen", 3)]


def test_reihenfolge_natuerlich_nach_familie_und_typ():
    typen = lg.sammle(ANSICHT_EG + ANSICHT_OG)
    namen = [(t.familie, t.name) for t in lg.reihenfolge(typen, [TUEREN])]
    # Schiebetür vor Tür 1-flg; T2 vor T10
    assert namen == [(u"Schiebetür", u"S1"), (u"Tür 1-flg", u"T2"),
                     (u"Tür 1-flg", u"T10")], namen
    alle = lg.reihenfolge(typen, [TUEREN, FENSTER])
    assert alle[0].kategorie == u"Fenster" and len(alle) == 4
    assert lg.reihenfolge(typen, []) == []


def test_beschriftung():
    typ = lg.Typ(TUEREN, u"Türen", 1, None, u"Tür 1-flg", u"T2")
    assert lg.beschriftung(typ, lg.BESCHRIFTUNG_KEINE) == u""
    assert lg.beschriftung(typ, lg.BESCHRIFTUNG_TYP) == u"T2"
    assert lg.beschriftung(typ, lg.BESCHRIFTUNG_FAMILIE_TYP) == u"Tür 1-flg: T2"
    # Systemfamilie, deren Familienname gleich dem Typ ist
    gleich = lg.Typ(TUEREN, u"Türen", 2, None, u"X", u"X")
    assert lg.beschriftung(gleich, lg.BESCHRIFTUNG_FAMILIE_TYP) == u"X"


def test_eingaben():
    assert lg.pruefe_massstab(u"50") == 50
    assert lg.pruefe_massstab(u"50,5") is None
    assert lg.pruefe_massstab(u"0") is None
    assert lg.pruefe_massstab(u"abc") is None
    assert lg.pruefe_abstand(u"7,5") == 7.5
    assert lg.pruefe_abstand(u"0") == 0.0
    assert lg.pruefe_abstand(u"-1") is None
    assert lg.pruefe_abstand(u"") is None


if __name__ == "__main__":
    for name, funktion in sorted(globals().items()):
        if name.startswith("test_") and callable(funktion):
            funktion()
            print("ok ", name)
