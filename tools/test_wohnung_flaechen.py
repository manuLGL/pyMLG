# -*- coding: utf-8 -*-
"""Offline-Test der Logik von ApartmentAreas (ohne Revit).

    python tools\test_wohnung_flaechen.py
"""
import os
import sys

os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from wohnung_flaechen import logik as lg  # noqa: E402


def test_zuordnung():
    erwartet = {
        u"1_W_WOHNEN": lg.WNFL,
        u"2_W_NEBENRAUM": lg.WNFL,
        u"2_W_SAN_GARDEROBE": lg.WNFL,
        u"2_W_SAN_WELLNESS": lg.WNFL,
        u"2_W_SANITÄR": lg.WNFL,
        u"2_W_VERKEHRSFLÄCHE": lg.WNFL,
        u"2_WA1_LOGGIA": lg.LOGGIA,
        u"2_WA2_BALKON": lg.BALKON,
        u"2_WA3_TERRASSE": lg.TERRASSE,
        u"2_WA4_TERRASSE (EG)": lg.TERRASSE_EG,
        u"2_WA5.1_TERRASSE (DACH)": lg.TERRASSE_DT,
        u"2_WA5.2_EIGENGARTEN (DACH)": None,
        u"2_WA6_EIGENGARTEN": None,
        u"2_WA_WINTERGARTEN": None,
        u"2_WE_EINLAGERUNG": None,
        u"3_GW_BESPRECHUNG": lg.GW_NFL,
        u"3_GW_ERSCHLIE�UNG": lg.GW_NFL,     # kaputtes ß
        u"3_GW_SANITÄR": lg.GW_NFL,
        u"4_GW_VERK_BÜRO": None,
        u"4_GW_VERK_ERSCHLIEßUNG": None,
        u"1_w_wohnen": lg.WNFL,
        u"": None,
        None: None,
        u"WOHNEN": None,
    }
    for kategorie, name in erwartet.items():
        assert lg.parameter_fuer(kategorie) == name, kategorie


def test_summen():
    werte, rest = lg.summen([
        (u"1_W_WOHNEN", 200.0), (u"2_W_SANITÄR", 50.0),
        (u"2_WA2_BALKON", 30.0), (u"2_WA6_EIGENGARTEN", 99.0),
        (u"", 5.0), (u"3_GW_BÜRO", 10.0)])
    assert werte[lg.WNFL] == 250.0
    assert werte[lg.BALKON] == 30.0
    assert werte[lg.GW_NFL] == 10.0
    assert werte[lg.LOGGIA] == 0.0
    assert set(werte) == set(lg.ALLE)
    assert rest == {u"2_WA6_EIGENGARTEN": 1, u"-": 1}


def test_qm():
    assert abs(lg.qm(100.0) - 9.290304) < 1e-9
    assert lg.qm_text(100.0) == u"9,29 m²"


if __name__ == "__main__":
    for name, funktion in sorted(globals().items()):
        if name.startswith("test_") and callable(funktion):
            funktion()
            print("ok ", name)
