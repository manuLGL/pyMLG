# -*- coding: utf-8 -*-
"""Offline-Test der Logik von GroupExport und GroupToLink (ohne Revit).

    python tools\\test_gruppen_export.py
"""
import os
import sys

os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from gruppen_export import logik as lg  # noqa: E402


def test_dateiname():
    assert lg.dateiname(u"WHG_TYP A2a - TU_OG07-OG12") == \
        u"WHG_TYP A2a - TU_OG07-OG12"
    assert lg.dateiname(u'Bad 1/2: "neu"?') == u"Bad 1_2_ _neu__"
    assert lg.dateiname(u"  Küche.  ") == u"Küche"
    assert lg.dateiname(u"") == lg.ERSATZ
    assert lg.dateiname(u" . ") == lg.ERSATZ
    assert lg.dateiname(u"con") == u"con_"
    assert lg.dateiname(u"LPT1.alt") == u"LPT1.alt_"
    assert lg.dateiname(u"Console") == u"Console"


def test_dateinamen_eindeutig():
    assert lg.dateinamen([u"A", u"a", u"B", u"A"]) == [
        u"A.rvt", u"a (2).rvt", u"B.rvt", u"A (3).rvt"]
    assert lg.dateinamen([u"X/1", u"X:1"]) == [u"X_1.rvt", u"X_1 (2).rvt"]


def test_zuordnung():
    dateien = [u"WHG_TYP A2a.rvt", u"WHG_TYP A2a.0001.rvt", u"bad 1_2.RVT",
               u"Notiz.txt", u"B.0003.rvt"]
    assert lg.zuordnung([u"whg_typ a2a", u"Bad 1/2", u"B", u"C"], dateien) == {
        u"whg_typ a2a": u"WHG_TYP A2a.rvt", u"Bad 1/2": u"bad 1_2.RVT"}
    assert lg.zuordnung([u"X"], []) == {}


if __name__ == "__main__":
    for name, funktion in sorted(globals().items()):
        if name.startswith("test_") and callable(funktion):
            funktion()
            print("ok  ", name)
