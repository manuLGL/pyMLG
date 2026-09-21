# -*- coding: utf-8 -*-
"""Offline-Test der Logik von TransferSingle (ohne Revit).

    python tools\\test_transfer_single.py
"""
import os
import sys

# Texte werden auf Deutsch verglichen
os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from filter_more import baum as bm  # noqa: E402
from transfer_single import logik as lg  # noqa: E402

DATEN = [
    (u"Wände – Typen", u"Basiswand: AW 36.5", 10),
    (u"Wände – Typen", u"Basiswand: AW 11.5", 11),
    (u"Ansichtsvorlagen", u"GR 1:100", 20),
    (u"Ansichtsvorlagen", u"GR 1:50", 21),
    (u"Materialien", u"Beton", 30),
]


def test_baum():
    w = lg.baue(DATEN)
    assert sorted(w.ids) == [10, 11, 20, 21, 30]
    assert [g.name for g in w.kinder] == [u"Ansichtsvorlagen",
                                          u"Materialien", u"Wände – Typen"]
    waende = w.kinder[2]
    assert [b.name for b in waende.kinder] == [u"Basiswand: AW 11.5",
                                               u"Basiswand: AW 36.5"]
    assert waende.kinder[0].ebene == lg.ELEMENT
    assert bm.anzeige(waende, {10}) == (None, u"1 / 2")
    assert len(lg.blaetter(w)) == 5


def test_suche():
    w = lg.baue(DATEN)
    # natürliche Sortierung: "GR 1:50" (21) vor "GR 1:100" (20)
    gruppe, blatt = lg.naechster_treffer(w, u"gr 1")
    assert blatt.element_id == 21 and gruppe.name == u"Ansichtsvorlagen"
    assert lg.naechster_treffer(w, u"gr 1", 21)[1].element_id == 20
    # am Ende wieder von vorn
    assert lg.naechster_treffer(w, u"gr 1", 20)[1].element_id == 21
    assert lg.naechster_treffer(w, u"gibtsnicht") is None
    assert lg.naechster_treffer(w, u"   ") is None


def test_umbenennen():
    N = lg.neuer_name
    assert N(u"Wand", lg.PRAEFIX, u"X_") == u"X_Wand"
    assert N(u"Wand", lg.SUFFIX, u"_neu") == u"Wand_neu"
    assert N(u"Wand äb", lg.GROSS) == u"WAND ÄB"
    assert N(u"WAND", lg.KLEIN) == u"wand"
    assert N(u"aUSSENWAND beton-kern 36.5", lg.ERSTER_GROSS) == \
        u"Aussenwand Beton-Kern 36.5"
    assert N(u"AW_36.5_AW", lg.ERSETZEN, u"AW", u"IW") == u"IW_36.5_IW"
    assert N(u"AW", lg.ERSETZEN, u"", u"x") == u"AW"
    assert N(u"EG 01 - Raum 1 / 11", lg.ZAHLEN, u"1", u"2") == \
        u"EG 02 - Raum 2 / 11"
    try:
        N(u"x", lg.ZAHLEN, u"a", u"2")
    except ValueError:
        pass
    else:
        raise AssertionError(u"ungültige Zahl muss ValueError werfen")


def test_freie_nummer():
    vergeben = {u"a101", u"a101-2"}
    assert lg.freie_nummer(u"A101", vergeben) == u"A101-3"
    assert lg.freie_nummer(u"A102", vergeben) == u"A102"
    assert lg.freie_nummer(u"A102", vergeben) == u"A102-2"


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items())
             if n.startswith("test_") and callable(f)]
    for name, funktion in tests:
        funktion()
        print("  ok:", name)
    print("%d Tests bestanden" % len(tests))
