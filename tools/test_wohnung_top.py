# -*- coding: utf-8 -*-
"""Offline-Test der Logik von ApartmentNumbering (ohne Revit).

    python tools\\test_wohnung_top.py
"""
import os
import sys

os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from stuetzen_nummer import logik as nummern  # noqa: E402
from wohnung_top import logik as lg  # noqa: E402


def test_name_zerlegen():
    assert lg.zerlege(u"WHG_TYP A2a - TU_OG07-OG12") == (
        u"A2a", u"TU", u"OG07-OG12")
    assert lg.zerlege(u"WHG_TYP B1d - TU_P10") == (u"B1d", u"TU", u"P10")
    for falsch in (u"WHG_TYP B1d - TU P10", u"WHG_TYP A2a", u"Bad 01",
                   u"WHG_TYP 2a - TU_OG01", u""):
        assert lg.zerlege(falsch) is None, falsch
    assert lg.ist_wohnung(u"WHG_TYP B1d - TU P10")
    assert not lg.ist_wohnung(u"Bad 01")


def test_werte():
    w = lg.werte(u"WHG_TYP A2a - TU_OG07-OG12", u" TU ", u"LBS 45")
    assert w == {lg.GRUPPE_TYP: u"WHG_TYP A2a - TU_OG07-OG12",
                 lg.TYPNAME: u"A2a", lg.KATEGORIE: u"A-TYP",
                 lg.GESCHOSS: u"OG07-OG12", lg.BAUTEIL: u"TU",
                 lg.STIEGE: u"LBS 45"}
    # Leere Eingaben werden nicht geschrieben
    w = lg.werte(u"WHG_TYP b1d - TU_EG", u"", u"  ")
    assert lg.BAUTEIL not in w and lg.STIEGE not in w
    assert w[lg.KATEGORIE] == u"B-TYP"
    assert lg.werte(u"WHG_TYP X") is None


QUADRAT = [(0.0, 0.0), (10.0, 0.0), (10.0, 10.0), (0.0, 10.0)]


def _verschoben(ring, dx):
    return [(x + dx, y) for x, y in ring]


def test_innen_mit_loch():
    loch = [(4.0, 4.0), (6.0, 4.0), (6.0, 6.0), (4.0, 6.0)]
    assert lg.innen((1.0, 1.0), [QUADRAT])
    assert not lg.innen((11.0, 1.0), [QUADRAT])
    assert not lg.innen((5.0, 5.0), [QUADRAT, loch])


def test_eintritt():
    seg = nummern.segmente([[(-5.0, 5.0), (25.0, 5.0)]])
    assert abs(lg.eintritt(seg, [QUADRAT]) - 5.0) < 1e-9
    # beginnt die Linie innen, zählt ihr Anfang
    seg_innen = nummern.segmente([[(2.0, 5.0), (25.0, 5.0)]])
    assert lg.eintritt(seg_innen, [QUADRAT]) == 0.0
    vorbei = nummern.segmente([[(-5.0, 20.0), (25.0, 20.0)]])
    assert lg.eintritt(vorbei, [QUADRAT]) is None
    # Wohnung aus zwei Räumen: der frühere zählt
    zwei = [[_verschoben(QUADRAT, 12.0)], [QUADRAT]]
    assert abs(lg.eintritt_mehrere(seg, zwei) - 5.0) < 1e-9


def test_reihenfolge_ueber_geschosse():
    # drei Wohnungen je Geschoss nebeneinander, Linie von rechts nach links
    seg = nummern.segmente([[(40.0, 5.0), (-5.0, 5.0)]])
    wohnungen = {u"A": [[QUADRAT]],
                 u"B": [[_verschoben(QUADRAT, 12.0)]],
                 u"C": [[_verschoben(QUADRAT, 24.0)]],
                 u"X": [[[(0.0, 30.0), (5.0, 30.0), (5.0, 35.0)]]]}
    treffer = []
    for hoehe, geschoss in ((10.0, u"OG1"), (0.0, u"EG")):
        for name, flaechen in sorted(wohnungen.items()):
            station = lg.eintritt_mehrere(seg, flaechen)
            if station is not None:
                treffer.append((hoehe, station, (geschoss, name)))
    folge = lg.reihenfolge(treffer)
    assert folge == [(u"EG", u"C"), (u"EG", u"B"), (u"EG", u"A"),
                     (u"OG1", u"C"), (u"OG1", u"B"), (u"OG1", u"A")]
    zaehler = nummern.Zaehler(u"01")
    nummeriert = [(g, zaehler.weiter()) for g, _ in folge]
    assert nummeriert[-1] == (u"OG1", u"06")
    assert lg.bereiche(nummeriert) == [(u"EG", u"01", u"03"),
                                       (u"OG1", u"04", u"06")]


def test_verbunden():
    # Zug aus drei Stücken, dazu eine fremde Linie und eine auf anderer Höhe
    enden = [((0.0, 0.0, 0.0), (10.0, 0.0, 0.0)),
             ((10.0, 5.0, 0.0), (10.0, 0.0, 0.0)),      # verkehrt gezeichnet
             ((10.0, 5.0, 0.0), (0.0, 5.0, 0.0)),
             ((50.0, 50.0, 0.0), (60.0, 50.0, 0.0)),
             ((0.0, 0.0, 10.0), (5.0, 0.0, 10.0))]
    assert lg.verbunden(enden, [2], 0.01) == [0, 1, 2]
    assert lg.verbunden(enden, [3], 0.01) == [3]
    assert lg.verbunden(enden, [0, 3], 0.01) == [0, 1, 2, 3]


if __name__ == "__main__":
    for name, funktion in sorted(globals().items()):
        if name.startswith("test_") and callable(funktion):
            funktion()
            print("ok ", name)
