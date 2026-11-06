# -*- coding: utf-8 -*-
"""Offline-Test von FilterLegend: Füllung aus Überschreibungen wählen,
Ansichten zusammenführen, sortieren, Raster (ohne Revit).

    python tools\test_filter_legende.py
"""
import os
import sys

os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from filter_legende import logik as lg  # noqa: E402

ROT, BLAU, GRAU = (255, 0, 0), (0, 0, 255), (128, 128, 128)
SCHRAFFUR = 4711


def test_farbe_ohne_muster_wird_vollton():
    grafik = lg.Grafik(oberflaeche=lg.Fuellung(vg_farbe=ROT))
    assert lg.darstellung(grafik) == lg.Fuellung(lg.VOLL, ROT)


def test_muster_ohne_farbe_wird_schwarz():
    grafik = lg.Grafik(oberflaeche=lg.Fuellung(vg_muster=SCHRAFFUR))
    assert lg.darstellung(grafik) == lg.Fuellung(SCHRAFFUR, lg.SCHWARZ)


def test_vordergrund_und_hintergrund():
    grafik = lg.Grafik(oberflaeche=lg.Fuellung(SCHRAFFUR, ROT, lg.VOLL, GRAU))
    assert lg.darstellung(grafik) == lg.Fuellung(SCHRAFFUR, ROT, lg.VOLL, GRAU)


def test_vollton_vorne_verdeckt_hintergrund():
    grafik = lg.Grafik(oberflaeche=lg.Fuellung(lg.VOLL, ROT, SCHRAFFUR, GRAU))
    assert lg.darstellung(grafik) == lg.Fuellung(lg.VOLL, ROT)


def test_automatisch_reihenfolge():
    nur_schnitt = lg.Grafik(schnitt=lg.Fuellung(lg.VOLL, BLAU), linie=ROT)
    assert lg.darstellung(nur_schnitt) == lg.Fuellung(lg.VOLL, BLAU)
    beides = lg.Grafik(oberflaeche=lg.Fuellung(lg.VOLL, ROT),
                       schnitt=lg.Fuellung(lg.VOLL, BLAU))
    assert lg.darstellung(beides) == lg.Fuellung(lg.VOLL, ROT)
    nur_linie = lg.Grafik(linie=GRAU)
    assert lg.darstellung(nur_linie) == lg.Fuellung(lg.VOLL, GRAU)
    nur_schnittlinie = lg.Grafik(schnittlinie=BLAU)
    assert lg.darstellung(nur_schnittlinie) == lg.Fuellung(lg.VOLL, BLAU)
    assert lg.darstellung(lg.Grafik()) is None
    assert lg.darstellung(lg.Grafik(oberflaeche=lg.Fuellung())) is None


def test_feste_quelle():
    grafik = lg.Grafik(oberflaeche=lg.Fuellung(lg.VOLL, ROT),
                       schnitt=lg.Fuellung(lg.VOLL, BLAU), schnittlinie=GRAU)
    assert lg.darstellung(grafik, lg.QUELLE_SCHNITT) == lg.Fuellung(lg.VOLL, BLAU)
    assert lg.darstellung(grafik, lg.QUELLE_OBERFLAECHE) == lg.Fuellung(lg.VOLL, ROT)
    nur_schnittlinie = lg.Grafik(oberflaeche=lg.Fuellung(lg.VOLL, ROT), schnittlinie=GRAU)
    assert lg.darstellung(nur_schnittlinie, lg.QUELLE_SCHNITT) == lg.Fuellung(lg.VOLL, GRAU)
    assert lg.darstellung(lg.Grafik(schnitt=lg.Fuellung(lg.VOLL, BLAU)),
                          lg.QUELLE_OBERFLAECHE) is None


def _eintrag(schluessel, name, farbe=None, ansicht=u"EG", position=0, **art):
    grafik = lg.Grafik(oberflaeche=lg.Fuellung(vg_farbe=farbe)) if farbe else lg.Grafik()
    eintrag = lg.Eintrag(schluessel, u"ref%d" % schluessel, name, grafik,
                         ansicht=ansicht, **art)
    eintrag.position = position
    return eintrag


def test_vorgewaehlt():
    assert lg.vorgewaehlt(_eintrag(1, u"Wand", ROT))
    assert not lg.vorgewaehlt(_eintrag(1, u"Wand"))
    assert not lg.vorgewaehlt(_eintrag(1, u"Wand", ROT, sichtbar=False))
    assert not lg.vorgewaehlt(_eintrag(1, u"Wand", ROT, aktiv=False))


def test_zusammenfuehren_erster_gilt_und_merkt_abweichung():
    eg = [_eintrag(1, u"Tragend", ROT), _eintrag(2, u"Bestand", GRAU, position=1)]
    og = [_eintrag(3, u"Abbruch", BLAU, u"OG"), _eintrag(1, u"Tragend", BLAU, u"OG"),
          _eintrag(2, u"Bestand", GRAU, u"OG")]
    alle = lg.zusammenfuehren([eg, og])
    assert [e.name for e in alle] == [u"Tragend", u"Bestand", u"Abbruch"]
    assert [e.position for e in alle] == [0, 1, 2]
    assert alle[0].abweichend == [u"OG"]
    assert alle[0].grafik.oberflaeche.vg_farbe == ROT
    assert alle[1].abweichend == []
    # erneutes Zusammenführen setzt abweichend zurück
    assert lg.zusammenfuehren([eg])[0].abweichend == []


def test_sortieren():
    eintraege = [_eintrag(1, u"F10", position=0), _eintrag(2, u"f2", position=1),
                 _eintrag(3, u"A", position=2)]
    assert [e.name for e in lg.sortiere(eintraege)] == [u"F10", u"f2", u"A"]
    assert [e.name for e in lg.sortiere(eintraege, lg.REIHENFOLGE_NAME)] == \
        [u"A", u"f2", u"F10"]


def test_raster():
    assert lg.raster(3, 1, 5.0, 2.0, 70.0) == [(0.0, 0.0), (0.0, -7.0), (0.0, -14.0)]
    # 5 Felder in 2 Spalten: 3 + 2
    assert lg.raster(5, 2, 5.0, 2.0, 70.0) == [(0, 0), (0, -7.0), (0, -14.0),
                                               (70.0, 0), (70.0, -7.0)]
    # mehr Spalten als Felder
    assert lg.raster(2, 4, 5.0, 0.0, 50.0) == [(0, 0), (50.0, 0)]
    assert lg.raster(0, 3, 5.0, 2.0, 70.0) == []


def test_beschriftung_und_eingaben():
    eintrag = _eintrag(1, u"AR_Wand_Tragend")
    assert lg.beschriftung(eintrag) == u"AR_Wand_Tragend"
    assert lg.beschriftung(eintrag, u"  ") == u"AR_Wand_Tragend"
    assert lg.beschriftung(eintrag, u" Tragende Wand ") == u"Tragende Wand"
    assert lg.pruefe(u"12,5", 0, 500) == 12.5
    assert lg.pruefe(u"501", 0, 500) is None
    assert lg.pruefe(u"x", 0, 500) is None
    assert lg.pruefe(u"2", 1, 20, ganz=True) == 2
    assert lg.pruefe(u"2.5", 1, 20, ganz=True) is None
    assert lg.hexfarbe((255, 0, 16)) == u"#FF0010"


if __name__ == "__main__":
    fehler = 0
    for name, funktion in sorted(globals().items()):
        if name.startswith("test_") and callable(funktion):
            try:
                funktion()
                print("ok    " + name)
            except AssertionError as ausnahme:
                fehler += 1
                print("FEHLER " + name + ": " + repr(ausnahme))
    print("%d Fehler" % fehler)
    sys.exit(1 if fehler else 0)
