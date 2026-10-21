# -*- coding: utf-8 -*-
"""Offline-Test der Regeln von AutoRooms (ohne Revit).

    python tools\test_raum_automatik.py
"""
import os
import sys

os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from raum_automatik import logik as lg  # noqa: E402


def test_merkmale():
    assert lg.merkmale(u"Bett Einzelbett 90x200", breite=0.9) == {lg.BETT}
    assert lg.merkmale(u"Bett 180x200", breite=1.8) == {lg.BETT,
                                                         lg.DOPPELBETT}
    assert lg.DOPPELBETT in lg.merkmale(u"M_Doppelbett")
    assert lg.BETT in lg.merkmale(u"M_Bed-Standard Queen")
    assert lg.BETT not in lg.merkmale(u"Bedienfeld")
    assert lg.merkmale(u"WC wandhängend", True) == {lg.SANITAER}
    assert lg.merkmale(u"Duschtasse 90", True) == {lg.SANITAER, lg.DUSCHE}
    assert lg.merkmale(u"Badewanne") == {lg.SANITAER, lg.DUSCHE}
    assert lg.merkmale(u"M_Bathtub Rectangular") == {lg.SANITAER, lg.DUSCHE}
    # Küchenspüle in der Kategorie Sanitär ist Küche, kein Sanitärobjekt
    assert lg.merkmale(u"Küchenspüle", True) == {lg.KUECHE}
    assert lg.merkmale(u"Waschmaschine", True) == set()
    # Dusche mit Ablauf bleibt Dusche
    assert lg.merkmale(u"Dusche bodengleich mit Ablauf", True) == {
        lg.SANITAER, lg.DUSCHE}
    assert lg.merkmale(u"Brausetasse 90/90") == {lg.SANITAER, lg.DUSCHE}
    assert lg.merkmale(u"Sofá 3 plazas") == {lg.SOFA}
    assert lg.merkmale(u"Tisch rund") == set()
    assert lg.merkmale(u"Einbauschrank 60") == {lg.SCHRANK}
    assert lg.merkmale(u"Garderobe") == {lg.SCHRANK}
    # Kühlschrank ist Küche, kein Schrank
    assert lg.merkmale(u"Kühlschrank") == {lg.KUECHE}


def test_art():
    assert lg.art({lg.KUECHE, lg.SOFA}, 1) == lg.WOHNKUECHE
    # Spüle der Küche macht die Wohnküche nicht zum WC
    assert lg.art({lg.KUECHE, lg.SOFA, lg.SANITAER}, 1) == lg.WOHNKUECHE
    assert lg.art({lg.KUECHE}, 3) is None
    # Einraumwohnung: Bett und Küche zusammen
    assert lg.art({lg.KUECHE, lg.BETT}, 1) == lg.WOHNKUECHE
    assert lg.art({lg.KUECHE, lg.BETT, lg.SOFA, lg.SCHRANK}, 3) ==         lg.WOHNKUECHE
    assert lg.art({lg.SOFA}, 1) is None
    assert lg.art({lg.SANITAER, lg.DUSCHE}, 1) == lg.BAD
    assert lg.art({lg.SANITAER}, 1) == lg.WC
    assert lg.art({lg.BETT}, 1) == lg.ZIMMER
    assert lg.art({lg.SCHRANK}, 3) == lg.GANG
    # ohne Schrank kein Gang
    assert lg.art(set(), 3) is None
    assert lg.art({lg.SOFA, lg.SCHRANK}, 4) is None
    # Abstellraum: nur Schrank, keine Tür
    assert lg.art({lg.SCHRANK}, 0) == lg.AR
    assert lg.art({lg.SCHRANK}, 1) == lg.AR
    assert lg.art({lg.SCHRANK}, 2) == lg.AR
    assert lg.art({lg.SCHRANK, lg.SOFA}, 0) is None


def _bereich(kennung, merkmale=(), tueren=1, flaeche=10.0):
    b = lg.Bereich(kennung, True, flaeche)
    b.merkmale = set(merkmale)
    b.tueren = tueren
    return b


def test_beispielwohnung():
    """Grundriss aus dem Beispiel: Wohnküche, Bad, WC, Gang, 2 Zimmer, AR,
    dazu Treppenhaus und eine zweite Wohnung."""
    wk = _bereich("wk", [lg.KUECHE, lg.SOFA], 1, 31.5)
    bad = _bereich("bad", [lg.SANITAER, lg.DUSCHE], 1, 5.6)
    wc = _bereich("wc", [lg.SANITAER], 1, 2.3)
    gang = _bereich("gang", [lg.SCHRANK], 6, 11.5)
    z1 = _bereich("z1", [lg.BETT, lg.DOPPELBETT], 1, 13.0)
    z2 = _bereich("z2", [lg.BETT], 1, 11.0)
    ar = _bereich("ar", [], 1, 2.9)
    th = lg.Bereich("th", False, 20.0, u"TREPPENHAUS")
    gang_b = _bereich("gang_b", [lg.SCHRANK], 4, 8.0)
    za = _bereich("za", [lg.BETT], 1, 14.0)
    zb = _bereich("zb", [lg.BETT], 1, 10.0)
    zc = lg.Bereich("zc", False, 12.0, u"Zimmer 1")
    alle = [wk, bad, wc, gang, z1, z2, ar, th, gang_b, za, zb, zc]
    verbindungen = [("gang", r) for r in ("wk", "bad", "wc", "z1", "z2",
                                         "ar")]
    verbindungen += [("gang", "th"), ("th", "gang_b"), ("gang_b", "za"),
                     ("gang_b", "zb"), ("gang_b", "zc")]
    lg.bewerte(alle)
    lg.nummeriere(alle, verbindungen)
    assert [b.raumname() for b in (wk, bad, wc, gang, z1, z2)] == [
        u"WOHNKÜCHE", u"BAD", u"WC", u"VR", u"ZIMMER 1", u"ZIMMER 2"]
    assert ar.art is None
    assert wk.kategorie() == u"1_W_WOHNEN"
    assert bad.kategorie() == wc.kategorie() == u"2_W_SANITÄR"
    assert gang.kategorie() == u"2_W_VERKEHRSFLÄCHE"
    assert (bad.boden(), wc.boden()) == (u"Fliesen", u"Fliesen")
    assert wk.boden() == gang.boden() == z1.boden() == u"Parkett"
    ar2 = _bereich("ar2", [lg.SCHRANK], 0)
    lg.bewerte([ar2])
    assert (ar2.raumname(), ar2.kategorie(), ar2.boden()) == (
        u"AR", u"2_W_NEBENRAUM", u"Parkett")
    # zweite Wohnung: ZIMMER 1 gibt es schon, grösseres Zimmer zuerst
    assert (za.nummer, zb.nummer) == (2, 3)
    # feste Raumnummern je Wohnung (VR = Vorraum)
    lg.raumnummern(alle, verbindungen)
    assert [b.raumnummer for b in (wk, bad, wc, gang, z1, z2)] == [
        u"005", u"003", u"005", u"001", u"006", u"007"]
    assert ar.raumnummer is None and th.raumnummer is None
    assert (gang_b.raumnummer, za.raumnummer, zb.raumnummer) == (
        u"001", u"007", u"008")


def test_ohne_doppelbett_und_mehrere():
    a = _bereich("a", [lg.BETT], 1, 9.0)
    b = _bereich("b", [lg.BETT], 1, 12.0)
    c = _bereich("c", [lg.BETT, lg.DOPPELBETT], 1, 8.0)
    flur = _bereich("f", [lg.SCHRANK], 3)
    alle = [a, b, c, flur]
    lg.bewerte(alle)
    lg.nummeriere(alle, [("f", "a"), ("f", "b"), ("f", "c")])
    assert (c.nummer, b.nummer, a.nummer) == (1, 2, 3)


def test_flur_ohne_schrank_trennt_wohnungen():
    """Treppenhaus und Wohnungsflure ohne Schrank sind keine Räume, trennen
    aber trotzdem die Wohnungen."""
    th = _bereich("th", [], 4, 100.0)
    f1 = _bereich("f1", [], 5)
    f2 = _bereich("f2", [], 5)
    z1 = _bereich("z1", [lg.BETT, lg.DOPPELBETT], 1)
    z2 = _bereich("z2", [lg.BETT, lg.DOPPELBETT], 1)
    alle = [th, f1, f2, z1, z2]
    lg.bewerte(alle)
    lg.nummeriere(alle, [("th", "f1"), ("th", "f2"), ("f1", "z1"),
                         ("f2", "z2")])
    assert (th.art, f1.art) == (None, None)
    assert (z1.nummer, z2.nummer) == (1, 1)


def test_schluessel_name():
    assert lg.schluessel_name(u"1_W_Wohnen") == lg.schluessel_name(
        u"1_W_WOHNEN")
    assert lg.schluessel_name(u"2_W_SANITÄR ") == lg.schluessel_name(
        u"2_w_sanitär")


if __name__ == "__main__":
    for name, funktion in sorted(globals().items()):
        if name.startswith("test_") and callable(funktion):
            funktion()
            print("ok", name)
