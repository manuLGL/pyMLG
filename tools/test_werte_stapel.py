# -*- coding: utf-8 -*-
"""Offline-Test der Logik von StackValues (ohne Revit).

    python tools\test_werte_stapel.py
"""
import os
import sys

os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from werte_stapel import logik as lg  # noqa: E402

TOL = 0.05 / 0.3048     # 5 cm in Fuss


def wand(a, b, m=None):
    if m is None:
        m = ((a[0] + b[0]) / 2.0, (a[1] + b[1]) / 2.0)
    return lg.linien_lage(a, b, m)


def test_lage():
    p = lg.punkt_lage((10.0, 5.0))
    assert lg.gleiche_lage(p, lg.punkt_lage((10.1, 5.0)), TOL)
    assert not lg.gleiche_lage(p, lg.punkt_lage((10.2, 5.0)), TOL)
    w = wand((0.0, 0.0), (20.0, 0.0))
    # umgekehrt gezeichnet: gleich
    assert lg.gleiche_lage(w, wand((20.0, 0.0), (0.0, 0.0)), TOL)
    # kürzer: nicht gleich
    assert not lg.gleiche_lage(w, wand((0.0, 0.0), (18.0, 0.0)), TOL)
    # gleiche Enden, aber Bogen statt Gerade: nicht gleich
    assert not lg.gleiche_lage(w, wand((0.0, 0.0), (20.0, 0.0), (10.0, 4.0)),
                               TOL)
    # Punkt und Linie nie gleich
    assert not lg.gleiche_lage(lg.punkt_lage((10.0, 0.0)), w, TOL)


def test_raster_grenzen():
    """Treffer über Zellgrenzen hinweg, auch mit negativen Koordinaten."""
    lagen = [lg.punkt_lage((x * 0.07, -0.1)) for x in range(-20, 21)]
    r = lg.Raster(lagen, TOL)
    for i, lage in enumerate(lagen):
        erwartet = [j for j, l2 in enumerate(lagen)
                    if lg.gleiche_lage(lage, l2, TOL)]
        assert r.treffer(lage) == erwartet, (i, r.treffer(lage), erwartet)


def _el(kennung, kategorie, lage, ebene=None, **werte):
    return lg.Element(kennung, kategorie, lage, ebene, werte)


def test_zuordnung_nur_gleiche_kategorie():
    quellen = [_el(1, "stuetze", lg.punkt_lage((0, 0)), "EG", nr=u"S1"),
               _el(2, "wand", wand((0, 0), (10, 0)), "EG", nr=u"W1")]
    ziele = [_el(10, "stuetze", lg.punkt_lage((0.05, 0)), "OG1"),
             _el(11, "wand", wand((10, 0), (0, 0)), "OG1"),
             _el(12, "unterzug", wand((0, 0), (10, 0)), "OG1"),
             _el(13, "stuetze", lg.punkt_lage((5, 5)), "OG1")]
    z = lg.zuordnen(quellen, ziele, TOL)
    assert z == {0: [0], 1: [1]}, z
    assert lg.gegenstuecke_je_ebene(z, ziele) == {"OG1": 2}


def test_plan():
    quellen = [_el(1, "s", lg.punkt_lage((0, 0)), "EG", nr=u"S1", tx=u"",
                   zahl=2.5),
               _el(2, "s", lg.punkt_lage((50, 0)), "EG", nr=u"S2"),
               _el(3, "s", lg.punkt_lage((99, 0)), "EG", nr=u"S3")]
    ziele = [
        # leer -> füllen
        _el(10, "s", lg.punkt_lage((0, 0)), "OG1", nr=u"", tx=u"x",
            zahl=None),
        # schon gleich
        _el(11, "s", lg.punkt_lage((0, 0)), "OG2", nr=u"S1", tx=u"",
            zahl=2.5 + 1e-12),
        # anderer Wert -> behalten (ohne Überschreiben)
        _el(12, "s", lg.punkt_lage((50, 0)), "OG1", nr=u"ALT"),
        # Parameter fehlt / schreibgeschützt
        _el(13, "s", lg.punkt_lage((50, 0)), "OG2",
            nr=lg.SCHREIBGESCHUETZT),
        # Ebene nicht gewählt
        _el(14, "s", lg.punkt_lage((0, 0)), "DG", nr=u""),
    ]
    z = lg.zuordnen(quellen, ziele, TOL)
    p = lg.plane(quellen, ziele, z, ["nr", "tx", "zahl"],
                 ebenen={"OG1", "OG2"})
    assert p.auftraege == [(0, "nr", u"S1"), (0, "zahl", 2.5)], p.auftraege
    # tx: Quelle leer bei 10 und 11; zahl/tx fehlen bei S2 als Quelle
    assert p.gleich == 2, p.gleich           # 11: nr, zahl
    assert p.behalten == [(2, "nr", u"ALT", u"S2")], p.behalten
    assert p.fehlt == [(3, "nr", lg.SCHREIBGESCHUETZT)], p.fehlt
    assert p.ohne_gegenstueck == [2], p.ohne_gegenstueck
    assert p.ziele() == [0]

    p = lg.plane(quellen, ziele, z, ["nr"], ebenen={"OG1"},
                 ueberschreiben=True)
    assert p.auftraege == [(0, "nr", u"S1"), (2, "nr", u"S2")], p.auftraege
    assert p.behalten == []

    # Ziel ohne Parameter
    p = lg.plane(quellen, [_el(20, "s", lg.punkt_lage((0, 0)), "OG1")],
                 {0: [0]}, ["nr"])
    assert p.fehlt == [(0, "nr", lg.FEHLT)]


def test_mehrdeutig():
    """Zwei gewählte Quellen an derselben Stelle (z. B. EG und OG1)."""
    quellen = [_el(1, "s", lg.punkt_lage((0, 0)), "EG", nr=u"S1", k=u"A"),
               _el(2, "s", lg.punkt_lage((0, 0)), "OG1", nr=u"S9", k=u"A")]
    ziele = [_el(10, "s", lg.punkt_lage((0, 0)), "OG2", nr=u"", k=u"")]
    p = lg.plane(quellen, ziele, lg.zuordnen(quellen, ziele, TOL),
                 ["nr", "k"])
    assert p.mehrdeutig == [(0, "nr", [u"S1", u"S9"])], p.mehrdeutig
    assert p.auftraege == [(0, "k", u"A")], p.auftraege


def test_gleiche_achse():
    lang = wand((0.0, 0.0), (20.0, 0.0))
    # oben kürzer, liegt ganz auf der Achse: zählt
    assert lg.gleiche_achse(lang, wand((15.0, 0.02), (2.0, 0.02)), TOL)
    # nur ein Stück (unter der Hälfte der kürzeren) gemeinsam: nicht
    assert not lg.gleiche_achse(lang, wand((18.0, 0.0), (30.0, 0.0)), TOL)
    # zur Hälfte gemeinsam: zählt
    assert lg.gleiche_achse(lang, wand((14.0, 0.0), (26.0, 0.0)), TOL)
    # parallel versetzt: nicht
    assert not lg.gleiche_achse(lang, wand((0.0, 1.0), (20.0, 1.0)), TOL)
    # Bogen nur bei gleicher Lage
    bogen = wand((0.0, 0.0), (20.0, 0.0), (10.0, 4.0))
    assert lg.gleiche_achse(bogen, bogen, TOL)
    assert not lg.gleiche_achse(lang, bogen, TOL)
    assert lg.ueberlappung(lang, wand((25.0, 0.0), (30.0, 0.0)), TOL) < 0


def test_zuordnung_achse():
    quellen = [_el(1, "w", wand((0, 0), (20, 0)), "EG", nr=u"W1")]
    ziele = [_el(10, "w", wand((5, 0), (12, 0)), "OG1"),
             _el(11, "w", wand((0, 0), (20, 0)), "OG2"),
             _el(12, "w", wand((5, 3), (12, 3)), "OG1")]
    assert lg.zuordnen(quellen, ziele, TOL) == {1: [0]}
    assert lg.zuordnen(quellen, ziele, TOL, achse=True) == {0: [0], 1: [0]}


def test_warum():
    quellen = [_el(1, "w", wand((0, 0), (20, 0)), "EG"),
               _el(2, "s", lg.punkt_lage((50, 50)), "EG"),
               _el(3, "s", lg.punkt_lage((80, 80)), "EG"),
               _el(4, "w", wand((100, 0), (110, 0)), "EG")]
    ziele = [_el(10, "w", wand((0, 0), (16, 0)), "OG1"),        # kürzer
             _el(11, "s", lg.punkt_lage((50.4, 50)), "OG1"),    # 12 cm
             _el(12, "s", lg.punkt_lage((80, 80)), "DG"),       # abgewählt
             _el(13, "w", wand((0, 0), (20, 0)), "DG")]
    z = lg.zuordnen(quellen, ziele, TOL)
    plan = lg.plane(quellen, ziele, z, [], ebenen={"OG1"})
    assert plan.ohne_gegenstueck == [0, 1, 2, 3]
    w = lg.warum(quellen, ziele, z, {"OG1"}, TOL, plan.ohne_gegenstueck)
    assert w[0][:2] == (lg.NICHT_GEWAEHLT, 3), w[0]
    assert w[1][:2] == (lg.DANEBEN, 1) and abs(w[1][2] - 0.4) < 1e-9, w[1]
    assert w[2][:2] == (lg.NICHT_GEWAEHLT, 2), w[2]
    assert w[3] == (lg.NICHTS, None, None), w[3]
    # ohne die DG-Wand: andere Länge
    z = lg.zuordnen(quellen, ziele[:3], TOL)
    w = lg.warum(quellen, ziele[:3], z, {"OG1"}, TOL, [0])
    assert w[0][:2] == (lg.ANDERE_LAENGE, 0) and abs(w[0][2] - 4) < 1e-9


def test_hilfen():
    assert lg.ist_leer(None) and lg.ist_leer(u"  ") and not lg.ist_leer(0)
    assert lg.gleich(1, 1) and not lg.gleich(u"1", 1)
    assert lg.zahl(u" 5,5 ") == 5.5


if __name__ == "__main__":
    for name, funktion in sorted(globals().items()):
        if name.startswith("test_") and callable(funktion):
            funktion()
            print("ok  ", name)
