# -*- coding: utf-8 -*-
"""Offline-Test der Logik von LinkKoerper (ohne Revit).

    python tools\\test_link_koerper.py
"""
import os
import sys

# Texte werden auf Deutsch verglichen
os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from link_koerper import logik as lg  # noqa: E402

WUERFEL = [(x, y, z) for x in (-0.5, 0.5) for y in (-0.5, 0.5)
           for z in (-0.5, 0.5)]


def test_familienname_und_schluessel():
    name = lg.familienname([123456], 98765, u"TGA: Haus|A.rvt", u"Rohre")
    assert name == u"pyMLG Link TGA_ Haus_A.rvt Rohre (123456-98765)"
    assert lg.schluessel_aus_name(name) == (u"123456", 98765)
    # verschachtelte Verknüpfung: Kette der IDs
    name = lg.familienname([1, 22], 5, u"Innen.rvt", u"Wände")
    assert lg.schluessel_aus_name(name) == lg.schluessel([1, 22], 5)


def test_fremde_familien_werden_ignoriert():
    assert lg.schluessel_aus_name(u"Tisch (1-2)") is None
    # Klammern im Linknamen stören die Kennung am Ende nicht
    name = lg.familienname([7], 8, u"TGA (alt).rvt", u"Rohre")
    assert u"TGA _alt_.rvt" in name
    assert lg.schluessel_aus_name(name) == (u"7", 8)
    assert lg.schluessel_aus_name(u"pyMLG Link ohne Kennung") is None
    assert lg.schluessel_aus_name(None) is None


def test_typname():
    assert lg.form_aus_typname(lg.typname(u"abc")) == u"abc"
    assert lg.form_aus_typname(u"Standard") is None


def test_mitte():
    punkte = [(x + 10.0, y, z - 2.0) for x, y, z in WUERFEL]
    assert lg.mitte(punkte) == (10.0, 0.0, -2.0)
    assert lg.mitte([]) is None


def test_formkennung_stabil_gegen_rechenrauschen():
    rauschen = [(x + 1e-7, y - 1e-7, z) for x, y, z in WUERFEL]
    assert lg.formkennung(WUERFEL, 1.0) == lg.formkennung(rauschen, 1.0)


def test_formkennung_erkennt_andere_form():
    laenger = [(x * 2.0, y, z) for x, y, z in WUERFEL]
    assert lg.formkennung(WUERFEL, 1.0) != lg.formkennung(laenger, 2.0)
    # gedrehter Quader: gleiche Masse, andere Kennung
    quader = [(x * 2.0, y, z) for x, y, z in WUERFEL]
    gedreht = [(y, x, z) for x, y, z in quader]
    assert lg.formkennung(quader, 2.0) != lg.formkennung(gedreht, 2.0)


def test_entscheide():
    mitte = (1.0, 2.0, 3.0)
    assert lg.entscheide(None, None, u"f", mitte) == lg.NEU
    assert lg.entscheide(u"f", mitte, u"f", mitte) == lg.GLEICH
    # unter 1 mm gilt als gleich
    assert lg.entscheide(u"f", (1.0005, 2.0, 3.0), u"f", mitte) == lg.GLEICH
    assert lg.entscheide(u"f", (1.5, 2.0, 3.0), u"f", mitte) == \
        lg.VERSCHIEBEN
    assert lg.entscheide(u"g", mitte, u"f", mitte) == lg.ERSETZEN
    # Familie ohne Instanz: neu aufbauen
    assert lg.entscheide(None, None, u"f", mitte) == lg.NEU
    assert lg.entscheide(u"f", None, u"f", mitte) == lg.ERSETZEN


EBENEN = [(u"UG", -3.0), (u"EG", 0.0), (u"OG 1", 3.2), (u"OG 2", 6.4)]


def test_ebene_nach_name():
    # Name gewinnt, auch wenn die Höhe im Link etwas abweicht
    assert lg.waehle_ebene(EBENEN, u" og 1 ", 3.25, 4.0) == 2


def test_ebene_nach_hoehe():
    # anderer Name, gleiche Höhe (Toleranz 5 cm)
    assert lg.waehle_ebene(EBENEN, u"1.OG", 3.18, 4.0) == 2


def test_ebene_rueckfall():
    # keine passende Höhe: höchste Ebene unter der Quellebene
    assert lg.waehle_ebene(EBENEN, u"Zwischen", 4.5, 5.0) == 2
    # ohne Quellebene: unter dem Körper
    assert lg.waehle_ebene(EBENEN, None, None, 7.0) == 3
    # alles unter der untersten Ebene: die unterste
    assert lg.waehle_ebene(EBENEN, None, None, -9.0) == 0
    assert lg.waehle_ebene([], u"EG", 0.0, 0.0) is None


def test_ergebnis_text():
    ergebnis = lg.Ergebnis()
    ergebnis.zaehle(lg.NEU)
    ergebnis.zaehle(lg.VERSCHIEBEN)
    ergebnis.ohne_geometrie = 2
    text = ergebnis.text()
    assert u"Neu erstellt: 1" in text
    assert u"Verschoben (Lage geändert): 1" in text
    assert u"Ohne Volumengeometrie übersprungen: 2" in text
    assert u"Fehler" not in text


if __name__ == "__main__":
    tests = [(name, funktion) for name, funktion in sorted(globals().items())
             if name.startswith("test_")]
    for name, funktion in tests:
        funktion()
        print("ok  ", name)
    print("%d Tests bestanden" % len(tests))
