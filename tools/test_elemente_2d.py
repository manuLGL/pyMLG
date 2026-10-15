# -*- coding: utf-8 -*-
"""Offline-Test der Eingrenzung und Gruppierung von Elements2D (ohne Revit).

    python tools\\test_elemente_2d.py
"""
import os
import re
import sys

# Texte werden auf Deutsch verglichen
os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from elemente_2d import logik as lg  # noqa: E402
from filter_more import baum as bm  # noqa: E402

E = lg.Eintrag
DATEN = [
    E(10, u"Linien", u"Linienstil", u"Dünn", None, u"EG", u"Grundrisse",
      platziert=True),
    E(11, u"Linien", u"Linienstil", u"Dünn", None, u"Detail 1",
      u"Zeichenansichten"),
    E(12, u"Detailgruppen", u"Detailgruppe", u"Treppe", u"Treppe",
      u"Detail 1", u"Zeichenansichten"),
    E(13, u"Textnotizen", u"Text", u"2.5mm", u"Achtung", u"A01 - Grundriss",
      u"Pläne", platziert=True, auf_plan=True),
]


def test_eingrenzen():
    ids = lambda liste: [e.element_id for e in liste]  # noqa: E731
    assert ids(lg.grenze_ein(DATEN)) == [10, 11, 12, 13]
    assert ids(lg.grenze_ein(DATEN, mit_plaenen=False)) == [10, 11, 12]
    # Elemente auf Plänen fallen bei "nur nicht platzierte" immer weg
    assert ids(lg.grenze_ein(DATEN, lg.UNPLATZIERTE)) == [11, 12]
    assert ids(lg.grenze_ein(DATEN, lg.UNPLATZIERTE,
                             mit_plaenen=False)) == [11, 12]
    # platzierte Ansichten; Elemente direkt auf Plänen sind "platziert"
    assert ids(lg.grenze_ein(DATEN, lg.PLATZIERTE)) == [10, 13]
    assert ids(lg.grenze_ein(DATEN, lg.PLATZIERTE,
                             mit_plaenen=False)) == [10]


def test_bestimmte_plaene():
    ids = lambda liste: [e.element_id for e in liste]  # noqa: E731
    daten = [
        E(1, u"Linien", u"L", u"L", None, u"EG", u"Grundrisse",
          plaene=[100, 101]),
        E(2, u"Linien", u"L", u"L", None, u"OG", u"Grundrisse",
          plaene=[102]),
        E(3, u"Texte", u"T", u"T", None, u"A100", u"Pläne", auf_plan=True,
          plaene=[100]),
        E(4, u"Linien", u"L", u"L", None, u"Detail", u"Zeichenansichten"),
    ]
    assert daten[0].platziert and not daten[3].platziert
    assert ids(lg.grenze_ein(daten, lg.PLATZIERTE, plaene={100})) == [1, 3]
    assert ids(lg.grenze_ein(daten, lg.PLATZIERTE, plaene={100},
                             mit_plaenen=False)) == [1]
    assert ids(lg.grenze_ein(daten, lg.PLATZIERTE, plaene={102, 999})) == [2]
    assert ids(lg.grenze_ein(daten, lg.PLATZIERTE, plaene=set())) == []
    # Planauswahl gilt nur bei "auf einem Plan"
    assert ids(lg.grenze_ein(daten, lg.ALLE, plaene={102})) == [1, 2, 3, 4]


def test_nach_kategorie():
    w = bm.baue(lg.datensaetze(DATEN, lg.NACH_KATEGORIE))
    assert [k.name for k in w.kinder] == [u"Detailgruppen", u"Linien",
                                          u"Textnotizen"]
    linien = w.kinder[1]
    duenn = linien.kinder[0].kinder[0]
    assert duenn.name == u"Dünn" and len(duenn.ids) == 2
    # ohne Name: Typname, dahinter Id und Ansicht
    assert duenn.kinder[0].name == u"Dünn  [10]  ·  EG"


def test_nach_ansicht():
    w = bm.baue(lg.datensaetze(DATEN, lg.NACH_ANSICHT))
    assert [k.name for k in w.kinder] == [u"Grundrisse", u"Pläne",
                                          u"Zeichenansichten"]
    detail = w.kinder[2].kinder[0]
    assert detail.name == u"Detail 1" and sorted(detail.ids) == [11, 12]
    assert [k.name for k in detail.kinder] == [u"Detailgruppen", u"Linien"]
    assert detail.kinder[0].kinder[0].name == u"Treppe  [12]"


def test_anzahl_ansichten_und_kurz():
    assert lg.anzahl_ansichten(DATEN) == 3
    assert lg.kurz(u"Zeile 1\n  Zeile 2") == u"Zeile 1 Zeile 2"
    assert lg.kurz(u"x" * 80, 10) == u"x" * 9 + u"…"
    assert lg.kurz(None) == u""


def test_xaml_texte_vollstaendig():
    """Jeder Platzhalter im XAML hat einen Text (ohne WPF geprüft)."""
    pfad = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..",
                        "pyMLG.extension", "lib", "elemente_2d",
                        "fenster.py")
    with open(pfad, encoding="utf-8") as datei:
        quelle = datei.read()
    xaml = quelle.split('XAML = u"""', 1)[1].split('"""', 1)[0]
    benutzt = set(re.findall(r"\{\{(\w+)\}\}", xaml))
    definiert = set(re.findall(r'^    "(t\d+)":', quelle, re.M))
    assert benutzt == definiert, (benutzt ^ definiert)


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items())
             if n.startswith("test_") and callable(f)]
    for name, funktion in tests:
        funktion()
        print("  ok:", name)
    print("%d Tests bestanden" % len(tests))
