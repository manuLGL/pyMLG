# -*- coding: utf-8 -*-
"""Offline-Test der Logik von LinkedIds (ohne Revit).

    python tools\\test_linked_ids.py
"""
import os
import sys

# Texte werden auf Deutsch verglichen
os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from linked_ids import logik as lg  # noqa: E402


def rohr(element_id, dokument=u"TGA.rvt"):
    return lg.Eintrag(dokument, element_id, kategorie=u"Rohre",
                      typ=u"Rundrohr: Standard", name=u"Rundrohr",
                      unique_id=u"uid-%d" % element_id,
                      verknuepfung=u"TGA.rvt : Position 1",
                      ist_verknuepfung=True)


def wand(element_id):
    return lg.Eintrag(u"Projekt.rvt", element_id, kategorie=u"Wände",
                      typ=u"Basiswand: AW 36.5")


def test_ergaenze_ohne_doppelte():
    eintraege = [rohr(100)]
    # gleiche ID in einem anderen Dokument ist ein anderes Element
    assert lg.ergaenze(eintraege, [rohr(100), rohr(100, u"ARCH.rvt"),
                                   rohr(101)]) == 2
    assert [e.element_id for e in eintraege] == [100, 100, 101]
    assert [e.dokument for e in eintraege] == [u"TGA.rvt", u"ARCH.rvt",
                                               u"TGA.rvt"]


def test_gruppiere_haelt_reihenfolge():
    gruppen = lg.gruppiere([rohr(100), wand(5), rohr(101)])
    assert [name for name, _ in gruppen] == [u"TGA.rvt", u"Projekt.rvt"]
    assert [e.element_id for e in gruppen[0][1]] == [100, 101]


def test_ids_text():
    # ein Dokument: nur die IDs, direkt für "Auswählen nach ID"
    assert lg.ids_text([rohr(100), rohr(101), rohr(100)]) == u"100, 101"
    # mehrere Dokumente: je Dokument ein Block mit Namen
    text = lg.ids_text([rohr(100), wand(5)])
    assert text == u"TGA.rvt:\r\n100\r\n\r\nProjekt.rvt:\r\n5"


def test_tabelle():
    zeilen = lg.tabelle([rohr(100)]).split(u"\r\n")
    assert zeilen[0].split(u"\t")[:2] == [u"Dokument", u"Verknüpfung"]
    spalten = zeilen[1].split(u"\t")
    assert spalten[0] == u"TGA.rvt" and spalten[5] == u"100"
    assert spalten[6] == u"uid-100"


def test_zusammenfassung():
    assert lg.zusammenfassung([rohr(100), rohr(101, u"ARCH.rvt")]) == \
        u"2 Element(e) aus 2 Verknüpfung(en)"
    assert lg.zusammenfassung([rohr(100), wand(5)]) == \
        u"1 Element(e) aus 1 Verknüpfung(en), 1 aus dem aktuellen Modell"
    assert lg.zusammenfassung([wand(5)]) == \
        u"1 Element(e) aus dem aktuellen Modell"


def test_beschreibung_und_herkunft():
    # der Name steckt schon im Typ und wird nicht wiederholt
    assert rohr(100).beschreibung == u"Rohre · Rundrohr: Standard"
    eigener_name = lg.Eintrag(u"TGA.rvt", 9, kategorie=u"Rohre",
                              typ=u"Rundrohr: Standard", name=u"Steigstrang")
    assert eigener_name.beschreibung == \
        u"Rohre · Rundrohr: Standard · Steigstrang"
    assert rohr(100).herkunft == u"TGA.rvt  (TGA.rvt : Position 1)"
    assert wand(5).herkunft == u"Aktuelles Modell"
    # nicht geladene Verknüpfung: nur die ID ist bekannt
    ungeladen = lg.Eintrag(u"ARCH.rvt", 7, ist_verknuepfung=True,
                           geladen=False)
    assert u"nicht geladen" in ungeladen.herkunft
    assert u"nicht geladenen Verknüpfung" in ungeladen.beschreibung


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items())
             if n.startswith("test_") and callable(f)]
    for name, funktion in tests:
        funktion()
        print("  ok:", name)
    print("%d Tests bestanden" % len(tests))
