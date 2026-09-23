# -*- coding: utf-8 -*-
"""Offline-Test der Logik des Ansichtsvorlagen-Managers (ohne Revit).

    python tools\\test_vorlagen_manager.py
"""
import os
import sys

# Texte werden auf Deutsch verglichen
os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from vorlagen_manager import logik as lg  # noqa: E402


def test_passt():
    assert lg.passt(u"360 acotado grundriss", [])
    assert lg.passt(u"360 acotado grundriss", [u"aco", u"grund"])
    assert not lg.passt(u"360 acotado grundriss", [u"aco", u"schnitt"])


def test_natuerlich():
    namen = [u"Plan 10", u"plan 2", u"Plan 1a", u""]
    assert sorted(namen, key=lg.natuerlich) == [u"", u"Plan 1a", u"plan 2",
                                                u"Plan 10"]


def test_vereine():
    assert lg.vereine([]) is None
    assert lg.vereine([3]) == 3
    assert lg.vereine([3, 3, 3]) == 3
    assert lg.vereine([3, 4]) is lg.VERSCHIEDEN
    # Auch None ist ein Wert und darf nicht mit "leer" verwechselt werden
    assert lg.vereine([None, None]) is None
    assert lg.vereine([None, 1]) is lg.VERSCHIEDEN


def test_vereine_flagge():
    assert lg.vereine_flagge([True, True]) is True
    assert lg.vereine_flagge([False, False]) is False
    assert lg.vereine_flagge([True, False]) is None
    assert lg.vereine_flagge([]) is None


def test_reihenfolge():
    # Bekannte eingebaute Parameter zuerst und in der Dialogreihenfolge,
    # danach die uebrigen eingebauten, zuletzt die Projektparameter
    zeilen = [
        (u"360_OrganizacionPlanos", lg.gruppe_und_rang(None, False)),
        (u"Detaillierungsgrad", lg.gruppe_und_rang("VIEW_DETAIL_LEVEL", True)),
        (u"Irgendwas", lg.gruppe_und_rang("VIEW_CAMERA_POSITION", True)),
        (u"Ansichtsmassstab", lg.gruppe_und_rang("VIEW_SCALE", True)),
        (u"Filter (V/G)", lg.gruppe_und_rang("VIS_GRAPHICS_FILTERS", True)),
    ]
    sortiert = [name for name, _ in sorted(
        zeilen, key=lambda z: (z[1][0], z[1][1], lg.natuerlich(z[0])))]
    assert sortiert == [u"Ansichtsmassstab", u"Detaillierungsgrad",
                        u"Filter (V/G)", u"Irgendwas",
                        u"360_OrganizacionPlanos"]


def test_pruefe_name():
    vorhandene = [u"360 BASE", u"360 ACOTADO"]
    assert lg.pruefe_name(vorhandene, u"  Neu  ") == u"Neu"
    # Der eigene alte Name bleibt erlaubt, auch mit anderer Schreibweise
    assert lg.pruefe_name(vorhandene, u"360 base", u"360 BASE") == u"360 base"
    for schlecht in (u"", u"   ", u"Mit|Strich", u"A{B}", u"C:D"):
        try:
            lg.pruefe_name(vorhandene, schlecht)
        except ValueError:
            continue
        raise AssertionError(u"haette scheitern muessen: %r" % schlecht)
    try:
        lg.pruefe_name(vorhandene, u"360 base")
    except ValueError:
        pass
    else:
        raise AssertionError(u"doppelter Name haette scheitern muessen")


def test_freier_name():
    assert lg.freier_name([], u"Vorlage") == u"Vorlage"
    assert lg.freier_name([u"Vorlage"], u"Vorlage") == u"Vorlage 2"
    assert lg.freier_name([u"vorlage", u"Vorlage 2"], u"Vorlage") == \
        u"Vorlage 3"


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items())
             if n.startswith("test_") and callable(f)]
    for name, funktion in tests:
        funktion()
        print("  ok:", name)
    print("%d Tests bestanden" % len(tests))
