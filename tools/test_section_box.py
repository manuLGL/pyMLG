# -*- coding: utf-8 -*-
"""Offline-Test des Schnittbox-Textformats (ohne Revit).

    python tools\test_section_box.py
"""
import os
import sys

# Texte werden auf Deutsch verglichen
os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from section_box import logik as lg  # noqa: E402

GEMEINSAM = ((12.345678, 6.789012, 1.5),
             (0.8, 0.6, 0.0), (-0.6, 0.8, 0.0), (0.0, 0.0, 1.0))
INTERN = ((2.0, 3.0, 1.5),
          (1.0, 0.0, 0.0), (0.0, 1.0, 0.0), (0.0, 0.0, 1.0))


def beispiel():
    return lg.Box(GEMEINSAM, INTERN, (-5.0, -3.0, 0.0), (5.0, 3.0, 4.0),
                  quelle=u"Koordination.rvt | {3D} | 2026-09-23 14:12")


def test_hin_und_zurueck():
    box = lg.lies_text(lg.erzeuge_text(beispiel()))
    for gelesen, erwartet in zip(box.gemeinsam, GEMEINSAM):
        assert max(abs(a - b) for a, b in zip(gelesen, erwartet)) < 1e-6
    for gelesen, erwartet in zip(box.intern, INTERN):
        assert max(abs(a - b) for a, b in zip(gelesen, erwartet)) < 1e-6
    assert box.min == (-5.0, -3.0, 0.0)
    assert box.max == (5.0, 3.0, 4.0)
    assert box.quelle == u"Koordination.rvt | {3D} | 2026-09-23 14:12"


def test_masse():
    assert beispiel().masse == (10.0, 6.0, 4.0)
    assert beispiel().masse_text() == u"10.00 \u00d7 6.00 \u00d7 4.00 m"


def test_kopf_und_kommentare():
    zeilen = lg.erzeuge_text(beispiel()).split(u"\r\n")
    assert zeilen[0] == lg.KOPF
    assert zeilen[1].startswith(u"# Koordination.rvt")
    assert u"SHARED-O" in zeilen[3]


def test_aus_mail_zitiert():
    """Text aus einer Mail: Zitatzeichen, Rand und Fliesstext davor."""
    text = lg.erzeuge_text(beispiel())
    zitat = u"\n".join([u"Hallo Manuel,", u"schau dir das bitte an:", u""]
                       + [u">  " + z for z in text.split(u"\r\n")]
                       + [u"", u"Gruss"])
    box = lg.lies_text(zitat)
    assert box.min == (-5.0, -3.0, 0.0)
    assert box.quelle == u"Koordination.rvt | {3D} | 2026-09-23 14:12"


def test_ohne_intern_gilt_gemeinsam():
    zeilen = [z for z in lg.erzeuge_text(beispiel()).split(u"\r\n")
              if not z.startswith(u"INTERN")]
    box = lg.lies_text(u"\r\n".join(zeilen))
    assert box.intern == box.gemeinsam


def fehler(text):
    try:
        lg.lies_text(text)
    except lg.Fehler as ausnahme:
        return u"%s" % ausnahme
    raise AssertionError(u"Fehler erwartet")


def test_fremder_text():
    assert u"keine pyMLG-Schnittbox" in fehler(u"Guten Morgen")
    assert u"keine pyMLG-Schnittbox" in fehler(u"")
    assert u"keine pyMLG-Schnittbox" in fehler(None)
    # Kopf ohne Werte reicht nicht
    assert u"keine pyMLG-Schnittbox" in fehler(lg.KOPF)


def test_unvollstaendig():
    zeilen = [z for z in lg.erzeuge_text(beispiel()).split(u"\r\n")
              if not z.startswith(u"MAX")]
    assert u"keine pyMLG-Schnittbox" in fehler(u"\r\n".join(zeilen))


def test_leere_box():
    box = lg.Box(GEMEINSAM, INTERN, (0.0, 0.0, 0.0), (0.0, 0.0, 0.0))
    assert u"leer" in fehler(lg.erzeuge_text(box))


def test_kaputte_achsen():
    schief = (GEMEINSAM[0], (2.0, 0.0, 0.0), GEMEINSAM[2], GEMEINSAM[3])
    box = lg.Box(schief, INTERN, (-1.0, -1.0, 0.0), (1.0, 1.0, 1.0))
    assert u"Achsen" in fehler(lg.erzeuge_text(box))


def test_muell_in_zahlen():
    text = lg.erzeuge_text(beispiel()).replace(u"MIN       -5.000000",
                                               u"MIN       x")
    assert u"keine pyMLG-Schnittbox" in fehler(text)


def test_huelle():
    punkte = [(1.0, 2.0, 0.0), (3.0, -1.0, 4.0), (2.0, 0.0, 1.0)]
    assert lg.huelle(punkte) == ((1.0, -1.0, 0.0), (3.0, 2.0, 4.0))
    assert lg.huelle(punkte, rand=0.5) == ((0.5, -1.5, -0.5), (3.5, 2.5, 4.5))
    assert lg.huelle([]) is None


def test_huelle_flaches_element():
    """Eine ebene Wandflaeche hat in einer Achse keine Dicke."""
    punkte = [(0.0, 5.0, 0.0), (4.0, 5.0, 3.0)]
    klein, gross = lg.huelle(punkte, rand=0.0, mindestmass=1.0)
    assert klein == (0.0, 4.5, 0.0)
    assert gross == (4.0, 5.5, 3.0)
    # ausreichend grosse Kanten bleiben, wie sie sind
    assert (gross[0] - klein[0], gross[2] - klein[2]) == (4.0, 3.0)


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items())
             if n.startswith("test_") and callable(f)]
    for name, funktion in tests:
        funktion()
        print("  ok:", name)
    print("%d Tests bestanden" % len(tests))
