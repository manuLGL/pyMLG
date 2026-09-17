# -*- coding: utf-8 -*-
"""Offline-Test der Zuordnung Kopie -> Original der Phasen-Werkzeuge (ohne Revit).

    python tools\\test_phasen_zuordnung.py
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from phasen.zuordnung import Eintrag, zuordnen  # noqa: E402

WAND, TUER = ("Wände", 1), ("Türen", 2)


def pruefe(name, bedingung):
    print(("  OK    " if bedingung else "  FEHLER") + "  " + name)
    return bedingung


def verschoben(eintrag, neue_id, v, stoerung=(0.0, 0.0, 0.0)):
    return Eintrag(neue_id, eintrag.schluessel,
                   [(x + v[0] + stoerung[0], y + v[1] + stoerung[1],
                     z + v[2] + stoerung[2]) for x, y, z in eintrag.punkte])


def main():
    ok = True
    v = (10.0, 0.0, 0.0)

    w1 = Eintrag(1, WAND, [(0, 0, 0), (20, 0, 0)])
    w2 = Eintrag(2, WAND, [(0, 5, 0), (20, 5, 0)])
    t1 = Eintrag(3, TUER, [(4, 0, 0)])

    # 1. Rückgabe in anderer Reihenfolge und mit zusätzlicher Tür
    kopien = [verschoben(t1, 13, v), verschoben(w2, 12, v), verschoben(w1, 11, v)]
    e = zuordnen([w1, w2, t1], kopien, v)
    ok &= pruefe("Reihenfolge egal", e == {11: 1, 12: 2, 13: 3})

    # 2. Gleiche Typen dicht nebeneinander werden nicht vertauscht
    a = Eintrag(1, WAND, [(0, 0, 0), (20, 0, 0)])
    b = Eintrag(2, WAND, [(0, 0.3, 0), (20, 0.3, 0)])
    kopien = [verschoben(b, 12, v), verschoben(a, 11, v)]
    e = zuordnen([a, b], kopien, v)
    ok &= pruefe("dichte Nachbarn", e == {11: 1, 12: 2})

    # 3. Kategorie muss passen, auch wenn die Lage stimmt
    e = zuordnen([t1], [Eintrag(99, WAND, [(14, 0, 0)])], v)
    ok &= pruefe("andere Kategorie ignoriert", e == {})

    # 4. Leichte Abweichung (Begrenzungsrahmen nach Verbindung) wird toleriert
    e = zuordnen([w1], [verschoben(w1, 11, v, (0.1, 0, 0))], v)
    ok &= pruefe("kleine Abweichung", e == {11: 1})

    # 5. Zu grosse Abweichung wird nicht zugeordnet
    e = zuordnen([w1], [verschoben(w1, 11, v, (2.0, 0, 0))], v)
    ok &= pruefe("grosse Abweichung", e == {})

    # 6. Vertikal (Einfügen auf anderer Ebene)
    vz = (0.0, 0.0, 9.84)
    e = zuordnen([w1, t1], [verschoben(w1, 11, vz), verschoben(t1, 13, vz)], vz)
    ok &= pruefe("vertikal", e == {11: 1, 13: 3})

    # 7. Ohne Referenzpunkte keine Zuordnung
    e = zuordnen([Eintrag(1, WAND, [])], [Eintrag(11, WAND, [])], v)
    ok &= pruefe("ohne Punkte", e == {})

    print("\nAlle Tests bestanden." if ok else "\nEs gibt Fehler.")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
