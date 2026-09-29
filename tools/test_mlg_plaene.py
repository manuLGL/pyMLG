# -*- coding: utf-8 -*-
"""Offline-Test der Plan-Werkzeuge (Reiter MLGplans) ohne Revit.

    python tools\\test_mlg_plaene.py
"""
import os
import sys

# Texte werden auf Deutsch verglichen
os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from mlg_plaene import logik as lg  # noqa: E402


def pruefe(name, bedingung):
    print(("  OK    " if bedingung else "  FEHLER") + "  " + name)
    return bedingung


def main():
    ok = True
    bereich = (0.0, 0.0, 10.0, 6.0)

    # --- Platz auf dem Plan
    r = lg.freie_position(bereich, [], 4, 2)
    ok &= pruefe("leerer Plan: oben links", r == (0, 4, 4, 6))

    belegt = [r]
    r2 = lg.freie_position(bereich, belegt, 4, 2, abstand=0.5)
    ok &= pruefe("zweite Ansicht rechts daneben", r2 == (4.5, 4, 8.5, 6))

    belegt.append(r2)
    r3 = lg.freie_position(bereich, belegt, 4, 2, abstand=0.5)
    ok &= pruefe("dritte Ansicht in die nächste Zeile", r3 == (0, 1.5, 4, 3.5))

    ok &= pruefe("zu gross -> None", lg.freie_position(bereich, [], 11, 2) is None)
    voll = [(0, 0, 10, 6)]
    ok &= pruefe("voller Plan -> None", lg.freie_position(bereich, voll, 1, 1) is None)

    # vorhandene Ansicht in der Mitte wird umgangen
    mitte = [(3, 2, 7, 4)]
    r = lg.freie_position(bereich, mitte, 3, 2)
    ok &= pruefe("vorhandene Ansicht umgangen",
                 r is not None and not lg.ueberlappt(r, mitte[0]))

    ok &= pruefe("Berührung ist keine Überlappung",
                 not lg.ueberlappt((0, 0, 1, 1), (1, 0, 2, 1)))
    ok &= pruefe("Abstand wird geprüft",
                 lg.ueberlappt((0, 0, 1, 1), (1.2, 0, 2, 1), abstand=0.5))

    ok &= pruefe("zentriert", lg.zentriert((0, 0, 2, 2), bereich) == (4, 2, 6, 4))
    ok &= pruefe("Zeichenfläche mit Rand",
                 lg.zeichenflaeche(bereich, 1, 2) == (1, 1, 7, 5))
    ok &= pruefe("Zeichenfläche: zu viel Rand -> ganzer Plankopf",
                 lg.zeichenflaeche(bereich, 4, 4) == bereich)

    # --- Überschneidungen erlaubt
    ok &= pruefe("beste Position = freie, wenn Platz",
                 lg.beste_position(bereich, [], 4, 2) == (0, 4, 4, 6))
    links = [(0, 0, 6, 6)]
    r = lg.beste_position(bereich, links, 6, 6)
    ok &= pruefe("kein Platz: im Bereich, kleinste Überschneidung",
                 lg.liegt_in(r, bereich) and r == (4, 0, 10, 6))
    r = lg.beste_position(bereich, [], 12, 8)
    ok &= pruefe("zu gross: oben links bündig", r == (0, -2, 12, 6))
    ok &= pruefe("Überschneidungsfläche", lg.ueberlappung((0, 0, 2, 2), (1, 1, 3, 3)) == 1)

    # --- Vorschau: ausrichten, verteilen, einrasten, anordnen
    a, b = (0, 0, 2, 2), (5, 1, 6, 4)
    ok &= pruefe("links bündig", lg.ausrichten([a, b], lg.LINKS) == [a, (0, 1, 1, 4)])
    ok &= pruefe("oben bündig", lg.ausrichten([a, b], lg.OBEN) == [(0, 2, 2, 4), b])
    ok &= pruefe("waagrechte Mitte", lg.ausrichten([a, b], lg.MITTE_H)
                 == [(2, 0, 4, 2), (2.5, 1, 3.5, 4)])
    ok &= pruefe("einzeln an Zeichenfläche", lg.ausrichten([a], lg.RECHTS, bereich)
                 == [(8, 0, 10, 2)])
    drei = [(0, 0, 2, 1), (3, 0, 4, 1), (9, 0, 10, 1)]
    ok &= pruefe("waagrecht verteilen",
                 lg.verteilen(drei) == [(0, 0, 2, 1), (5, 0, 6, 1), (9, 0, 10, 1)])
    senkrecht = [(0, 9, 1, 10), (0, 8, 1, 8.5), (0, 0, 1, 1)]
    ok &= pruefe("senkrecht verteilen",
                 lg.verteilen(senkrecht, waagrecht=False)
                 == [(0, 9, 1, 10), (0, 4.75, 1, 5.25), (0, 0, 1, 1)])
    ok &= pruefe("verteilen unter drei unverändert", lg.verteilen([a, b]) == [a, b])
    dx, dy = lg.einrasten((2.1, 3.0, 4.1, 4.0), [(0, 0, 2, 1)], 0.2)
    ok &= pruefe("einrasten an Kante", abs(dx + 0.1) < 1e-9 and dy == 0.0)
    ok &= pruefe("nicht einrasten ausserhalb Toleranz",
                 lg.einrasten((3, 5, 4, 6), [(0, 0, 2, 1)], 0.2) == (0.0, 0.0))
    angeordnet = lg.anordnen(bereich, [], [(0, 0, 4, 2), (0, 0, 4, 2)], 0.5)
    ok &= pruefe("automatisch anordnen",
                 angeordnet == [(0, 4, 4, 6), (4.5, 4, 8.5, 6)])
    ok &= pruefe("Überschneidungen finden",
                 lg.ueberschneidungen([(0, 0, 2, 2), (1, 1, 3, 3), (5, 5, 6, 6)]) == {0, 1})

    # --- Nummern
    ok &= pruefe("nummer_plus", lg.nummer_plus(u"A-101") == u"A-102")
    ok &= pruefe("nummer_plus mit Nullen", lg.nummer_plus(u"01.009") == u"01.010")
    ok &= pruefe("nummer_plus letzte Zahl", lg.nummer_plus(u"2-A-09b") == u"2-A-10b")
    ok &= pruefe("nummer_plus ohne Zahl", lg.nummer_plus(u"B") == u"B-1")
    ok &= pruefe("nummer_plus Schritt 10", lg.nummer_plus(u"A-100", 10) == u"A-110")
    try:
        lg.nummer_plus(u"A-001", -5)
        ok &= pruefe("negative Nummer abgelehnt", False)
    except ValueError:
        ok &= pruefe("negative Nummer abgelehnt", True)

    vergeben = lg.nummern_menge([u"A-101", u"a-102"])
    ok &= pruefe("nächste freie Nummer (Gross/Klein egal)",
                 lg.naechste_freie_nummer(u"A-101", vergeben) == u"A-103")
    ok &= pruefe("... und vergeben", u"a-103" in vergeben)

    ok &= pruefe("natürliche Sortierung",
                 sorted([u"A-10", u"A-2", u"A-1"], key=lg.natuerlich)
                 == [u"A-1", u"A-2", u"A-10"])

    werte = {u"Ansicht": u"EG Grundriss", u"Ebene": u"EG"}
    text, fehlend = lg.muster_anwenden(u"AP-{Ebene}-###", werte.get, 7)
    ok &= pruefe("Muster mit Zähler", text == u"AP-EG-007" and not fehlend)
    text, fehlend = lg.muster_anwenden(u"{Ansicht} {Gibtsnicht}", werte.get)
    ok &= pruefe("fehlender Platzhalter", text == u"EG Grundriss " and fehlend == [u"Gibtsnicht"])
    text, _ = lg.muster_anwenden(u"{Ansicht}", {u"Ansicht": u"Raum #1"}.get)
    ok &= pruefe("Raute im Wert bleibt", text == u"Raum #1")

    vergeben = lg.nummern_menge([u"AP-001", u"AP-002"])
    nummer, _ = lg.nummer_aus_muster(u"AP-###", werte.get, vergeben)
    ok &= pruefe("Muster: erste freie Nummer", nummer == u"AP-003")
    vergeben = lg.nummern_menge([u"EG"])
    nummer, _ = lg.nummer_aus_muster(u"{Ebene}", werte.get, vergeben)
    ok &= pruefe("Muster ohne Zähler bei Belegung", nummer == u"EG-2")

    ok &= pruefe("Umnummerierung",
                 lg.umnummerierung(3, u"B-100", 10) == [u"B-100", u"B-110", u"B-120"])
    ok &= pruefe("Konflikte mit fremden Plänen",
                 lg.konflikte([u"B-100", u"B-110"], lg.nummern_menge([u"b-110"]))
                 == [u"B-110"])

    # --- Ausrichten: Zuordnung der Ansichtsfenster
    F = lg.Fenster
    vorbild = [F(1, "plan", 100, "EG", (30, 30)), F(2, "schnitt", 50, "S-A", (70, 40)),
               F(3, "legende", 0, "LEG", (70, 10))]
    ziel = [F(13, "legende", 0, "LEG", (5, 5)), F(11, "plan", 100, "OG", (35, 28)),
            F(12, "schnitt", 50, "S-B", (65, 45))]
    ok &= pruefe("drei Ansichten je Art zugeordnet",
                 sorted(lg.fenster_zuordnen(vorbild, ziel)) == [(1, 11), (2, 12), (3, 13)])

    vorbild = [F(1, "schnitt", 50, "A", (10, 10)), F(2, "schnitt", 50, "B", (60, 10))]
    ziel = [F(12, "schnitt", 50, "D", (58, 12)), F(11, "schnitt", 50, "C", (12, 9))]
    ok &= pruefe("zwei Schnitte nach Lage",
                 sorted(lg.fenster_zuordnen(vorbild, ziel)) == [(1, 11), (2, 12)])

    vorbild = [F(1, "plan", 100, "EG", (10, 10)), F(2, "plan", 50, "Detail", (60, 10))]
    ziel = [F(11, "plan", 50, "Detail OG", (12, 10)), F(12, "plan", 100, "OG", (60, 10))]
    ok &= pruefe("gleicher Maßstab vor Lage",
                 sorted(lg.fenster_zuordnen(vorbild, ziel)) == [(1, 12), (2, 11)])

    ok &= pruefe("ohne Gegenstück bleibt frei",
                 lg.fenster_zuordnen([F(1, "plan", 100, "EG", (0, 0))],
                                     [F(11, "schnitt", 100, "S", (0, 0))]) == [])

    print("\nAlle Tests bestanden." if ok else "\nEs gibt Fehler.")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
