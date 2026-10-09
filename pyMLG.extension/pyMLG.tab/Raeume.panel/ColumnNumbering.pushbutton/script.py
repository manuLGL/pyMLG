#! python3
# -*- coding: utf-8 -*-
"""ColumnNumbering: Stützen nach dem Plan des Statikers durchnummerieren.

Die Logik liegt in lib/stuetzen_nummer (Fenster: fenster.py, Revit:
revit.py, Pfad und Nummern: logik.py).

CPython-Besonderheiten dieses Repos (siehe ScheduleSync.panel/README.md):
pyrevit.forms ist nicht nutzbar - das Fenster ist eigenes WPF per XAML.
Kein script.exit() - der Ablauf endet nur über 'return'.
"""

import io
import os
import sys
import traceback

_EXT = os.path.dirname(os.path.abspath(__file__))
while not _EXT.endswith(".extension") and os.path.dirname(_EXT) != _EXT:
    _EXT = os.path.dirname(_EXT)
if os.path.join(_EXT, "lib") not in sys.path:
    sys.path.append(os.path.join(_EXT, "lib"))

from mlg_sprache import t  # noqa: E402

TITEL = t(u"Stützen nummerieren", u"Number Columns", u"Numerar pilares")

# So viele Einträge werden in Meldungen einzeln genannt
MAX_ZEILEN = 8


def _liste(eintraege):
    zeilen = [u"• %s" % e for e in eintraege[:MAX_ZEILEN]]
    if len(eintraege) > MAX_ZEILEN:
        zeilen.append(t(u"… und %d weitere", u"… and %d more",
                        u"… y %d más") % (len(eintraege) - MAX_ZEILEN))
    return u"\n".join(zeilen)


def _klicken(uidoc, lauf, zaehler, einst):
    from stuetzen_nummer import revit as rv

    hinweis = u""
    while True:
        text = hinweis + t(u"Element für %s anklicken - ESC = fertig",
                           u"Pick the element for %s - ESC = done",
                           u"Seleccione el elemento para %s - ESC = "
                           u"terminar") \
            % zaehler.aktuell
        stuetze = rv.waehle_stuetze(uidoc, einst.kategorien, text)
        if stuetze is None:
            return
        schon = lauf.vergeben(stuetze)
        if schon is not None:
            hinweis = t(u"Hat schon %s. ", u"Already has %s. ",
                        u"Ya tiene %s. ") % schon
            continue
        hinweis = u""
        lauf.nummeriere([stuetze], zaehler)


def _pfad(uidoc, lauf, zaehler, einst, linien):
    from schedule_sync import ui
    from stuetzen_nummer import logik as lg
    from stuetzen_nummer import revit as rv

    doc = uidoc.Document
    kandidaten = rv.stuetzen(doc, einst.kategorien, doc.ActiveView)
    segmente = lg.segmente(rv.pfad(linien))
    reihenfolge = lg.entlang(segmente,
                             [rv.grundriss(e) for e in kandidaten],
                             einst.abstand)
    if not reihenfolge:
        ui.meldung(t(u"Kein sichtbares Element liegt nahe genug an der "
                     u"Linie (bei Wänden zählt die Wandmitte). Abstand im "
                     u"Fenster vergrößern?",
                     u"No visible element is close enough to the line "
                     u"(for walls the wall midpoint counts). Increase the "
                     u"distance in the window?",
                     u"Ningún elemento visible está lo bastante cerca de la "
                     u"línea (en los muros cuenta el punto medio). "
                     u"¿Aumentar la distancia en la ventana?"),
                   titel=TITEL, hauptzeile=t(u"Keine Elemente gefunden",
                                             u"No elements found",
                                             u"No se han encontrado "
                                             u"elementos"),
                   warnung=True)
        return
    lauf.nummeriere([kandidaten[i] for i in reihenfolge], zaehler)


def _linien(uidoc):
    """Vorausgewählte Linien, wenn sie einen zusammenhängenden Pfad bilden,
    sonst anklicken lassen."""
    from stuetzen_nummer import revit as rv

    doc = uidoc.Document
    linien = rv.gewaehlte_linien(doc, uidoc.Selection.GetElementIds())
    if linien and len(rv.pfad(linien)) == 1:
        return linien
    return rv.waehle_linien(uidoc)


def _abschluss(lauf, zaehler):
    """Zusammenfassung zeigen, übernehmen oder verwerfen."""
    from schedule_sync import ui
    from stuetzen_nummer import fenster
    from stuetzen_nummer import revit as rv

    if not lauf.reihenfolge:
        lauf.verwerfen()
        return
    erste, letzte = lauf.reihenfolge[0][0], lauf.reihenfolge[-1][0]
    gestapelt = sum(anzahl - 1 for _, _, anzahl in lauf.reihenfolge)
    zeilen = [t(u"%d Nummern: %s … %s", u"%d numbers: %s … %s",
                u"%d números: %s … %s") % (len(lauf.reihenfolge), erste,
                                            letzte)]
    if gestapelt:
        zeilen.append(t(u"Dazu %d Stützen darüber/darunter mit gleicher "
                        u"Nummer.",
                        u"Plus %d columns above/below with the same number.",
                        u"Además %d pilares encima/debajo con el mismo "
                        u"número.") % gestapelt)
    if lauf.fehler:
        zeilen += [u"", t(u"Nicht geschrieben:", u"Not written:",
                          u"No escritos:"),
                   _liste([u"%s: %s" % (rv.beschreibe(e), g)
                           for e, g in lauf.fehler])]
    doppelt = lauf.doppelte()
    if doppelt:
        zeilen += [u"", t(u"Diese Nummern tragen schon andere Elemente:",
                          u"These numbers are already used by other "
                          u"elements:",
                          u"Estos números ya los usan otros elementos:"),
                   _liste([u"%s: %s" % (n, u", ".join(rv.beschreibe(e)
                                                     for e in el[:3]))
                           for n, el in sorted(doppelt.items())])]
    zeilen += [u"", t(u"Nein = alles zurücknehmen.",
                      u"No = undo everything.",
                      u"No = deshacer todo.")]
    try:
        behalten = ui.frage(u"\n".join(zeilen), titel=TITEL,
                            standard_ja=True,
                            warnung=bool(lauf.fehler or doppelt),
                            hauptzeile=t(u"Nummern übernehmen?",
                                         u"Keep the numbers?",
                                         u"¿Conservar los números?"))
    except Exception:
        lauf.verwerfen()
        raise
    if behalten:
        lauf.uebernehmen()
        fenster.merke_naechste(zaehler.aktuell)
    else:
        lauf.verwerfen()


def main():
    from schedule_sync import ui

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        ui.meldung(t(u"Stützen nummerieren geht nur in Projektdateien.",
                     u"Numbering columns only works in project files.",
                     u"Numerar pilares solo funciona en archivos de "
                     u"proyecto."),
                   titel=TITEL, hauptzeile=t(u"Kein Projekt geöffnet",
                                             u"No project open",
                                             u"No hay ningún proyecto "
                                             u"abierto"),
                   warnung=True)
        return
    if doc.IsReadOnly:
        ui.meldung(t(u"Das Dokument ist schreibgeschützt.",
                     u"The document is read-only.",
                     u"El documento es de solo lectura."), titel=TITEL,
                   hauptzeile=t(u"Stützen können nicht nummeriert werden",
                                u"Columns cannot be numbered",
                                u"No se pueden numerar los pilares"),
                   warnung=True)
        return

    from stuetzen_nummer import fenster
    from stuetzen_nummer import logik as lg
    from stuetzen_nummer import revit as rv

    einst = fenster.frage_einstellungen(__revit__, doc)
    if einst is None:
        return

    linien = None
    if einst.modus == fenster.PFAD:
        linien = _linien(uidoc)
        if not linien:
            return

    zaehler = lg.Zaehler(einst.start, einst.schritt)
    lauf = rv.Lauf(uidoc, einst.schluessel, einst.kategorien, einst.stapeln)
    try:
        if linien is None:
            _klicken(uidoc, lauf, zaehler, einst)
        else:
            _pfad(uidoc, lauf, zaehler, einst, linien)
    except Exception:
        lauf.verwerfen()
        raise
    _abschluss(lauf, zaehler)


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "ColumnNumbering_Fehler.log")
    try:
        if not os.path.isdir(os.path.dirname(protokoll)):
            os.makedirs(os.path.dirname(protokoll))
        with io.open(protokoll, "a", encoding="utf-8") as datei:
            datei.write(spur + u"\n" + u"-" * 70 + u"\n")
    except Exception:
        protokoll = None
    try:
        from schedule_sync import ui
        letzte_zeile = (spur.strip().splitlines() or [u""])[-1]
        ui.meldung(letzte_zeile + (t(u"\n\nProtokoll: %s", u"\n\nLog: %s",
                                     u"\n\nRegistro: %s") % protokoll
                                   if protokoll else u""),
                   titel=TITEL,
                   hauptzeile=t(u"Stützen nummerieren ist fehlgeschlagen",
                                u"Number Columns failed",
                                u"Numerar pilares ha fallado"),
                   warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
