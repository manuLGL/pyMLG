#! python3
# -*- coding: utf-8 -*-
"""FilterLegend: Legende aus den Ansichtsfiltern - je Filter ein Farbfeld
mit Muster und Farbe der Überschreibung und daneben der Name.

Die Logik liegt in lib/filter_legende (logik.py, revit.py, fenster.py).

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

TITEL = t(u"Filterlegende", u"Filter legend", u"Leyenda de filtros")

# so viele Filter werden im Bericht einzeln genannt
NAMEN_IM_BERICHT = 15


def _namen(namen):
    zeilen = [u"• " + name for name in namen[:NAMEN_IM_BERICHT]]
    if len(namen) > NAMEN_IM_BERICHT:
        zeilen.append(t(u"… und %d weitere", u"… and %d more", u"… y %d más")
                      % (len(namen) - NAMEN_IM_BERICHT))
    return zeilen


def _bericht(ergebnis):
    zeilen = [t(u"%d Farbfelder, %d Beschriftungen, %d neue Füllbereichstypen",
                u"%d swatches, %d labels, %d new filled region types",
                u"%d muestras, %d etiquetas, %d tipos de región rellenada nuevos")
              % (ergebnis.felder, ergebnis.texte, ergebnis.neue_typen)]
    if ergebnis.ersetzt:
        zeilen.insert(0, t(u"Vorhandene Legende ersetzt (%d alte Elemente gelöscht).",
                           u"Existing legend replaced (%d old elements deleted).",
                           u"Leyenda existente reemplazada (%d elementos antiguos borrados).")
                      % ergebnis.geloescht)
    for liste, kopf in (
            (ergebnis.ohne_farbe, t(u"Ohne Farbüberschreibung, übersprungen:",
                                    u"No colour override, skipped:",
                                    u"Sin color modificado, omitidos:")),
            (ergebnis.ersatzmuster, t(u"Muster nicht für Füllbereiche nutzbar - "
                                      u"als Vollton gezeigt:",
                                      u"Pattern not usable for filled regions - "
                                      u"shown as solid fill:",
                                      u"Patrón no utilizable en regiones rellenadas - "
                                      u"se muestra como relleno sólido:")),
            (ergebnis.fehler, t(u"Farbfeld konnte nicht erstellt werden:",
                                u"Swatch could not be created:",
                                u"No se pudo crear la muestra:"))):
        if liste:
            zeilen.append(u"")
            zeilen.append(kopf)
            zeilen.extend(_namen(liste))
    return u"\n".join(zeilen)


def main():
    from schedule_sync import ui

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        ui.meldung(t(u"Legenden gibt es nur in Projektdateien.",
                     u"Legends only exist in project files.",
                     u"Las leyendas solo existen en archivos de proyecto."),
                   titel=TITEL,
                   hauptzeile=t(u"Kein Projekt geöffnet", u"No project open",
                                u"No hay ningún proyecto abierto"),
                   warnung=True)
        return
    if doc.IsReadOnly:
        ui.meldung(t(u"Das Dokument ist schreibgeschützt.",
                     u"The document is read-only.",
                     u"El documento es de solo lectura."),
                   titel=TITEL,
                   hauptzeile=t(u"Legende kann nicht erstellt werden",
                                u"Legend cannot be created",
                                u"No se puede crear la leyenda"),
                   warnung=True)
        return

    from filter_legende import revit as rv
    if not rv.quellen(doc):
        ui.meldung(t(u"Keine Ansicht und keine Ansichtsvorlage im Projekt hat Filter.",
                     u"No view and no view template in the project has filters.",
                     u"Ninguna vista ni plantilla de vista del proyecto tiene filtros."),
                   titel=TITEL,
                   hauptzeile=t(u"Keine Filter gefunden", u"No filters found",
                                u"No se encontraron filtros"),
                   warnung=True)
        return

    from filter_legende import fenster
    auswahl = fenster.starte(__revit__, doc, uidoc.ActiveView)
    if not auswahl:
        return

    try:
        ergebnis = rv.erzeuge(doc, auswahl[u"posten"], auswahl[u"name"],
                              auswahl[u"massstab"], auswahl[u"feld_breite"],
                              auswahl[u"feld_hoehe"], auswahl[u"abstand"],
                              spalten=auswahl[u"spalten"],
                              spaltenbreite=auswahl[u"spaltenbreite"],
                              quelle=auswahl[u"quelle"],
                              texttyp_id=auswahl[u"texttyp"],
                              rahmen=auswahl[u"rahmen"],
                              ersetzen=auswahl[u"ersetzen"])
    except rv.LegendenFehler as fehler:
        ui.meldung(u"%s" % fehler, titel=TITEL,
                   hauptzeile=t(u"Legende nicht erstellt", u"Legend not created",
                                u"Leyenda no creada"),
                   warnung=True)
        return

    try:
        uidoc.ActiveView = ergebnis.ansicht
    except Exception:
        pass
    ui.meldung(_bericht(ergebnis), titel=TITEL,
               hauptzeile=t(u"\"%s\" erstellt", u"\"%s\" created", u"\"%s\" creada")
               % ergebnis.ansicht.Name,
               warnung=bool(ergebnis.fehler))


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "FilterLegend_Fehler.log")
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
                   hauptzeile=t(u"Filterlegende ist fehlgeschlagen",
                                u"Filter legend failed",
                                u"Leyenda de filtros ha fallado"),
                   warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
