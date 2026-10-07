#! python3
# -*- coding: utf-8 -*-
"""DxfLegend: Legende aus einer DXF-Datei als Revit-Legende nachbauen -
Linien mit Farbe, Muster und Stärke, Texte, Schraffuren.

Die Logik liegt in lib/dxf_legende (Leser: dxf.py, Umsetzung: logik.py,
Revit: revit.py, Dialog: fenster.py).

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

TITEL = t(u"Legende aus DXF", u"Legend from DXF", u"Leyenda desde DXF")


def _bericht(ergebnis):
    zeilen = [t(u"%d Linien, %d Texte, %d Füllbereiche",
                u"%d lines, %d texts, %d filled regions",
                u"%d líneas, %d textos, %d regiones rellenadas")
              % (ergebnis.linien, ergebnis.texte, ergebnis.flaechen)]
    neu = ergebnis.neu
    zeilen.append(t(u"Neu angelegt: %d Linienstile, %d Linienmuster, %d Texttypen, "
                    u"%d Füllbereichstypen, %d Füllmuster",
                    u"Created: %d line styles, %d line patterns, %d text types, "
                    u"%d filled region types, %d fill patterns",
                    u"Creados: %d estilos de línea, %d patrones de línea, %d tipos "
                    u"de texto, %d tipos de región, %d patrones de relleno")
                  % (neu[u"stile"], neu[u"muster"], neu[u"texttypen"],
                     neu[u"fuelltypen"], neu[u"fuellmuster"]))
    geaendert = ergebnis.geaendert
    if any(geaendert.values()):
        zeilen.append(t(u"Überschrieben: %d Linienstile, %d Linienmuster, %d Texttypen, "
                        u"%d Füllbereichstypen, %d Füllmuster",
                        u"Overwritten: %d line styles, %d line patterns, %d text types, "
                        u"%d filled region types, %d fill patterns",
                        u"Sobrescritos: %d estilos de línea, %d patrones de línea, %d tipos "
                        u"de texto, %d tipos de región, %d patrones de relleno")
                      % (geaendert[u"stile"], geaendert[u"muster"], geaendert[u"texttypen"],
                         geaendert[u"fuelltypen"], geaendert[u"fuellmuster"]))
    if ergebnis.ersetzt:
        zeilen.insert(0, t(u"Vorhandene Ansicht ersetzt (%d alte Elemente gelöscht).",
                           u"Existing view replaced (%d old elements deleted).",
                           u"Vista existente reemplazada (%d elementos antiguos borrados).")
                      % ergebnis.geloescht)
    if ergebnis.fehler:
        zeilen.append(u"")
        zeilen.append(t(u"Nicht erstellt:", u"Not created:", u"No creados:"))
        for art, anzahl in sorted(ergebnis.fehler.items()):
            zeilen.append(u"• %s: %d" % (art, anzahl))
        protokoll = ergebnis.schreibe_protokoll()
        if protokoll:
            zeilen.append(t(u"Protokoll: %s", u"Log: %s", u"Registro: %s") % protokoll)
    zeilen.append(u"")
    zeilen.append(t(u"Linienstärken sind Revit-Stifte - die Breite hängt von den "
                    u"Linienstärken des Projekts ab.",
                    u"Line weights are Revit pens - the width depends on the "
                    u"project's line weights.",
                    u"Los grosores son plumas de Revit - el ancho depende de los "
                    u"grosores de línea del proyecto."))
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

    from dxf_legende import fenster
    auswahl = fenster.starte(__revit__, doc)
    if not auswahl:
        return

    from dxf_legende import revit as rv
    try:
        ergebnis = rv.erzeuge(doc, auswahl[u"plan"], auswahl[u"name"],
                              auswahl[u"massstab"], auswahl[u"vorlage"],
                              ueberschreiben=auswahl[u"ueberschreiben"],
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
    protokoll = os.path.join(basis, "pyMLG", "DxfLegend_Fehler.log")
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
                   hauptzeile=t(u"Legende aus DXF ist fehlgeschlagen",
                                u"Legend from DXF failed",
                                u"Leyenda desde DXF ha fallado"),
                   warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
