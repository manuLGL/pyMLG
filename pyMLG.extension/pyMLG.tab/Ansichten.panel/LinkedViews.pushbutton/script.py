#! python3
# -*- coding: utf-8 -*-
"""Verknüpfte Ansichten: die verknüpfte Ansicht von RVT-Verknüpfungen suchen
und in Ansichten bzw. deren Vorlagen einstellen.

Die Logik liegt in lib/link_ansichten (Fenster: fenster.py, Revit:
revit.py, Regeln: logik.py).

CPython-Besonderheiten dieses Repos (siehe ScheduleSync.panel/README.md):
pyrevit.forms ist nicht nutzbar - das Fenster ist eigenes WPF per XAML.
Kein script.exit() - der Ablauf endet nur über 'return'. Ausgaben ins
pyRevit-Ausgabefenster werden vermieden (print_md stürzte bei vielen
Zeilen mit NullReferenceException ab).
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

TITEL = t(u"Verknüpfte Ansichten", u"Linked Views",
          u"Vistas vinculadas")


def main():
    from schedule_sync import ui

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        ui.meldung(t(u"RVT-Verknüpfungen gibt es nur in Projektdateien.",
                     u"RVT links only exist in project files.",
                     u"Los vínculos RVT solo existen en archivos de "
                     u"proyecto."),
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
                   hauptzeile=t(u"Ansichten können nicht geändert werden",
                                u"Views cannot be changed",
                                u"No se pueden cambiar las vistas"),
                   warnung=True)
        return

    from link_ansichten import revit as rv
    verknuepfungen = rv.verknuepfungen(doc)
    if not verknuepfungen:
        ui.meldung(t(u"Im Projekt ist keine RVT-Verknüpfung platziert.",
                     u"No RVT link is placed in the project.",
                     u"No hay ningún vínculo RVT colocado en el proyecto."),
                   titel=TITEL,
                   hauptzeile=t(u"Keine Verknüpfungen", u"No links",
                                u"No hay vínculos"),
                   warnung=True)
        return

    from link_ansichten import fenster
    fenster.starte(__revit__, doc, verknuepfungen)


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "LinkedViews_Fehler.log")
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
                   hauptzeile=t(u"Verknüpfte Ansichten ist fehlgeschlagen",
                                u"Linked Views failed",
                                u"Vistas vinculadas ha fallado"),
                   warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
