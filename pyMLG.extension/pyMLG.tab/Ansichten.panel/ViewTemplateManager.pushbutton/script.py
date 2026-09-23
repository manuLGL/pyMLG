#! python3
# -*- coding: utf-8 -*-
"""Ansichtsvorlagen-Manager: Vorlagen suchen und mehrere zugleich bearbeiten.

Die Logik liegt in lib/vorlagen_manager (Fenster: fenster.py, Revit:
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

TITEL = t(u"Ansichtsvorlagen-Manager", u"View Template Manager",
          u"Gestor de plantillas de vista")


def main():
    from schedule_sync import ui

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        ui.meldung(t(u"Ansichtsvorlagen gibt es nur in Projektdateien.",
                     u"View templates only exist in project files.",
                     u"Las plantillas de vista solo existen en archivos de "
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
                   hauptzeile=t(u"Vorlagen können nicht bearbeitet werden",
                                u"Templates cannot be edited",
                                u"No se pueden editar las plantillas"),
                   warnung=True)
        return

    from vorlagen_manager import fenster
    fenster.starte(__revit__, doc)


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "ViewTemplateManager_Fehler.log")
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
                   hauptzeile=t(u"Ansichtsvorlagen-Manager ist fehlgeschlagen",
                                u"View Template Manager failed",
                                u"El gestor de plantillas de vista ha "
                                u"fallado"),
                   warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
