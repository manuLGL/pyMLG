#! python3
# -*- coding: utf-8 -*-
"""LevelAutoSet: Elemente auf Ebenen setzen, ohne sie zu verschieben.

Die Logik liegt in lib/level_auto_set (Fenster: fenster.py, Revit: revit.py,
Regeln: logik.py).

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

TITEL = u"LevelAutoSet"


def main():
    from schedule_sync import ui

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        ui.meldung(t(u"Ebenen gibt es nur in Projektdateien.", u"Levels only exist in project files.", u"Los niveles solo existen en archivos de proyecto."),
                   titel=TITEL, hauptzeile=t(u"Kein Projekt geöffnet", u"No project open", u"No hay ningún proyecto abierto"),
                   warnung=True)
        return
    if doc.IsReadOnly:
        ui.meldung(t(u"Das Dokument ist schreibgeschützt.", u"The document is read-only.", u"El documento es de solo lectura."), titel=TITEL,
                   hauptzeile=t(u"Ebenen können nicht gesetzt werden", u"Levels cannot be set", u"No se pueden asignar niveles"),
                   warnung=True)
        return

    from level_auto_set import fenster
    fenster.starte(__revit__, uidoc)


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "LevelAutoSet_Fehler.log")
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
        ui.meldung(letzte_zeile + (t(u"\n\nProtokoll: %s", u"\n\nLog: %s", u"\n\nRegistro: %s") % protokoll
                                   if protokoll else u""),
                   titel=TITEL, hauptzeile=t(u"LevelAutoSet ist fehlgeschlagen", u"LevelAutoSet failed", u"LevelAutoSet ha fallado"),
                   warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
