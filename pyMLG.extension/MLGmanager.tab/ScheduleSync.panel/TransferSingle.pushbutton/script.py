#! python3
# -*- coding: utf-8 -*-
"""TransferSingle: einzelne Elemente von einem Projekt in andere übertragen.

Die Logik liegt in lib/transfer_single (Fenster: fenster.py, Revit: revit.py,
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

TITEL = u"TransferSingle"


def main():
    from schedule_sync import ui

    projekte = [d for d in __revit__.Application.Documents
                if not d.IsFamilyDocument and not d.IsLinked]
    if not projekte:
        ui.meldung(t(u"Bitte zuerst ein Projekt öffnen.",
                     u"Please open a project first.",
                     u"Abra primero un proyecto."),
                   titel=TITEL,
                   hauptzeile=t(u"Kein Projekt geöffnet", u"No project open",
                                u"No hay ningún proyecto abierto"),
                   warnung=True)
        return

    from transfer_single import fenster
    fenster.starte(__revit__)


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "TransferSingle_Fehler.log")
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
                   hauptzeile=t(u"TransferSingle ist fehlgeschlagen",
                                u"TransferSingle failed",
                                u"TransferSingle ha fallado"),
                   warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
