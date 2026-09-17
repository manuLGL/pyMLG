#! python3
# -*- coding: utf-8 -*-
"""JoinMultiple: viele Elemente nach Kategorie-Priorität verbinden.

Die Logik liegt in lib/join_multiple (Fenster: fenster.py, Revit: revit.py,
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

TITEL = u"JoinMultiple"


def main():
    from schedule_sync import ui

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        ui.meldung(u"Elemente verbinden geht nur in Projektdateien.",
                   titel=TITEL, hauptzeile=u"Kein Projekt geöffnet",
                   warnung=True)
        return
    if doc.IsReadOnly:
        ui.meldung(u"Das Dokument ist schreibgeschützt.", titel=TITEL,
                   hauptzeile=u"Elemente können nicht verbunden werden",
                   warnung=True)
        return

    from join_multiple import fenster
    fenster.starte(__revit__, uidoc)


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "JoinMultiple_Fehler.log")
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
        ui.meldung(letzte_zeile + (u"\n\nProtokoll: %s" % protokoll
                                   if protokoll else u""),
                   titel=TITEL, hauptzeile=u"JoinMultiple ist fehlgeschlagen",
                   warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
