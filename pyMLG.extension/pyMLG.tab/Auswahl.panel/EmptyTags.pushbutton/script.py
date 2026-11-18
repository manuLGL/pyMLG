#! python3
# -*- coding: utf-8 -*-
"""EmptyTags: alle "?"-Beschriftungen eines Typs in der aktiven Ansicht wählen.

Die Logik liegt in lib/leere_beschriftung (Revit: revit.py, Regeln: logik.py).

CPython-Besonderheiten dieses Repos (siehe ScheduleSync.panel/README.md):
pyrevit.forms ist nicht nutzbar. Kein script.exit() - der Ablauf endet nur
über 'return'.
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

TITEL = u"Empty Tags"


def _shift_gedrueckt():
    try:
        from System.Windows.Input import Keyboard, ModifierKeys
        return bool(Keyboard.Modifiers & ModifierKeys.Shift)
    except Exception:
        return False


def main():
    from schedule_sync import ui

    ganze_familie = _shift_gedrueckt()

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        ui.meldung(t(u"Empty Tags funktioniert nur in Projektdateien.",
                     u"Empty Tags only works in project files.",
                     u"Empty Tags solo funciona en archivos de proyecto."),
                   titel=TITEL,
                   hauptzeile=t(u"Kein Projekt geöffnet", u"No project open",
                                u"No hay ningún proyecto abierto"),
                   warnung=True)
        return

    from leere_beschriftung import revit as rv

    ansicht = doc.ActiveView
    beispiele = rv.beschriftungen_der_auswahl(uidoc)
    if not beispiele:
        beispiele = rv.waehle_muster(uidoc)
        if not beispiele:
            return

    gefunden = rv.passende(doc, ansicht, beispiele, ganze_familie)
    if not gefunden:
        ui.meldung(t(u"In der aktiven Ansicht zeigt keine Beschriftung "
                     u"denselben Text wie das Muster.",
                     u"No tag in the active view shows the same text as "
                     u"the sample.",
                     u"Ninguna etiqueta de la vista activa muestra el mismo "
                     u"texto que la muestra."),
                   titel=TITEL,
                   hauptzeile=t(u"Nichts gefunden", u"Nothing found",
                                u"No se encontró nada"))
        return
    rv.setze_auswahl(uidoc, gefunden)


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "EmptyTags_Fehler.log")
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
                   hauptzeile=t(u"Empty Tags ist fehlgeschlagen",
                                u"Empty Tags failed",
                                u"Empty Tags ha fallado"),
                   warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
