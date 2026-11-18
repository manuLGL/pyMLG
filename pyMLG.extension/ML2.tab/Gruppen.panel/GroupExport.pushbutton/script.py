#! python3
# -*- coding: utf-8 -*-
"""GroupExport: gewählte Modellgruppen als eigene Revit-Modelle (.rvt) in
einen Ordner speichern - Dateiname = Gruppenname.

Die Logik liegt in lib/gruppen_export (Revit: revit.py, Dateinamen:
logik.py).

CPython-Besonderheiten dieses Repos (siehe ScheduleSync.panel/README.md):
pyrevit.forms ist nicht nutzbar - Dialoge kommen aus schedule_sync.ui.
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

TITEL = t(u"Gruppen exportieren", u"Export Groups", u"Exportar grupos")

# So viele Einträge werden in Meldungen einzeln genannt
MAX_ZEILEN = 8


def _liste(eintraege):
    zeilen = [u"• %s" % e for e in eintraege[:MAX_ZEILEN]]
    if len(eintraege) > MAX_ZEILEN:
        zeilen.append(t(u"… und %d weitere", u"… and %d more",
                        u"… y %d más") % (len(eintraege) - MAX_ZEILEN))
    return u"\n".join(zeilen)


def _warnung(hauptzeile, text):
    from schedule_sync import ui
    ui.meldung(text, titel=TITEL, hauptzeile=hauptzeile, warnung=True)


def _gruppen(uidoc):
    """Vorausgewählte Modellgruppen, sonst Auswahl aus einer Liste.
    Rückgabe {Gruppenname: Instanz} oder None."""
    from gruppen_export import revit as rv
    from schedule_sync import ui

    doc = uidoc.Document
    gruppen = rv.gewaehlte_gruppen(doc, uidoc.Selection.GetElementIds())
    if gruppen:
        return gruppen
    alle = rv.alle_gruppen(doc)
    if not alle:
        _warnung(t(u"Keine Modellgruppen gefunden",
                   u"No model groups found",
                   u"No se han encontrado grupos de modelo"),
                 t(u"Im Projekt ist keine Modellgruppe platziert.",
                   u"No model group is placed in the project.",
                   u"No hay ningún grupo de modelo colocado en el "
                   u"proyecto."))
        return None
    namen = ui.waehle_mehrfach(
        sorted(alle, key=lambda n: n.lower()), titel=TITEL,
        hinweis=t(u"Gruppen wählen - jede wird als eigenes Modell mit ihrem "
                  u"Namen gespeichert:",
                  u"Choose groups - each is saved as its own model with its "
                  u"name:",
                  u"Elija grupos: cada uno se guarda como modelo propio con "
                  u"su nombre:"),
        schaltflaeche=t(u"Weiter", u"Next", u"Siguiente"))
    if not namen:
        return None
    return dict((n, alle[n]) for n in namen)


def _startordner(doc):
    pfad = doc.PathName
    if pfad and os.path.isdir(os.path.dirname(pfad)):
        return os.path.dirname(pfad)
    return u""


def main():
    from schedule_sync import ui

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        _warnung(t(u"Kein Projekt geöffnet", u"No project open",
                   u"No hay ningún proyecto abierto"),
                 t(u"Gruppen exportieren geht nur in Projektdateien.",
                   u"Exporting groups only works in project files.",
                   u"Exportar grupos solo funciona en archivos de "
                   u"proyecto."))
        return

    from gruppen_export import logik as lg
    from gruppen_export import revit as rv

    gruppen = _gruppen(uidoc)
    if not gruppen:
        return
    ordner = ui.ordner_waehlen(
        _startordner(doc),
        titel=t(u"Ordner für die exportierten Gruppen wählen",
                u"Choose the folder for the exported groups",
                u"Elija la carpeta para los grupos exportados"))
    if not ordner:
        return

    namen = sorted(gruppen, key=lambda n: n.lower())
    auftraege = [(n, os.path.join(ordner, datei))
                 for n, datei in zip(namen, lg.dateinamen(namen))]
    vorhanden = [os.path.basename(p) for _, p in auftraege
                 if os.path.exists(p)]
    if vorhanden and not ui.frage(
            t(u"Diese Dateien gibt es im Ordner schon:",
              u"These files already exist in the folder:",
              u"Estos archivos ya existen en la carpeta:")
            + u"\n" + _liste(vorhanden) + u"\n\n"
            + t(u"Überschreiben? Nein = diese Gruppen auslassen.",
                u"Overwrite? No = skip these groups.",
                u"¿Sobrescribir? No = omitir estos grupos."),
            titel=TITEL, standard_ja=False, warnung=True,
            hauptzeile=t(u"Dateien überschreiben?", u"Overwrite files?",
                         u"¿Sobrescribir archivos?")):
        auftraege = [(n, p) for n, p in auftraege if not os.path.exists(p)]
        if not auftraege:
            return

    app = __revit__.Application
    fertig, fehler = [], []
    for name, pfad in auftraege:
        try:
            rv.exportiere(app, doc, gruppen[name], pfad)
            fertig.append(os.path.basename(pfad))
        except Exception as ausnahme:
            fehler.append(u"%s: %s" % (name, ausnahme))

    zeilen = [t(u"%d Gruppen gespeichert in:", u"%d groups saved in:",
                u"%d grupos guardados en:") % len(fertig), ordner]
    if fertig:
        zeilen += [u"", _liste(fertig)]
    if fehler:
        zeilen += [u"", t(u"Nicht exportiert:", u"Not exported:",
                          u"No exportados:"), _liste(fehler)]
    ui.meldung(u"\n".join(zeilen), titel=TITEL, warnung=bool(fehler),
               hauptzeile=t(u"Export abgeschlossen", u"Export finished",
                            u"Exportación terminada") if not fehler
               else t(u"Export mit Fehlern", u"Export with errors",
                      u"Exportación con errores"))


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "GroupExport_Fehler.log")
    try:
        if not os.path.isdir(os.path.dirname(protokoll)):
            os.makedirs(os.path.dirname(protokoll))
        with io.open(protokoll, "a", encoding="utf-8") as datei:
            datei.write(spur + u"\n" + u"-" * 70 + u"\n")
    except Exception:
        protokoll = None
    try:
        letzte_zeile = (spur.strip().splitlines() or [u""])[-1]
        _warnung(t(u"Gruppen exportieren ist fehlgeschlagen",
                   u"Export Groups failed",
                   u"Exportar grupos ha fallado"),
                 letzte_zeile + (t(u"\n\nProtokoll: %s", u"\n\nLog: %s",
                                   u"\n\nRegistro: %s") % protokoll
                                 if protokoll else u""))
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
