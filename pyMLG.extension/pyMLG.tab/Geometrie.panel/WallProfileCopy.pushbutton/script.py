#! python3
# -*- coding: utf-8 -*-
"""WallProfileCopy: bearbeitetes Wandprofil auf gleiche Wände übertragen.

Die Logik liegt in lib/wand_profil (Revit: revit.py, Lage: logik.py).

CPython-Besonderheiten dieses Repos (siehe ScheduleSync.panel/README.md):
pyrevit.forms ist nicht nutzbar - Dialoge über schedule_sync.ui.
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

TITEL = t(u"Profil übertragen", u"Transfer profile", u"Transferir perfil")

# So viele Wände werden in Meldungen einzeln genannt
MAX_ZEILEN = 8


def _liste(eintraege):
    zeilen = [u"• %s: %s" % e for e in eintraege[:MAX_ZEILEN]]
    if len(eintraege) > MAX_ZEILEN:
        zeilen.append(t(u"… und %d weitere", u"… and %d more",
                        u"… y %d más") % (len(eintraege) - MAX_ZEILEN))
    return u"\n".join(zeilen)


def main():
    from schedule_sync import ui

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        ui.meldung(t(u"Wandprofile gibt es nur in Projektdateien.", u"Wall profiles only exist in project files.", u"Los perfiles de muro solo existen en archivos de proyecto."),
                   titel=TITEL, hauptzeile=t(u"Kein Projekt geöffnet", u"No project open", u"No hay ningún proyecto abierto"),
                   warnung=True)
        return
    if doc.IsReadOnly:
        ui.meldung(t(u"Das Dokument ist schreibgeschützt.", u"The document is read-only.", u"El documento es de solo lectura."), titel=TITEL,
                   hauptzeile=t(u"Profil kann nicht übertragen werden", u"Profile cannot be transferred", u"No se puede transferir el perfil"),
                   warnung=True)
        return

    from wand_profil import revit as rv

    quelle, ziele = rv.aufteilen(doc, uidoc.Selection.GetElementIds())
    try:
        if quelle is None:
            quelle = rv.waehle_quelle(uidoc)
            ziele = [w for w in ziele if w.Id != quelle.Id]
        if not ziele:
            ziele = rv.waehle_ziele(uidoc, quelle)
    except rv.Abbruch:
        return

    passend, uebersprungen = rv.pruefe_ziele(doc, quelle, ziele)
    uebersprungen = [(rv.beschreibe(w), g) for w, g in uebersprungen]
    if not passend:
        ui.meldung(_liste(uebersprungen) or t(u"Keine Zielwände gewählt.", u"No target walls selected.", u"No se han seleccionado muros de destino."),
                   titel=TITEL, hauptzeile=t(u"Keine passende Zielwand", u"No suitable target wall", u"Ningún muro de destino adecuado"),
                   warnung=True)
        return

    text = t(u"Quelle: %s", u"Source: %s", u"Origen: %s") % rv.beschreibe(quelle)
    if uebersprungen:
        text += u"\n\n" + t(u"Übersprungen:", u"Skipped:", u"Omitidos:") + u"\n" + _liste(uebersprungen)
    if not ui.frage(text, titel=TITEL, standard_ja=True,
                    hauptzeile=t(u"Profil auf %d Wände übertragen?", u"Transfer profile to %d walls?", u"¿Transferir el perfil a %d muros?") % len(passend)):
        return

    fertig, fehler = rv.uebertrage(doc, quelle, passend)
    fehler = [(rv.beschreibe(w), g) for w, g in fehler]
    probleme = uebersprungen + fehler
    hauptzeile = t(u"Profil auf %d von %d Wänden übertragen", u"Profile transferred to %d of %d walls", u"Perfil transferido a %d de %d muros") % (
        len(fertig), len(passend) + len(uebersprungen))
    if probleme:
        ui.meldung(t(u"Nicht übertragen:", u"Not transferred:", u"No transferidos:") + u"\n" + _liste(probleme)
                   + (t(u"\n\nProtokoll: %s", u"\n\nLog: %s", u"\n\nRegistro: %s") % rv.PROTOKOLL if fehler else u""),
                   titel=TITEL, hauptzeile=hauptzeile, warnung=True)
    else:
        ui.meldung(t(u"Alle Wände haben jetzt das Profil der Quellwand.", u"All walls now have the profile of the source wall.", u"Todos los muros tienen ahora el perfil del muro de origen."),
                   titel=TITEL, hauptzeile=hauptzeile)


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "WallProfileCopy_Fehler.log")
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
                   titel=TITEL, hauptzeile=t(u"Profil übertragen ist fehlgeschlagen", u"Transfer profile failed", u"Transferir perfil ha fallado"),
                   warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
