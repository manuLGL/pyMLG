#! python3
# -*- coding: utf-8 -*-
"""AutoRooms: Räume im Geschoss des aktiven Grundrisses nach der
Einrichtung anlegen - Bett, Sanitär, Küche + Sofa, Türen.

Die Logik liegt in lib/raum_automatik (Regeln: logik.py, Revit: revit.py).

CPython-Besonderheiten dieses Repos (siehe ScheduleSync.panel/README.md):
pyrevit.forms ist nicht nutzbar - Dialoge aus schedule_sync.ui.
Kein script.exit() - der Ablauf endet nur über 'return'.
"""

import io
import os
import sys
import time
import traceback

_EXT = os.path.dirname(os.path.abspath(__file__))
while not _EXT.endswith(".extension") and os.path.dirname(_EXT) != _EXT:
    _EXT = os.path.dirname(_EXT)
if os.path.join(_EXT, "lib") not in sys.path:
    sys.path.append(os.path.join(_EXT, "lib"))

from mlg_sprache import t  # noqa: E402

TITEL = t(u"Räume automatisch", u"Auto Rooms", u"Habitaciones automáticas")

MAX_ZEILEN = 15


class _Uhr(object):
    """Dauer der Schritte - nach %LOCALAPPDATA%/pyMLG/AutoRooms_Ablauf.log."""

    def __init__(self):
        self.start = self.letzte = time.time()
        self.zeilen = []

    def schritt(self, text):
        jetzt = time.time()
        self.zeilen.append(u"%7.1f s  %s" % (jetzt - self.letzte, text))
        self.letzte = jetzt

    def schreibe(self):
        basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
        pfad = os.path.join(basis, "pyMLG", "AutoRooms_Ablauf.log")
        try:
            if not os.path.isdir(os.path.dirname(pfad)):
                os.makedirs(os.path.dirname(pfad))
            with io.open(pfad, "w", encoding="utf-8") as datei:
                datei.write(u"\n".join(self.zeilen + [
                    u"%7.1f s  gesamt" % (time.time() - self.start)]))
        except Exception:
            pass


def _liste(eintraege):
    zeilen = [u"• %s" % e for e in eintraege[:MAX_ZEILEN]]
    if len(eintraege) > MAX_ZEILEN:
        zeilen.append(t(u"… und %d weitere", u"… and %d more",
                        u"… y %d más") % (len(eintraege) - MAX_ZEILEN))
    return zeilen


def _bericht(ergebnis):
    from raum_automatik import logik as lg

    zaehler = {}
    for _, bereich in ergebnis.angelegt:
        zaehler[bereich.art] = zaehler.get(bereich.art, 0) + 1
    zeilen = [t(u"%d Räume angelegt:", u"%d rooms created:",
                u"%d habitaciones creadas:") % len(ergebnis.angelegt)]
    zeilen += [u"• %s: %d" % (lg.NAMEN[a].replace(u" %d", u""), zaehler[a])
               for a in lg.REIHENFOLGE if a in zaehler]
    zeilen.append(t(u"%d freie Bereiche ohne passende Regel bleiben leer.",
                    u"%d free areas without a matching rule stay empty.",
                    u"%d zonas libres sin regla aplicable quedan vacías.")
                  % ergebnis.verworfen)
    if ergebnis.fehlende_parameter:
        zeilen += [u"", t(u"Parameter fehlt oder ist schreibgeschützt: %s",
                          u"Parameter missing or read-only: %s",
                          u"Parámetro inexistente o de solo lectura: %s")
                   % u", ".join(sorted(ergebnis.fehlende_parameter))]
    if ergebnis.ohne_schluessel:
        zeilen += [u"", t(u"Schlüssel fehlt in der Schlüsselbauteilliste "
                          u"Raumkategorie: %s",
                          u"Key missing in the Raumkategorie key schedule: %s",
                          u"Falta la clave en la tabla de claves "
                          u"Raumkategorie: %s")
                   % u", ".join(sorted(ergebnis.ohne_schluessel))]
    if ergebnis.fehler:
        zeilen += [u"", t(u"Fehler:", u"Errors:", u"Errores:")]
        zeilen += _liste(ergebnis.fehler)
    return zeilen


def main():
    from schedule_sync import ui

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument or doc.IsReadOnly:
        ui.meldung(t(u"Das Werkzeug geht nur in beschreibbaren Projektdateien.",
                     u"The tool only works in writable project files.",
                     u"La herramienta solo funciona en proyectos editables."),
                   titel=TITEL, warnung=True)
        return

    from Autodesk.Revit.DB import Transaction, TransactionGroup
    from raum_automatik import revit as rv

    ebene, phase = rv.ebene_und_phase(doc, uidoc.ActiveView)
    if ebene is None or phase is None:
        ui.meldung(t(u"Bitte einen Grundriss des Geschosses öffnen - die "
                     u"Räume entstehen auf seiner Ebene und in seiner Phase.",
                     u"Please open a floor plan of the level - the rooms are "
                     u"created on its level and in its phase.",
                     u"Abra una planta del nivel: las habitaciones se crean "
                     u"en su nivel y su fase."),
                   titel=TITEL, hauptzeile=t(u"Kein Grundriss aktiv",
                                             u"No floor plan active",
                                             u"Ninguna planta activa"),
                   warnung=True)
        return

    uhr = _Uhr()
    schluessel_ids = rv.schluessel(doc)
    uhr.schritt(u"Schlüssel Raumkategorie (%d)" % len(schluessel_ids))
    gegenstaende = rv.einrichtung(doc, ebene)
    uhr.schritt(u"Einrichtung (%d)" % len(gegenstaende))
    tueren = rv.tuer_punkte(doc, ebene)
    uhr.schritt(u"Türen (%d)" % len(tueren))

    gruppe = TransactionGroup(doc, TITEL)
    gruppe.Start()
    try:
        transaktion = Transaction(doc, TITEL)
        transaktion.Start()
        try:
            ergebnis = rv.lege_an(doc, ebene, phase, gegenstaende, tueren,
                                  schluessel_ids)
            uhr.schritt(u"Räume anlegen und prüfen (%d behalten, %d "
                        u"verworfen)" % (len(ergebnis.angelegt),
                                         ergebnis.verworfen))
            transaktion.Commit()
            uhr.schritt(u"Transaktion abschliessen")
            uhr.zeilen += [u"", u"Freie Bereiche (Fläche, Türen, Ergebnis, "
                                u"erkannte Gegenstände):"]
            uhr.zeilen += [u"%7.2f m²  %d Türen  %-10s  %s" % (
                b.flaeche, b.tueren, b.raumname() or u"-",
                u", ".join(sorted(set(b.inhalt))) or u"(nichts erkannt)")
                for b in sorted(ergebnis.bereiche, key=lambda b: -b.flaeche)]
            uhr.schreibe()
        except Exception:
            if transaktion.HasStarted() and not transaktion.HasEnded():
                transaktion.RollBack()
            raise
        if not ergebnis.angelegt:
            gruppe.RollBack()
            ui.meldung(t(u"Auf %s greift in keinem freien Bereich eine "
                         u"Regel (%d freie Bereiche, %d erkannte "
                         u"Einrichtungsgegenstände, %d Türen).",
                         u"On %s no rule matches any free area (%d free "
                         u"areas, %d furnishing items found, %d doors).",
                         u"En %s ninguna regla se aplica a una zona libre "
                         u"(%d zonas libres, %d objetos de mobiliario, "
                         u"%d puertas).")
                       % (ebene.Name, ergebnis.verworfen, len(gegenstaende),
                          len(tueren)),
                       titel=TITEL, hauptzeile=t(u"Keine Räume angelegt",
                                                 u"No rooms created",
                                                 u"No se crearon "
                                                 u"habitaciones"))
            return
        zeilen = _bericht(ergebnis)
        if not schluessel_ids:
            zeilen += [u"", t(u"Keine Schlüsselbauteilliste für "
                              u"\"Raumkategorie\" gefunden.",
                              u"No key schedule for \"Raumkategorie\" "
                              u"found.",
                              u"No se encontró la tabla de claves "
                              u"\"Raumkategorie\".")]
        zeilen += [u"", t(u"Die Räume sind in der Ansicht zu sehen. "
                          u"Behalten? Nein = alles zurücknehmen.",
                          u"The rooms are visible in the view. Keep them? "
                          u"No = undo everything.",
                          u"Las habitaciones se ven en la vista. "
                          u"¿Conservarlas? No = deshacer todo.")]
        uidoc.RefreshActiveView()
        behalten = ui.frage(u"\n".join(zeilen), titel=TITEL,
                            standard_ja=True,
                            warnung=bool(ergebnis.fehler
                                         or ergebnis.fehlende_parameter
                                         or ergebnis.ohne_schluessel),
                            hauptzeile=t(u"Räume übernehmen? (%s)",
                                         u"Keep the rooms? (%s)",
                                         u"¿Conservar las habitaciones? (%s)")
                            % ebene.Name)
    except Exception:
        if gruppe.HasStarted() and not gruppe.HasEnded():
            gruppe.RollBack()
        raise
    if behalten:
        gruppe.Assimilate()
    else:
        gruppe.RollBack()


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "AutoRooms_Fehler.log")
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
                   hauptzeile=t(u"Räume automatisch ist fehlgeschlagen",
                                u"Auto Rooms failed",
                                u"Habitaciones automáticas ha fallado"),
                   warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
