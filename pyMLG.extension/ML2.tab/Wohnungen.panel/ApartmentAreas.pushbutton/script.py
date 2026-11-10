#! python3
# -*- coding: utf-8 -*-
"""ApartmentAreas: In jeden Eingangspfeil (Familie aus ApartmentNumbering)
die Flächen der Räume seiner Gruppe schreiben, summiert nach
Raumkategorie: WNFL, Loggia, Balkon, Terrassen, Gewerbe-NFL.

Die Logik liegt in lib/wohnung_flaechen (Regeln: logik.py, Revit:
revit.py). Geschrieben wird nur in die Pfeile.

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

TITEL = t(u"Wohnungsflächen", u"Apartment Areas", u"Superficies de vivienda")

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


def _auswahl(doc, ids):
    """Gewählte Elemente samt den Mitgliedern gewählter Gruppen."""
    from Autodesk.Revit.DB import Group
    ergebnis, offen = [], list(ids)
    while offen:
        i = offen.pop()
        ergebnis.append(i)
        element = doc.GetElement(i)
        if isinstance(element, Group):
            offen.extend(element.GetMemberIds())
    return ergebnis


def _parameter_bereit(lauf, liste):
    """(Namen zum Schreiben, fehlende Namen) oder None = abbrechen."""
    from schedule_sync import ui
    from wohnung_flaechen import logik as lg
    from wohnung_flaechen import revit as rv

    fehlend, fest = rv.pruefe_parameter([p.element for p in liste], lg.ALLE)
    namen = [n for n in lg.ALLE if n not in fehlend]
    if not namen:
        _warnung(t(u"Parameter fehlen", u"Parameters missing",
                   u"Faltan parámetros"),
                 t(u"Keinen der Parameter gibt es an den Eingangspfeilen:",
                   u"The entrance arrows have none of the parameters:",
                   u"Las flechas de entrada no tienen ninguno de los "
                   u"parámetros:") + u"\n" + _liste(list(lg.ALLE)))
        return None
    if not fest:
        return namen, fehlend
    namen_fest = sorted(fest)
    if not ui.frage(
            t(u"Bei diesen Parametern ist 'Werte je nach Gruppeninstanz "
              u"unterschiedlich' aus:",
              u"'Values can vary by group instance' is off for these "
              u"parameters:",
              u"'Los valores pueden variar por instancia de grupo' está "
              u"desactivado en estos parámetros:")
            + u"\n" + _liste(namen_fest) + u"\n\n"
            + t(u"Ohne diese Einstellung können Eingangspfeile in Gruppen "
                u"nicht je Wohnung verschiedene Flächen haben. Jetzt "
                u"einschalten? Nein = abbrechen, nichts wird geändert.",
                u"Without this setting, entrance arrows in groups cannot "
                u"hold different areas per apartment. Turn it on now? "
                u"No = cancel, nothing is changed.",
                u"Sin este ajuste, las flechas de entrada en grupos no "
                u"pueden tener superficies distintas por vivienda. "
                u"¿Activarlo ahora? No = cancelar, no se cambia nada."),
            titel=TITEL, standard_ja=True,
            hauptzeile=t(u"Werte je Gruppeninstanz erlauben?",
                         u"Allow values per group instance?",
                         u"¿Permitir valores por instancia de grupo?")):
        return None
    lauf.erlaube_varianz([fest[n] for n in namen_fest])
    return namen, fehlend


def _abschluss(lauf, liste, ohne_gruppe, namen, fehlend):
    """Zusammenfassung zeigen, übernehmen oder verwerfen."""
    from schedule_sync import ui
    from wohnung_flaechen import logik as lg
    from wohnung_flaechen import revit as rv

    gesamt = dict((n, sum(p.werte[n] for p in liste)) for n in namen)
    zeilen = [t(u"%d Eingangspfeile gefüllt. Summe aller Pfeile:",
                u"%d entrance arrows filled. Total of all arrows:",
                u"%d flechas de entrada rellenadas. Suma de todas:")
              % lauf.pfeile]
    zeilen += [u"   %s:  %s" % (n, lg.qm_text(gesamt[n])) for n in namen]
    beispiel = next((p for p in liste if any(p.werte[n] for n in namen)),
                    liste[0])
    zeilen += [u"", t(u"Beispiel %s:", u"Example %s:", u"Ejemplo %s:")
               % beispiel.beschreibung()]
    zeilen += [u"   %s:  %s" % (n, lg.qm_text(beispiel.werte[n]))
               for n in namen if beispiel.werte[n]]
    if fehlend:
        zeilen += [u"", t(u"Nicht am Pfeil vorhanden, daher nicht gefüllt: ",
                          u"Not on the arrow, so not filled: ",
                          u"No existen en la flecha, por eso no se "
                          u"rellenan: ") + u", ".join(fehlend)]
    # Nicht gezählte Raumkategorien - je Gruppe nur einmal
    ungezaehlt, gesehen = {}, set()
    for p in liste:
        schluessel = rv.id_wert(p.gruppe.Id)
        if schluessel in gesehen:
            continue
        gesehen.add(schluessel)
        for k, z in p.ungezaehlt.items():
            ungezaehlt[k] = ungezaehlt.get(k, 0) + z
    if ungezaehlt:
        zeilen += [u"", t(u"Nicht gezählte Räume (Raumkategorie):",
                          u"Rooms not counted (room category):",
                          u"Habitaciones no contadas (categoría):"),
                   _liste([u"%s (%d)" % e for e in
                           sorted(ungezaehlt.items())])]
    warnung = False
    leer = [p.beschreibung() for p in liste if p.raeume == 0]
    if leer:
        warnung = True
        zeilen += [u"", t(u"Gruppe ohne Räume - alles 0:",
                          u"Group without rooms - all 0:",
                          u"Grupo sin habitaciones: todo 0:"), _liste(leer)]
    if ohne_gruppe:
        warnung = True
        zeilen += [u"", t(u"Übersprungen - Pfeil in keiner Gruppe:",
                          u"Skipped - arrow not in a group:",
                          u"Omitidas: flecha fuera de grupo:"),
                   _liste([u"Id %d" % rv.id_wert(e.Id)
                           for e in ohne_gruppe])]
    if lauf.fehler:
        warnung = True
        anzahl = {}
        for _, name, grund in lauf.fehler:
            anzahl[(name, grund)] = anzahl.get((name, grund), 0) + 1
        zeilen += [u"", t(u"Nicht geschrieben:", u"Not written:",
                          u"No escritos:"),
                   _liste([u"%s: %s (%d)" % (n, g, z)
                           for (n, g), z in sorted(anzahl.items())])]
    zeilen += [u"", t(u"Nein = alles zurücknehmen.",
                      u"No = undo everything.",
                      u"No = deshacer todo.")]
    try:
        behalten = ui.frage(u"\n".join(zeilen), titel=TITEL,
                            standard_ja=True, warnung=warnung,
                            hauptzeile=t(u"Werte übernehmen?",
                                         u"Keep the values?",
                                         u"¿Conservar los valores?"))
    except Exception:
        lauf.verwerfen()
        raise
    if behalten:
        lauf.uebernehmen()
    else:
        lauf.verwerfen()


def main():
    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        _warnung(t(u"Kein Projekt geöffnet", u"No project open",
                   u"No hay ningún proyecto abierto"),
                 t(u"Wohnungsflächen geht nur in Projektdateien.",
                   u"Apartment Areas only works in project files.",
                   u"Superficies de vivienda solo funciona en archivos de "
                   u"proyecto."))
        return
    if doc.IsReadOnly:
        _warnung(t(u"Dokument schreibgeschützt", u"Document read-only",
                   u"Documento de solo lectura"),
                 t(u"Das Dokument ist schreibgeschützt.",
                   u"The document is read-only.",
                   u"El documento es de solo lectura."))
        return

    from wohnung_flaechen import revit as rv
    from wohnung_top import logik as top

    # Vorher gewählte Pfeile/Gruppen: nur diese - sonst alle
    auswahl = list(uidoc.Selection.GetElementIds())
    liste, ohne_gruppe = [], []
    if auswahl:
        liste, ohne_gruppe = rv.pfeile(doc, _auswahl(doc, auswahl))
    if not liste and not ohne_gruppe:
        liste, ohne_gruppe = rv.pfeile(doc)
    if not liste:
        _warnung(t(u"Keine Eingangspfeile", u"No entrance arrows",
                   u"No hay flechas de entrada"),
                 t(u"Es gibt keinen Eingangspfeil '%s' in einer Gruppe.",
                   u"There is no entrance arrow '%s' in a group.",
                   u"No hay ninguna flecha de entrada '%s' en un grupo.")
                 % top.MARKER_FAMILIE)
        return

    lauf = rv.Lauf(doc)
    try:
        bereit = _parameter_bereit(lauf, liste)
        if bereit is None:
            lauf.verwerfen()
            return
        namen, fehlend = bereit
        lauf.schreibe(liste, namen)
    except rv.GruppenFehler as fehler:
        lauf.verwerfen()
        _warnung(t(u"Revit hat das Schreiben abgelehnt - nichts geändert",
                   u"Revit rejected the change - nothing was changed",
                   u"Revit ha rechazado el cambio; no se ha cambiado nada"),
                 u"%s" % fehler)
        return
    except Exception:
        lauf.verwerfen()
        raise
    _abschluss(lauf, liste, ohne_gruppe, namen, fehlend)


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "ApartmentAreas_Fehler.log")
    try:
        if not os.path.isdir(os.path.dirname(protokoll)):
            os.makedirs(os.path.dirname(protokoll))
        with io.open(protokoll, "a", encoding="utf-8") as datei:
            datei.write(spur + u"\n" + u"-" * 70 + u"\n")
    except Exception:
        protokoll = None
    try:
        letzte_zeile = (spur.strip().splitlines() or [u""])[-1]
        _warnung(t(u"Wohnungsflächen ist fehlgeschlagen",
                   u"Apartment Areas failed",
                   u"Superficies de vivienda ha fallado"),
                 letzte_zeile + (t(u"\n\nProtokoll: %s", u"\n\nLog: %s",
                                   u"\n\nRegistro: %s") % protokoll
                                 if protokoll else u""))
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
