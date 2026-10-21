#! python3
# -*- coding: utf-8 -*-
"""ApartmentNumbering: Wohnungsgruppen (WHG_TYP ...) entlang einer Linie
über mehrere Geschosse nummerieren (TOP) und die Wohnungsparameter ihrer
Räume aus dem Gruppentypnamen füllen.

Die Logik liegt in lib/wohnung_top (Fenster: fenster.py, Revit: revit.py,
Namen und Geometrie: logik.py). Linie und Pfad kommen von ColumnNumbering
(lib/stuetzen_nummer).

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

TITEL = t(u"Wohnungen nummerieren", u"Number Apartments",
          u"Numerar viviendas")

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


def _linien(uidoc):
    """Vorausgewählte Linien oder ein Klick auf ein Stück der Linie - die
    damit verbundenen Stücke kommen von selbst dazu."""
    from stuetzen_nummer import revit as sn
    from wohnung_top import revit as rv

    doc = uidoc.Document
    linien = sn.gewaehlte_linien(doc, uidoc.Selection.GetElementIds())
    if not linien:
        linie = sn.waehle_linie(uidoc, t(
            u"Linie durch die Wohnungen anklicken",
            u"Pick the line through the apartments",
            u"Seleccione la línea a través de las viviendas"))
        if linie is None:
            return None
        linien = [linie]
    return rv.linienzug(doc, linien)


class Plan(object):
    """Was der Lauf schreiben wird."""

    def __init__(self):
        self.auftraege = []        # [(Wohnung, {Parameter: Wert})]
        self.nummeriert = []       # [(Geschossname, TOP)]
        self.falscher_name = []    # [Wohnung]
        self.ohne_raeume = []      # [Wohnung]
        self.daneben = 0           # Wohnungen der Geschosse ohne Linie
        self.fehlend = []          # Parameter, die es an Räumen nicht gibt


def _plane(wohnungen, einst, segmente):
    from stuetzen_nummer import logik as nummern
    from wohnung_top import logik as lg
    from wohnung_top import revit as rv

    ebenen = dict((e.schluessel, e) for e in einst.ebenen)
    plan = Plan()
    treffer = []
    for w in wohnungen:
        ebene = ebenen.get(rv.id_wert(w.ebene_id))
        if ebene is None:
            continue
        station = lg.eintritt_mehrere(segmente, w.flaechen())
        if station is None:
            plan.daneben += 1
        elif lg.werte(w.name) is None:
            plan.falscher_name.append(w)
        elif not w.raeume:
            plan.ohne_raeume.append(w)
        else:
            treffer.append((ebene.hoehe, station, (ebene, w)))
    zaehler = nummern.Zaehler(einst.start, 1)
    for ebene, w in lg.reihenfolge(treffer):
        werte = lg.werte(w.name, einst.bauteil, einst.stiege)
        werte[lg.TOP] = zaehler.weiter()
        plan.auftraege.append((w, werte))
        plan.nummeriert.append((ebene.name, werte[lg.TOP]))
    return plan


def _parameter_bereit(lauf, plan):
    """Fehlende Parameter melden, fehlende Gruppenvarianz nach Rückfrage
    einschalten. False = abbrechen."""
    from schedule_sync import ui
    from wohnung_top import logik as lg
    from wohnung_top import revit as rv

    raeume = [r for w, _ in plan.auftraege for r in w.raeume]
    namen = [n for n in lg.ALLE
             if any(n in werte for _, werte in plan.auftraege)]
    fehlend, fest = rv.pruefe_parameter(raeume, namen)
    # Parameter, die es an den Räumen nicht gibt, werden nicht gefüllt
    plan.fehlend = fehlend
    for _, werte in plan.auftraege:
        for name in fehlend:
            werte.pop(name, None)
    if len(fehlend) == len(namen):
        _warnung(t(u"Parameter fehlen", u"Parameters missing",
                   u"Faltan parámetros"),
                 t(u"Keinen der Parameter gibt es an den Räumen:",
                   u"The rooms have none of the parameters:",
                   u"Las habitaciones no tienen ninguno de los "
                   u"parámetros:")
                 + u"\n" + _liste(fehlend))
        return False
    if not fest:
        return True
    namen_fest = sorted(fest)
    if not ui.frage(
            t(u"Bei diesen Parametern ist 'Werte je nach Gruppeninstanz "
              u"unterschiedlich' aus:",
              u"'Values can vary by group instance' is off for these "
              u"parameters:",
              u"'Los valores pueden variar por instancia de grupo' está "
              u"desactivado en estos parámetros:")
            + u"\n" + _liste(namen_fest) + u"\n\n"
            + t(u"Ohne diese Einstellung können Räume in Gruppen nicht je "
                u"Wohnung verschiedene Werte haben (z. B. TOP). Jetzt "
                u"einschalten? Nein = abbrechen, nichts wird geändert.",
                u"Without this setting, rooms in groups cannot hold "
                u"different values per apartment (e.g. TOP). Turn it on "
                u"now? No = cancel, nothing is changed.",
                u"Sin este ajuste, las habitaciones en grupos no pueden "
                u"tener valores distintos por vivienda (p. ej. TOP). "
                u"¿Activarlo ahora? No = cancelar, no se cambia nada."),
            titel=TITEL, standard_ja=True,
            hauptzeile=t(u"Werte je Gruppeninstanz erlauben?",
                         u"Allow values per group instance?",
                         u"¿Permitir valores por instancia de grupo?")):
        return False
    lauf.erlaube_varianz([fest[n] for n in namen_fest])
    return True


def _abschluss(lauf, plan, einst):
    """Zusammenfassung zeigen, übernehmen oder verwerfen."""
    from schedule_sync import ui
    from wohnung_top import logik as lg

    zeilen = [t(u"%d Wohnungen, %d Räume:", u"%d apartments, %d rooms:",
                u"%d viviendas, %d habitaciones:")
              % (len(plan.auftraege), lauf.raeume)]
    for geschoss, erste, letzte in lg.bereiche(plan.nummeriert):
        zeilen.append(u"   %s:  TOP %s" % (geschoss, erste) if erste == letzte
                      else u"   %s:  TOP %s … %s" % (geschoss, erste, letzte))
    geschrieben = plan.auftraege[0][1]
    namen = [u"%s = %s" % (n, geschrieben[n])
             if n in (lg.BAUTEIL, lg.STIEGE) else n
             for n in lg.ALLE if n in geschrieben]
    zeilen += [u"", t(u"Geschrieben: ", u"Written: ", u"Escritos: ")
               + u", ".join(namen)]
    beispiel = lg.werte(plan.auftraege[0][0].name)
    zeilen.append(t(u"Beispiel: %s -> %s, %s, %s",
                    u"Example: %s -> %s, %s, %s",
                    u"Ejemplo: %s -> %s, %s, %s")
                  % (beispiel[lg.GRUPPE_TYP], beispiel[lg.TYPNAME],
                     beispiel[lg.KATEGORIE], beispiel[lg.GESCHOSS]))
    zeilen.append(t(u"Raumnummer nach Tabelle (Vorraum 001, Bad 003 ...): "
                    u"%d Räume", u"Room number from the table (Vorraum 001, "
                    u"Bad 003 ...): %d rooms", u"Número de habitación según "
                    u"la tabla (Vorraum 001, Bad 003 ...): %d habitaciones")
                  % lauf.nummern)
    if lauf.ohne_nummer:
        zeilen.append(t(u"Ohne Nummer (nicht in der Tabelle): ",
                        u"Without number (not in the table): ",
                        u"Sin número (no están en la tabla): ")
                      + u", ".join(u"%s (%d)" % e for e in
                                   sorted(lauf.ohne_nummer.items())))
    if plan.fehlend:
        zeilen.append(t(u"Nicht vorhanden, daher nicht gefüllt: ",
                        u"Not present, so not filled: ",
                        u"No existen, por eso no se rellenan: ")
                      + u", ".join(plan.fehlend))
    warnung = False
    if plan.falscher_name:
        warnung = True
        zeilen += [u"", t(u"Übersprungen - Name passt nicht zu "
                          u"'WHG_TYP A2a - TU_OG07-OG12':",
                          u"Skipped - name does not match "
                          u"'WHG_TYP A2a - TU_OG07-OG12':",
                          u"Omitidos: el nombre no corresponde a "
                          u"'WHG_TYP A2a - TU_OG07-OG12':"),
                   _liste([w.beschreibung() for w in plan.falscher_name])]
    if plan.ohne_raeume:
        warnung = True
        zeilen += [u"", t(u"Übersprungen - keine Räume in der Gruppe:",
                          u"Skipped - no rooms in the group:",
                          u"Omitidos: el grupo no tiene habitaciones:"),
                   _liste([w.beschreibung() for w in plan.ohne_raeume])]
    if plan.daneben:
        zeilen += [u"", t(u"%d Wohnungsgruppen der gewählten Geschosse "
                          u"liegen nicht auf der Linie.",
                          u"%d apartment groups on the chosen levels are not "
                          u"on the line.",
                          u"%d grupos de vivienda de los niveles elegidos no "
                          u"están en la línea.") % plan.daneben]
    if lauf.fehler:
        warnung = True
        anzahl = {}
        for _, _, name, grund in lauf.fehler:
            anzahl[(name, grund)] = anzahl.get((name, grund), 0) + 1
        zeilen += [u"", t(u"Nicht geschrieben:", u"Not written:",
                          u"No escritos:"),
                   _liste([t(u"%s: %s (%d Räume)", u"%s: %s (%d rooms)",
                             u"%s: %s (%d habitaciones)") % (n, g, z)
                           for (n, g), z in sorted(anzahl.items())])]
    zeilen += [u"", t(u"Die TOP-Nummern des aktiven Grundrisses stehen als "
                      u"Vorschau in der Ansicht. Nein = alles zurücknehmen.",
                      u"The TOP numbers of the active floor plan are shown "
                      u"as a preview in the view. No = undo everything.",
                      u"Los números TOP de la planta activa se muestran como "
                      u"vista previa. No = deshacer todo.")]
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
                 t(u"Wohnungen nummerieren geht nur in Projektdateien.",
                   u"Numbering apartments only works in project files.",
                   u"Numerar viviendas solo funciona en archivos de "
                   u"proyecto."))
        return
    if doc.IsReadOnly:
        _warnung(t(u"Wohnungen können nicht nummeriert werden",
                   u"Apartments cannot be numbered",
                   u"No se pueden numerar las viviendas"),
                 t(u"Das Dokument ist schreibgeschützt.",
                   u"The document is read-only.",
                   u"El documento es de solo lectura."))
        return

    from stuetzen_nummer import logik as nummern
    from stuetzen_nummer import revit as sn
    from wohnung_top import fenster
    from wohnung_top import revit as rv

    wohnungen = rv.wohnungen(doc)
    if not wohnungen:
        _warnung(t(u"Keine Wohnungsgruppen gefunden",
                   u"No apartment groups found",
                   u"No se han encontrado grupos de vivienda"),
                 t(u"Es gibt keine Modellgruppe, deren Typname mit WHG_TYP "
                   u"beginnt.",
                   u"There is no model group whose type name starts with "
                   u"WHG_TYP.",
                   u"No hay ningún grupo de modelo cuyo nombre de tipo "
                   u"empiece por WHG_TYP."))
        return

    einst = fenster.frage_einstellungen(__revit__, rv.ebenen(doc, wohnungen))
    if einst is None:
        return
    linien = _linien(uidoc)
    if not linien:
        return

    segmente = nummern.segmente(sn.pfad(linien))
    plan = _plane(wohnungen, einst, segmente)
    if not plan.auftraege:
        text = t(u"Die Linie geht auf den gewählten Geschossen durch keine "
                 u"passende Wohnungsgruppe (es zählen die Räume der Gruppe).",
                 u"On the chosen levels the line does not pass through any "
                 u"matching apartment group (the rooms of the group count).",
                 u"En los niveles elegidos la línea no atraviesa ningún "
                 u"grupo de vivienda válido (cuentan las habitaciones del "
                 u"grupo).")
        uebersprungen = plan.falscher_name + plan.ohne_raeume
        if uebersprungen:
            text += u"\n\n" + t(u"Übersprungen:", u"Skipped:",
                                u"Omitidos:") + u"\n" + _liste(
                [w.beschreibung() for w in uebersprungen])
        _warnung(t(u"Keine Wohnungen auf der Linie",
                   u"No apartments on the line",
                   u"No hay viviendas en la línea"), text)
        return

    lauf = rv.Lauf(uidoc)
    try:
        if not _parameter_bereit(lauf, plan):
            lauf.verwerfen()
            return
        lauf.schreibe(plan.auftraege, rv.vorschau_ebene(doc.ActiveView))
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
    _abschluss(lauf, plan, einst)


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "ApartmentNumbering_Fehler.log")
    try:
        if not os.path.isdir(os.path.dirname(protokoll)):
            os.makedirs(os.path.dirname(protokoll))
        with io.open(protokoll, "a", encoding="utf-8") as datei:
            datei.write(spur + u"\n" + u"-" * 70 + u"\n")
    except Exception:
        protokoll = None
    try:
        letzte_zeile = (spur.strip().splitlines() or [u""])[-1]
        _warnung(t(u"Wohnungen nummerieren ist fehlgeschlagen",
                   u"Number Apartments failed",
                   u"Numerar viviendas ha fallado"),
                 letzte_zeile + (t(u"\n\nProtokoll: %s", u"\n\nLog: %s",
                                   u"\n\nRegistro: %s") % protokoll
                                 if protokoll else u""))
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
