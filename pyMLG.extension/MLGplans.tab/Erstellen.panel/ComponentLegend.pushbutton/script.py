#! python3
# -*- coding: utf-8 -*-
"""ComponentLegend: Legende aus den Familientypen gewählter Ansichten -
ein Legendenbauteil je Typ, untereinander, in der gewählten Ansichtsrichtung.

Die Logik liegt in lib/bauteil_legende (logik.py, revit.py, fenster.py).

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

TITEL = t(u"Legende aus Ansichten", u"Legend from views", u"Leyenda desde vistas")

# so viele Typen werden im Bericht einzeln genannt
NAMEN_IM_BERICHT = 15


def _typliste(typen):
    zeilen = [u"• %s: %s" % (typ.kategorie, typ.familie + u": " + typ.name
                             if typ.familie and typ.familie != typ.name else typ.name)
              for typ in typen[:NAMEN_IM_BERICHT]]
    if len(typen) > NAMEN_IM_BERICHT:
        zeilen.append(t(u"… und %d weitere", u"… and %d more", u"… y %d más")
                      % (len(typen) - NAMEN_IM_BERICHT))
    return zeilen


def _bericht(ergebnis):
    zeilen = [t(u"%d Legendenbauteile platziert, %d Beschriftungen",
                u"%d legend components placed, %d labels",
                u"%d componentes de leyenda colocados, %d etiquetas")
              % (ergebnis.platziert, ergebnis.texte)]
    if ergebnis.ketten or ergebnis.ketten_fehler:
        zeilen.append(t(u"%d Maßketten", u"%d dimensions", u"%d cotas") % ergebnis.ketten
                      + (t(u" (%d nicht möglich)", u" (%d not possible)",
                           u" (%d no posibles)") % ergebnis.ketten_fehler
                         if ergebnis.ketten_fehler else u""))
    if ergebnis.ersetzt:
        zeilen.insert(0, t(u"Vorhandene Legende ersetzt (%d alte Elemente gelöscht).",
                           u"Existing legend replaced (%d old elements deleted).",
                           u"Leyenda existente reemplazada (%d elementos antiguos borrados).")
                      % ergebnis.geloescht)
    if ergebnis.abgelehnt:
        zeilen.append(u"")
        zeilen.append(t(u"Revit kann diese Typen nicht als Legendenbauteil zeigen:",
                        u"Revit cannot show these types as legend components:",
                        u"Revit no puede mostrar estos tipos como componente de leyenda:"))
        zeilen.extend(_typliste(ergebnis.abgelehnt))
    if ergebnis.ohne_richtung:
        zeilen.append(u"")
        zeilen.append(t(u"Gewählte Ansichtsrichtung gibt es für diese Typen nicht - "
                        u"im Grundriss gezeigt:",
                        u"The chosen view direction does not exist for these types - "
                        u"shown in floor plan:",
                        u"La dirección de vista elegida no existe para estos tipos - "
                        u"se muestran en planta:"))
        zeilen.extend(_typliste(ergebnis.ohne_richtung))
    return u"\n".join(zeilen)


def _starte_ersteinrichtung(uidoc, ui):
    """Ohne Legendenbauteil im Projekt: Legende öffnen und Revits Befehl
    "Legendenbauteil" starten (bzw. "Legende", wenn es noch keine gibt).
    Die API kann beides nicht selbst anlegen - ein Klick genügt, danach
    platziert das Tool alle Bauteile."""
    from Autodesk.Revit.UI import PostableCommand, RevitCommandId
    from bauteil_legende import revit as rv

    legenden = rv.dxf_rv.legenden(uidoc.Document)
    if legenden:
        befehl = PostableCommand.LegendComponent
        hauptzeile = t(u"Einmalig: ein Legendenbauteil klicken",
                       u"One time: click one legend component",
                       u"Una sola vez: haga clic en un componente de leyenda")
        text = t(u"Revit kann über die API kein allererstes Legendenbauteil anlegen, "
                 u"nur vorhandene kopieren. Nach \"Schliessen\" öffnet das Tool die "
                 u"Legende \"%s\" und startet \"Legendenbauteil\".\n\n"
                 u"1. Einmal irgendwo in die Legende klicken (Typ egal), Esc.\n"
                 u"2. Diesen Button erneut drücken - ab dann setzt das Tool alle "
                 u"Familien automatisch.\n\n"
                 u"Das Bauteil darf danach bleiben oder gelöscht werden, sobald die "
                 u"neue Legende es enthält.",
                 u"Revit's API cannot create the very first legend component, it can "
                 u"only copy existing ones. After \"Close\" the tool opens the legend "
                 u"\"%s\" and starts \"Legend Component\".\n\n"
                 u"1. Click once anywhere in the legend (any type), Esc.\n"
                 u"2. Press this button again - from then on the tool places all "
                 u"families automatically.\n\n"
                 u"The component may stay or be deleted once the new legend "
                 u"contains one.",
                 u"La API de Revit no puede crear el primer componente de leyenda, "
                 u"solo copiar los existentes. Tras \"Cerrar\" la herramienta abre la "
                 u"leyenda \"%s\" e inicia \"Componente de leyenda\".\n\n"
                 u"1. Haga clic una vez en la leyenda (tipo indiferente), Esc.\n"
                 u"2. Pulse este botón otra vez - a partir de ahí la herramienta "
                 u"coloca todas las familias automáticamente.\n\n"
                 u"El componente puede quedarse o borrarse en cuanto la nueva "
                 u"leyenda contenga uno.") % rv.name_von(legenden[0])
    else:
        befehl = PostableCommand.Legend
        hauptzeile = t(u"Einmalig: eine Legende anlegen",
                       u"One time: create a legend",
                       u"Una sola vez: cree una leyenda")
        text = t(u"Im Projekt gibt es noch keine Legende, und die API kann keine "
                 u"anlegen. Nach \"Schliessen\" startet das Tool den Revit-Befehl "
                 u"\"Legende\".\n\n"
                 u"1. Im Revit-Dialog mit OK bestätigen.\n"
                 u"2. Diesen Button erneut drücken - dann folgt noch ein Klick für "
                 u"das erste Legendenbauteil, danach läuft alles automatisch.",
                 u"The project has no legend yet, and the API cannot create one. "
                 u"After \"Close\" the tool starts Revit's \"Legend\" command.\n\n"
                 u"1. Confirm the Revit dialog with OK.\n"
                 u"2. Press this button again - one more click for the first legend "
                 u"component follows, then everything runs automatically.",
                 u"El proyecto aún no tiene ninguna leyenda y la API no puede crearla. "
                 u"Tras \"Cerrar\" la herramienta inicia el comando \"Leyenda\" de "
                 u"Revit.\n\n"
                 u"1. Confirme el diálogo de Revit con Aceptar.\n"
                 u"2. Pulse este botón otra vez - sigue un clic para el primer "
                 u"componente de leyenda y luego todo es automático.")
    ui.meldung(text, titel=TITEL, hauptzeile=hauptzeile)
    if legenden:
        uidoc.ActiveView = legenden[0]
    __revit__.PostCommand(RevitCommandId.LookupPostableCommandId(befehl))


def main():
    from schedule_sync import ui

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        ui.meldung(t(u"Legenden gibt es nur in Projektdateien.",
                     u"Legends only exist in project files.",
                     u"Las leyendas solo existen en archivos de proyecto."),
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
                   hauptzeile=t(u"Legende kann nicht erstellt werden",
                                u"Legend cannot be created",
                                u"No se puede crear la leyenda"),
                   warnung=True)
        return

    from bauteil_legende import revit as rv
    if rv.vorlage_bauteil(doc) is None:
        _starte_ersteinrichtung(uidoc, ui)
        return

    from bauteil_legende import fenster
    auswahl = fenster.starte(__revit__, doc, uidoc.ActiveView)
    if not auswahl:
        return

    try:
        ergebnis = rv.erzeuge(doc, auswahl[u"typen"], auswahl[u"name"],
                              auswahl[u"massstab"], auswahl[u"abstand"],
                              beschriftung=auswahl[u"beschriftung"],
                              texttyp_id=auswahl[u"texttyp"],
                              ersetzen=auswahl[u"ersetzen"],
                              richtung=auswahl[u"richtung"],
                              masse=auswahl[u"masse"],
                              mass_art=auswahl[u"mass_art"])
    except rv.LegendenFehler as fehler:
        ui.meldung(u"%s" % fehler, titel=TITEL,
                   hauptzeile=t(u"Legende nicht erstellt", u"Legend not created",
                                u"Leyenda no creada"),
                   warnung=True)
        return

    try:
        uidoc.ActiveView = ergebnis.ansicht
    except Exception:
        pass
    ui.meldung(_bericht(ergebnis), titel=TITEL,
               hauptzeile=t(u"\"%s\" erstellt", u"\"%s\" created", u"\"%s\" creada")
               % ergebnis.ansicht.Name,
               warnung=bool(ergebnis.abgelehnt))


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "ComponentLegend_Fehler.log")
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
                   hauptzeile=t(u"Legende aus Ansichten ist fehlgeschlagen",
                                u"Legend from views failed",
                                u"Leyenda desde vistas ha fallado"),
                   warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
