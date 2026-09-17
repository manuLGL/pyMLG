#! python3
# -*- coding: utf-8 -*-
"""Importiert bearbeitete Excel-Werte zurück in das Revit-Modell.

Zwei Besonderheiten der CPython-Engine, die den Aufbau dieser Datei erklären:

1. pyrevit.forms ist unter CPython nicht verfügbar - jeder Zugriff darauf wirft
   PyRevitCPythonNotSupported. Die Dialoge kommen deshalb aus schedule_sync.ui.
2. script.exit() ist nichts anderes als sys.exit(). Das dabei ausgelöste
   SystemExit verlässt den CPython-Host und endet in Revit als nichtssagendes
   "Object reference not set to an instance of an object". Deshalb steckt der
   gesamte Ablauf in main() und beendet sich ausschliesslich über 'return'.

Damit auch Fehler beim Import der Bibliotheken sichtbar werden, liegen diese
Importe bewusst innerhalb von main() und damit innerhalb des Fehlerfangs.
"""

import io
import os
import sys
import traceback

# Absicherung, falls pyRevit den lib-Ordner nicht bereits in sys.path gelegt hat
_EXT = os.path.dirname(os.path.abspath(__file__))
while not _EXT.endswith(".extension") and os.path.dirname(_EXT) != _EXT:
    _EXT = os.path.dirname(_EXT)
if os.path.join(_EXT, "lib") not in sys.path:
    sys.path.append(os.path.join(_EXT, "lib"))

from pyrevit import script

output = script.get_output()


def main():
    from collections import OrderedDict

    from Autodesk.Revit.DB import (
        ElementId,
        Transaction,
        TransactionGroup,
        TransactionStatus,
    )

    from schedule_sync import deps, ui
    from schedule_sync.excel_io import lese_arbeitsmappe
    from schedule_sync.importer import Bericht, analysiere, wende_an

    # Ohne geöffnetes Projekt ist ActiveUIDocument null. Ein Zugriff darauf
    # ergäbe eine NullReferenceException, die Revit nur als "Object reference
    # not set to an instance of an object" anzeigt - deshalb hier abfangen.
    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None:
        ui.meldung(u"Es ist kein Projekt geöffnet. Bitte zuerst eine "
                   u"Revit-Projektdatei öffnen.",
                   hauptzeile=u"Kein aktives Dokument", warnung=True)
        return
    if doc.IsFamilyDocument:
        ui.meldung(u"Bauteillisten gibt es nur in Projektdateien, nicht im "
                   u"Familieneditor.",
                   hauptzeile=u"Familiendokument wird nicht unterstützt",
                   warnung=True)
        return

    def elementlink(element_id):
        """Klickbarer Link auf ein Element in der Ausgabe (mit Fallback)."""
        if element_id is None:
            return u"-"
        try:
            return output.linkify(ElementId(element_id))
        except Exception:
            return str(element_id)

    # -----------------------------------------------------------------------
    # 1. Vorbedingungen
    # -----------------------------------------------------------------------
    if doc.IsReadOnly:
        ui.meldung(u"Das aktive Dokument ist schreibgeschützt geöffnet. "
                   u"Ein Import ist nicht möglich.",
                   hauptzeile=u"Import nicht möglich", warnung=True)
        return

    try:
        openpyxl = deps.lade_openpyxl()
    except deps.AbhaengigkeitFehlt as fehler:
        ui.meldung(str(fehler), hauptzeile=u"openpyxl fehlt", warnung=True)
        return

    # -----------------------------------------------------------------------
    # 2. Excel-Datei wählen und lesen
    # -----------------------------------------------------------------------
    quelldatei = ui.datei_oeffnen(
        startordner=os.path.dirname(doc.PathName) if doc.PathName else u"",
        titel=u"Bearbeitete Excel-Datei auswählen",
    )
    if not quelldatei:
        return

    # Bei OneDrive/SharePoint kann die lokale Datei älter sein als die online
    # bearbeitete Fassung. Der Zeitstempel macht sofort sichtbar, welcher Stand
    # tatsächlich eingelesen wurde.
    try:
        import datetime
        stand = datetime.datetime.fromtimestamp(
            os.path.getmtime(quelldatei)).strftime(u"%d.%m.%Y %H:%M:%S")
    except Exception:
        stand = u"unbekannt"

    try:
        blaetter, lese_warnungen = lese_arbeitsmappe(openpyxl, quelldatei)
    except PermissionError:
        ui.meldung(u"Vermutlich ist die Datei gerade in Excel geöffnet. "
                   u"Bitte schliessen und erneut versuchen.\n\n%s" % quelldatei,
                   hauptzeile=u"Datei konnte nicht gelesen werden", warnung=True)
        return
    except Exception as fehler:
        ui.meldung(u"%s" % fehler,
                   hauptzeile=u"Die Excel-Datei konnte nicht gelesen werden",
                   warnung=True)
        return

    if not blaetter:
        text = u"\n".join(u"• %s" % warnung for warnung in lese_warnungen)
        ui.meldung(u"Erwartet wird ein Blatt, dessen Zelle A1 'UniqueId' "
                   u"enthält (so wie vom Export erzeugt).\n\n%s" % text,
                   hauptzeile=u"Kein auswertbares Tabellenblatt gefunden",
                   warnung=True)
        return

    # -----------------------------------------------------------------------
    # 3. Analyse (liest nur, schreibt noch nichts)
    # -----------------------------------------------------------------------
    bericht = Bericht()
    gesehen = {}

    with ui.Fortschritt(output, len(blaetter)) as fortschritt:
        for index, blatt in enumerate(blaetter):
            fortschritt.aktualisiere(index + 1)
            try:
                gesehen = analysiere(doc, blatt, bericht, gesehen)
            except Exception as fehler:
                bericht.notiere_fehler(
                    blatt["blattname"], 0, None, u"-", u"-",
                    u"Blatt konnte nicht ausgewertet werden: %s" % fehler)

    if not bericht.aenderungen:
        meldung = [u"Geprüfte Zeilen: %d" % bericht.zeilen_gesamt,
                   u"Bereits übereinstimmend: %d Wert(e)" % bericht.unveraendert,
                   u"Ignoriert (gesperrt/berechnet): %d Zelle(n)"
                   % bericht.ignoriert_gesperrt]
        if bericht.uebersprungen:
            meldung.append(u"Übersprungene Zeilen: %d"
                           % len(bericht.uebersprungen))
        if bericht.fehler:
            meldung.append(u"Fehler: %d (Details in der Ausgabe)"
                           % len(bericht.fehler))
        meldung.append(u"")
        meldung.append(u"Gelesener Dateistand: %s" % stand)
        meldung.append(u"Wurde die Datei online (OneDrive/SharePoint) "
                       u"bearbeitet? Dann erst die Synchronisierung abwarten.")
        meldung.append(u"")
        meldung.append(u"Die Spaltenanalyse im pyRevit-Ausgabefenster zeigt, "
                       u"welche Spalten geprüft wurden.")
        ui.meldung(u"\n".join(meldung),
                   hauptzeile=u"Keine schreibbaren Abweichungen gefunden")
        # Kein vorzeitiges Ende: die Auswertung unten wird immer ausgegeben,
        # sonst lässt sich "nichts zu tun" nicht von "nichts erkannt" trennen.

    # -----------------------------------------------------------------------
    # 4. Bestätigung einholen
    # -----------------------------------------------------------------------
    if bericht.aenderungen:
        betroffene = len(set(a.uid for a in bericht.aenderungen))
        typaenderungen = len([a for a in bericht.aenderungen
                              if a.ist_typparameter])

        text = []
        if typaenderungen:
            text.append(u"Darunter %d Typparameter - diese wirken auf ALLE "
                        u"Instanzen des jeweiligen Typs." % typaenderungen)
        if bericht.uebersprungen:
            text.append(u"%d Zeile(n) werden übersprungen."
                        % len(bericht.uebersprungen))
        if bericht.fehler:
            text.append(u"%d Zelle(n) konnten nicht verarbeitet werden."
                        % len(bericht.fehler))
        if text:
            text.append(u"")
        text.append(u"Alle Änderungen werden zusammengefasst und lassen sich "
                    u"mit einem einzigen Rückgängig-Schritt zurücknehmen.")
        text.append(u"")
        text.append(u"Jetzt in das Modell schreiben?")

        if not ui.frage(u"\n".join(text),
                        hauptzeile=u"%d Wert(e) an %d Element(en) ändern"
                                   % (len(bericht.aenderungen), betroffene),
                        standard_ja=True):
            return

        # -------------------------------------------------------------------
        # 5. Schreiben - alles in einer TransactionGroup (ein Undo-Schritt)
        # -------------------------------------------------------------------
        gruppen = OrderedDict()
        for aenderung in bericht.aenderungen:
            gruppen.setdefault(aenderung.blatt, []).append(aenderung)

        transaktionsgruppe = TransactionGroup(doc, u"ScheduleSync: Excel-Import")
        transaktionsgruppe.Start()
        try:
            with ui.Fortschritt(output, len(gruppen)) as fortschritt:
                for index, (blattname, aenderungen) in enumerate(
                        gruppen.items()):
                    fortschritt.aktualisiere(index + 1)
                    transaktion = Transaction(doc,
                                              u"Import: %s" % blattname[:100])
                    transaktion.Start()
                    try:
                        wende_an(doc, aenderungen, bericht)
                        transaktion.Commit()
                    except Exception:
                        if transaktion.GetStatus() == TransactionStatus.Started:
                            transaktion.RollBack()
                        raise
            transaktionsgruppe.Assimilate()
        except Exception as fehler:
            if transaktionsgruppe.GetStatus() == TransactionStatus.Started:
                transaktionsgruppe.RollBack()
            ui.meldung(u"%s" % fehler,
                       hauptzeile=u"Import abgebrochen und vollständig "
                                  u"zurückgenommen", warnung=True)
            return

    # -----------------------------------------------------------------------
    # 6. Zusammenfassung
    # -----------------------------------------------------------------------
    output.print_md(u"# ScheduleSync - Import")
    output.print_md(u"**Datei:** `%s`" % quelldatei)
    output.print_md(u"**Dateistand:** %s" % stand)

    output.print_table(
        table_data=[
            [u"Geänderte Werte", bericht.angewendete_aenderungen],
            [u"Geänderte Elemente", bericht.betroffene_elemente],
            [u"Unverändert (Wert stimmte bereits)", bericht.unveraendert],
            [u"Ignoriert (gesperrt/berechnet)", bericht.ignoriert_gesperrt],
            [u"Übersprungene Zeilen", len(bericht.uebersprungen)],
            [u"Fehler", len(bericht.fehler)],
        ],
        title=u"Zusammenfassung",
        columns=[u"Kennzahl", u"Anzahl"],
    )

    # Spaltendiagnose: beantwortet die Frage "warum passiert bei meiner
    # geänderten Spalte nichts?" ohne Rätselraten.
    if bericht.spalten:
        output.print_table(
            table_data=[[
                e["blatt"], e["spalte"], e["name"], e["quelle"], e["status"],
                e["verglichen"], e["gleich"], e["geplant"], e["gesperrt"],
                e["uebersprungen"], e["fehler"],
            ] for e in bericht.spalten.values()],
            title=u"Spaltenanalyse",
            columns=[u"Blatt", u"Sp.", u"Feld", u"Zuordnung", u"Status",
                     u"geprüft", u"gleich", u"zu ändern", u"gesperrt",
                     u"übersprungen", u"Fehler"],
        )

    geaendert = [a for a in bericht.aenderungen if a.angewendet]
    if geaendert:
        output.print_table(
            table_data=[[
                a.blatt, a.excel_zeile, elementlink(a.element_id), a.feldname,
                a.alt_anzeige or u"(leer)",
                u"Typparameter" if a.ist_typparameter else u"Instanz",
            ] for a in geaendert],
            title=u"Geänderte Werte",
            columns=[u"Blatt", u"Zeile", u"Element", u"Feld", u"vorher", u"Art"],
        )

    if bericht.uebersprungen:
        output.print_table(
            table_data=[[e["blatt"], e["zeile"], e["uid"], e["grund"]]
                        for e in bericht.uebersprungen],
            title=u"Übersprungene Zeilen",
            columns=[u"Blatt", u"Zeile", u"UniqueId", u"Grund"],
        )

    if bericht.fehler:
        output.print_table(
            table_data=[[e["blatt"], e["zeile"], elementlink(e["element_id"]),
                         e["element"], e["feld"], e["meldung"]]
                        for e in bericht.fehler],
            title=u"Fehler",
            columns=[u"Blatt", u"Zeile", u"Element", u"Bezeichnung", u"Feld",
                     u"Meldung"],
        )

    if lese_warnungen:
        output.print_md(u"## Hinweise zur Datei")
        for warnung in lese_warnungen:
            output.print_md(u"- %s" % warnung)


def _zeige_fehler(spur, titel):
    """Fehler sichtbar machen - und dabei selbst nicht scheitern koennen."""
    protokoll = os.path.join(os.environ.get("TEMP", "."),
                             "pyMLG_ScheduleSync_Fehler.log")
    try:
        with io.open(protokoll, "w", encoding="utf-8") as datei:
            datei.write(spur)
    except Exception:
        protokoll = None

    try:
        output.print_md(u"# " + titel)
        print(spur)
        if protokoll:
            print(u"Protokoll: " + protokoll)
    except Exception:
        pass

    try:
        from schedule_sync import ui as _ui
        letzte_zeile = (spur.strip().splitlines() or [u""])[-1]
        zusatz = (u" und in " + protokoll) if protokoll else u""
        _ui.meldung(letzte_zeile
                    + u"\n\nEinzelheiten im pyRevit-Ausgabefenster"
                    + zusatz + u".",
                    hauptzeile=titel, warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    # Keine stillen Abstuerze: den Traceback zeigen, statt ihn als
    # nichtssagenden Revit-Fehlerdialog ("Object reference not set to an
    # instance of an object") enden zu lassen.
    _zeige_fehler(traceback.format_exc(), u"ScheduleSync - Import ist fehlgeschlagen")
