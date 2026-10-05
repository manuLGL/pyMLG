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

from mlg_sprache import t  # noqa: E402

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
        ui.meldung(t(u"Es ist kein Projekt geöffnet. Bitte zuerst eine "
                   u"Revit-Projektdatei öffnen.", u"No project is open. Please open a Revit project file first.", u"No hay ningún proyecto abierto. Abra primero un archivo de proyecto de Revit."),
                   hauptzeile=t(u"Kein aktives Dokument", u"No active document", u"No hay documento activo"), warnung=True)
        return
    if doc.IsFamilyDocument:
        ui.meldung(t(u"Bauteillisten gibt es nur in Projektdateien, nicht im "
                   u"Familieneditor.", u"Schedules only exist in project files, not in the family editor.", u"Las tablas de planificación solo existen en archivos de proyecto, no en el editor de familias."),
                   hauptzeile=t(u"Familiendokument wird nicht unterstützt", u"Family document not supported", u"No se admiten documentos de familia"),
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
        ui.meldung(t(u"Das aktive Dokument ist schreibgeschützt geöffnet. "
                   u"Ein Import ist nicht möglich.", u"The active document is open read-only. Import is not possible.", u"El documento activo está abierto como solo lectura. No es posible importar."),
                   hauptzeile=t(u"Import nicht möglich", u"Import not possible", u"Importación no posible"), warnung=True)
        return

    try:
        openpyxl = deps.lade_openpyxl()
    except deps.AbhaengigkeitFehlt as fehler:
        ui.meldung(str(fehler), hauptzeile=t(u"openpyxl fehlt", u"openpyxl missing", u"Falta openpyxl"), warnung=True)
        return

    # -----------------------------------------------------------------------
    # 2. Excel-Datei wählen und lesen
    # -----------------------------------------------------------------------
    quelldatei = ui.datei_oeffnen(
        startordner=os.path.dirname(doc.PathName) if doc.PathName else u"",
        titel=t(u"Bearbeitete Excel-Datei auswählen", u"Select the edited Excel file", u"Seleccione el archivo de Excel editado"),
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
        ui.meldung(t(u"Vermutlich ist die Datei gerade in Excel geöffnet. "
                   u"Bitte schliessen und erneut versuchen.\n\n%s", u"The file is probably open in Excel. Please close it and try again.\n\n%s", u"Probablemente el archivo está abierto en Excel. Ciérrelo e inténtelo de nuevo.\n\n%s") % quelldatei,
                   hauptzeile=t(u"Datei konnte nicht gelesen werden", u"File could not be read", u"No se pudo leer el archivo"), warnung=True)
        return
    except Exception as fehler:
        ui.meldung(u"%s" % fehler,
                   hauptzeile=t(u"Die Excel-Datei konnte nicht gelesen werden", u"The Excel file could not be read", u"No se pudo leer el archivo de Excel"),
                   warnung=True)
        return

    if not blaetter:
        text = u"\n".join(u"• %s" % warnung for warnung in lese_warnungen)
        ui.meldung(t(u"Erwartet wird ein Blatt, dessen Zelle A1 'UniqueId' "
                   u"enthält (so wie vom Export erzeugt).\n\n%s", u"Expected a sheet whose cell A1 contains 'UniqueId' (as created by the export).\n\n%s", u"Se espera una hoja cuya celda A1 contenga 'UniqueId' (como la genera la exportación).\n\n%s") % text,
                   hauptzeile=t(u"Kein auswertbares Tabellenblatt gefunden", u"No usable worksheet found", u"No se encontró ninguna hoja utilizable"),
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
                    t(u"Blatt konnte nicht ausgewertet werden: %s", u"Sheet could not be evaluated: %s", u"No se pudo evaluar la hoja: %s") % fehler)

    if not bericht.aenderungen:
        meldung = [t(u"Geprüfte Zeilen: %d", u"Rows checked: %d", u"Filas comprobadas: %d") % bericht.zeilen_gesamt,
                   t(u"Bereits übereinstimmend: %d Wert(e)", u"Already matching: %d value(s)", u"Ya coincidentes: %d valor(es)") % bericht.unveraendert,
                   t(u"Ignoriert (gesperrt/berechnet): %d Zelle(n)", u"Ignored (locked/calculated): %d cell(s)", u"Ignoradas (bloqueadas/calculadas): %d celda(s)")
                   % bericht.ignoriert_gesperrt]
        if bericht.uebersprungen:
            meldung.append(t(u"Übersprungene Zeilen: %d", u"Skipped rows: %d", u"Filas omitidas: %d")
                           % len(bericht.uebersprungen))
        if bericht.fehler:
            meldung.append(t(u"Fehler: %d (Details in der Ausgabe)", u"Errors: %d (details in the output)", u"Errores: %d (detalles en la salida)")
                           % len(bericht.fehler))
        meldung.append(u"")
        meldung.append(t(u"Gelesener Dateistand: %s", u"File version read: %s", u"Versión del archivo leída: %s") % stand)
        meldung.append(t(u"Wurde die Datei online (OneDrive/SharePoint) "
                       u"bearbeitet? Dann erst die Synchronisierung abwarten.", u"Was the file edited online (OneDrive/SharePoint)? Then wait for synchronisation first.", u"¿Se editó el archivo en línea (OneDrive/SharePoint)? Espere primero a la sincronización."))
        meldung.append(u"")
        meldung.append(t(u"Die Spaltenanalyse im pyRevit-Ausgabefenster zeigt, "
                       u"welche Spalten geprüft wurden.", u"The column analysis in the pyRevit output window shows which columns were checked.", u"El análisis de columnas en la ventana de salida de pyRevit muestra qué columnas se comprobaron."))
        ui.meldung(u"\n".join(meldung),
                   hauptzeile=t(u"Keine schreibbaren Abweichungen gefunden", u"No writable differences found", u"No se encontraron diferencias que escribir"))
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
            text.append(t(u"Darunter %d Typparameter - diese wirken auf ALLE "
                        u"Instanzen des jeweiligen Typs.", u"Including %d type parameters - these affect ALL instances of the respective type.", u"Entre ellos %d parámetros de tipo: afectan a TODOS los ejemplares del tipo.") % typaenderungen)
        if bericht.uebersprungen:
            text.append(t(u"%d Zeile(n) werden übersprungen.", u"%d row(s) will be skipped.", u"Se omitirán %d fila(s).")
                        % len(bericht.uebersprungen))
        if bericht.fehler:
            text.append(t(u"%d Zelle(n) konnten nicht verarbeitet werden.", u"%d cell(s) could not be processed.", u"No se pudieron procesar %d celda(s).")
                        % len(bericht.fehler))
        if text:
            text.append(u"")
        text.append(t(u"Alle Änderungen werden zusammengefasst und lassen sich "
                    u"mit einem einzigen Rückgängig-Schritt zurücknehmen.", u"All changes are combined and can be reverted with a single undo step.", u"Todos los cambios se agrupan y se pueden deshacer con un solo paso."))
        text.append(u"")
        text.append(t(u"Jetzt in das Modell schreiben?", u"Write to the model now?", u"¿Escribir ahora en el modelo?"))

        if not ui.frage(u"\n".join(text),
                        hauptzeile=t(u"%d Wert(e) an %d Element(en) ändern", u"Change %d value(s) on %d element(s)", u"Cambiar %d valor(es) en %d elemento(s)")
                                   % (len(bericht.aenderungen), betroffene),
                        standard_ja=True):
            return

        # -------------------------------------------------------------------
        # 5. Schreiben - alles in einer TransactionGroup (ein Undo-Schritt)
        # -------------------------------------------------------------------
        gruppen = OrderedDict()
        for aenderung in bericht.aenderungen:
            gruppen.setdefault(aenderung.blatt, []).append(aenderung)

        transaktionsgruppe = TransactionGroup(doc, t(u"ScheduleSync: Excel-Import", u"ScheduleSync: Excel import", u"ScheduleSync: importar Excel"))
        transaktionsgruppe.Start()
        try:
            with ui.Fortschritt(output, len(gruppen)) as fortschritt:
                for index, (blattname, aenderungen) in enumerate(
                        gruppen.items()):
                    fortschritt.aktualisiere(index + 1)
                    transaktion = Transaction(doc,
                                              t(u"Import: %s", u"Import: %s", u"Importar: %s") % blattname[:100])
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
                       hauptzeile=t(u"Import abgebrochen und vollständig "
                                  u"zurückgenommen", u"Import cancelled and fully reverted", u"Importación cancelada y revertida por completo"), warnung=True)
            return

    # -----------------------------------------------------------------------
    # 6. Zusammenfassung
    # -----------------------------------------------------------------------
    output.print_md(t(u"# ScheduleSync - Import", u"# ScheduleSync - Import", u"# ScheduleSync - Importar"))
    output.print_md(t(u"**Datei:** `%s`", u"**File:** `%s`", u"**Archivo:** `%s`") % quelldatei)
    output.print_md(t(u"**Dateistand:** %s", u"**File version:** %s", u"**Versión del archivo:** %s") % stand)

    output.print_table(
        table_data=[
            [t(u"Geänderte Werte", u"Changed values", u"Valores cambiados"), bericht.angewendete_aenderungen],
            [t(u"Geänderte Elemente", u"Changed elements", u"Elementos cambiados"), bericht.betroffene_elemente],
            [t(u"Unverändert (Wert stimmte bereits)", u"Unchanged (value already matched)", u"Sin cambios (el valor ya coincidía)"), bericht.unveraendert],
            [t(u"Ignoriert (gesperrt/berechnet)", u"Ignored (locked/calculated)", u"Ignorados (bloqueados/calculados)"), bericht.ignoriert_gesperrt],
            [t(u"Übersprungene Zeilen", u"Skipped rows", u"Filas omitidas"), len(bericht.uebersprungen)],
            [t(u"Fehler", u"Errors", u"Errores"), len(bericht.fehler)],
        ],
        title=t(u"Zusammenfassung", u"Summary", u"Resumen"),
        columns=[t(u"Kennzahl", u"Metric", u"Indicador"), t(u"Anzahl", u"Count", u"Cantidad")],
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
            title=t(u"Spaltenanalyse", u"Column analysis", u"Análisis de columnas"),
            columns=[t(u"Blatt", u"Sheet", u"Hoja"), u"Sp.", t(u"Feld", u"Field", u"Campo"), t(u"Zuordnung", u"Mapping", u"Asignación"), t(u"Status", u"Status", u"Estado"),
                     t(u"geprüft", u"checked", u"comprobados"), u"gleich", t(u"zu ändern", u"to change", u"a cambiar"), u"gesperrt",
                     t(u"übersprungen", u"skipped", u"omitidos"), t(u"Fehler", u"Errors", u"Errores")],
        )

    geaendert = [a for a in bericht.aenderungen if a.angewendet]
    if geaendert:
        output.print_table(
            table_data=[[
                a.blatt, a.excel_zeile, elementlink(a.element_id), a.feldname,
                a.alt_anzeige or u"(leer)",
                t(u"Typparameter", u"Type parameter", u"Parámetro de tipo") if a.ist_typparameter else t(u"Instanz", u"Instance", u"Ejemplar"),
            ] for a in geaendert],
            title=t(u"Geänderte Werte", u"Changed values", u"Valores cambiados"),
            columns=[t(u"Blatt", u"Sheet", u"Hoja"), t(u"Zeile", u"Row", u"Fila"), t(u"Element", u"Element", u"Elemento"), t(u"Feld", u"Field", u"Campo"), u"vorher", u"Art"],
        )

    if bericht.uebersprungen:
        output.print_table(
            table_data=[[e["blatt"], e["zeile"], e["uid"], e["grund"]]
                        for e in bericht.uebersprungen],
            title=t(u"Übersprungene Zeilen", u"Skipped rows", u"Filas omitidas"),
            columns=[t(u"Blatt", u"Sheet", u"Hoja"), t(u"Zeile", u"Row", u"Fila"), u"UniqueId", t(u"Grund", u"Reason", u"Motivo")],
        )

    if bericht.fehler:
        output.print_table(
            table_data=[[e["blatt"], e["zeile"], elementlink(e["element_id"]),
                         e["element"], e["feld"], e["meldung"]]
                        for e in bericht.fehler],
            title=t(u"Fehler", u"Errors", u"Errores"),
            columns=[t(u"Blatt", u"Sheet", u"Hoja"), t(u"Zeile", u"Row", u"Fila"), t(u"Element", u"Element", u"Elemento"), t(u"Bezeichnung", u"Description", u"Descripción"), t(u"Feld", u"Field", u"Campo"),
                     t(u"Meldung", u"Message", u"Mensaje")],
        )

    if lese_warnungen:
        output.print_md(t(u"## Hinweise zur Datei", u"## Notes on the file", u"## Notas sobre el archivo"))
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
            print(t(u"Protokoll: ", u"Log: ", u"Registro: ") + protokoll)
    except Exception:
        pass

    try:
        from schedule_sync import ui as _ui
        letzte_zeile = (spur.strip().splitlines() or [u""])[-1]
        zusatz = (t(u" und in ", u" and in ", u" y en ") + protokoll) if protokoll else u""
        _ui.meldung(letzte_zeile
                    + t(u"\n\nEinzelheiten im pyRevit-Ausgabefenster", u"\n\nDetails in the pyRevit output window", u"\n\nDetalles en la ventana de salida de pyRevit")
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
    _zeige_fehler(traceback.format_exc(), t(u"ScheduleSync - Import ist fehlgeschlagen", u"ScheduleSync - Import failed", u"ScheduleSync - La importación ha fallado"))
