#! python3
# -*- coding: utf-8 -*-
"""Exportiert Revit-Bauteillisten nach Excel (openpyxl, CPython3-Engine).

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

# Der lib-Ordner der Extension wird von pyRevit automatisch in sys.path gelegt.
# Die folgende Zeile ist nur eine Absicherung, falls das Skript aus einem
# ungewöhnlichen Kontext gestartet wird.
_EXT = os.path.dirname(os.path.abspath(__file__))
while not _EXT.endswith(".extension") and os.path.dirname(_EXT) != _EXT:
    _EXT = os.path.dirname(_EXT)
if os.path.join(_EXT, "lib") not in sys.path:
    sys.path.append(os.path.join(_EXT, "lib"))

from mlg_sprache import t  # noqa: E402

from pyrevit import script

output = script.get_output()


def main():
    import datetime
    import re

    from schedule_sync import deps, ui
    from schedule_sync.excel_io import schreibe_arbeitsmappe
    from schedule_sync.schedule_model import (
        lese_ansicht,
        lese_felder,
        lese_sortierung,
        lese_zeilen,
        pruefe_schedule,
        sammle_schedules,
        spaltenstatus,
        SCHREIBBAR,
    )

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

    ungueltige_dateizeichen = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

    def dateiname_teil(text, standard=t(u"Unbenannt", u"Untitled", u"Sin nombre")):
        """Macht einen Revit-Namen als Teil eines Dateinamens verwendbar."""
        sauber = ungueltige_dateizeichen.sub(u"_", (text or u"").strip())
        sauber = sauber.strip(u". ")
        return sauber or standard

    def projektname():
        """Projektname aus den Projektinformationen, sonst Dateiname."""
        try:
            name = doc.ProjectInformation.Name
            if name and name.strip():
                return name.strip()
        except Exception:
            pass
        if doc.PathName:
            return os.path.splitext(os.path.basename(doc.PathName))[0]
        return doc.Title or t(u"Projekt", u"Project", u"Proyecto")

    def startordner():
        """Ordner der Revit-Datei als Vorgabe für den Speichern-Dialog."""
        if doc.PathName:
            ordner = os.path.dirname(doc.PathName)
            if os.path.isdir(ordner):
                return ordner
        return u""

    # -----------------------------------------------------------------------
    # 1. Abhängigkeiten prüfen
    # -----------------------------------------------------------------------
    try:
        openpyxl = deps.lade_openpyxl()
    except deps.AbhaengigkeitFehlt as fehler:
        ui.meldung(str(fehler), hauptzeile=t(u"openpyxl fehlt", u"openpyxl missing", u"Falta openpyxl"), warnung=True)
        return

    # -----------------------------------------------------------------------
    # 2. Bauteillisten auswählen
    # -----------------------------------------------------------------------
    alle_schedules = sammle_schedules(doc)
    if not alle_schedules:
        ui.meldung(t(u"In diesem Projekt wurden keine Bauteillisten gefunden.", u"No schedules were found in this project.", u"No se encontraron tablas de planificación en este proyecto."),
                   hauptzeile=t(u"Nichts zu exportieren", u"Nothing to export", u"Nada que exportar"))
        return

    nach_name = {}
    for schedule in alle_schedules:
        # Ansichtsnamen sind in Revit eindeutig; die Absicherung kostet nichts.
        name = schedule.Name
        while name in nach_name:
            name += u" "
        nach_name[name] = schedule

    auswahl = ui.waehle_mehrfach(
        sorted(nach_name.keys()),
        titel=t(u"ScheduleSync - Export", u"ScheduleSync - Export", u"ScheduleSync - Exportar"),
        hinweis=t(u"Bauteillisten auswählen, die nach Excel exportiert werden "
                u"sollen:", u"Select the schedules to export to Excel:", u"Seleccione las tablas que se exportarán a Excel:"),
        schaltflaeche=t(u"Nach Excel exportieren", u"Export to Excel", u"Exportar a Excel"),
    )
    if not auswahl:
        return

    # -----------------------------------------------------------------------
    # 3. Exportierbarkeit prüfen
    # -----------------------------------------------------------------------
    exportierbar = []
    abgelehnt = []
    warnungen = []

    for name in auswahl:
        schedule = nach_name[name]
        ok, grund, schedule_warnungen = pruefe_schedule(schedule)
        if ok:
            exportierbar.append((name, schedule))
            for warnung in schedule_warnungen:
                warnungen.append(u"%s: %s" % (name, warnung))
        else:
            abgelehnt.append((name, grund))

    if not exportierbar:
        text = u"\n".join(u"• %s\n   %s" % (name, grund)
                          for name, grund in abgelehnt)
        ui.meldung(text, hauptzeile=t(u"Keine der gewählten Bauteillisten kann "
                                    u"exportiert werden", u"None of the selected schedules can be exported", u"No se puede exportar ninguna de las tablas seleccionadas"), warnung=True)
        return

    if abgelehnt:
        text = u"\n".join(u"• %s\n   %s" % (name, grund)
                          for name, grund in abgelehnt)
        weiter = ui.frage(
            t(u"%s\n\nDie übrigen %d Liste(n) trotzdem exportieren?", u"%s\n\nExport the remaining %d schedule(s) anyway?", u"%s\n\n¿Exportar de todos modos las %d tabla(s) restantes?")
            % (text, len(exportierbar)),
            hauptzeile=t(u"Nicht unterstützte Bauteillisten werden übersprungen", u"Unsupported schedules will be skipped", u"Se omitirán las tablas no admitidas"),
            warnung=True, standard_ja=True)
        if not weiter:
            return

    # -----------------------------------------------------------------------
    # 4. Daten einlesen
    # -----------------------------------------------------------------------
    bloecke = []
    lesefehler = []

    with ui.Fortschritt(output, len(exportierbar)) as fortschritt:
        for index, (name, schedule) in enumerate(exportierbar):
            fortschritt.aktualisiere(index + 1)
            try:
                felder = lese_felder(doc, schedule)
                if not felder:
                    lesefehler.append((name, t(u"Die Bauteilliste enthält keine "
                                             u"sichtbaren Felder.", u"The schedule has no visible fields.", u"La tabla no tiene campos visibles.")))
                    continue
                sortierung = lese_sortierung(doc, schedule, felder)
                zeilen, uebersprungen = lese_zeilen(doc, schedule, felder,
                                                    sortierung)
                if uebersprungen:
                    warnungen.append(
                        t(u"%s: %d Element(e) konnten nicht gelesen werden.", u"%s: %d element(s) could not be read.", u"%s: no se pudieron leer %d elemento(s).")
                        % (name, uebersprungen))
                bloecke.append({
                    "name": name,
                    "uid": schedule.UniqueId,
                    "felder": felder,
                    "zeilen": zeilen,
                    "spaltenstatus": spaltenstatus(felder, zeilen),
                    # Tabelle wie in Revit, inkl. Gruppen und Summen
                    "ansicht": lese_ansicht(schedule),
                })
            except Exception as fehler:
                lesefehler.append((name, str(fehler)))

    if not bloecke:
        text = u"\n".join(u"• %s: %s" % (name, grund)
                          for name, grund in lesefehler)
        ui.meldung(text, hauptzeile=t(u"Es konnten keine Daten gelesen werden", u"No data could be read", u"No se pudieron leer datos"),
                   warnung=True)
        return

    # -----------------------------------------------------------------------
    # 5. Speicherort wählen und schreiben
    # -----------------------------------------------------------------------
    if len(bloecke) == 1:
        kern = u"%s_%s" % (dateiname_teil(projektname()),
                           dateiname_teil(bloecke[0]["name"]))
    else:
        kern = u"%s_Bauteillisten" % dateiname_teil(projektname())

    vorgabe = u"%s_%s.xlsx" % (kern, datetime.date.today().strftime("%Y-%m-%d"))

    zielpfad = ui.datei_speichern(vorgabename=vorgabe,
                                  startordner=startordner(),
                                  titel=t(u"Excel-Datei speichern", u"Save Excel file", u"Guardar archivo de Excel"))
    if not zielpfad:
        return

    try:
        schreibe_arbeitsmappe(openpyxl, zielpfad, bloecke, doc.Title)
    except PermissionError:
        ui.meldung(t(u"Vermutlich ist die Datei gerade in Excel geöffnet. "
                   u"Bitte schliessen und erneut versuchen.\n\n%s", u"The file is probably open in Excel. Please close it and try again.\n\n%s", u"Probablemente el archivo está abierto en Excel. Ciérrelo e inténtelo de nuevo.\n\n%s") % zielpfad,
                   hauptzeile=t(u"Datei konnte nicht geschrieben werden", u"File could not be written", u"No se pudo escribir el archivo"),
                   warnung=True)
        return
    except Exception as fehler:
        ui.meldung(u"%s" % fehler,
                   hauptzeile=t(u"Die Excel-Datei konnte nicht geschrieben werden", u"The Excel file could not be written", u"No se pudo escribir el archivo de Excel"),
                   warnung=True)
        return

    # -----------------------------------------------------------------------
    # 6. Zusammenfassung
    # -----------------------------------------------------------------------
    output.print_md(t(u"# ScheduleSync - Export", u"# ScheduleSync - Export", u"# ScheduleSync - Exportar"))
    output.print_md(t(u"**Datei:** `%s`", u"**File:** `%s`", u"**Archivo:** `%s`") % zielpfad)

    tabelle = []
    for block in bloecke:
        status = block["spaltenstatus"]
        schreibbar = len([s for s in status.values() if s in SCHREIBBAR])
        tabelle.append([block["name"], len(block["zeilen"]),
                        len(block["felder"]), schreibbar,
                        len(block["felder"]) - schreibbar,
                        u"ja" if block.get("ansicht") else u"-"])

    output.print_table(
        table_data=tabelle,
        title=t(u"Exportierte Bauteillisten", u"Exported schedules", u"Tablas exportadas"),
        columns=[t(u"Bauteilliste", u"Schedule", u"Tabla"), t(u"Elemente", u"Elements", u"Elementos"), t(u"Spalten", u"Columns", u"Columnas"), t(u"davon änderbar", u"of which editable", u"de ellas editables"),
                 t(u"gesperrt (grau)", u"locked (grey)", u"bloqueadas (gris)"), t(u"Blatt 'Ansicht'", u"Sheet 'View'", u"Hoja 'Vista'")],
    )

    if warnungen:
        output.print_md(t(u"## Hinweise", u"## Notes", u"## Notas"))
        for warnung in warnungen:
            output.print_md(u"- %s" % warnung)

    if lesefehler:
        output.print_md(t(u"## Nicht exportiert", u"## Not exported", u"## No exportado"))
        for name, grund in lesefehler:
            output.print_md(u"- **%s**: %s" % (name, grund))

    if abgelehnt:
        output.print_md(t(u"## Übersprungen", u"## Skipped", u"## Omitido"))
        for name, grund in abgelehnt:
            output.print_md(u"- **%s**: %s" % (name, grund))

    output.print_md(
        t(u"---\n"
        u"Graue Zellen sind schreibgeschützt oder berechnet und werden beim "
        u"Import ignoriert. Orange Zellen sind **Typparameter** - eine Änderung "
        u"dort wirkt auf alle Instanzen des Typs. Spalte A (UniqueId) ist "
        u"ausgeblendet und darf nicht verändert werden.\n\n"
        u"Das Datenblatt enthält immer **eine Zeile je Element**, sortiert wie "
        u"in Revit. Die Tabelle genau wie in Revit - mit Gruppen, "
        u"zusammengefassten Zeilen und Summen - steht im Blatt "
        u"**'Ansicht - …'**; es dient nur zum Lesen und wird beim Import "
        u"ignoriert.", u"---\nGrey cells are read-only or calculated and are ignored on import. Orange cells are **type parameters** - a change there affects all instances of the type. Column A (UniqueId) is hidden and must not be changed.\n\nThe data sheet always contains **one row per element**, sorted as in Revit. The table exactly as in Revit - with groups, combined rows and totals - is in the sheet **'View - …'**; it is for reading only and is ignored on import.", u"---\nLas celdas grises son de solo lectura o calculadas y se ignoran al importar. Las celdas naranjas son **parámetros de tipo**: un cambio afecta a todos los ejemplares del tipo. La columna A (UniqueId) está oculta y no debe modificarse.\n\nLa hoja de datos contiene siempre **una fila por elemento**, ordenada como en Revit. La tabla exactamente como en Revit, con grupos, filas agrupadas y totales, está en la hoja **'Vista - …'**; es solo de lectura y se ignora al importar."))

    # -----------------------------------------------------------------------
    # 7. Datei öffnen (optional)
    # -----------------------------------------------------------------------
    if ui.frage(t(u"Datei jetzt in Excel öffnen?\n\n%s", u"Open the file in Excel now?\n\n%s", u"¿Abrir el archivo en Excel ahora?\n\n%s") % zielpfad,
                hauptzeile=t(u"Export abgeschlossen", u"Export complete", u"Exportación completada"), standard_ja=True):
        try:
            os.startfile(zielpfad)
        except Exception as fehler:
            ui.meldung(u"%s\n\nPfad:\n%s" % (fehler, zielpfad),
                       hauptzeile=t(u"Die Datei konnte nicht geöffnet werden "
                                  u"(ist Excel installiert?)", u"The file could not be opened (is Excel installed?)", u"No se pudo abrir el archivo (¿está instalado Excel?)"),
                       warnung=True)


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
    _zeige_fehler(traceback.format_exc(), t(u"ScheduleSync - Export ist fehlgeschlagen", u"ScheduleSync - Export failed", u"ScheduleSync - La exportación ha fallado"))
