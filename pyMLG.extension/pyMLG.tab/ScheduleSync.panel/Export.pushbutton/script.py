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
_EXT = os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.dirname(os.path.abspath(__file__)))))
if os.path.join(_EXT, "lib") not in sys.path:
    sys.path.append(os.path.join(_EXT, "lib"))

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

    ungueltige_dateizeichen = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

    def dateiname_teil(text, standard=u"Unbenannt"):
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
        return doc.Title or u"Projekt"

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
        ui.meldung(str(fehler), hauptzeile=u"openpyxl fehlt", warnung=True)
        return

    # -----------------------------------------------------------------------
    # 2. Bauteillisten auswählen
    # -----------------------------------------------------------------------
    alle_schedules = sammle_schedules(doc)
    if not alle_schedules:
        ui.meldung(u"In diesem Projekt wurden keine Bauteillisten gefunden.",
                   hauptzeile=u"Nichts zu exportieren")
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
        titel=u"ScheduleSync - Export",
        hinweis=u"Bauteillisten auswählen, die nach Excel exportiert werden "
                u"sollen:",
        schaltflaeche=u"Nach Excel exportieren",
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
        ui.meldung(text, hauptzeile=u"Keine der gewählten Bauteillisten kann "
                                    u"exportiert werden", warnung=True)
        return

    if abgelehnt:
        text = u"\n".join(u"• %s\n   %s" % (name, grund)
                          for name, grund in abgelehnt)
        weiter = ui.frage(
            u"%s\n\nDie übrigen %d Liste(n) trotzdem exportieren?"
            % (text, len(exportierbar)),
            hauptzeile=u"Nicht unterstützte Bauteillisten werden übersprungen",
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
                    lesefehler.append((name, u"Die Bauteilliste enthält keine "
                                             u"sichtbaren Felder."))
                    continue
                sortierung = lese_sortierung(doc, schedule, felder)
                zeilen, uebersprungen = lese_zeilen(doc, schedule, felder,
                                                    sortierung)
                if uebersprungen:
                    warnungen.append(
                        u"%s: %d Element(e) konnten nicht gelesen werden."
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
        ui.meldung(text, hauptzeile=u"Es konnten keine Daten gelesen werden",
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
                                  titel=u"Excel-Datei speichern")
    if not zielpfad:
        return

    try:
        schreibe_arbeitsmappe(openpyxl, zielpfad, bloecke, doc.Title)
    except PermissionError:
        ui.meldung(u"Vermutlich ist die Datei gerade in Excel geöffnet. "
                   u"Bitte schliessen und erneut versuchen.\n\n%s" % zielpfad,
                   hauptzeile=u"Datei konnte nicht geschrieben werden",
                   warnung=True)
        return
    except Exception as fehler:
        ui.meldung(u"%s" % fehler,
                   hauptzeile=u"Die Excel-Datei konnte nicht geschrieben werden",
                   warnung=True)
        return

    # -----------------------------------------------------------------------
    # 6. Zusammenfassung
    # -----------------------------------------------------------------------
    output.print_md(u"# ScheduleSync - Export")
    output.print_md(u"**Datei:** `%s`" % zielpfad)

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
        title=u"Exportierte Bauteillisten",
        columns=[u"Bauteilliste", u"Elemente", u"Spalten", u"davon änderbar",
                 u"gesperrt (grau)", u"Blatt 'Ansicht'"],
    )

    if warnungen:
        output.print_md(u"## Hinweise")
        for warnung in warnungen:
            output.print_md(u"- %s" % warnung)

    if lesefehler:
        output.print_md(u"## Nicht exportiert")
        for name, grund in lesefehler:
            output.print_md(u"- **%s**: %s" % (name, grund))

    if abgelehnt:
        output.print_md(u"## Übersprungen")
        for name, grund in abgelehnt:
            output.print_md(u"- **%s**: %s" % (name, grund))

    output.print_md(
        u"---\n"
        u"Graue Zellen sind schreibgeschützt oder berechnet und werden beim "
        u"Import ignoriert. Orange Zellen sind **Typparameter** - eine Änderung "
        u"dort wirkt auf alle Instanzen des Typs. Spalte A (UniqueId) ist "
        u"ausgeblendet und darf nicht verändert werden.\n\n"
        u"Das Datenblatt enthält immer **eine Zeile je Element**, sortiert wie "
        u"in Revit. Die Tabelle genau wie in Revit - mit Gruppen, "
        u"zusammengefassten Zeilen und Summen - steht im Blatt "
        u"**'Ansicht - …'**; es dient nur zum Lesen und wird beim Import "
        u"ignoriert.")

    # -----------------------------------------------------------------------
    # 7. Datei öffnen (optional)
    # -----------------------------------------------------------------------
    if ui.frage(u"Datei jetzt in Excel öffnen?\n\n%s" % zielpfad,
                hauptzeile=u"Export abgeschlossen", standard_ja=True):
        try:
            os.startfile(zielpfad)
        except Exception as fehler:
            ui.meldung(u"%s\n\nPfad:\n%s" % (fehler, zielpfad),
                       hauptzeile=u"Die Datei konnte nicht geöffnet werden "
                                  u"(ist Excel installiert?)",
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
    _zeige_fehler(traceback.format_exc(), u"ScheduleSync - Export ist fehlgeschlagen")
