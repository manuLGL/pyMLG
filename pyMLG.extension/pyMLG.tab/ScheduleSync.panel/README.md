# ScheduleSync – Bauteillisten zwischen Revit und Excel abgleichen

Zwei Schaltflächen im Ribbon `pyMLG ▸ Schedule Sync`:

| Schaltfläche | Funktion |
|---|---|
| **Schedules nach Excel** | Exportiert ausgewählte Bauteillisten in eine `.xlsx`-Datei |
| **Excel zurück lesen** | Schreibt die in Excel geänderten Werte zurück ins Modell |

Die Zuordnung zwischen Excel-Zeile und Revit-Element läuft über die **UniqueId**
des Elements. Diese bleibt über Sitzungen, Synchronisierungen und
Arbeitsteilung hinweg stabil – im Gegensatz zur ElementId, die sich ändern kann.

---

## 1. Voraussetzungen

* **Revit 2026**
* **pyRevit** (aktuelle Version), Engine **CPython 3**
  Beide Skripte beginnen mit `#! python3`; pyRevit wählt dadurch automatisch
  die CPython-Engine. Es ist keine weitere Einstellung nötig.
* **openpyxl** in der CPython-Umgebung (siehe nächster Abschnitt)

### Dialoge unter CPython

pyRevit stellt `pyrevit.forms` unter der CPython-Engine nur als Platzhalter
bereit – jeder Zugriff darauf (auch `alert`, `save_file` oder `pick_file`)
scheitert mit `PyRevitCPythonNotSupported`, weil die Formulare über den
IronPython-XAML-Loader gebaut werden, den es unter pythonnet nicht gibt.

ScheduleSync bringt deshalb eine eigene Dialogschicht mit
(`lib/schedule_sync/ui.py`):

| Zweck | Umsetzung |
|---|---|
| Meldungen und Ja/Nein-Rückfragen | `Autodesk.Revit.UI.TaskDialog` (reine Revit-API) |
| Auswahl der Bauteillisten | WinForms-Fenster mit Kontrollkästchen und Suchfeld |
| Datei öffnen/speichern | WinForms-Dateidialoge, ersatzweise `Microsoft.Win32` |
| Fortschritt | Fortschrittsbalken des pyRevit-Ausgabefensters |

Ein Unterschied zu `forms.ProgressBar`: Der Fortschrittsbalken des
Ausgabefensters bietet **keinen Abbrechen-Knopf**. Lange Exporte laufen also
durch – abbrechen lässt sich nur der Dialog davor.

Der Code kommt ohne XAML aus und läuft deshalb auch unter IronPython.

**Zweite Stolperfalle:** `script.exit()` ist in pyRevit nichts anderes als
`sys.exit()`. Das dabei ausgelöste `SystemExit` verlässt den CPython-Host und
endet in Revit als nichtssagendes *„Object reference not set to an instance of
an object"*. Beide Skripte verzichten deshalb vollständig darauf: Der Ablauf
steckt in einer Funktion `main()` und endet ausschließlich über `return`.

Umschlossen wird `main()` von einem Fehlerfang, der den vollständigen Traceback
in das pyRevit-Ausgabefenster **und** nach `%TEMP%\pyMLG_ScheduleSync_Fehler.log`
schreibt. Ein unerwarteter Fehler ist damit lesbar, statt im Revit-Dialog zu
verschwinden.

## 2. Installation

### 2.1 Extension registrieren

Falls der Ordner `pyMLG.extension` noch nicht bei pyRevit hinterlegt ist:

1. pyRevit-Ribbon ▸ **pyRevit ▸ Settings ▸ Custom Extension Directories**
2. Das Verzeichnis eintragen, in dem `pyMLG.extension` liegt
   (also den **übergeordneten** Ordner, nicht `pyMLG.extension` selbst)
3. **Save Settings and Reload**

### 2.2 openpyxl bereitstellen

openpyxl ist in pyRevit nicht enthalten und muss einmalig installiert werden.
Empfohlen wird die Installation direkt in die Extension – dann wandert die
Bibliothek mit dem Repository mit und jeder Arbeitsplatz hat sie automatisch:

```bat
pip install --target "<Pfad>\pyMLG.extension\site-packages" openpyxl
```

Das Skript durchsucht beim Start automatisch:

1. den normalen `sys.path` der pyRevit-CPython-Engine
2. `pyMLG.extension\site-packages`
3. `.venv\Lib\site-packages` neben dem Extension-Ordner (Entwicklungs-Setup)

**Welche Python-Version?** openpyxl ist reines Python, die Nebenversion ist
daher unkritisch – ein `pip` ab Python 3.8 genügt. Zur Orientierung:
pyRevit 5.x nutzt CPython 3.12, pyRevit 4.8.x nutzt CPython 3.8
(`pyrevit --version` zeigt die installierte pyRevit-Version).

**Alternative ohne Kopie in der Extension:** openpyxl an einen beliebigen Ort
installieren und den Ordner unter **pyRevit ▸ Settings ▸ Custom Search Paths**
eintragen (das ist der von pyRevit vorgesehene Weg, den `PYTHONPATH` zu
erweitern). Danach Revit neu starten.

Fehlt openpyxl, meldet das Werkzeug das beim Start mit der passenden
`pip`-Zeile inklusive Zielpfad – es stürzt nicht ab.

---

## 3. Export: Schedules nach Excel

1. Schaltfläche **Schedules nach Excel** anklicken
2. Eine oder mehrere Bauteillisten auswählen (Mehrfachauswahl möglich)
3. Speicherort bestätigen – vorgeschlagen werden der Ordner der Revit-Datei und
   der Dateiname `<Projektname>_<Bauteilliste>_<Datum>.xlsx`
4. Auf Wunsch wird die Datei anschließend in Excel geöffnet

### Aufbau der Excel-Datei

Je Bauteilliste entstehen **zwei Tabellenblätter** (für Excel unzulässige
Zeichen im Namen werden ersetzt, Länge auf 31 Zeichen gekürzt):

| Blatt | Inhalt | Zweck |
|---|---|---|
| **`<Bauteilliste>`** | eine Zeile **je Element** | zum Bearbeiten und Zurückschreiben |
| **`Ansicht - <Bauteilliste>`** (graue Registerkarte) | die Tabelle **genau wie in Revit** – Titel, Gruppenköpfe, zusammengefasste Zeilen, Summen, formatierte Werte | nur zum Lesen, wird beim Import ignoriert |

Warum zwei? Revit darf gleiche Elemente zu einer Zeile zusammenfassen („12 ×
Fahrradbügel A"). Eine solche Zeile lässt sich nicht eindeutig zurückschreiben.
Das Datenblatt bricht sie deshalb immer in einzelne Elemente auf, das
Ansichtsblatt zeigt die Originaldarstellung.

**Datenblatt:**

| Spalte | Inhalt | Hinweis |
|---|---|---|
| **A** | `UniqueId` | **ausgeblendet**, Schlüssel für den Rückimport – nicht ändern |
| **B** | `ElementId` | nur zur Orientierung, wird beim Import ignoriert |
| **C** | `Kategorie` | nur zur Orientierung |
| **D …** | je eine sichtbare Spalte der Bauteilliste | Überschrift wie in Revit |

Die Zeilen sind **so sortiert wie in Revit** (alle Sortier-/Gruppierfelder, auch
ausgeblendete, auf- oder absteigend, Texte „natürlich": EG 2 vor EG 10). Die
Kopfzeile ist fett und fixiert, die Spaltenbreiten sind angepasst, ein
Autofilter ist gesetzt.

**Farbcodierung der Zellen:**

| Darstellung | Bedeutung |
|---|---|
| weiß | Instanzparameter, frei editierbar |
| **orange** | Typparameter – eine Änderung wirkt auf **alle Instanzen** dieses Typs |
| **grau** | schreibgeschützt, berechnet (z. B. Fläche, Volumen, Anzahl, Formeln) oder am Element nicht vorhanden → wird beim Import **ignoriert** |

Zusätzlich ist ein Blattschutz ohne Passwort aktiv: Graue Zellen sind gesperrt,
editierbare Zellen sind freigegeben. Der Schutz ist als Leitplanke gedacht und
kann in Excel jederzeit aufgehoben werden – am Verhalten des Imports ändert das
nichts, dieser prüft den Schreibschutz erneut am Modell.

Ein ausgeblendetes Blatt `_pyMLG_Meta` enthält die technische Zuordnung der
Spalten zu den Revit-Parametern (BuiltInParameter-Id bzw. Shared-Parameter-GUID).
**Bitte nicht löschen oder bearbeiten.** Fehlt es, funktioniert der Import
weiterhin, sucht die Parameter dann aber nur noch über den Spaltennamen.

### Welche Bauteillisten werden unterstützt?

**Alle** Bauteillisten, die im Projektbrowser auftauchen – unabhängig von
Gruppierung, Sortierung, Summen oder der Option *„Jede Instanz aufführen"*.

Die Elemente des Datenblatts werden **nicht** aus der Revit-Tabelle gelesen,
sondern über `FilteredElementCollector(doc, schedule.Id)` ermittelt. Damit sind
exakt die Elemente enthalten, die die Bauteilliste inklusive ihrer Filter
umfasst – egal, wie Revit sie darstellt.

Einige Fälle bringen Besonderheiten mit, auf die das Werkzeug nach dem Export
im Ausgabefenster hinweist:

| Fall | Datenblatt | Ansichtsblatt |
|---|---|---|
| Zusammengefasst („Jede Instanz aufführen" aus) | eine Zeile je Element | zusammengefasst wie in Revit |
| Gruppenköpfe, Zwischen- und Gesamtsummen | gleich sortiert, ohne Summenzeilen | vollständig |
| Materialauszug | eine Zeile je Element, Materialspalten grau | eine Zeile je Element und Material |
| Elemente aus verknüpften Modellen | nicht enthalten (nicht beschreibbar) | enthalten |
| Berechnete Spalten (Formel, Prozent) | leer und grau | mit Werten |
| Anzahl | immer 1 (eine Zeile = ein Element), grau | wie in Revit |
| Schlüsselliste | die Schlüssel selbst | wie in Revit |

Nicht angeboten werden nur Vorlagen, interne Schlüsseltext-Listen und
Revisionslisten von Plankopf-Familien – diese sind keine eigenständigen
Bauteillisten.

---

## 4. Import: Excel zurück lesen

1. Schaltfläche **Excel zurück lesen** anklicken
2. Die bearbeitete `.xlsx`-Datei auswählen
3. Das Werkzeug prüft zunächst nur (es wird noch nichts geschrieben) und zeigt,
   wie viele Werte an wie vielen Elementen geändert würden
4. Nach Bestätigung werden alle Änderungen geschrieben
5. Zum Abschluss erscheint eine Auswertung im pyRevit-Ausgabefenster

### Was beim Import passiert

* Es werden **nur Werte geschrieben, die sich tatsächlich unterscheiden**.
  Unveränderte Zellen erzeugen keine Modelländerung.
* Zahlenwerte werden von der Projekteinheit in die interne Revit-Einheit
  zurückgerechnet (Revit rechnet intern in Fuß-basierten Einheiten). In Excel
  wird also z. B. in Millimetern gearbeitet, nicht in Fuß.
  Sowohl `2750.5` als auch `2.750,5` werden akzeptiert.
* Ja/Nein-Parameter akzeptieren `Ja`/`Nein`, `Yes`/`No`, `Wahr`/`Falsch`,
  `X`, `1`/`0`.
* Parameter, die auf ein anderes Element verweisen (z. B. Ebene, Material),
  werden über den Namen aufgelöst. Geschrieben wird nur bei **genau einem**
  eindeutigen Treffer, sonst erscheint ein Fehlereintrag.
* Alle Änderungen laufen in **einer TransactionGroup** – ein einziges
  *Rückgängig* nimmt den kompletten Import zurück.

### Was übersprungen wird (ohne Abbruch)

| Situation | Verhalten |
|---|---|
| Element existiert nicht mehr | Zeile wird mit Grund protokolliert |
| Zelle grau / Parameter schreibgeschützt | Zelle wird ignoriert, gezählt |
| Wert nicht konvertierbar (z. B. Text in Zahlenspalte) | Fehlereintrag mit Element und Feldname |
| Element in der Zentraldatei von jemand anderem ausgeliehen | Zeile wird übersprungen, Besitzer wird genannt |
| Blatt `Ansicht - …` | wird still übersprungen (reine Lesekopie) |
| Tabellenblatt ohne `UniqueId` in Zelle A1 | Blatt wird mit Warnung übersprungen |
| Derselbe Typparameter in mehreren Zeilen mit **gleichem** Wert | wird einmal geschrieben |
| Derselbe Typparameter in mehreren Zeilen mit **verschiedenen** Werten | Fehlereintrag, nichts wird geschrieben |

Zeilen dürfen in Excel gelöscht, sortiert und gefiltert werden – gelöschte
Zeilen bedeuten schlicht „nicht importieren". Neue Zeilen ohne gültige UniqueId
werden ignoriert; das Werkzeug legt **keine** Elemente an.

---

## 5. Hinweise zu Arbeitsteilung und verknüpften Modellen

* **Zentraldatei/Worksharing:** Vor dem Schreiben wird geprüft, ob ein Element
  von einem anderen Benutzer ausgeliehen ist. Solche Zeilen werden übersprungen,
  statt die gesamte Transaktion scheitern zu lassen. Nach dem Import wie
  gewohnt synchronisieren.
* **Verknüpfte Modelle:** Es werden ausschließlich Elemente des aktiven
  Dokuments verarbeitet. Elemente aus Links lassen sich ohnehin nicht
  beschreiben und sind daher bewusst nicht enthalten.
* Löst eine Änderung eine Revit-Warnung aus (z. B. doppelte Werte bei als
  eindeutig markierten Parametern), erscheint der übliche Revit-Warndialog.

---

## 6. Fehlerbehebung

| Meldung | Ursache und Abhilfe |
|---|---|
| *„openpyxl fehlt"* | Abschnitt 2.2 ausführen. Die Meldung enthält den fertigen `pip`-Befehl mit Zielpfad. |
| *„Datei gesperrt"* | Die Excel-Datei ist geöffnet. Excel schließen und erneut versuchen. |
| Datenblatt hat mehr Zeilen als die Tabelle in Revit | Gewollt: Revit fasst Elemente zusammen, das Datenblatt listet jedes Element einzeln. Die zusammengefasste Darstellung steht im Blatt *„Ansicht - …"*. |
| *„Kein Metablatt gefunden"* | Das Blatt `_pyMLG_Meta` wurde gelöscht. Der Import läuft über die Spaltennamen weiter; für volle Eindeutigkeit neu exportieren. |
| *„Parameter am Element nicht gefunden"* | Der Parameter ist an diesem Element nicht (mehr) vorhanden, z. B. weil die Familie getauscht wurde. |
| *„… is not currently supported under CPython"* | Ein Dialog greift noch auf `pyrevit.forms` zu. Alle Dialoge müssen aus `schedule_sync.ui` kommen (siehe 1., Abschnitt *Dialoge unter CPython*). |
| *„This property must be set before runtime is initialized"* | Die CPython-Engine lässt sich nicht mehr initialisieren – meist Folge eines vorherigen harten Absturzes. **Revit komplett neu starten**; ein pyRevit-Reload genügt nicht, weil die Laufzeit im Revit-Prozess lebt. |
| *„Revit konnte den externen Befehl nicht abschließen … Object reference not set to an instance of an object"* | Ein `SystemExit` ist aus dem Skript herausgelaufen (`script.exit()` bzw. `sys.exit()`). Im Skript ausschließlich `return` verwenden. Tritt es dennoch auf: Traceback in `%TEMP%\pyMLG_ScheduleSync_Fehler.log`. |
| Das Werkzeug erscheint nicht im Ribbon | Extension-Pfad prüfen (2.1), anschließend **Reload** in den pyRevit-Einstellungen. |

---

## 7. Aufbau im Repository

```
pyMLG.extension/
├─ lib/schedule_sync/          gemeinsame Logik beider Schaltflächen
│  ├─ deps.py                  openpyxl laden / Installationshinweis
│  ├─ ui.py                    Dialoge (CPython-tauglich, ohne pyrevit.forms)
│  ├─ revit_helpers.py         Parameter, Einheiten, Worksharing
│  ├─ schedule_model.py        Bauteillisten lesen und prüfen
│  ├─ excel_io.py              Excel schreiben und lesen
│  └─ importer.py              Rückimport: Analyse und Anwendung
└─ pyMLG.tab/ScheduleSync.panel/
   └─ Excel.stack/
      ├─ Export.pushbutton/    script.py, bundle.yaml, icon.png
      └─ Import.pushbutton/    script.py, bundle.yaml, icon.png
```

Der Ordner `lib` wird von pyRevit automatisch in den Suchpfad gelegt, beide
Skripte greifen also auf dieselben Module zu.
