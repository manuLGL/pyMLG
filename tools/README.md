# tools

Hilfsskripte, die **nicht** Teil der pyRevit-Extension sind.

## make_icons.py

Erzeugt alle `icon.png` der Werkzeuge (96x96 RGBA) aus Code – ohne externe
Bibliotheken, gezeichnet mit 4x-Supersampling.

```bat
python tools\make_icons.py pyMLG.extension\pyMLG.tab
python tools\make_icons.py pyMLG.extension\MLGmanager.tab
python tools\make_icons.py pyMLG.extension\MLGplans.tab
```

Optional lassen sich einzelne Werkzeuge neu rendern:

```bat
python tools\make_icons.py pyMLG.extension\pyMLG.tab WorksetON TagDistance
```

Bildsprache: dunkelblaue Kontur, weisses "Papier", eine Akzentfarbe je
Funktion (blau = Ansichten/Plaene, gruen = hinzufuegen/an, orange = aendern,
rot = aus, violett = Phasen). Bewusst wenige grosse Formen, damit die Symbole
im Ribbon bei 32 px lesbar bleiben.

Ein neues Werkzeug bekommt eine Zeichenfunktion und einen Eintrag im Dict
`SYMBOLE` (Reiter pyMLG), `MANAGER` (Reiter MLGmanager) bzw. `PLAENE`
(Reiter MLGplans) am Ende der Datei.
