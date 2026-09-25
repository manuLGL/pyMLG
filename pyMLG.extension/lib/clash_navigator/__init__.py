# -*- coding: utf-8 -*-
"""Clash Navigator: Kollisionen aus Navisworks in Revit abarbeiten.

Ein andockbares Panel (IronPython, registriert in startup.py) liest einen
Kollisionsbericht aus Navisworks (XML oder CSV), zeigt die Kollisionen als
Baum oder Tabelle und legt beim Anklicken eine Schnittbox um die beteiligten
Elemente - auch um Elemente aus Verknüpfungen.

bericht.py   XML- und CSV-Bericht lesen (Revit-frei, offline testbar)
logik.py     Status, Gruppierung, Filter, Kastenrechnung (Revit-frei)
speicher.py  Bewertungen als JSON neben dem Bericht, Rückgängig (Revit-frei)
revit.py     Elemente finden, Schnittbox, Clash-Ansicht, Auswahl
panel.py     das andockbare Panel (WPF) und das ExternalEvent

NWD/NWF/NWC lassen sich ohne Navisworks nicht lesen - in Navisworks unter
Clash Detective > Berichte das Format "XML" wählen und bei den Inhalten
"Element-ID" bzw. die Elementeigenschaften einschliessen.
"""

# Fest vergebene Kennung des andockbaren Panels - nie ändern, Revit merkt sich
# darunter Lage und Grösse.
PANEL_ID = "5c1f3a52-8d8e-4f6b-9a61-2f0f7c3d9b17"
