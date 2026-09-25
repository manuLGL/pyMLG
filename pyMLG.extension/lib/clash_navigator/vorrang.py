# -*- coding: utf-8 -*-
"""Vorrangregeln: welches Gewerk bleibt liegen, welches weicht aus?

Die Reihenfolge ist einstellbar, oben steht, was liegen bleibt:

    Bau > Abwasser > Lüftung > Sprinkler > Heizung > Kälte > Trinkwasser
        > Sonstige > Elektro

Das Gewerk eines Elements ergibt sich aus
  1. seiner Kategorie (Kanal -> Lüftung, Kabeltrasse -> Elektro, Wand/Decke/
     Träger -> Bau),
  2. bei Rohren aus Stichwörtern im Namen des Rohrsystems ("REFRIG_Retorno"
     -> Kälte) - Firmen benennen ihre Systeme meist sprechend,
  3. sonst aus der Systemklassifizierung von Revit (PipeSystemType).

Ohne Revit - offline testbar. Die Revit-Seite liefert nur Kategorie,
Klassifizierung (Name des Enum-Werts) und Systemname.
"""

import re

from mlg_sprache import tt

BAU = u"bau"
ABWASSER = u"abwasser"
LUEFTUNG = u"lueftung"
SPRINKLER = u"sprinkler"
HEIZUNG = u"heizung"
KAELTE = u"kaelte"
TRINKWASSER = u"trinkwasser"
SONSTIGE = u"sonstige"
ELEKTRO = u"elektro"

STANDARD = (BAU, ABWASSER, LUEFTUNG, SPRINKLER, HEIZUNG, KAELTE,
            TRINKWASSER, SONSTIGE, ELEKTRO)

TEXTE = {
    BAU: (u"Bau (Wand, Decke, Träger)", u"Structure (wall, floor, beam)",
          u"Obra (muro, forjado, viga)"),
    ABWASSER: (u"Abwasser (Gefälle)", u"Drainage (gravity)",
               u"Saneamiento (pendiente)"),
    LUEFTUNG: (u"Lüftung", u"Ventilation", u"Ventilación"),
    SPRINKLER: (u"Sprinkler / Löschwasser", u"Sprinkler / fire",
                u"Rociadores / PCI"),
    HEIZUNG: (u"Heizung", u"Heating", u"Calefacción"),
    KAELTE: (u"Kälte", u"Cooling", u"Refrigeración"),
    TRINKWASSER: (u"Trinkwasser", u"Domestic water",
                  u"Agua sanitaria (AFS/ACS)"),
    SONSTIGE: (u"Sonstige Rohre", u"Other pipes", u"Otras tuberías"),
    ELEKTRO: (u"Elektro (Trassen, Rohre)", u"Electrical (trays, conduit)",
              u"Electricidad (bandejas, tubos)"),
}

# Kategorien (Namen der BuiltInCategory) mit festem Gewerk
_KATEGORIEN = {
    u"OST_DuctCurves": LUEFTUNG, u"OST_DuctFitting": LUEFTUNG,
    u"OST_DuctAccessory": LUEFTUNG, u"OST_FlexDuctCurves": LUEFTUNG,
    u"OST_DuctTerminal": LUEFTUNG,
    u"OST_CableTray": ELEKTRO, u"OST_CableTrayFitting": ELEKTRO,
    u"OST_Conduit": ELEKTRO, u"OST_ConduitFitting": ELEKTRO,
    u"OST_Sprinklers": SPRINKLER,
    u"OST_Walls": BAU, u"OST_Floors": BAU, u"OST_Roofs": BAU,
    u"OST_Ceilings": BAU, u"OST_StructuralFraming": BAU,
    u"OST_StructuralColumns": BAU, u"OST_Columns": BAU,
    u"OST_StructuralFoundation": BAU, u"OST_Stairs": BAU,
}

# Stichwörter im Systemnamen - Reihenfolge zählt: "KALTWASSER" ist
# Trinkwasser, nicht Kälte; deshalb steht Trinkwasser vor Kälte.
# Nicht "VENT": auf Spanisch heisst Lüftung "ventilación".
_STICHWOERTER = (
    (SPRINKLER, (u"SPRINK", u"ROCIA", u"PCI", u"INCEND", u"FIRE",
                 u"LOESCH", u"LOSCH", u"BIE")),
    (ABWASSER, (u"ABWASSER", u"SCHMUTZ", u"REGENW", u"FALLROHR", u"DRAIN",
                u"SANEA", u"PLUVIA", u"FECAL", u"RESIDUAL", u"WASTE",
                u"SEWER", u"SANITARY")),
    (TRINKWASSER, (u"TRINK", u"KALTWASSER", u"WARMWASSER", u"ZIRKUL",
                   u"AFS", u"ACS", u"POTABLE", u"DOMESTIC", u"TWK", u"TWW")),
    (KAELTE, (u"KAELTE", u"KALTE", u"REFRIG", u"FRIO", u"CHILL", u"COOL",
              u"KUEHL", u"KUHL", u"CLIMA")),
    (HEIZUNG, (u"HEIZ", u"CALEF", u"HEAT", u"CALOR", u"CALDERA")),
)

# Systemklassifizierung von Revit (Name des PipeSystemType-Werts)
_KLASSIFIZIERUNG = {
    u"Sanitary": ABWASSER, u"Vent": ABWASSER,
    u"DomesticHotWater": TRINKWASSER, u"DomesticColdWater": TRINKWASSER,
    u"SupplyHydronic": HEIZUNG, u"ReturnHydronic": HEIZUNG,
    u"FireProtectWet": SPRINKLER, u"FireProtectDry": SPRINKLER,
    u"FireProtectPreaction": SPRINKLER, u"FireProtectOther": SPRINKLER,
}


def _grossbuchstaben(text):
    text = (text or u"").upper()
    for alt, neu in ((u"Ä", u"AE"), (u"Ö", u"OE"), (u"Ü", u"UE"),
                     (u"Á", u"A"), (u"É", u"E"), (u"Í", u"I"),
                     (u"Ó", u"O"), (u"Ú", u"U"), (u"Ñ", u"N")):
        text = text.replace(alt, neu)
    return text


def gewerk(kategorie=u"", klassifizierung=u"", systemname=u""):
    """Gewerk aus Kategorie, Klassifizierung und Systemname."""
    fest = _KATEGORIEN.get(kategorie or u"")
    if fest is not None:
        return fest
    name = _grossbuchstaben(systemname)
    if name:
        for code, woerter in _STICHWOERTER:
            for wort in woerter:
                # kurze Kürzel (AFS, PCI ...) nur als eigenes Wort
                if len(wort) <= 4:
                    if re.search(r"(^|[^A-Z])" + wort + r"($|[^A-Z])", name):
                        return code
                elif wort in name:
                    return code
    return _KLASSIFIZIERUNG.get(klassifizierung or u"", SONSTIGE)


def text(code):
    return tt(TEXTE.get(code, (code, code, code)))


def reihenfolge(gespeichert):
    """Gespeicherte Reihenfolge, ergänzt um neue bzw. bereinigt von
    unbekannten Gewerken."""
    liste = [code for code in (gespeichert or []) if code in STANDARD]
    for code in STANDARD:
        if code not in liste:
            liste.append(code)
    return liste


def rang(code, liste):
    """Je kleiner, desto eher bleibt das Gewerk liegen."""
    try:
        return liste.index(code)
    except ValueError:
        return len(liste)


def weicht_aus(code_a, code_b, liste):
    """True, wenn a vor b ausweichen sollte (niedrigerer Vorrang)."""
    return rang(code_a, liste) > rang(code_b, liste)
