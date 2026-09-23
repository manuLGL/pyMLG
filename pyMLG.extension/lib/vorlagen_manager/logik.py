# -*- coding: utf-8 -*-
"""Revit-freie Logik des Ansichtsvorlagen-Managers.

Hier steht alles, was ohne Revit prüfbar ist: die Wortsuche, das
Zusammenfassen gleicher/verschiedener Werte mehrerer markierter Vorlagen,
die Reihenfolge der Parameterzeilen und die Namensprüfung.

Reihenfolge der Parameter: Der native Dialog "Ansichtsvorlagen" zeigt sie in
einer festen Reihenfolge, die die API nicht liefert (GetTemplateParameterIds
ist unsortiert). Deshalb sortiert RANG die geläufigen eingebauten Parameter
wie im Dialog; alle übrigen eingebauten folgen alphabetisch, Projekt- und
gemeinsam genutzte Parameter zuletzt - dort stehen sie auch im Dialog.
"""

import re

from mlg_sprache import t

# Gruppen für die Sortierung
GRUPPE_BEKANNT = 0        # eingebaut, Reihenfolge wie im nativen Dialog
GRUPPE_EINGEBAUT = 1      # eingebaut, hier nicht einsortiert
GRUPPE_PROJEKT = 2        # Projekt- bzw. gemeinsam genutzter Parameter

# In Ansichts- und Vorlagennamen verboten (wie im nativen Dialog)
VERBOTENE_ZEICHEN = u"\\:{}[]|;<>?`~"


class _Verschieden(object):
    """Platzhalter: Die markierten Vorlagen haben unterschiedliche Werte."""

    def __repr__(self):
        return "<verschieden>"


VERSCHIEDEN = _Verschieden()


# Reihenfolge wie im Dialog "Ansichtsvorlagen"; Namen der BuiltInParameter
RANG = [
    "VIEW_SCALE",
    "VIEW_SCALE_PULLDOWN_METRIC",
    "VIEW_SCALE_PULLDOWN_IMPERIAL",
    "VIEW_MODEL_DISPLAY_MODE",
    "VIEW_DETAIL_LEVEL",
    "VIEW_PARTS_VISIBILITY",
    "VIS_GRAPHICS_MODEL",
    "VIS_GRAPHICS_ANNOTATION",
    "VIS_GRAPHICS_ANALYTICAL_MODEL",
    "VIS_GRAPHICS_IMPORT",
    "VIS_GRAPHICS_FILTERS",
    "VIS_GRAPHICS_WORKSETS",
    "VIS_GRAPHICS_RVT_LINKS",
    "VIS_GRAPHICS_POINT_CLOUDS",
    "VIS_GRAPHICS_COORDINATION_MODEL",
    "VIS_GRAPHICS_DESIGNOPTIONS",
    "MODEL_GRAPHICS_STYLE",
    "GRAPHIC_DISPLAY_OPTIONS_MODEL",
    "GRAPHIC_DISPLAY_OPTIONS_SHADOWS",
    "GRAPHIC_DISPLAY_OPTIONS_SKETCHY_LINES",
    "GRAPHIC_DISPLAY_OPTIONS_LIGHTING",
    "GRAPHIC_DISPLAY_OPTIONS_PHOTO_EXPOSURE",
    "GRAPHIC_DISPLAY_OPTIONS_BACKGROUND",
    "GRAPHIC_DISPLAY_OPTIONS_FOG",
    "VIEW_UNDERLAY_ORIENTATION",
    "VIEW_UNDERLAY_BOTTOM_ID",
    "VIEW_UNDERLAY_TOP_ID",
    "PLAN_VIEW_RANGE",
    "PLAN_VIEW_NORTH",
    "VIEW_PHASE_FILTER",
    "VIEW_PHASE",
    "VIEW_DISCIPLINE",
    "VIEW_SHOW_HIDDEN_LINES",
    "COLOR_SCHEME_LOCATION",
    "VIEW_BACK_CLIPPING",
    "VIEWER_VOLUME_OF_INTEREST_CROP",
    "VIEW_SHOW_GRIDS",
    "VIEW_CLEAN_JOINS",
    "VIEW_DESCRIPTION",
]

_RANG = dict((name, nummer) for nummer, name in enumerate(RANG))


def gruppe_und_rang(bip_name, eingebaut):
    """(Gruppe, Rang) für die Sortierung einer Parameterzeile.

    bip_name   Name des BuiltInParameter oder None
    eingebaut  True bei eingebauten Parametern
    """
    if bip_name and bip_name in _RANG:
        return GRUPPE_BEKANNT, _RANG[bip_name]
    if eingebaut:
        return GRUPPE_EINGEBAUT, 0
    return GRUPPE_PROJEKT, 0


def passt(suchtext, woerter):
    """True, wenn alle Suchwörter im (kleingeschriebenen) Text vorkommen."""
    return all(wort in suchtext for wort in woerter)


_ZAHL = re.compile(r"(\d+)")


def natuerlich(text):
    """Sortierschlüssel, der Zahlen als Zahlen vergleicht:
    "Plan 2" kommt vor "Plan 10"."""
    teile = _ZAHL.split((text or u"").lower())
    return [(1, int(teil)) if teil.isdigit() else (0, teil)
            for teil in teile if teil != u""]


def vereine(werte):
    """Einen Wert aus mehreren machen: der Wert selbst oder VERSCHIEDEN.

    Eine leere Folge ergibt None.
    """
    erster = None
    gesetzt = False
    for wert in werte:
        if not gesetzt:
            erster, gesetzt = wert, True
        elif wert != erster:
            return VERSCHIEDEN
    return erster if gesetzt else None


def vereine_flagge(flaggen):
    """True, False oder None (gemischt) - für dreiwertige Kontrollkästchen."""
    wert = vereine(flaggen)
    return None if wert is VERSCHIEDEN else wert


def pruefe_name(vorhandene, name, alter_name=None):
    """Namen bereinigen und prüfen. Wirft ValueError mit Klartext.

    vorhandene   Folge bereits vergebener Namen
    alter_name   beim Umbenennen der bisherige Name (darf bleiben)
    """
    name = (name or u"").strip()
    if not name:
        raise ValueError(t(u"Der Name darf nicht leer sein.",
                           u"The name must not be empty.",
                           u"El nombre no puede estar vacío."))
    schlecht = sorted(set(z for z in name if z in VERBOTENE_ZEICHEN))
    if schlecht:
        raise ValueError(t(u"Nicht erlaubte Zeichen: %s",
                           u"Characters not allowed: %s",
                           u"Caracteres no permitidos: %s")
                         % u" ".join(schlecht))
    klein = name.lower()
    alt = (alter_name or u"").lower()
    for vorhanden in vorhandene:
        if vorhanden.lower() == klein and vorhanden.lower() != alt:
            raise ValueError(t(u"Der Name \"%s\" ist bereits vergeben.",
                               u"The name \"%s\" is already in use.",
                               u"El nombre \"%s\" ya está en uso.") % name)
    return name


def freier_name(vorhandene, basis):
    """basis, sonst "basis 2", "basis 3", ... - der erste freie Name."""
    vergeben = set((v or u"").lower() for v in vorhandene)
    if basis.lower() not in vergeben:
        return basis
    nummer = 2
    while (u"%s %d" % (basis, nummer)).lower() in vergeben:
        nummer += 1
    return u"%s %d" % (basis, nummer)
