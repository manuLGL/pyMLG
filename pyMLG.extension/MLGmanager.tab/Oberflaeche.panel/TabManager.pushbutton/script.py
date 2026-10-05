# -*- coding: utf-8 -*-
__title__ = "Tab\nManager"
__doc__ = "Verwalte sichtbare Ribbon-Tabs"
__author__ = "Manuel"

import json
import os

import clr

clr.AddReference('AdWindows')
from Autodesk.Windows import ComponentManager  # noqa: E402

from pyrevit import forms  # noqa: E402
from mlg_sprache import t  # noqa: E402

TITEL = t(u"Tab Manager", u"Tab Manager", u"Gestor de fichas")

config_file = os.path.join(os.getenv('APPDATA'), 'pyRevit', 'ribbon_settings.json')

# Gespeichert wird je Tab die interne Id, nicht der angezeigte Titel: der
# Titel hängt von der Revit-Sprache ab ("Architektur" / "Architecture"),
# die Id nicht. Geschützte Tabs werden über Id oder Titel erkannt (bei
# Add-in-Tabs sind beide gleich).
PROTECTED_TABS = ["pyMLG", "pyRevit"]


class TabGruppe(object):
    """Alle Ribbon-Tabs mit derselben Id (Revit legt manche doppelt an)."""

    def __init__(self, tab_id):
        self.tab_id = tab_id
        self.tabs = []
        self.name = u""

    @property
    def titel(self):
        for tab in self.tabs:
            if tab.Title:
                return tab.Title
        return self.tab_id

    @property
    def sichtbar(self):
        return any(tab.IsVisible for tab in self.tabs)


def load_settings():
    if not os.path.exists(config_file):
        return {}
    try:
        with open(config_file, 'r') as f:
            return json.load(f)
    except Exception as e:
        forms.alert(t(u"Einstellungen konnten nicht geladen werden.", u"Settings could not be loaded.", u"No se pudo cargar la configuración."),
                    sub_msg=str(e), title=TITEL)
        return {}


def save_settings(settings):
    try:
        ordner = os.path.dirname(config_file)
        if not os.path.isdir(ordner):
            os.makedirs(ordner)
        with open(config_file, 'w') as f:
            json.dump(settings, f, indent=2)
    except Exception as e:
        forms.alert(t(u"Einstellungen konnten nicht gespeichert werden.", u"Settings could not be saved.", u"No se pudo guardar la configuración."),
                    sub_msg=str(e), title=TITEL)


def tab_gruppen():
    gruppen = {}
    for tab in ComponentManager.Ribbon.Tabs:
        # Kontextabhängige Tabs ("Ändern | Wände") blendet Revit selbst ein
        if getattr(tab, "IsContextualTab", False):
            continue
        tab_id = tab.Id or tab.Title
        if not tab_id:
            continue
        gruppen.setdefault(tab_id, TabGruppe(tab_id)).tabs.append(tab)

    # Anzeigename = Titel; bei gleichem Titel verschiedener Tabs die Id anhängen
    anzahl = {}
    for gruppe in gruppen.values():
        anzahl[gruppe.titel] = anzahl.get(gruppe.titel, 0) + 1
    for gruppe in gruppen.values():
        gruppe.name = gruppe.titel
        if anzahl[gruppe.titel] > 1:
            gruppe.name = u"{} ({})".format(gruppe.titel, gruppe.tab_id)
    return gruppen


def ist_geschuetzt(gruppe):
    return gruppe.tab_id in PROTECTED_TABS or gruppe.titel in PROTECTED_TABS


def gespeichert_sichtbar(settings, gruppe):
    """True/False aus den Einstellungen, None wenn der Tab dort fehlt.

    Ältere Einstellungsdateien sind noch nach Titel gespeichert.
    """
    if gruppe.tab_id in settings:
        return settings[gruppe.tab_id]
    return settings.get(gruppe.titel)


def main():
    gruppen = tab_gruppen()
    waehlbar = sorted((g for g in gruppen.values() if not ist_geschuetzt(g)),
                      key=lambda g: g.name.lower())
    settings = load_settings()

    eintraege = []
    for gruppe in waehlbar:
        sichtbar = gespeichert_sichtbar(settings, gruppe)
        if sichtbar is None:
            sichtbar = gruppe.sichtbar
        eintraege.append(forms.TemplateListItem(gruppe, checked=bool(sichtbar)))

    auswahl = forms.SelectFromList.show(
        eintraege,
        title=t(u"Wähle sichtbare Ribbon-Tabs", u"Choose visible ribbon tabs", u"Elija las fichas visibles de la cinta"),
        width=500,
        height=600,
        button_name=t(u"Anwenden", u"Apply", u"Aplicar"),
        multiselect=True,
    )
    if auswahl is None:
        return

    gewaehlt = set(g.tab_id for g in auswahl)
    neue_settings = {}
    for gruppe in gruppen.values():
        sichtbar = ist_geschuetzt(gruppe) or gruppe.tab_id in gewaehlt
        neue_settings[gruppe.tab_id] = sichtbar
        for tab in gruppe.tabs:
            tab.IsVisible = sichtbar
    save_settings(neue_settings)


main()
