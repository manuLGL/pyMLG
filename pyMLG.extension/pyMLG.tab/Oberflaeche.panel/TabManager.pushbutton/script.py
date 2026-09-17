# -*- coding: utf-8 -*-
__title__ = "Tab\nManager"
__doc__ = "Verwalte sichtbare Ribbon-Tabs"
__author__ = "Manuel"

import clr
import sys
import os
import json

try:
    clr.AddReference('AdWindows')
    from Autodesk.Windows import ComponentManager
except Exception as e:
    print("FEHLER beim Laden von AdWindows:")
    print(str(e))
    sys.exit()

from pyrevit import forms

config_file = os.path.join(
    os.getenv('APPDATA'),
    'pyRevit',
    'ribbon_settings.json'
)

PROTECTED_TABS = ["pyMLG", "pyRevit"]


def load_settings():
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r') as f:
                settings = json.load(f)
                print("✓ Settings geladen aus: " + config_file)
                return settings
        except Exception as e:
            print("⚠ Fehler beim Laden: " + str(e))
            return {}
    print("→ Keine Settings-Datei gefunden")
    return {}


def save_settings(settings):
    try:
        os.makedirs(os.path.dirname(config_file), exist_ok=True)
        with open(config_file, 'w') as f:
            json.dump(settings, f, indent=2)
        print("✓ Settings gespeichert")
    except Exception as e:
        print("⚠ Fehler beim Speichern: " + str(e))


try:
    ribbon = ComponentManager.Ribbon
except Exception as e:
    print("FEHLER: Kann Ribbon nicht laden!")
    print(str(e))
    sys.exit()

tab_groups = {}
for tab in ribbon.Tabs:
    title = tab.Title
    if title not in tab_groups:
        tab_groups[title] = []
    tab_groups[title].append(tab)

all_tab_names = sorted(tab_groups.keys())
selectable_tab_names = [name for name in all_tab_names if name not in PROTECTED_TABS]

print("\n=== DEBUG INFO ===")
print("Alle Tabs: " + str(len(all_tab_names)))
print("Wählbare Tabs: " + str(len(selectable_tab_names)))
print("Geschützte Tabs: " + str(PROTECTED_TABS))

saved_settings = load_settings()

print("\n=== GESPEICHERTE SETTINGS ===")
if saved_settings:
    for name, visible in saved_settings.items():
        if visible:
            print("  ✓ " + name)
else:
    print("  (keine)")

preselected = []

if saved_settings:
    print("\n=== NUTZE GESPEICHERTE SETTINGS ===")
    for tab_name in selectable_tab_names:
        if saved_settings.get(tab_name, False):
            preselected.append(tab_name)
            print("  → Vorauswahl: " + tab_name)
else:
    print("\n=== NUTZE AKTUELLEN STATUS ===")
    for tab_name in selectable_tab_names:
        is_visible = any(tab.IsVisible for tab in tab_groups[tab_name])
        if is_visible:
            preselected.append(tab_name)
            print("  → Vorauswahl: " + tab_name + " (aktuell sichtbar)")

print("\n=== VORAUSWAHL FINAL ===")
print("Anzahl: " + str(len(preselected)))
for name in preselected:
    print("  ✓ " + name)

print("\n→ Öffne GUI...")

# Alternative: Liste mit Checkboxen erstellen
from pyrevit import forms


class TabItem:
    def __init__(self, name, checked=False):
        self.name = name
        self.checked = checked

    def __str__(self):
        return self.name

    def __repr__(self):
        return self.name


# Items erstellen mit Vorauswahl
tab_items = []
for tab_name in selectable_tab_names:
    is_checked = tab_name in preselected
    item = TabItem(tab_name, is_checked)
    tab_items.append(item)
    if is_checked:
        print("  PRE-CHECK: " + tab_name)

# Alte Methode (funktioniert nicht immer)
# selected_tabs = forms.SelectFromList.show(...)

# Neue Methode mit CommandSwitchWindow
from pyrevit.forms import SelectFromList

selected_tabs = SelectFromList.show(
    context=tab_items,
    title='Wähle sichtbare Ribbon-Tabs',
    width=500,
    height=600,
    button_name='Anwenden',
    multiselect=True
)

# Konvertiere zurück zu Namen
if selected_tabs:
    selected_tabs = [item.name for item in selected_tabs]

print("\n=== USER AUSWAHL ===")
if selected_tabs is not None:
    print("Ausgewählt: " + str(len(selected_tabs)))
    for name in selected_tabs:
        print("  ✓ " + name)

    new_settings = {}
    for name in selectable_tab_names:
        new_settings[name] = (name in selected_tabs)

    for protected in PROTECTED_TABS:
        if protected in tab_groups:
            new_settings[protected] = True

    for tab_name in all_tab_names:
        should_be_visible = new_settings.get(tab_name, True)
        for tab in tab_groups[tab_name]:
            tab.IsVisible = should_be_visible

    save_settings(new_settings)

    print("\n✓ FERTIG!")
else:
    print("→ Abgebrochen")