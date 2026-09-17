# -*- coding: utf-8 -*-
"""Sprache der pyMLG-Werkzeuge: Deutsch, Englisch oder Spanisch.

Die Sprache folgt der Revit-Sprache (Application.Language):
    German  -> Deutsch
    Spanish -> Spanisch
    sonst   -> Englisch

Übersteuern (z.B. zum Testen): Umgebungsvariable PYMLG_SPRACHE=de|en|es.

Texte stehen direkt im Code, deutsch zuerst:

    from mlg_sprache import t
    t(u"Schließen", u"Close", u"Cerrar")

Für XAML: Platzhalter {{schluessel}} im Text und ein Wörterbuch
{schluessel: (de, en, es)} -> uebersetze_xaml(). Doppelte geschweifte
Klammern kommen in XAML sonst nicht vor ({StaticResource ...} hat eine).

Hinweis: Die Ribbon-Beschriftungen (bundle.yaml) wählt pyRevit selbst -
nach seiner Einstellung "Sprache" (user_locale), nicht nach Revit.

Läuft unter IronPython und CPython.
"""

import os
import re

DE = "de"
EN = "en"
ES = "es"

_sprache = []


def sprache():
    """'de', 'en' oder 'es' - einmal ermittelt, dann zwischengespeichert."""
    if not _sprache:
        _sprache.append(_ermittle())
    return _sprache[0]


def setze_sprache(kuerzel):
    """Nur für Tests: Sprache fest einstellen."""
    del _sprache[:]
    _sprache.append(kuerzel)


def _aus_name(name):
    name = (name or "").lower()
    if "german" in name or name.startswith("de"):
        return DE
    if "spanish" in name or name.startswith("es"):
        return ES
    return EN


def _ermittle():
    vorgabe = os.environ.get("PYMLG_SPRACHE", "").strip().lower()[:2]
    if vorgabe in (DE, EN, ES):
        return vorgabe
    try:
        from pyrevit import HOST_APP
        return _aus_name(str(HOST_APP.language))
    except Exception:
        return EN


def t(de, en=None, es=None):
    """Text in der aktuellen Sprache. Fehlt eine Übersetzung: Englisch,
    dann Deutsch."""
    kuerzel = sprache()
    if kuerzel == ES and es:
        return es
    if kuerzel in (EN, ES) and en:
        return en
    return de


def tt(texte):
    """t() für ein Tupel (de, en, es)."""
    return t(*texte)


def _xml(text):
    return (text.replace(u"&", u"&amp;").replace(u"<", u"&lt;")
            .replace(u">", u"&gt;").replace(u'"', u"&quot;"))


def uebersetze_xaml(xaml, texte):
    """Ersetzt {{schluessel}} durch den übersetzten, XML-sicheren Text.
    Unbekannte Schlüssel bleiben sichtbar stehen (fällt beim Testen auf)."""
    def ersetzen(treffer):
        schluessel = treffer.group(1)
        if schluessel not in texte:
            return treffer.group(0)
        return _xml(tt(texte[schluessel]))
    return re.sub(r"\{\{(\w+)\}\}", ersetzen, xaml)
