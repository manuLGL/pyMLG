# -*- coding: utf-8 -*-
"""Revit-freie Logik von TransferSingle.

Baum:  Alle > Gruppe (z.B. "Wände – Typen", "Ansichtsvorlagen") > Element.
Die Knoten stammen aus filter_more.baum; Markierung, Zählung und Zustand
funktionieren dort unabhängig von der Tiefe.
"""

import re

from filter_more import baum as bm

WURZEL = 0
GRUPPE = 1
ELEMENT = 2

# Umbenennen
PRAEFIX = "praefix"
SUFFIX = "suffix"
GROSS = "gross"
KLEIN = "klein"
ERSTER_GROSS = "erster_gross"
ERSETZEN = "ersetzen"
ZAHLEN = "zahlen"

_WORT = re.compile(u"[^\\W\\d_]+", re.UNICODE)
_ZAHL = re.compile(r"\d+")


def baue(datensaetze, wurzelname=u"Alle"):
    """[(Gruppe, Bezeichnung, Id)] -> Wurzelknoten (Knoten aus filter_more)."""
    wurzel = bm.Knoten(wurzelname, WURZEL)
    for gruppe, bezeichnung, element_id in datensaetze:
        knoten = wurzel.kind(gruppe, GRUPPE)
        blatt = bm.Knoten(bezeichnung, ELEMENT)
        blatt.element_id = element_id
        blatt.ids.append(element_id)
        knoten.kinder.append(blatt)
        knoten.ids.append(element_id)
        wurzel.ids.append(element_id)
    for knoten in [wurzel] + wurzel.kinder:
        knoten.kinder.sort(key=lambda k: (bm.natuerlich(k.name),
                                          k.element_id or 0))
    return wurzel


def blaetter(wurzel):
    return [b for g in wurzel.kinder for b in g.kinder]


def naechster_treffer(wurzel, text, nach_id=None):
    """(Gruppe, Blatt) des nächsten Elements, dessen Name den Text enthält -
    nach dem Element nach_id, am Ende wieder von vorn. None ohne Treffer."""
    text = (text or u"").strip().casefold()
    if not text:
        return None
    paare = [(g, b) for g in wurzel.kinder for b in g.kinder]
    start = 0
    if nach_id is not None:
        for i, (_g, b) in enumerate(paare):
            if b.element_id == nach_id:
                start = i + 1
                break
    for i in range(len(paare)):
        gruppe, blatt = paare[(start + i) % len(paare)]
        if text in blatt.name.casefold():
            return gruppe, blatt
    return None


def erster_gross(name):
    """Jedes Wort mit großem Anfangsbuchstaben, Rest klein."""
    return _WORT.sub(lambda m: m.group(0)[:1].upper() + m.group(0)[1:].lower(),
                     name)


def ersetze_zahl(name, alt, neu):
    """Ganze Zahlen mit dem Wert alt durch neu ersetzen. Führende Nullen der
    Fundstelle bleiben erhalten ("EG 01" mit 1 -> 2 wird "EG 02")."""
    alt, neu = (alt or u"").strip(), (neu or u"").strip()
    if not alt.isdigit() or not neu.isdigit():
        raise ValueError(u"ganze Zahlen erwartet")
    ziel = int(alt)

    def ersetzen(treffer):
        teil = treffer.group(0)
        if int(teil) != ziel:
            return teil
        if len(teil) > 1 and teil.startswith(u"0"):
            return neu.zfill(len(teil))
        return neu
    return _ZAHL.sub(ersetzen, name)


def neuer_name(name, aktion, a=u"", b=u""):
    name = name or u""
    if aktion == PRAEFIX:
        return (a or u"") + name
    if aktion == SUFFIX:
        return name + (a or u"")
    if aktion == GROSS:
        return name.upper()
    if aktion == KLEIN:
        return name.lower()
    if aktion == ERSTER_GROSS:
        return erster_gross(name)
    if aktion == ERSETZEN:
        return name.replace(a, b or u"") if a else name
    if aktion == ZAHLEN:
        return ersetze_zahl(name, a, b)
    raise ValueError(aktion)


def freie_nummer(nummer, vergeben):
    """Plannummer, die im Ziel noch frei ist: 'A101', sonst 'A101-2', ...
    Die gefundene Nummer wird in vergeben eingetragen."""
    kandidat = nummer
    zaehler = 2
    while kandidat.casefold() in vergeben:
        kandidat = u"%s-%d" % (nummer, zaehler)
        zaehler += 1
    vergeben.add(kandidat.casefold())
    return kandidat
