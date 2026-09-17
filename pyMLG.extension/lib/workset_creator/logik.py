# -*- coding: utf-8 -*-
"""Revit-freie Logik des Workset-Creators.

Namensvergleiche ignorieren Groß-/Kleinschreibung: "Wände" und "WÄNDE" als
zwei Worksets sind praktisch nie gewollt. Ob Revit einen Namen wirklich
annimmt, prüft beim Erstellen zusätzlich WorksetTable.IsWorksetNameUnique.
"""

import re

# Zeichen, die Revit in Namen ablehnt (wie NamingUtils.IsValidName)
VERBOTENE_ZEICHEN = u'\\:{}[]|;<>?`~'

NEU = u"neu"
UMBENANNT = u"umbenannt"
UEBERSPRUNGEN = u"übersprungen"
UNGUELTIG = u"ungültig"

MUSTER_UMBENENNEN = u"{name} ({n})"


def schluessel(name):
    return name.strip().casefold()


def bereinige(name):
    """Leerraum und umschließende Anführungszeichen (Excel) entfernen."""
    name = (name or u"").replace(u"﻿", u"").strip()
    if len(name) >= 2 and name[0] == name[-1] == u'"':
        name = name[1:-1].replace(u'""', u'"').strip()
    return name


def namen_aus_text(text):
    """Namen aus Zwischenablage oder Eingabe: eine Zeile oder eine
    Tabellenzelle (Tabulator, z.B. aus Excel) je Name. Leere Einträge und
    Doppelte fallen weg, die Reihenfolge bleibt."""
    namen = []
    for zeile in (text or u"").splitlines():
        for zelle in zeile.split(u"\t"):
            name = bereinige(zelle)
            if name:
                namen.append(name)
    return ohne_doppelte(namen)


def ohne_doppelte(namen):
    gesehen = set()
    ergebnis = []
    for name in namen:
        k = schluessel(name)
        if k not in gesehen:
            gesehen.add(k)
            ergebnis.append(name)
    return ergebnis


def zusammenfuehren(bisher, neu, anhaengen=True):
    """Neue Namen an die Liste anhängen (ohne Doppelte) oder sie ersetzen.
    Rückgabe: (Liste, Anzahl tatsächlich hinzugekommener Namen)."""
    basis = list(bisher) if anhaengen else []
    vorhanden = set(schluessel(n) for n in basis)
    hinzu = 0
    for name in ohne_doppelte(neu):
        if schluessel(name) not in vorhanden:
            vorhanden.add(schluessel(name))
            basis.append(name)
            hinzu += 1
    return basis, hinzu


def ist_gueltig(name):
    return bool(name.strip()) and not any(z in VERBOTENE_ZEICHEN
                                          for z in name)


def filterfunktion(muster, regex=False):
    """Prüffunktion name -> bool.

    Ohne Regex müssen alle Suchwörter im Namen vorkommen. Mit Regex wird
    re.search ohne Groß-/Kleinschreibung verwendet; ein ungültiger Ausdruck
    wirft re.error.
    """
    muster = muster or u""
    if regex:
        if not muster:
            return lambda name: True
        ausdruck = re.compile(muster, re.IGNORECASE)
        return lambda name: ausdruck.search(name) is not None
    woerter = muster.casefold().split()
    return lambda name: all(w in name.casefold() for w in woerter)


def eindeutiger_name(name, belegt, muster=MUSTER_UMBENENNEN):
    """Erster freier Name nach muster ({name}, {n} ab 2)."""
    n = 2
    while True:
        kandidat = muster.format(name=name, n=n)
        if schluessel(kandidat) not in belegt:
            return kandidat
        n += 1


def plane(namen, vorhandene, umbenennen, gueltig=ist_gueltig,
          muster=MUSTER_UMBENENNEN):
    """Was beim Erstellen mit jedem Namen passiert.

    namen       gewünschte Worksets in Listenreihenfolge
    vorhandene  Namen der Worksets im Projekt
    umbenennen  True: bei Namensgleichheit freien Namen suchen,
                False: überspringen
    Rückgabe:   [(Name, Zielname oder None, Status)]

    Bereits geplante Namen gelten für die folgenden als belegt, damit zwei
    Listeneinträge nie denselben Zielnamen bekommen.
    """
    belegt = set(schluessel(n) for n in vorhandene)
    plan = []
    for name in namen:
        name = bereinige(name)
        if not gueltig(name):
            plan.append((name, None, UNGUELTIG))
            continue
        if schluessel(name) not in belegt:
            ziel, status = name, NEU
        elif umbenennen:
            ziel, status = eindeutiger_name(name, belegt, muster), UMBENANNT
            if not gueltig(ziel):
                plan.append((name, None, UNGUELTIG))
                continue
        else:
            plan.append((name, None, UEBERSPRUNGEN))
            continue
        belegt.add(schluessel(ziel))
        plan.append((name, ziel, status))
    return plan
