# -*- coding: utf-8 -*-
"""Dateinamen für GroupExport und GroupToLink - ohne Revit."""

import re

ENDUNG = u".rvt"

# In Windows-Dateinamen verboten
_VERBOTEN = re.compile(u'[<>:"/\\\\|?*\x00-\x1f]')

# Geräte-Namen, die Windows als Dateinamen nicht zulässt
_RESERVIERT = set([u"CON", u"PRN", u"AUX", u"NUL"]
                  + [u"COM%d" % i for i in range(1, 10)]
                  + [u"LPT%d" % i for i in range(1, 10)])

ERSATZ = u"Gruppe"


def dateiname(gruppenname):
    """Gruppenname -> gültiger Dateiname ohne Endung."""
    name = _VERBOTEN.sub(u"_", gruppenname or u"").strip()
    # Windows schneidet Punkte und Leerzeichen am Ende ab
    name = name.rstrip(u". ")
    if not name:
        return ERSATZ
    if name.split(u".")[0].upper() in _RESERVIERT:
        name += u"_"
    return name


def dateinamen(gruppennamen):
    """Liste von Gruppennamen -> Dateinamen mit Endung, eindeutig ohne
    Rücksicht auf Gross-/Kleinschreibung ("A (2).rvt" beim zweiten)."""
    vergeben = set()
    ergebnis = []
    for gruppenname in gruppennamen:
        basis = dateiname(gruppenname)
        name = basis
        nummer = 1
        while name.lower() in vergeben:
            nummer += 1
            name = u"%s (%d)" % (basis, nummer)
        vergeben.add(name.lower())
        ergebnis.append(name + ENDUNG)
    return ergebnis


# Revit-Sicherungskopien: "Name.0001.rvt"
_SICHERUNG = re.compile(r"\.\d{4}$")


def zuordnung(gruppennamen, dateien):
    """Gruppenname -> .rvt-Datei gleichen Namens (ohne Rücksicht auf
    Gross-/Kleinschreibung; Gruppennamen mit verbotenen Zeichen passen auf
    den Dateinamen, den GroupExport daraus macht). Sicherungskopien zählen
    nicht. Rückgabe {Gruppenname: Datei}."""
    nach_stamm = {}
    for datei in dateien:
        if not datei.lower().endswith(ENDUNG):
            continue
        stamm = datei[:-len(ENDUNG)]
        if _SICHERUNG.search(stamm):
            continue
        nach_stamm.setdefault(stamm.lower(), datei)
    ergebnis = {}
    for name in gruppennamen:
        datei = (nach_stamm.get((name or u"").lower())
                 or nach_stamm.get(dateiname(name).lower()))
        if datei:
            ergebnis[name] = datei
    return ergebnis
