# -*- coding: utf-8 -*-
"""Feste Raumnummern innerhalb einer Wohnung - ohne Revit-API.

Gilt für AutoRooms und ApartmentNumbering (Reiter ML2):

    V    Vorraum            001
    G    Gang               002   (001, wenn die Wohnung keinen Vorraum hat)
    B    Bad / Bad 1        003
    B    Bad 2              004
    WKÜ  Wohnküche          005
    WC   WC                 005
    ZI 1 Zimmer 1           006
    ZI 2 Zimmer 2           007
    ZI 3 Zimmer 3           008
    AR   Abstellraum        009
    SR   Schrankraum        010

Erkannt wird der Raumname, kurz oder lang, gross oder klein ("ZI 1",
"ZIMMER 1", "Zimmer1"; "VR" ist der Vorraum von AutoRooms). Bad und Zimmer
ohne Nummer ("BAD", "ZIMMER") bekommen die freien Nummern ihrer Reihe, die
grössten Räume zuerst. Räume ausserhalb der Tabelle (oder ZI 4, Bad 3 ...)
bekommen keine Nummer.
"""

import re
import unicodedata

VORRAUM = u"vorraum"
GANG = u"gang"
BAD = u"bad"
WOHNKUECHE = u"wohnkueche"
WC = u"wc"
ZIMMER = u"zimmer"
AR = u"ar"
SR = u"sr"

# Raumart -> Nummer
NUMMERN = {
    VORRAUM: u"001",
    GANG: u"002",
    WOHNKUECHE: u"005",
    WC: u"005",
    AR: u"009",
    SR: u"010",
}
# Gang in einer Wohnung ohne Vorraum
GANG_OHNE_VORRAUM = u"001"
# Raumarten mit eigener Zählung: n-ter Raum -> Nummer
REIHEN = {
    BAD: (u"003", u"004"),
    ZIMMER: (u"006", u"007", u"008"),
}

# Schreibweisen (normiert, ohne Leerzeichen) je Raumart
_NAMEN = (
    (VORRAUM, (u"v", u"vr", u"vorraum")),
    (GANG, (u"g", u"gang")),
    (BAD, (u"b", u"bad")),
    (WOHNKUECHE, (u"wkue", u"wku", u"wohnkueche")),
    (WC, (u"wc",)),
    (ZIMMER, (u"zi", u"zimmer")),
    (AR, (u"ar", u"abstellraum")),
    (SR, (u"sr", u"schrankraum")),
)
_FORM = re.compile(u"^([a-z]+)(\\d*)$")


def _normiere(text):
    text = (text or u"").lower()
    for alt, neu in ((u"ä", u"ae"), (u"ö", u"oe"), (u"ü", u"ue"),
                     (u"ß", u"ss")):
        text = text.replace(alt, neu)
    text = unicodedata.normalize("NFKD", text)
    text = u"".join(z for z in text if not unicodedata.combining(z))
    return re.sub(u"[\\s._-]+", u"", text)


def erkenne(name):
    """Raumname -> (Raumart, Index oder None) oder None. Einen Index haben
    nur Bad und Zimmer ("BAD 2", "ZI 1")."""
    treffer = _FORM.match(_normiere(name))
    if not treffer:
        return None
    wort, ziffern = treffer.groups()
    for art, schreibweisen in _NAMEN:
        if wort in schreibweisen:
            if ziffern and art not in REIHEN:
                return None
            return art, (int(ziffern) if ziffern else None)
    return None


def nummern(raeume):
    """Nummern der Räume einer Wohnung.

    raeume: [(Schlüssel, Raumname, Fläche)]. Rückgabe {Schlüssel: Nummer};
    Räume ohne Nummer nach der Tabelle fehlen darin.
    """
    erkannt = [(s, erkenne(n), f) for s, n, f in raeume]
    erkannt = [(s, e, f) for s, e, f in erkannt if e is not None]
    hat_vorraum = any(e[0] == VORRAUM for _, e, _ in erkannt)
    ergebnis = {}
    for art, reihe in REIHEN.items():
        belegt = set()
        ohne_index = []
        for s, (a, index), f in erkannt:
            if a != art:
                continue
            if index is None:
                ohne_index.append((s, f))
            elif 1 <= index <= len(reihe):
                ergebnis[s] = reihe[index - 1]
                belegt.add(index)
        frei = [i for i in range(1, len(reihe) + 1) if i not in belegt]
        ohne_index.sort(key=lambda e: -e[1])
        for (s, _), index in zip(ohne_index, frei):
            ergebnis[s] = reihe[index - 1]
    for s, (art, _), _ in erkannt:
        if art == GANG and not hat_vorraum:
            ergebnis[s] = GANG_OHNE_VORRAUM
        elif art in NUMMERN:
            ergebnis[s] = NUMMERN[art]
    return ergebnis
