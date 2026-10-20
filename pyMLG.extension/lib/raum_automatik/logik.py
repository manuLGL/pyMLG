# -*- coding: utf-8 -*-
"""Regeln der Raum-Automatik - ohne Revit-API.

Ein freier Bereich (geschlossen, noch ohne Raum) bekommt nur dann einen Raum,
wenn eine Regel greift - in dieser Reihenfolge:

    1. Küche und Sofa oder Bett   -> WOHNKÜCHE   1_W_WOHNEN
    2. Sanitärobjekt mit Dusche
       oder Wanne                 -> BAD         2_W_SANITÄR
       Sanitärobjekt ohne         -> WC          2_W_SANITÄR
    3. Bett                       -> ZIMMER n    1_W_WOHNEN
    4. Schrank, mehr als 2 Türen,
       weder Küche noch Sofa      -> VR          2_W_VERKEHRSFLÄCHE
    5. nur Schrank, bis 2 Türen   -> AR          2_W_NEBENRAUM

Fussboden: BAD und WC Fliesen, alle anderen Parkett.

Zimmernummern zählen je Wohnung: Zimmer mit Doppelbett zuerst (ZIMMER 1),
dann die übrigen nach Fläche. Eine Wohnung sind die über Türen verbundenen
Räume; eine Tür zwischen zwei Verkehrsflächen (Wohnungsgang - Treppenhaus)
trennt Wohnungen. Nummern vorhandener Räume "ZIMMER n" bleiben belegt.

Die Einrichtung wird am Familien- und Typnamen erkannt (Deutsch, Englisch,
Spanisch), Sanitärobjekte zusätzlich an der Kategorie.
"""

import re
import unicodedata

WOHNKUECHE = u"wohnkueche"
BAD = u"bad"
WC = u"wc"
ZIMMER = u"zimmer"
GANG = u"gang"
AR = u"ar"

# Raumname (ZIMMER mit Nummer) und Schlüssel der Raumkategorie je Art
NAMEN = {
    WOHNKUECHE: u"WOHNKÜCHE",
    BAD: u"BAD",
    WC: u"WC",
    ZIMMER: u"ZIMMER %d",
    GANG: u"VR",
    AR: u"AR",
}
KATEGORIEN = {
    WOHNKUECHE: u"1_W_WOHNEN",
    BAD: u"2_W_SANITÄR",
    WC: u"2_W_SANITÄR",
    ZIMMER: u"1_W_WOHNEN",
    GANG: u"2_W_VERKEHRSFLÄCHE",
    AR: u"2_W_NEBENRAUM",
}
FLIESEN = u"Fliesen"
PARKETT = u"Parkett"
BOEDEN = {BAD: FLIESEN, WC: FLIESEN}
REIHENFOLGE = (WOHNKUECHE, ZIMMER, BAD, WC, GANG, AR)

# Merkmale eines Bereichs
BETT = u"bett"
DOPPELBETT = u"doppelbett"
SANITAER = u"sanitaer"
DUSCHE = u"dusche"
KUECHE = u"kueche"
SOFA = u"sofa"
SCHRANK = u"schrank"

# ab dieser Breite (kürzere Seite, m) ist ein Bett ein Doppelbett
DOPPELBETT_BREITE = 1.35
# Flur: mehr als so viele Türen
MIN_TUEREN = 2

# Teile der normierten Namen (klein, ohne Umlaute/Akzente)
_BETT = (u"bett", u"cama")
_BETT_WORT = (u"bed", u"beds")
_DOPPEL = (u"doppel", u"double", u"queen", u"king", u"matrimonio")
_SANITAER = (u"wc", u"toilet", u"brause", u"klosett", u"dusch", u"shower", u"ducha",
             u"wanne", u"bathtub", u"banera", u"waschtisch", u"waschbecken",
             u"handwasch", u"lavabo", u"urinal", u"bidet", u"inodoro")
_DUSCHE = (u"dusch", u"brause", u"shower", u"ducha", u"wanne", u"bathtub",
           u"banera")
_DUSCHE_WORT = (u"tub",)
# zählen nicht als Sanitärobjekt, auch wenn sie in der Kategorie stehen
_KEIN_SANITAER = (u"spuel", u"kueche", u"kitchen", u"cocina", u"fregadero",
                  u"waschmasch", u"trockner", u"washer", u"dryer",
                  u"lavadora", u"boiler", u"speicher", u"therme", u"ablauf",
                  u"drain", u"gully", u"sumidero", u"anschluss")
_KUECHE = (u"kueche", u"kitchen", u"cocina", u"herd", u"kochfeld",
           u"spuele", u"kuehlschrank", u"backofen", u"dunstabzug",
           u"fregadero", u"nevera", u"vitro")
_SOFA = (u"sofa", u"couch", u"canape")
_SCHRANK = (u"schrank", u"garderobe", u"wardrobe", u"closet", u"cupboard",
            u"armario", u"ropero")

_WORT = re.compile(u"[a-z0-9]+")


def normiere(text):
    """Klein, Umlaute ausgeschrieben, ohne Akzente."""
    text = (text or u"").lower()
    for alt, neu in ((u"ä", u"ae"), (u"ö", u"oe"), (u"ü", u"ue"),
                     (u"ß", u"ss")):
        text = text.replace(alt, neu)
    text = unicodedata.normalize("NFKD", text)
    return u"".join(z for z in text if not unicodedata.combining(z))


def _hat(text, teile):
    return any(teil in text for teil in teile)


def _wort(text, woerter):
    return any(w in woerter for w in _WORT.findall(text))


def merkmale(name, sanitaer_kategorie=False, breite=None):
    """Merkmale eines Einrichtungsgegenstands.

    name: Familien- und Typname, breite: kürzere Seite im Grundriss in m.
    """
    text = normiere(name)
    ergebnis = set()
    if _hat(text, _BETT) or _wort(text, _BETT_WORT):
        ergebnis.add(BETT)
        if _hat(text, _DOPPEL) or (breite is not None
                                   and breite >= DOPPELBETT_BREITE):
            ergebnis.add(DOPPELBETT)
    # Dusche/Wanne zählt immer, auch "Dusche mit Ablauf"
    dusche = _hat(text, _DUSCHE) or _wort(text, _DUSCHE_WORT)
    if dusche or ((sanitaer_kategorie or _hat(text, _SANITAER))
                  and not _hat(text, _KEIN_SANITAER)):
        ergebnis.add(SANITAER)
        if dusche:
            ergebnis.add(DUSCHE)
    if _hat(text, _KUECHE):
        ergebnis.add(KUECHE)
    if _hat(text, _SOFA):
        ergebnis.add(SOFA)
    if _hat(text.replace(u"kuehlschrank", u""), _SCHRANK):
        ergebnis.add(SCHRANK)
    return ergebnis


def art(merkmal_menge, tueren):
    """Raumart nach den Regeln oder None (kein Raum)."""
    m = merkmal_menge
    # Wohnküche, auch Einraumwohnung mit Bett und Küche zusammen
    if KUECHE in m and (SOFA in m or BETT in m):
        return WOHNKUECHE
    if SANITAER in m:
        return BAD if DUSCHE in m else WC
    if BETT in m:
        return ZIMMER
    if (SCHRANK in m and KUECHE not in m and SOFA not in m
            and tueren > MIN_TUEREN):
        return GANG
    # Abstellraum: nur Schrank, höchstens 2 Türen
    if m == set([SCHRANK]):
        return AR
    return None


def ist_verkehr(name):
    """Name eines vorhandenen Raums - Verkehrsfläche?"""
    text = normiere(name).strip()
    return (_hat(text, (u"gang", u"flur", u"vorraum", u"diele", u"korridor",
                        u"treppe", u"stiege", u"verkehr", u"pasillo",
                        u"corridor", u"hall", u"stair"))
            or text in (u"vr", u"stgh", u"th"))


_ZIMMER_NR = re.compile(u"^\\s*zimmer\\s*(\\d+)\\s*$")


def zimmer_nummer(name):
    """Nummer aus "ZIMMER n" oder None."""
    treffer = _ZIMMER_NR.match(normiere(name))
    return int(treffer.group(1)) if treffer else None


class Bereich(object):
    """Ein Raum des Geschosses - neu (frei) oder vorhanden.

    kennung: beliebig, eindeutig. Bei neuen Bereichen füllt
    bewerte() art; bei vorhandenen zählen name (Verkehr, Zimmernummer).
    """

    def __init__(self, kennung, neu=True, flaeche=0.0, name=u""):
        self.kennung = kennung
        self.neu = neu
        self.flaeche = flaeche
        self.name = name
        self.merkmale = set()
        self.inhalt = []         # erkannte Gegenstände (Namen)
        self.tueren = 0
        self.art = None
        self.nummer = None

    def verkehr(self):
        """Trennt Wohnungen - auch ein Flur ohne Schrank (kein Raum,
        aber mehr als 2 Türen), sonst verbindet er alle Wohnungen."""
        if self.neu:
            return self.art == GANG or (self.art is None
                                        and self.tueren > MIN_TUEREN)
        return ist_verkehr(self.name)

    def raumname(self):
        if self.art == ZIMMER:
            return NAMEN[ZIMMER] % self.nummer
        return NAMEN.get(self.art)

    def kategorie(self):
        return KATEGORIEN.get(self.art)

    def boden(self):
        return BOEDEN.get(self.art, PARKETT)


def bewerte(bereiche):
    for b in bereiche:
        if b.neu:
            b.art = art(b.merkmale, b.tueren)


def wohnungen(bereiche, verbindungen):
    """Gruppen von Bereichen (Listen) - über Türen verbunden, aber nicht
    über eine Tür zwischen zwei Verkehrsflächen."""
    nach_kennung = dict((b.kennung, b) for b in bereiche)
    eltern = dict((k, k) for k in nach_kennung)

    def wurzel(k):
        while eltern[k] != k:
            eltern[k] = eltern[eltern[k]]
            k = eltern[k]
        return k

    for a, b in verbindungen:
        if a not in nach_kennung or b not in nach_kennung or a == b:
            continue
        if nach_kennung[a].verkehr() and nach_kennung[b].verkehr():
            continue
        eltern[wurzel(a)] = wurzel(b)
    gruppen = {}
    for b in bereiche:
        gruppen.setdefault(wurzel(b.kennung), []).append(b)
    return list(gruppen.values())


def nummeriere(bereiche, verbindungen):
    """Zimmernummern je Wohnung: Doppelbett zuerst, dann nach Fläche
    absteigend; belegte Nummern vorhandener Räume werden übersprungen."""
    for gruppe in wohnungen(bereiche, verbindungen):
        belegt = set(n for n in (zimmer_nummer(b.name) for b in gruppe
                                 if not b.neu) if n is not None)
        zimmer = [b for b in gruppe if b.neu and b.art == ZIMMER]
        zimmer.sort(key=lambda b: (DOPPELBETT not in b.merkmale,
                                   -b.flaeche))
        nummer = 1
        for b in zimmer:
            while nummer in belegt:
                nummer += 1
            b.nummer = nummer
            belegt.add(nummer)


def schluessel_name(text):
    """Vergleichsform eines Schlüsselnamens (1_W_Wohnen = 1_W_WOHNEN)."""
    return normiere(text).strip()
