# -*- coding: utf-8 -*-
"""Wohnfläche nach WoFlV - ohne Revit-API.

Anrechnung (Wohnflächenverordnung §4):
    * Räume mit lichter Höhe ab 2 m: voll
    * lichte Höhe 1 m bis unter 2 m (Dachschrägen): zur Hälfte
    * lichte Höhe unter 1 m: gar nicht
    * Balkone, Loggien, Dachgärten, Terrassen: in der Regel zu einem Viertel
      (BALKON_FAKTOR)
Nicht zur Wohnfläche gehören (§2 Abs. 3, §3 Abs. 3) Keller, Heizungs-/
Wasch-/Trockenräume und Treppen mit mehr als drei Steigungen - sie stehen
mit 0 m² in der Tabelle. Gemeinschaftsflächen ausserhalb der Wohnung
(Treppenhaus, Aufzug, Garage, Fahrrad-/Müllraum ...) erscheinen gar nicht. Die Raumart wird am Raumnamen erkannt (Deutsch, Englisch,
Spanisch).

Der Raumparameter "Faktor Anrechenbarkeit" (auch jeder andere Name mit
"anrechenbar", ersatzweise "WoFlV-Faktor")
bestimmt die Anrechnung der Innenräume, 0..1 oder in Prozent (z. B. 50) -
dann wird keine Dachschräge gemessen. Ausnahmen: Balkone/Terrassen zählen
immer mit BALKON_FAKTOR (im Modell steht dort teils 0,5), Gemeinschafts-
flächen nie.

Flächen sind hier in m².
"""

import re

from mlg_sprache import sprache, t

BALKON_FAKTOR = 0.25

INNEN = u"innen"
AUSSEN = u"aussen"
NICHT = u"nicht"
# Gemeinschaftsflächen ausserhalb der Wohnung - erscheinen gar nicht
ALLGEMEIN = u"allgemein"

# Untertypen der Aussenflächen - bestimmen die Beschriftung der Tabelle
BALKON = u"balkon"
LOGGIA = u"loggia"
TERRASSE = u"terrasse"

FAKTOR_PARAMETER = (u"Faktor Anrechenbarkeit", u"Anrechenbarkeitsfaktor",
                    u"WoFlV-Faktor",
                    u"WoFlV_Faktor", u"WoFlV Faktor")

_AUSSEN_WORTE = (
    (u"loggia", LOGGIA),
    (u"balkon", BALKON), (u"balcon", BALKON), (u"balcony", BALKON),
    (u"terrass", TERRASSE), (u"terraz", TERRASSE), (u"terrace", TERRASSE),
    (u"dachgarten", TERRASSE), (u"freisitz", TERRASSE),
)
_ALLGEMEIN_WORTE = (
    u"treppenhaus", u"aufzug", u"hausanschluss", u"fahrrad", u"muell",
    u"garage", u"carport", u"schacht", u"gemeinschaft",
    u"stairwell", u"stair core", u"elevator", u"lift", u"parking",
    u"bike", u"refuse", u"common",
    u"caja de escalera", u"nucleo", u"ascensor", u"garaje", u"basura",
    u"bicicleta", u"zona comun",
)
_NICHT_WORTE = (
    u"keller", u"heizraum", u"heizung", u"waschkueche", u"trockenraum",
    u"treppe",
    u"basement", u"cellar", u"stair", u"boiler",
    u"sotano", u"trastero", u"escalera",
)
# Kein Zimmer, auch wenn "zimmer" im Namen steckt (Badezimmer)
_KEIN_ZIMMER = (
    u"bad", u"wc", u"dusch", u"kueche", u"flur", u"diele", u"abstell",
    u"hwr", u"hauswirtschaft", u"ankleide", u"garderobe", u"speis",
    u"vorraum", u"eingang", u"gang",
    u"bath", u"toilet", u"kitchen", u"hall", u"corridor", u"storage",
    u"closet", u"utility", u"laundry", u"pantry", u"entrance",
    u"bano", u"aseo", u"cocina", u"pasillo", u"distribuidor", u"recibidor",
    u"vestidor", u"lavadero", u"despensa",
)
_ZIMMER = (
    u"zimmer", u"wohn", u"schlaf", u"kind", u"arbeit", u"buero", u"gast",
    u"studio", u"living", u"bedroom", u"bed room", u"study", u"office",
    u"dining room", u"salon", u"estar", u"dormitorio", u"habitacion",
    u"despacho",
)
# Wohnküche, Living/Kitchen: der Wohnteil macht es zum Zimmer
_WOHNTEIL = (u"wohn", u"living", u"salon", u"estar")

_ERSATZ = ((u"ä", u"ae"), (u"ö", u"oe"), (u"ü", u"ue"), (u"ß", u"ss"),
           (u"á", u"a"), (u"é", u"e"), (u"í", u"i"), (u"ó", u"o"),
           (u"ú", u"u"), (u"ñ", u"n"))


def normiere(name):
    text = (name or u"").lower()
    for alt, neu in _ERSATZ:
        text = text.replace(alt, neu)
    return text


def art(name):
    """(INNEN|AUSSEN|NICHT, Untertyp oder None) aus dem Raumnamen."""
    text = normiere(name)
    for wort, untertyp in _AUSSEN_WORTE:
        if wort in text:
            return AUSSEN, untertyp
    if any(wort in text for wort in _ALLGEMEIN_WORTE):
        return ALLGEMEIN, None
    if any(wort in text for wort in _NICHT_WORTE):
        return NICHT, None
    return INNEN, None


def ist_zimmer(name):
    text = normiere(name)
    if art(name)[0] != INNEN:
        return False
    wohnteil = any(w in text for w in _WOHNTEIL)
    if not wohnteil and any(w in text for w in _KEIN_ZIMMER):
        return False
    return wohnteil or any(w in text for w in _ZIMMER)


def _kompakt(name):
    text = normiere(name).strip()
    for zeichen in (u" ", u"-", u"_"):
        text = text.replace(zeichen, u"")
    return text


def faktor_rang(name):
    """Vorrang eines Parameternamens für die Anrechenbarkeit: 0 = "Faktor
    Anrechenbarkeit", dann die übrigen FAKTOR_PARAMETER, dann jeder Name mit
    "anrechenbar"; None = kein solcher Parameter. Gross-/Kleinschreibung,
    Leerzeichen, Binde- und Unterstriche egal."""
    text = _kompakt(name)
    namen = [_kompakt(n) for n in FAKTOR_PARAMETER]
    if text in namen:
        return namen.index(text)
    if u"anrechenbar" in text:
        return len(namen)
    return None


def ist_faktor_name(name):
    return faktor_rang(name) is not None


def faktor_aus_parameter(wert):
    """Wert des Parameters WoFlV-Faktor als 0..1, sonst None."""
    if wert is None:
        return None
    if not isinstance(wert, (int, float)):
        text = (u"%s" % wert).replace(u"%", u"").replace(u",", u".").strip()
        if not text:
            return None
        try:
            if u"/" in text:                    # Bruch wie "1/2"
                zaehler, nenner = text.split(u"/", 1)
                wert = float(zaehler) / float(nenner)
            else:
                wert = float(text)
        except (ValueError, ZeroDivisionError):
            return None
    if wert < 0:
        return None
    return wert / 100.0 if wert > 1.0 else float(wert)


class Raum(object):
    """Ein Revit-Raum, so weit ihn die Berechnung braucht.

    anteil_1m / anteil_2m: Anteil der Grundfläche mit lichter Höhe ab
    1 m bzw. 2 m (0..1), None = nicht gemessen (wie voll hoch).
    """

    def __init__(self, name, flaeche, abteilung=u"", nummer=u"",
                 faktor=None, anteil_1m=None, anteil_2m=None, ref=None,
                 belegung=u"", parameter_hinweis=None):
        self.name = name or u""
        self.nummer = nummer or u""
        self.flaeche = float(flaeche)
        self.abteilung = (abteilung or u"").strip()
        self.belegung = (belegung or u"").strip()
        # was im Parameter Anrechenbarkeitsfaktor steht (für den Bericht)
        self.parameter_hinweis = parameter_hinweis
        self.faktor_vorgabe = faktor
        self.anteil_1m = anteil_1m
        self.anteil_2m = anteil_2m
        self.ref = ref
        self.art, self.untertyp = art(name)

    def faktor(self):
        """(Faktor, Abzug für Dachschrägen in m²)."""
        if self.art == ALLGEMEIN:
            return 0.0, 0.0
        if self.art == AUSSEN:
            return BALKON_FAKTOR, 0.0
        if self.faktor_vorgabe is not None:
            return self.faktor_vorgabe, 0.0
        if self.art == NICHT:
            return 0.0, 0.0
        if self.anteil_2m is None or self.anteil_1m is None:
            return 1.0, 0.0
        a1 = min(max(self.anteil_1m, 0.0), 1.0)
        a2 = min(max(self.anteil_2m, 0.0), a1)
        faktor = a2 + 0.5 * (a1 - a2)
        return faktor, self.flaeche * (1.0 - faktor)

    def text(self):
        return (u"%s %s" % (self.nummer, self.name)).strip()


class Wohnung(object):
    """Räume einer Wohnung auf dem Plan.

    name        - Belegung (bzw. Abteilung) der Wohnung, "" wenn unbekannt
    raeume      - die gezählten Räume
    andere      - Räume anderer Wohnungen (andere Belegung)
    ohne        - Räume ohne Belegung (Treppenhaus, Nachbarflächen ...)
    nach_belegung - False: kein Raum hatte eine Belegung, ersatzweise nach
                  Abteilung gefiltert
    """

    def __init__(self, name, raeume, andere, ohne, nach_belegung):
        self.name = name
        self.raeume = raeume
        self.andere = andere
        self.ohne = ohne
        self.nach_belegung = nach_belegung


def _groesste(raeume, schluessel):
    summen = {}
    for raum in raeume:
        wert = schluessel(raum)
        if wert:
            summen[wert] = summen.get(wert, 0.0) + raum.flaeche
    if not summen:
        return u""
    return sorted(summen.items(), key=lambda p: (-p[1], p[0]))[0][0]


def eine_wohnung(raeume):
    """Die Wohnung des Plans: die Belegung mit der grössten Fläche, Räume
    ohne Belegung zählen nicht. Hat kein Raum eine Belegung: die Abteilung
    mit der grössten Fläche, Räume ohne Abteilung bleiben dabei."""
    haupt = _groesste(raeume, lambda r: r.belegung)
    if haupt:
        return Wohnung(haupt, [r for r in raeume if r.belegung == haupt],
                       [r for r in raeume if r.belegung not in (u"", haupt)],
                       [r for r in raeume if not r.belegung], True)
    haupt = _groesste(raeume, lambda r: r.abteilung)
    if not haupt:
        return Wohnung(u"", list(raeume), [], [], False)
    return Wohnung(haupt, [r for r in raeume if r.abteilung in (u"", haupt)],
                   [r for r in raeume if r.abteilung not in (u"", haupt)],
                   [], False)


class Ergebnis(object):
    """Summen einer Wohnung. Die Wohnfläche ist die Summe der einzeln auf
    0,01 m² gerundeten Raumwerte - so geht die Aufgliederung in der Tabelle
    genau auf."""

    def __init__(self, raeume):
        # Treppenhaus & Co. gehören nicht zur Wohnung
        self.allgemein = [r for r in raeume if r.art == ALLGEMEIN]
        self.raeume = [r for r in raeume if r not in self.allgemein]
        self.wohnflaeche = 0.0
        self.aussen = 0.0
        self.aussen_angerechnet = 0.0
        self.aussen_faktoren = set()
        self.schraegen_abzug = 0.0
        self.nicht_angerechnet = []
        self.gewerbe = False            # setzt berechne()
        # Zimmer nach Namen: ein unter der Dachschräge geteilter Raum (zwei
        # Räume "Wohnen") ist ein Zimmer - wie eine Zeile in der Tabelle
        zimmer_namen = set()
        for raum in self.raeume:
            faktor, abzug = raum.faktor()
            angerechnet = angerechnete_flaeche(raum)
            self.wohnflaeche += angerechnet
            self.schraegen_abzug += abzug
            if raum.art == AUSSEN:
                self.aussen += raum.flaeche
                self.aussen_angerechnet += angerechnet
                self.aussen_faktoren.add(round(faktor, 4))
            if faktor <= 0.0:
                self.nicht_angerechnet.append(raum)
            if ist_zimmer(raum.name):
                zimmer_namen.add(raum.name.strip())
        self.zimmer = len(zimmer_namen)


def angerechnete_flaeche(raum):
    """Angerechnete Fläche eines Raums, auf 0,01 m² gerundet."""
    faktor = raum.faktor()[0]
    return round(raum.flaeche * faktor + 1e-9, 2)


# Gewerbe statt Wohnung: Raumnamen bzw. Belegung/Abteilung mit diesen Wörtern
_GEWERBE_RAUM = (
    u"buero", u"archiv", u"edv", u"server", u"besprechung", u"konferenz",
    u"empfang", u"foyer", u"praxis", u"behandlung", u"laden", u"verkauf",
    u"werkstatt", u"office", u"meeting", u"reception", u"oficina",
    u"sala de reuniones", u"recepcion",
)
_GEWERBE_NAME = (u"gewerbe", u"buero", u"praxis", u"laden", u"office",
                 u"commercial", u"oficina", u"local")
# Räume, die es nur in Wohnungen gibt
_WOHN_RAUM = (u"schlaf", u"kinder", u"wohnen", u"wohnzimmer", u"bedroom",
              u"living", u"dormitorio", u"salon")


def ist_gewerbe(raeume, name=u""):
    """Gewerbeeinheit statt Wohnung? Belegung/Name wie "Büro"/"Gewerbe",
    oder Büro-/Praxisräume und kein Wohn-/Schlafraum."""
    if any(w in normiere(name) for w in _GEWERBE_NAME):
        return True
    namen = [normiere(r.name) for r in raeume]
    gewerbe = any(any(w in n for w in _GEWERBE_RAUM) for n in namen)
    wohnen = any(any(w in n for w in _WOHN_RAUM) for n in namen)
    return gewerbe and not wohnen


def berechne(raeume, name=u""):
    ergebnis = Ergebnis(raeume)
    ergebnis.gewerbe = ist_gewerbe(ergebnis.raeume, name)
    return ergebnis


# ---------------------------------------------------------------- Ausgabe

def zahl(wert, stellen=2):
    text = u"%.*f" % (stellen, wert)
    return text if sprache() == u"en" else text.replace(u".", u",")


def m2(wert):
    return zahl(wert) + u" m²"


def prozent(wert):
    if abs(wert - round(wert)) < 1e-6:
        return u"%d %%" % int(round(wert))
    return zahl(wert, 1) + u" %"


def titel(wohnung, zimmer):
    if zimmer <= 0:
        return wohnung
    art_text = t(u"%d-Zimmer-Wohnung", u"%d-room apartment",
                 u"vivienda de %d habitaciones") % zimmer
    if zimmer == 1:
        art_text = t(u"1-Zimmer-Wohnung", u"1-room apartment",
                     u"vivienda de 1 habitación")
    return u"%s · %s" % (wohnung, art_text) if wohnung else art_text


def _rubrik(raum):
    """Wie der Raum angerechnet wird - bestimmt die Klammer hinter dem Namen."""
    faktor, abzug = raum.faktor()
    if faktor <= 0.0:
        return (u"nicht",)
    if abzug > 0.0 and raum.faktor_vorgabe is None:
        return (u"schraege",)
    if abs(faktor - 1.0) > 1e-6:
        return (u"anteil", round(faktor, 4))
    return (u"voll",)


def _beschriftung(name, rubrik, flaeche):
    """Raumname, bei nicht voller Anrechnung mit der Rechnung dahinter."""
    if rubrik[0] == u"nicht":
        return t(u"%s (nicht angerechnet, %s)", u"%s (not credited, %s)",
                 u"%s (no computa, %s)") % (name, m2(flaeche))
    if rubrik[0] == u"schraege":
        return t(u"%s (Dachschräge, Grundfläche %s)",
                 u"%s (sloped ceiling, floor area %s)",
                 u"%s (techo inclinado, superficie %s)") % (name, m2(flaeche))
    if rubrik[0] == u"anteil":
        return t(u"%s (%s von %s)", u"%s (%s of %s)", u"%s (%s de %s)") % (
            name, prozent(rubrik[1] * 100.0), m2(flaeche))
    return name


def raum_beschriftung(raum):
    return _beschriftung(raum.name, _rubrik(raum), raum.flaeche)


def gruppen(raeume):
    """Räume gleichen Namens zusammengefasst, auch bei unterschiedlicher
    Anrechnung: [(beschriftung, angerechnete Fläche), ...] in der Reihenfolge
    des ersten Auftretens. Die Fläche ist die Summe der gerundeten Raumwerte -
    so geht die Wohnfläche weiter genau auf. Bei gemischter Anrechnung steht
    die gesamte Grundfläche dahinter, die Einzelwerte im Bericht."""
    reihe, je_name = [], {}
    for raum in raeume:
        name = raum.name.strip()
        if name not in je_name:
            je_name[name] = []
            reihe.append(name)
        je_name[name].append(raum)
    ergebnis = []
    for name in reihe:
        teile = je_name[name]
        flaeche = sum(r.flaeche for r in teile)
        rubriken = set(_rubrik(r) for r in teile)
        if len(rubriken) == 1:
            text = _beschriftung(name, rubriken.pop(), flaeche)
        else:
            text = t(u"%s (teils angerechnet, Grundfläche %s)",
                     u"%s (partly credited, floor area %s)",
                     u"%s (computa en parte, superficie %s)") % (name,
                                                                 m2(flaeche))
        ergebnis.append((text, sum(angerechnete_flaeche(r) for r in teile)))
    return ergebnis


_REIHENFOLGE = {INNEN: 0, AUSSEN: 1, NICHT: 2}


def reihenfolge(raeume):
    """Innenräume, dann Balkone/Terrassen, dann nicht angerechnete - sonst
    in der gegebenen Reihenfolge."""
    def schluessel(paar):
        index, raum = paar
        voll = 2 if raum.faktor()[0] <= 0.0 else _REIHENFOLGE.get(raum.art, 0)
        return (voll, index)
    return [r for _, r in sorted(enumerate(raeume), key=schluessel)]


# Zeilenarten der Tabelle
TITEL = u"titel"            # "Parameterwerte" mittig, Linie über ganze Breite
KOPF = u"kopf"
ZEILE = u"zeile"
TRENNER = u"trenner"        # Linie innerhalb der Ränder
FUSS = u"fuss"
VOLL = u"voll"              # Linie über die ganze Breite
MIETE = u"miete"            # Platz für Miete kalt / NK / Miete warm


def parameter_kopf():
    return t(u"Parameterwerte", u"Parameter values", u"Valores de parámetros")


def kopf_text(zimmer):
    if zimmer <= 0:
        return t(u"Wohnung", u"Apartment", u"Vivienda")
    if zimmer == 1:
        return t(u"1-Zimmer-Wohnung", u"1-room apartment",
                 u"Vivienda de 1 habitación")
    return t(u"%d-Zimmer-Wohnung", u"%d-room apartment",
             u"Vivienda de %d habitaciones") % zimmer


def euro_mit_leerzeichen(text):
    """"460€" -> "460 €" (auch "460  €" -> "460 €")."""
    return re.sub(u"(\\d)\\s*€", u"\\1 €", text or u"")


def miete_vorgabe():
    """Leerer Mietblock - wird in Revit von Hand ausgefüllt."""
    return u"\r".join([t(u"Miete kalt:", u"Rent excl. utilities:",
                         u"Alquiler sin gastos:"),
                       t(u"NK:", u"Utilities:", u"Gastos:"),
                       t(u"Miete warm:", u"Rent incl. utilities:",
                         u"Alquiler con gastos:")])


_MIETE_WORTE = (u"miete", u"nebenkosten", u"rent", u"utilities",
                u"alquiler", u"gastos")
# vom Werkzeug geschriebene Kopfzeilen (auch der älteren Fassung)
_KOPF_WORTE = (u"zimmer-wohnung", u"fakten:", u"room apartment",
               u"apartment facts", u"datos de la vivienda", u"vivienda de ",
               u"parameterwerte", u"parameter values",
               u"valores de parametros", u"gewerbeflaeche",
               u"commercial area", u"superficie comercial")


def ist_miete_text(text):
    """Gehört die Textnotiz zum Mietblock (Miete kalt, NK, Miete warm)?"""
    zeilen = [normiere(z).strip() for z in (text or u"").splitlines()]
    return any(any(w in z for w in _MIETE_WORTE) or z.startswith(u"nk")
               for z in zeilen if z)


def ist_werkzeug_kopf(text):
    """Stammt der Text vom Werkzeug (z. B. "3-Zimmer-Wohnung Fakten:")? Dann
    ist er keine von Hand eingetragene Überschrift."""
    text = normiere(text)
    return any(w in text for w in _KOPF_WORTE)


def tabelle(ergebnis, miete=None):
    """Zeilen der Flächentabelle: [(art, links, rechts, fett), ...].
    miete: Text des Mietblocks (None = leere Vorgabe). Der Titel der Wohnung
    steht nicht in der Tabelle, sondern gross links oben auf dem Plan."""
    e = ergebnis
    gewerbe = getattr(e, "gewerbe", False)
    flaeche_name = (t(u"Gewerbefläche", u"Commercial area",
                      u"Superficie comercial") if gewerbe else
                    t(u"Wohnfläche", u"Living area", u"Superficie útil"))
    zeilen = [(TITEL, parameter_kopf(), u"", True), (VOLL, u"", u"", False),
              (KOPF, flaeche_name if gewerbe else kopf_text(e.zimmer), u"",
               True),
              (TRENNER, u"", u"", False)]
    zeilen += [(ZEILE, text, m2(wert), False)
               for text, wert in gruppen(reihenfolge(e.raeume))]
    zeilen += [(TRENNER, u"", u"", False),
               (ZEILE, flaeche_name, m2(e.wohnflaeche), True)]
    if gewerbe:
        # Gewerbe: keine Zimmerzahl, keine WoFlV
        fuss = [t(u"Summe der angerechneten Raumflächen.",
                  u"Sum of the credited room areas.",
                  u"Suma de las superficies computadas.")]
    else:
        zeilen.append((ZEILE, t(u"Zimmer", u"Rooms", u"Habitaciones"),
                       u"%d" % e.zimmer, True))
        fuss = [t(u"Wohnfläche nach WoFlV, Summe der angerechneten "
                  u"Raumflächen.",
                  u"Living area according to WoFlV, sum of the credited room "
                  u"areas.",
                  u"Superficie útil según WoFlV, suma de las superficies "
                  u"computadas.")]
    if e.aussen > 0.0:
        if len(e.aussen_faktoren) == 1:
            fuss.append(t(u"Balkone/Terrassen zu %s angerechnet.",
                          u"Balconies/terraces credited at %s.",
                          u"Balcones/terrazas computados al %s.")
                        % prozent(list(e.aussen_faktoren)[0] * 100.0))
        else:
            fuss.append(t(u"Balkone/Terrassen anteilig angerechnet.",
                          u"Balconies/terraces credited in part.",
                          u"Balcones/terrazas computados en parte."))
    if e.schraegen_abzug >= 0.005:
        fuss.append(t(u"Dachschrägen: lichte Höhe 1 bis 2 m zur Hälfte, unter "
                      u"1 m nicht angerechnet.",
                      u"Sloped ceilings: clear height 1 to 2 m credited at "
                      u"half, under 1 m not at all.",
                      u"Techos inclinados: altura libre de 1 a 2 m al 50 %, "
                      u"menos de 1 m no computa."))
    zeilen += [(FUSS, u" ".join(fuss), u"", False), (VOLL, u"", u"", False),
               (MIETE, euro_mit_leerzeichen(miete) if miete
                else miete_vorgabe(), u"", False)]
    return zeilen


def innen_links(linien, x0, y0, y1, meter=1.0):
    """Linke Kante des inneren Plankopfrahmens.

    linien: [[(x, y), ...], ...] (Plankopflinien), meter: Länge von 1 m in
    der Einheit der Punkte. Jedes Teilstück zählt einzeln; senkrechte Stücke
    3 mm bis 6 cm rechts vom Aussenrand x0 werden je x aufsummiert, die am
    weitesten rechts liegende Kante mit zusammen mehr als 30 % der Höhe gilt.
    Ohne Treffer: x0 + 2 cm (üblicher Heftrand)."""
    hoehe = y1 - y0
    je_x = {}
    for linie in linien:
        for a, b in zip(linie, linie[1:]):
            if abs(a[0] - b[0]) > 0.0005 * meter:      # nicht senkrecht
                continue
            x = (a[0] + b[0]) / 2.0
            if x0 + 0.003 * meter < x < x0 + 0.06 * meter:
                schluessel = round(x / meter, 4)
                je_x[schluessel] = je_x.get(schluessel, 0.0) + abs(a[1] - b[1])
    kanten = [x * meter for x, laenge in je_x.items() if laenge > 0.3 * hoehe]
    return max(kanten) if kanten else x0 + 0.02 * meter


def wohnungs_name(wohnung, planname):
    """Überschrift: Belegung (bzw. Abteilung) der Räume, sonst Planname."""
    return wohnung or planname or u""
