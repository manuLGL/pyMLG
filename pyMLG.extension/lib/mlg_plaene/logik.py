# -*- coding: utf-8 -*-
"""Revit-freie Logik der Plan-Werkzeuge: Platz auf dem Plan finden,
Plannummern bilden und umnummerieren.

Bewusst ohne Revit-API, damit sie ausserhalb von Revit prüfbar ist
(siehe tools/test_mlg_plaene.py).

Rechtecke sind Tupel (x0, y0, x1, y1) in Plankoordinaten (interne Fuss),
y zeigt nach oben - "oben links" heisst also grosses y, kleines x.
"""

import re

EPS = 1e-6


# ---------------------------------------------------------------- Rechtecke

def breite(r):
    return r[2] - r[0]


def hoehe(r):
    return r[3] - r[1]


def mitte(r):
    return ((r[0] + r[2]) / 2.0, (r[1] + r[3]) / 2.0)


def verschoben(r, dx, dy):
    return (r[0] + dx, r[1] + dy, r[2] + dx, r[3] + dy)


def zentriert(r, ziel):
    """r so verschoben, dass seine Mitte auf der von ziel liegt."""
    (mx, my), (zx, zy) = mitte(r), mitte(ziel)
    return verschoben(r, zx - mx, zy - my)


def ueberlappt(a, b, abstand=0.0):
    """True, wenn a und b näher als abstand beieinander liegen."""
    return (a[0] < b[2] + abstand - EPS and b[0] < a[2] + abstand - EPS
            and a[1] < b[3] + abstand - EPS and b[1] < a[3] + abstand - EPS)


def liegt_in(r, bereich):
    return (r[0] >= bereich[0] - EPS and r[1] >= bereich[1] - EPS
            and r[2] <= bereich[2] + EPS and r[3] <= bereich[3] + EPS)


def passt_in(b, h, bereich):
    return b <= breite(bereich) + EPS and h <= hoehe(bereich) + EPS


def zeichenflaeche(plankopf, rand=0.0, rechts_frei=0.0):
    """Nutzbare Fläche innerhalb des Plankopfs.

    rand wird rundum abgezogen, rechts_frei zusätzlich (z.B. für ein
    Schriftfeld am rechten Rand). Bleibt nichts übrig, gilt der ganze
    Plankopf.
    """
    x0, y0, x1, y1 = plankopf
    flaeche = (x0 + rand, y0 + rand,
               x1 - rand - rechts_frei, y1 - rand)
    if breite(flaeche) <= EPS or hoehe(flaeche) <= EPS:
        return tuple(plankopf)
    return flaeche


def freie_position(bereich, belegt, b, h, abstand=0.0):
    """Rechteck b x h an der ersten freien Stelle in Leserichtung.

    Kandidaten sind die linke obere Ecke des Bereichs und die Kanten der
    belegten Rechtecke (rechts daneben bzw. darunter). Gewählt wird der
    höchste, dann der am weitesten links liegende Platz, der im Bereich
    liegt und zu allen belegten Rechtecken den Abstand hält. Liefert das
    Rechteck oder None, wenn kein Platz ist.
    """
    xs = [bereich[0]] + [r[2] + abstand for r in belegt]
    obere = [bereich[3]] + [r[1] - abstand for r in belegt]
    kandidaten = sorted(set((round(o, 9), round(x, 9)) for o in obere for x in xs),
                        key=lambda k: (-k[0], k[1]))
    for oben, x in kandidaten:
        r = (x, oben - h, x + b, oben)
        if not liegt_in(r, bereich):
            continue
        if any(ueberlappt(r, anderes, abstand) for anderes in belegt):
            continue
        return r
    return None


def ueberlappung(a, b):
    """Fläche, die a und b gemeinsam haben."""
    dx = min(a[2], b[2]) - max(a[0], b[0])
    dy = min(a[3], b[3]) - max(a[1], b[1])
    return dx * dy if dx > 0 and dy > 0 else 0.0


def beste_position(bereich, belegt, b, h, abstand=0.0):
    """Wie freie_position(), aber immer mit Ergebnis: ist kein Platz frei,
    die Stelle im Bereich mit der kleinsten Überschneidung (bei Gleichstand
    in Leserichtung). Ist das Rechteck grösser als der Bereich, liegt es
    oben links bündig."""
    frei = freie_position(bereich, belegt, b, h, abstand)
    if frei is not None:
        return frei
    xs = [bereich[0], bereich[2] - b] + [r[2] + abstand for r in belegt]         + [r[0] - abstand - b for r in belegt]
    obere = [bereich[3], bereich[1] + h] + [r[1] - abstand for r in belegt]         + [r[3] + abstand + h for r in belegt]
    beste = None
    for oben in obere:
        for x in xs:
            # in den Bereich schieben, bei Übergrösse oben links bündig
            x = max(bereich[0], min(x, bereich[2] - b))
            oben = min(bereich[3], max(oben, bereich[1] + h))
            r = (x, oben - h, x + b, oben)
            wert = (round(sum(ueberlappung(r, anderes) for anderes in belegt), 9),
                    -round(oben, 9), round(x, 9))
            if beste is None or wert < beste[0]:
                beste = (wert, r)
    return beste[1]


# ---------------------------------------------------------------- Nummern

def norm(nummer):
    """Vergleichsform einer Plannummer (Revit unterscheidet keine
    Gross-/Kleinschreibung)."""
    return (nummer or u"").strip().lower()


def nummern_menge(nummern):
    return set(norm(n) for n in nummern)


def natuerlich(text):
    """Sortierschlüssel, bei dem A-2 vor A-10 kommt."""
    teile = re.split(r"(\d+)", (text or u"").lower())
    return [(0, int(teil), u"") if teil.isdigit() else (1, 0, teil)
            for teil in teile if teil]


_LETZTE_ZAHL = re.compile(r"^(.*?)(\d+)(\D*)$")


def nummer_plus(nummer, schritt=1):
    """Zählt die letzte Zahl in der Nummer weiter, führende Nullen bleiben.

    A-101 -> A-102, 01.009 -> 01.010, B -> B-1. Wirft ValueError, wenn die
    Zahl negativ würde.
    """
    treffer = _LETZTE_ZAHL.match(nummer)
    if not treffer:
        if schritt < 1:
            raise ValueError(nummer)
        return u"{}-{}".format(nummer, schritt)
    vorne, ziffern, hinten = treffer.groups()
    wert = int(ziffern) + schritt
    if wert < 0:
        raise ValueError(nummer)
    return u"{}{}{}".format(vorne, str(wert).zfill(len(ziffern)), hinten)


def naechste_freie_nummer(nummer, vergeben):
    """Nächste Nummer nach nummer, die nicht in vergeben (normiert) steht.
    Die gefundene Nummer wird in vergeben eingetragen."""
    kandidat = nummer_plus(nummer)
    while norm(kandidat) in vergeben:
        kandidat = nummer_plus(kandidat)
    vergeben.add(norm(kandidat))
    return kandidat


_MUSTER = re.compile(r"\{([^{}]+)\}|(#+)")


def muster_anwenden(muster, wert, zaehler=1):
    """Setzt Platzhalter in ein Muster ein.

    {Name}  -> wert(Name); None zählt als fehlend und wird leer eingesetzt
    ###     -> zaehler mit so vielen Stellen wie Rauten
    Liefert (Text, Liste fehlender Platzhalter).
    """
    fehlend = []

    def ersetzen(treffer):
        name, rauten = treffer.group(1), treffer.group(2)
        if rauten:
            return str(zaehler).zfill(len(rauten))
        text = wert(name.strip())
        if text is None:
            fehlend.append(name.strip())
            return u""
        return text

    return _MUSTER.sub(ersetzen, muster), fehlend


def hat_zaehler(muster):
    return any(m.group(2) for m in _MUSTER.finditer(muster))


def nummer_aus_muster(muster, wert, vergeben):
    """Erste freie Nummer nach dem Muster; wird in vergeben eingetragen.

    Mit ### wird hochgezählt (ab 1). Ohne Zähler wird bei Belegung -2, -3 ...
    angehängt. Liefert (Nummer, fehlende Platzhalter).
    """
    if hat_zaehler(muster):
        zaehler = 1
        while True:
            nummer, fehlend = muster_anwenden(muster, wert, zaehler)
            if norm(nummer) not in vergeben:
                break
            zaehler += 1
    else:
        basis, fehlend = muster_anwenden(muster, wert)
        nummer, zusatz = basis, 2
        while norm(nummer) in vergeben or not nummer.strip():
            nummer = u"{}-{}".format(basis, zusatz)
            zusatz += 1
    vergeben.add(norm(nummer))
    return nummer, fehlend


def umnummerierung(anzahl, start, schritt=1):
    """Neue Nummern für anzahl Pläne: start, start+schritt, ...

    Wirft ValueError bei Schritt 0 oder wenn eine Nummer negativ würde.
    """
    if schritt == 0:
        raise ValueError(u"schritt")
    neue, nummer = [], start
    for i in range(anzahl):
        if i:
            nummer = nummer_plus(nummer, schritt)
        neue.append(nummer)
    return neue


def konflikte(neue, fremde):
    """Neue Nummern, die schon ein nicht umnummerierter Plan trägt
    (fremde: normierte Menge)."""
    return [n for n in neue if norm(n) in fremde]


# ---------------------------------------------------------------- Ausrichten

class Fenster(object):
    """Ein Ansichtsfenster für die Zuordnung zwischen zwei Plänen.

    kennung: beliebig (in Revit die ElementId)
    art:     Ansichtsart (Grundriss, Schnitt ...) - nur gleiche passen
    massstab, ansicht: Maßstab und Kennung der Ansicht
    mitte:   (x, y) auf dem Plan
    """

    def __init__(self, kennung, art, massstab, ansicht, mitte):
        self.kennung = kennung
        self.art = art
        self.massstab = massstab
        self.ansicht = ansicht
        self.mitte = mitte


# Ein anderer Maßstab wiegt schwerer als jede Entfernung auf dem Plan
MASSSTAB_STRAFE = 1e6


def fenster_zuordnen(vorbilder, ziele):
    """Liefert [(Vorbild-Kennung, Ziel-Kennung)], jedes höchstens einmal.

    Nur gleiche Ansichtsarten passen. Vorrang hat dieselbe Ansicht (z.B.
    eine Legende auf beiden Plänen), dann gleicher Maßstab, dann die
    kürzeste Entfernung zwischen den Mitten.
    """
    paare = []
    for vi, v in enumerate(vorbilder):
        for zi, z in enumerate(ziele):
            if v.art != z.art:
                continue
            if v.ansicht == z.ansicht:
                kosten = -1.0
            else:
                kosten = ((v.mitte[0] - z.mitte[0]) ** 2
                          + (v.mitte[1] - z.mitte[1]) ** 2) ** 0.5
                if v.massstab != z.massstab:
                    kosten += MASSSTAB_STRAFE
            paare.append((kosten, vi, zi))
    paare.sort(key=lambda p: p[0])

    ergebnis, vergeben_v, vergeben_z = [], set(), set()
    for _, vi, zi in paare:
        if vi in vergeben_v or zi in vergeben_z:
            continue
        vergeben_v.add(vi)
        vergeben_z.add(zi)
        ergebnis.append((vorbilder[vi].kennung, ziele[zi].kennung))
    return ergebnis


# ---------------------------------------------------------------- Vorschau

def huelle(rechtecke):
    """Kleinstes Rechteck um alle Rechtecke."""
    return (min(r[0] for r in rechtecke), min(r[1] for r in rechtecke),
            max(r[2] for r in rechtecke), max(r[3] for r in rechtecke))


LINKS, RECHTS, OBEN, UNTEN = "links", "rechts", "oben", "unten"
MITTE_H, MITTE_V = "mitte_h", "mitte_v"


def ausrichten(rechtecke, art, referenz=None):
    """Richtet Rechtecke bündig aus (links, rechts, oben, unten, mitte_h =
    gleiche waagrechte Mitte, mitte_v = gleiche senkrechte Mitte).

    Bezug ist die Hülle der Rechtecke, bei einem einzelnen sinnvollerweise
    die Zeichenfläche (referenz). Liefert die Rechtecke in gleicher
    Reihenfolge.
    """
    ref = referenz if referenz is not None else huelle(rechtecke)
    ergebnis = []
    for r in rechtecke:
        dx = dy = 0.0
        if art == LINKS:
            dx = ref[0] - r[0]
        elif art == RECHTS:
            dx = ref[2] - r[2]
        elif art == MITTE_H:
            dx = mitte(ref)[0] - mitte(r)[0]
        elif art == OBEN:
            dy = ref[3] - r[3]
        elif art == UNTEN:
            dy = ref[1] - r[1]
        elif art == MITTE_V:
            dy = mitte(ref)[1] - mitte(r)[1]
        ergebnis.append(verschoben(r, dx, dy))
    return ergebnis


def verteilen(rechtecke, waagrecht=True):
    """Gleiche Abstände zwischen den Rechtecken; die äussersten bleiben.
    Ab drei Rechtecken, sonst unverändert."""
    if len(rechtecke) < 3:
        return list(rechtecke)
    ergebnis = list(rechtecke)
    if waagrecht:
        folge = sorted(range(len(rechtecke)), key=lambda i: rechtecke[i][0])
        erstes, letztes = rechtecke[folge[0]], rechtecke[folge[-1]]
        luecke = ((letztes[2] - erstes[0]) - sum(breite(r) for r in rechtecke)) \
            / (len(rechtecke) - 1)
        x = erstes[0]
        for i in folge:
            r = rechtecke[i]
            ergebnis[i] = verschoben(r, x - r[0], 0.0)
            x += breite(r) + luecke
    else:
        folge = sorted(range(len(rechtecke)), key=lambda i: -rechtecke[i][3])
        erstes, letztes = rechtecke[folge[0]], rechtecke[folge[-1]]
        luecke = ((erstes[3] - letztes[1]) - sum(hoehe(r) for r in rechtecke)) \
            / (len(rechtecke) - 1)
        oben = erstes[3]
        for i in folge:
            r = rechtecke[i]
            ergebnis[i] = verschoben(r, 0.0, oben - r[3])
            oben -= hoehe(r) + luecke
    return ergebnis


def einrasten(r, ziele, toleranz):
    """(dx, dy), um r an Kanten oder Mitten der Ziele einrasten zu lassen.

    Verglichen werden linke Kante, Mitte und rechte Kante (bzw. unten,
    Mitte, oben); es gilt die kleinste Abweichung innerhalb der Toleranz,
    sonst 0.
    """
    def linien_x(q):
        return (q[0], (q[0] + q[2]) / 2.0, q[2])

    def linien_y(q):
        return (q[1], (q[1] + q[3]) / 2.0, q[3])

    def bester(eigene, fremde):
        wahl = None
        for e in eigene:
            for f in fremde:
                d = f - e
                if abs(d) <= toleranz and (wahl is None or abs(d) < abs(wahl)):
                    wahl = d
        return wahl or 0.0

    fremde_x = [x for z in ziele for x in linien_x(z)]
    fremde_y = [y for z in ziele for y in linien_y(z)]
    return bester(linien_x(r), fremde_x), bester(linien_y(r), fremde_y)


def anordnen(bereich, fest, rechtecke, abstand=0.0):
    """Ordnet Rechtecke nacheinander in Leserichtung an, um die festen
    herum; ohne freien Platz an die Stelle mit der kleinsten Überschneidung.
    Liefert die neuen Rechtecke in gleicher Reihenfolge."""
    belegt = list(fest)
    ergebnis = []
    for r in rechtecke:
        neu = beste_position(bereich, belegt, breite(r), hoehe(r), abstand)
        belegt.append(neu)
        ergebnis.append(neu)
    return ergebnis


def ueberschneidungen(rechtecke):
    """Indizes der Rechtecke, die sich mit einem anderen überschneiden."""
    treffer = set()
    for i in range(len(rechtecke)):
        for j in range(i + 1, len(rechtecke)):
            if ueberlappung(rechtecke[i], rechtecke[j]) > EPS:
                treffer.add(i)
                treffer.add(j)
    return treffer
