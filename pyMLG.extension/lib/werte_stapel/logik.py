# -*- coding: utf-8 -*-
"""Revit-freie Logik von StackValues.

Bewusst ohne Revit-API, damit sie ausserhalb von Revit prüfbar ist
(siehe tools/test_werte_stapel.py).

Eine Lage ist der Grundriss eines Elements, Z spielt keine Rolle:

    (PUNKT, (x, y))                Stütze: Einfügepunkt
    (LINIE, a, b, m)               Wand/Unterzug: Anfang, Ende, Mitte der
                                   Achse - die Mitte unterscheidet Bögen

Zwei Linien liegen gleich, wenn Mitte und beide Enden innerhalb der
Toleranz liegen; die Zeichenrichtung ist egal.
"""

import math

PUNKT = "punkt"
LINIE = "linie"


class _Marke(object):
    def __init__(self, name):
        self.name = name

    def __repr__(self):
        return self.name


# Wert eines Zielparameters, der nicht geschrieben werden kann
FEHLT = _Marke("FEHLT")
SCHREIBGESCHUETZT = _Marke("SCHREIBGESCHUETZT")


# ---------------------------------------------------------------- Lage

def _abstand(a, b):
    return math.hypot(a[0] - b[0], a[1] - b[1])


def punkt_lage(p):
    return (PUNKT, (p[0], p[1]))


def linien_lage(a, b, m):
    return (LINIE, (a[0], a[1]), (b[0], b[1]), (m[0], m[1]))


def bezugspunkt(lage):
    """Punkt, nach dem einsortiert wird: Einfügepunkt bzw. Mitte."""
    return lage[1] if lage[0] == PUNKT else lage[3]


def gleiche_lage(a, b, toleranz):
    if a[0] != b[0]:
        return False
    if a[0] == PUNKT:
        return _abstand(a[1], b[1]) <= toleranz
    _, a0, a1, am = a
    _, b0, b1, bm = b
    if _abstand(am, bm) > toleranz:
        return False
    return ((_abstand(a0, b0) <= toleranz and _abstand(a1, b1) <= toleranz)
            or (_abstand(a0, b1) <= toleranz
                and _abstand(a1, b0) <= toleranz))


def _laenge(lage):
    return _abstand(lage[1], lage[2])


def _quer_und_laengs(p, a, b):
    """(Abstand von p zur Geraden durch a, b; Lage entlang a->b in Fuss)."""
    laenge = _abstand(a, b)
    ux, uy = (b[0] - a[0]) / laenge, (b[1] - a[1]) / laenge
    dx, dy = p[0] - a[0], p[1] - a[1]
    return abs(dx * uy - dy * ux), dx * ux + dy * uy


def _gerade(lage, toleranz):
    """Linie ohne Bogen: die Mitte liegt auf der Sehne."""
    return (_laenge(lage) > toleranz
            and _quer_und_laengs(lage[3], lage[1], lage[2])[0] <= toleranz)


def ueberlappung(a, b, toleranz):
    """Gemeinsame Länge zweier gerader Linien auf derselben Achse (Fuss),
    sonst None (nicht gerade, nicht auf einer Achse)."""
    if a[0] != LINIE or b[0] != LINIE:
        return None
    if not (_gerade(a, toleranz) and _gerade(b, toleranz)):
        return None
    q0, s0 = _quer_und_laengs(b[1], a[1], a[2])
    q1, s1 = _quer_und_laengs(b[2], a[1], a[2])
    if q0 > toleranz or q1 > toleranz:
        return None
    return min(_laenge(a), max(s0, s1)) - max(0.0, min(s0, s1))


# Mindestens dieser Anteil der kürzeren Linie muss gemeinsam sein
MIN_UEBERLAPPUNG = 0.5


def gleiche_achse(a, b, toleranz):
    """Gerade Linien auf derselben Achse, die sich mindestens zur Hälfte
    (der kürzeren) überlappen - die Länge darf abweichen. Sonst wie
    gleiche_lage (Punkte, Bögen)."""
    if gleiche_lage(a, b, toleranz):
        return True
    gemeinsam = ueberlappung(a, b, toleranz)
    if gemeinsam is None:
        return False
    return gemeinsam >= MIN_UEBERLAPPUNG * min(_laenge(a), _laenge(b))


def abweichung(a, b):
    """Grösster Abstand der einander entsprechenden Punkte (Fuss) - bei
    Linien in der besser passenden Richtung. None bei Punkt/Linie."""
    if a[0] != b[0]:
        return None
    if a[0] == PUNKT:
        return _abstand(a[1], b[1])
    mitte = _abstand(a[3], b[3])
    return max(mitte, min(max(_abstand(a[1], b[1]), _abstand(a[2], b[2])),
                          max(_abstand(a[1], b[2]), _abstand(a[2], b[1]))))


def _box(lage, rand=0.0):
    punkte = [lage[1]] if lage[0] == PUNKT else list(lage[1:])
    return (min(p[0] for p in punkte) - rand, min(p[1] for p in punkte) - rand,
            max(p[0] for p in punkte) + rand, max(p[1] for p in punkte) + rand)


class BoxRaster(object):
    """Lagen nach ihrem umschliessenden Rechteck einsortiert - für
    Vergleiche, bei denen die Mitten weit auseinander liegen dürfen."""

    def __init__(self, lagen, zelle=10.0):
        self.lagen = list(lagen)
        self.zelle = zelle
        self.zellen = {}
        for i, lage in enumerate(self.lagen):
            for schluessel in self._zellen(_box(lage)):
                self.zellen.setdefault(schluessel, []).append(i)

    def _zellen(self, box):
        x0, y0 = (int(math.floor(v / self.zelle)) for v in box[:2])
        x1, y1 = (int(math.floor(v / self.zelle)) for v in box[2:])
        for x in range(x0, x1 + 1):
            for y in range(y0, y1 + 1):
                yield (x, y)

    def kandidaten(self, lage, rand):
        """Indizes der Lagen, deren Rechteck das von lage (um rand
        vergrössert) berühren könnte."""
        ergebnis = set()
        for schluessel in self._zellen(_box(lage, rand)):
            ergebnis.update(self.zellen.get(schluessel, ()))
        return sorted(ergebnis)


class Raster(object):
    """Lagen nach ihrem Bezugspunkt in Zellen einsortiert - so wird nicht
    jedes Element mit jedem verglichen."""

    def __init__(self, lagen, toleranz):
        self.toleranz = toleranz
        self.zelle = max(toleranz, 1e-3)
        self.lagen = list(lagen)
        self.zellen = {}
        for i, lage in enumerate(self.lagen):
            self.zellen.setdefault(self._schluessel(lage), []).append(i)

    def _schluessel(self, lage):
        x, y = bezugspunkt(lage)
        return (int(math.floor(x / self.zelle)),
                int(math.floor(y / self.zelle)))

    def treffer(self, lage):
        """Indizes der Lagen, die mit lage übereinstimmen."""
        zx, zy = self._schluessel(lage)
        ergebnis = []
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for i in self.zellen.get((zx + dx, zy + dy), ()):
                    if gleiche_lage(lage, self.lagen[i], self.toleranz):
                        ergebnis.append(i)
        return sorted(ergebnis)


# ---------------------------------------------------------------- Zuordnung

class Element(object):
    """Quelle oder Ziel, wie die Logik sie sieht.

    werte: {Parameterschlüssel: Wert} - Wert ist None (leer), Text, Zahl
    oder FEHLT / SCHREIBGESCHUETZT.
    """

    def __init__(self, kennung, kategorie, lage, ebene=None, werte=None):
        self.kennung = kennung
        self.kategorie = kategorie
        self.lage = lage
        self.ebene = ebene
        self.werte = werte if werte is not None else {}


def zuordnen(quellen, ziele, toleranz, achse=False):
    """{Ziel-Index: [Quell-Indizes]} - nur gleiche Kategorie und Lage.
    Ziele ohne Quelle fehlen im Ergebnis.

    achse: Linien gelten auch als gleich, wenn sie auf derselben Achse
    liegen und sich ausreichend überlappen (gleiche_achse).
    """
    je_kategorie = {}
    for i, q in enumerate(quellen):
        je_kategorie.setdefault(q.kategorie, []).append(i)
    raster = dict((k, Raster([quellen[i].lage for i in idx], toleranz))
                  for k, idx in je_kategorie.items())
    boxen = {}
    ergebnis = {}
    for zi, ziel in enumerate(ziele):
        r = raster.get(ziel.kategorie)
        if r is None:
            continue
        idx = je_kategorie[ziel.kategorie]
        if achse and ziel.lage[0] == LINIE:
            b = boxen.get(ziel.kategorie)
            if b is None:
                b = boxen[ziel.kategorie] = BoxRaster(r.lagen)
            treffer = [idx[i] for i in b.kandidaten(ziel.lage, toleranz)
                       if gleiche_achse(ziel.lage, r.lagen[i], toleranz)]
        else:
            treffer = [idx[i] for i in r.treffer(ziel.lage)]
        if treffer:
            ergebnis[zi] = treffer
    return ergebnis


# ---------------------------------------------------------------- Diagnose

# Gründe, warum eine Quelle kein Gegenstück hat
NICHT_GEWAEHLT = "nicht_gewaehlt"   # Gegenstück nur auf abgewählter Ebene
ANDERE_LAENGE = "andere_laenge"     # gleiche Achse, Enden weichen ab
DANEBEN = "daneben"                 # knapp neben der Toleranz
NICHTS = "nichts"                   # nichts in der Nähe

# So weit (Fuss, ca. 1 m) wird nach knappen Fehlgriffen gesucht
SUCHRADIUS = 1.0 / 0.3048


def warum(quellen, ziele, zuordnung, ebenen, toleranz, quell_indizes):
    """{Quell-Index: (Grund, Ziel-Index oder None, Abweichung in Fuss)}
    für Quellen ohne Gegenstück auf den gewählten Ebenen."""
    zu_quelle = {}
    for zi, qis in zuordnung.items():
        for qi in qis:
            zu_quelle.setdefault(qi, []).append(zi)
    je_kategorie = {}
    for zi, ziel in enumerate(ziele):
        je_kategorie.setdefault(ziel.kategorie, []).append(zi)
    boxen = {}

    ergebnis = {}
    for qi in quell_indizes:
        quelle = quellen[qi]
        treffer = zu_quelle.get(qi)
        if treffer:
            ergebnis[qi] = (NICHT_GEWAEHLT, treffer[0], 0.0)
            continue
        idx = je_kategorie.get(quelle.kategorie, [])
        b = boxen.get(quelle.kategorie)
        if b is None:
            b = boxen[quelle.kategorie] = BoxRaster(
                [ziele[zi].lage for zi in idx])
        bester = None
        for i in b.kandidaten(quelle.lage, SUCHRADIUS):
            zi = idx[i]
            lage = ziele[zi].lage
            if ebenen is not None and ziele[zi].ebene not in ebenen:
                continue
            gemeinsam = ueberlappung(quelle.lage, lage, toleranz)
            weg = abweichung(quelle.lage, lage)
            if gemeinsam is not None and gemeinsam > toleranz:
                kandidat = (0, weg, ANDERE_LAENGE, zi)
            elif weg is not None and weg <= SUCHRADIUS:
                kandidat = (1, weg, DANEBEN, zi)
            else:
                continue
            if bester is None or kandidat[:2] < bester[:2]:
                bester = kandidat
        if bester is None:
            ergebnis[qi] = (NICHTS, None, None)
        else:
            ergebnis[qi] = (bester[2], bester[3], bester[1])
    return ergebnis


def gegenstuecke_je_ebene(zuordnung, ziele):
    """{Ebene: Anzahl Ziele mit Quelle}."""
    ergebnis = {}
    for zi in zuordnung:
        ebene = ziele[zi].ebene
        ergebnis[ebene] = ergebnis.get(ebene, 0) + 1
    return ergebnis


# ---------------------------------------------------------------- Plan

def ist_leer(wert):
    if wert is None:
        return True
    try:
        return not wert.strip()
    except AttributeError:
        return False


def gleich(a, b):
    if isinstance(a, float) or isinstance(b, float):
        try:
            return abs(a - b) <= 1e-9 * max(1.0, abs(a), abs(b))
        except TypeError:
            return False
    return a == b


def _eindeutig(werte):
    ergebnis = []
    for w in werte:
        if not any(gleich(w, e) for e in ergebnis):
            ergebnis.append(w)
    return ergebnis


class Plan(object):
    def __init__(self):
        self.auftraege = []         # [(Ziel-Index, Schlüssel, neuer Wert)]
        self.gleich = 0             # Ziel hatte den Wert schon
        self.quelle_leer = 0        # Quelle ohne Wert - nichts übertragen
        self.behalten = []          # [(Ziel-Index, Schlüssel, alt, neu)]
        self.mehrdeutig = []        # [(Ziel-Index, Schlüssel, [Werte])]
        self.fehlt = []             # [(Ziel-Index, Schlüssel, FEHLT/...)]
        self.ohne_gegenstueck = []  # [Quell-Index]

    def ziele(self):
        """Ziel-Indizes, die etwas geschrieben bekommen."""
        return sorted(set(zi for zi, _, _ in self.auftraege))


def plane(quellen, ziele, zuordnung, schluessel, ebenen=None,
          ueberschreiben=False):
    """Was in welches Ziel geschrieben wird.

    ebenen:          erlaubte Ebenen der Ziele (None = alle)
    ueberschreiben:  False = nur leere Zielwerte füllen
    Leere Quellwerte werden nie übertragen (sie würden Ziele leeren).
    """
    plan = Plan()
    erreicht = set()
    for zi in sorted(zuordnung):
        ziel = ziele[zi]
        if ebenen is not None and ziel.ebene not in ebenen:
            continue
        erreicht.update(zuordnung[zi])
        for s in schluessel:
            werte = [quellen[q].werte.get(s) for q in zuordnung[zi]]
            werte = _eindeutig([w for w in werte
                                if w is not FEHLT
                                and w is not SCHREIBGESCHUETZT
                                and not ist_leer(w)])
            if not werte:
                plan.quelle_leer += 1
                continue
            if len(werte) > 1:
                plan.mehrdeutig.append((zi, s, werte))
                continue
            neu = werte[0]
            alt = ziel.werte.get(s, FEHLT)
            if alt is FEHLT or alt is SCHREIBGESCHUETZT:
                plan.fehlt.append((zi, s, alt))
            elif ist_leer(alt) or (ueberschreiben and not gleich(alt, neu)):
                plan.auftraege.append((zi, s, neu))
            elif gleich(alt, neu):
                plan.gleich += 1
            else:
                plan.behalten.append((zi, s, alt, neu))
    plan.ohne_gegenstueck = [i for i in range(len(quellen))
                             if i not in erreicht]
    return plan


def zahl(text):
    """Dezimalzahl mit Komma oder Punkt -> float (ValueError sonst)."""
    return float((text or u"").strip().replace(u",", u"."))
