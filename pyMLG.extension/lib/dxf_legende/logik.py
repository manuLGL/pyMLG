# -*- coding: utf-8 -*-
"""Revit-freie Umsetzung DXF -> Legendenplan (offline testbar).

Der Legendenplan beschreibt alles in Papier-Millimetern, Ursprung unten
links an der Legende. revit.py rechnet mit dem Ansichtsmassstab in
Modellkoordinaten um. Massstab: faktor = Papier-mm je Zeichnungseinheit,
vorgeschlagen über die häufigste Texthöhe.

Farben: Weiss (ACI 7) und fast Weiss werden schwarz - auf weissem Papier
wären sie unsichtbar. Optional werden helle Farben abgedunkelt.

Linienstärken: Revit kennt nur Stiftnummern 1..16. STIFTE_MM ist die
Stifttabelle der metrischen Revit-Vorlage (Modelllinien ~1:50) - die
tatsächliche Breite hängt von der Tabelle des Projekts ab.
"""

import math
import re

from dxf_legende import geometrie as geo

# Stiftnummer 1..16 -> Breite in mm (Näherung)
STIFTE_MM = (0.18, 0.25, 0.35, 0.5, 0.7, 1.0, 1.4, 2.0, 2.8, 4.0, 5.0,
             6.0, 7.0, 8.0, 9.0, 10.0)
# CAD-Vorgabe-Linienstärke
VORGABE_STAERKE = 0.25
# Kürzester Strich bzw. Lücke eines Revit-Musters
MIN_SEGMENT_MM = 0.25
# Kürzester Punkt (Revit-"Punkt" hat keine Länge)
VORGABE_SCHRIFT = u"Arial"
# Ab dieser Helligkeit (0..1) wird beim Abdunkeln nachgedunkelt
HELL_GRENZE = 0.6
# Breite eines Zeichens in Texthöhen (für Grenzen und Vorschau)
ZEICHENBREITE = 0.62

DURCHGEZOGEN = u"CONTINUOUS"

# Zeichen, die Revit in Namen von Linienstilen und Typen ablehnt
_VERBOTEN = re.compile(u"[\\\\:{}\\[\\]|;<>?`~]")


# ---------------------------------------------------------------------------
# Kleinteile
# ---------------------------------------------------------------------------

def revit_name(text, laenge=80):
    text = _VERBOTEN.sub(u"_", u"%s" % text).strip()
    return text[:laenge].strip() or u"_"


def zahl_text(wert, stellen=2):
    """2.5 -> '2.5', 3.0 -> '3', 0.125 -> '0.13'."""
    text = (u"%%.%df" % stellen) % wert
    if u"." in text:
        text = text.rstrip(u"0").rstrip(u".")
    return text


def hexfarbe(rgb):
    return u"%02X%02X%02X" % tuple(int(k) for k in rgb)


def helligkeit(rgb):
    r, g, b = [k / 255.0 for k in rgb]
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def papierfarbe(rgb, abdunkeln=False):
    """CAD-Farbe (dunkler Hintergrund) -> Farbe auf weissem Papier."""
    rgb = tuple(int(k) for k in rgb)
    if min(rgb) >= 235:
        return (0, 0, 0)
    if abdunkeln:
        hell = helligkeit(rgb)
        if hell > HELL_GRENZE:
            faktor = HELL_GRENZE / hell
            rgb = tuple(int(round(k * faktor)) for k in rgb)
    return rgb


def stift(mm):
    """Breite in mm -> nächste Revit-Stiftnummer 1..16."""
    beste, abstand = 1, None
    for nummer, breite in enumerate(STIFTE_MM, 1):
        d = abs(math.log(max(mm, 0.01) / breite))
        if abstand is None or d < abstand:
            beste, abstand = nummer, d
    return beste


def staerke_mm(kurve, faktor):
    """Linienstärke auf dem Papier: CAD-Stärke oder Polylinienbreite."""
    basis = kurve.staerke_mm if kurve.staerke_mm else VORGABE_STAERKE
    return max(basis, kurve.breite * faktor)


# ---------------------------------------------------------------------------
# Muster
# ---------------------------------------------------------------------------

def _normalisiere(striche_mm):
    """DXF-Striche (positiv Strich, negativ Lücke, 0 Punkt; schon in mm) ->
    ([("dash"|"dot"|"space", mm)], vorlauf).

    Revit verlangt: beginnt mit Strich/Punkt, Strich und Lücke wechseln
    sich ab, endet mit Lücke. vorlauf = Länge, die vorne abgeschnitten und
    hinten angehängt wurde (verschiebt den Musteranfang). Leere Liste =
    durchgezogen."""
    teile = []
    for s in striche_mm:
        art = u"space" if s < 0 else (u"dot" if s == 0 else u"dash")
        laenge = abs(s)
        if teile and (teile[-1][0] == u"space") == (art == u"space"):
            # Gleiche Sorte zusammenfassen (Punkt + Strich -> Strich)
            vorher = teile[-1]
            neu_art = u"space" if art == u"space" else (
                u"dot" if vorher[0] == art == u"dot" else u"dash")
            teile[-1] = (neu_art, vorher[1] + laenge)
        else:
            teile.append((art, laenge))
    if not any(a == u"space" for a, _l in teile) or \
            all(a == u"space" for a, _l in teile):
        return [], 0.0
    # Ringförmig: erstes und letztes Teil gleicher Sorte zusammenlegen
    if len(teile) > 1 and (teile[0][0] == u"space") == (teile[-1][0] == u"space"):
        letztes = teile.pop()
        art = teile[0][0] if teile[0][0] == letztes[0] else u"dash"
        teile[0] = (art, teile[0][1] + letztes[1])
        vorlauf = -letztes[1]
    else:
        vorlauf = 0.0
    if teile[0][0] == u"space":
        vorlauf += teile[0][1]
        teile = teile[1:] + teile[:1]
    ergebnis = []
    for art, laenge in teile:
        if art != u"dot":
            laenge = max(MIN_SEGMENT_MM, laenge)
        else:
            laenge = 0.0
        ergebnis.append((art, round(laenge, 3)))
    return ergebnis, vorlauf


def linienmuster_mm(striche, faktor):
    """Linientyp (DXF-Striche in Zeichnungseinheiten) -> Revit-Segmente in
    Papier-mm als Tupel, () = durchgezogen."""
    segmente, _vorlauf = _normalisiere([s * faktor for s in striche])
    return tuple(segmente)


def fuellgitter_mm(muster, faktor, ursprung):
    """Musterlinien einer Schraffur (Weltkoordinaten) -> Revit-Füllgitter
    in Papier-mm: [(winkel, ox, oy, abstand, versatz, (segmente…))].
    segmente: abwechselnd Strich, Lücke (leer = durchgezogen)."""
    gitter = []
    for winkel, (bx, by), (dx, dy), striche in muster:
        ux, uy = math.cos(winkel), math.sin(winkel)
        versatz = (dx * ux + dy * uy) * faktor
        abstand = (dy * ux - dx * uy) * faktor
        if abstand < 0:
            abstand, versatz = -abstand, -versatz
        if abstand < 0.05:
            continue
        segmente, vorlauf = _normalisiere([s * faktor for s in striche])
        laengen = tuple(max(MIN_SEGMENT_MM, l) for _a, l in segmente)
        ox = (bx - ursprung[0]) * faktor + ux * vorlauf
        oy = (by - ursprung[1]) * faktor + uy * vorlauf
        gitter.append((winkel, ox, oy, abstand, versatz, laengen))
    return gitter


# ---------------------------------------------------------------------------
# Massstab
# ---------------------------------------------------------------------------

def typische_texthoehe(texte):
    """Häufigste Texthöhe (gewichtet nach Zeichenzahl) oder None."""
    zaehler = {}
    for text in texte:
        if text.hoehe <= 0:
            continue
        schluessel = float(u"%.4g" % text.hoehe)
        zaehler[schluessel] = zaehler.get(schluessel, 0) + len(text.inhalt.strip())
    if not zaehler:
        return None
    return max(sorted(zaehler), key=lambda h: zaehler[h])


def grenzen(zeichnung):
    """(xmin, ymin, xmax, ymax) der Zeichnung in Zeichnungseinheiten."""
    listen = [geo.abtasten(k.punkte, k.geschlossen) for k in zeichnung.kurven]
    for flaeche in zeichnung.flaechen:
        listen.extend(geo.abtasten(s, True) for s in flaeche.schleifen)
    listen.extend(textrahmen(t_) for t_ in zeichnung.texte)
    return geo.grenzen(listen)


def textrahmen(text):
    """Geschätzte Ecken eines Textes (für Grenzen und Vorschau)."""
    zeilen = text.inhalt.split(u"\n")
    breite = max(len(z) for z in zeilen) * text.hoehe * ZEICHENBREITE \
        * (text.breitenfaktor or 1.0)
    zeilenhoehe = text.hoehe * 1.6
    hoehe = text.hoehe + zeilenhoehe * (len(zeilen) - 1)
    x0 = {u"links": 0.0, u"mitte": -breite / 2.0, u"rechts": -breite}[text.h_ausr]
    if text.v_ausr == u"mitte":
        y1 = hoehe / 2.0 if len(zeilen) > 1 else text.hoehe / 2.0
    elif text.v_ausr == u"oben":
        y1 = 0.0
    else:
        y1 = hoehe
    ecken = [(x0, y1), (x0 + breite, y1), (x0 + breite, y1 - hoehe),
             (x0, y1 - hoehe)]
    c, s = math.cos(text.drehung), math.sin(text.drehung)
    return [(text.x + x * c - y * s, text.y + x * s + y * c) for x, y in ecken]


def vorschlag(zeichnung, ziel_mm=2.5):
    """(faktor, texthöhe_mm) als Vorgabe für den Dialog.

    Ist die Zeichnung in mm und die Schrift schon papiergross (1,5-5 mm),
    bleibt sie 1:1. Ohne Texte: Legende wird 180 mm breit."""
    h = typische_texthoehe(zeichnung.texte)
    if h:
        if zeichnung.einheiten == 4 and 1.5 <= h <= 5.0:
            return 1.0, h
        return ziel_mm / h, ziel_mm
    rahmen = grenzen(zeichnung)
    breite = (rahmen[2] - rahmen[0]) if rahmen else 0.0
    return (180.0 / breite if breite > 0 else 1.0), None


def vorschlag_name(zeichnung, dateiname):
    """Grösster einzeiliger Text (Titel), sonst der Dateiname."""
    h = typische_texthoehe(zeichnung.texte) or 0.0
    kandidaten = [t_ for t_ in zeichnung.texte
                  if u"\n" not in t_.inhalt and t_.hoehe > h * 1.15]
    if kandidaten:
        titel = max(kandidaten, key=lambda t_: t_.hoehe)
        return titel.inhalt.strip()
    return dateiname


# ---------------------------------------------------------------------------
# Legendenplan
# ---------------------------------------------------------------------------

class LinienStil(object):

    def __init__(self, name, farbe, stift_nr, muster_name, muster):
        self.name = name
        self.farbe = farbe
        self.stift = stift_nr
        self.muster_name = muster_name      # None = durchgezogen
        self.muster = muster                # ((art, mm), …)


class TextTyp(object):

    def __init__(self, name, schrift, hoehe_mm, breitenfaktor, farbe):
        self.name = name
        self.schrift = schrift
        self.hoehe_mm = hoehe_mm
        self.breitenfaktor = breitenfaktor
        self.farbe = farbe


class FuellTyp(object):

    def __init__(self, name, farbe, muster_name, gitter):
        self.name = name
        self.farbe = farbe
        self.muster_name = muster_name      # None = Vollton
        self.gitter = gitter


class Plan(object):

    def __init__(self):
        self.linien = []        # (stil_schlüssel, punkte_mm, geschlossen)
        self.texte = []         # TextEintrag
        self.flaechen = []      # (typ_schlüssel, [schleife_mm])
        self.stile = {}
        self.texttypen = {}
        self.fuelltypen = {}
        self.breite_mm = 0.0
        self.hoehe_mm = 0.0


class TextEintrag(object):

    def __init__(self, typ, inhalt, x, y, drehung, h_ausr, v_ausr):
        self.typ = typ
        self.inhalt = inhalt
        self.x = x
        self.y = y
        self.drehung = drehung
        self.h_ausr = h_ausr
        self.v_ausr = v_ausr


class _Namen(object):
    """Vergibt eindeutige Namen: gleicher Name für anderen Inhalt -> _2."""

    def __init__(self):
        self.nach_inhalt = {}
        self.vergeben = set()

    def name(self, wunsch, inhalt):
        if inhalt in self.nach_inhalt:
            return self.nach_inhalt[inhalt]
        wunsch = revit_name(wunsch)
        name, nummer = wunsch, 2
        while name.lower() in self.vergeben:
            name = u"%s_%d" % (wunsch, nummer)
            nummer += 1
        self.vergeben.add(name.lower())
        self.nach_inhalt[inhalt] = name
        return name


def erstelle_plan(zeichnung, faktor, praefix=u"LEY", abdunkeln=False):
    """Zeichnung -> Plan in Papier-mm (Ursprung unten links)."""
    plan = Plan()
    rahmen = grenzen(zeichnung)
    if rahmen is None:
        return plan
    ox, oy = rahmen[0], rahmen[1]
    plan.breite_mm = (rahmen[2] - rahmen[0]) * faktor
    plan.hoehe_mm = (rahmen[3] - rahmen[1]) * faktor
    praefix = revit_name(praefix or u"LEY", 20)
    musternamen = _Namen()
    namen = _Namen()

    def mm(p):
        return ((p[0] - ox) * faktor, (p[1] - oy) * faktor,
                p[2] if len(p) > 2 else 0.0)

    for kurve in zeichnung.kurven:
        farbe = papierfarbe(kurve.farbe, abdunkeln)
        nr = stift(staerke_mm(kurve, faktor))
        lt = kurve.linientyp or DURCHGEZOGEN
        striche = zeichnung.linientypen.get(lt, [])
        muster = linienmuster_mm(
            striche, faktor * zeichnung.ltscale * (kurve.lt_faktor or 1.0))
        muster_name = None
        if muster:
            muster_name = musternamen.name(u"%s_%s" % (praefix, lt), (u"m",) + muster)
        name = namen.name(u"%s_%s_%s_%d" % (praefix, hexfarbe(farbe),
                                            lt if muster else u"SOLID", nr),
                          (u"l", farbe, nr, muster))
        if name not in plan.stile:
            plan.stile[name] = LinienStil(name, farbe, nr, muster_name, muster)
        plan.linien.append((name, [mm(p) for p in kurve.punkte], kurve.geschlossen))

    for text in zeichnung.texte:
        farbe = papierfarbe(text.farbe, abdunkeln)
        schrift = text.schrift or VORGABE_SCHRIFT
        hoehe = round(text.hoehe * faktor, 2)
        bf = round(text.breitenfaktor or 1.0, 2)
        wunsch = u"%s_%smm_%s_%s" % (praefix, zahl_text(hoehe), schrift, hexfarbe(farbe))
        if abs(bf - 1.0) > 0.005:
            wunsch += u"_x%s" % zahl_text(bf)
        name = namen.name(wunsch, (u"t", schrift, hoehe, bf, farbe))
        if name not in plan.texttypen:
            plan.texttypen[name] = TextTyp(name, schrift, hoehe, bf, farbe)
        x, y, _b = mm((text.x, text.y))
        plan.texte.append(TextEintrag(name, text.inhalt, x, y, text.drehung,
                                      text.h_ausr, text.v_ausr))

    for flaeche in zeichnung.flaechen:
        farbe = papierfarbe(flaeche.farbe, abdunkeln)
        if flaeche.muster:
            gitter = fuellgitter_mm(flaeche.muster, faktor, (ox, oy))
            # Ursprung spielt für die Gleichheit keine Rolle
            inhalt = (u"g",) + tuple((round(g[0], 4), round(g[3], 3), round(g[4], 3), g[5])
                                     for g in gitter)
            muster_name = musternamen.name(u"%s_%s" % (praefix, flaeche.muster_name),
                                           inhalt) if gitter else None
        else:
            gitter, muster_name = [], None
        wunsch = u"%s_%s_%s" % (praefix, flaeche.muster_name if muster_name else u"SOLID",
                                hexfarbe(farbe))
        name = namen.name(wunsch, (u"f", farbe, muster_name))
        if name not in plan.fuelltypen:
            plan.fuelltypen[name] = FuellTyp(name, farbe, muster_name, gitter)
        plan.flaechen.append((name, [[mm(p) for p in s] for s in flaeche.schleifen]))
    return plan
