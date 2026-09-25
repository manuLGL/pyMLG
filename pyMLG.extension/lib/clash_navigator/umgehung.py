# -*- coding: utf-8 -*-
"""Umgehungen für Rohre und Luftkanäle berechnen - ohne Revit.

Eine gerade Leitung S -> E kreuzt ein Hindernis (Hüllkasten). Die Leitung
wird vor dem Hindernis aufgetrennt, weicht in eine Richtung aus und kehrt
dahinter zurück:

              P1o ________________ P2o          Seitenansicht, Richtung "oben"
                 /                \\            Winkel 45°: schräge Stücke
    S ________ P1   [ Hindernis ]  P2 ________ E
                 |<- t0      t1 ->|

Alle Längen in Metern, Punkte als (x, y, z) im Wirtssystem. Der Versatz
ergibt sich aus der Oberkante des Hindernisses (in Ausweichrichtung), dem
Sicherheitsabstand und der halben Leitung inkl. Dämmung.

Die Kollisionsprüfung der Varianten gegen die Umgebung macht
umgehung_revit.py - hier entstehen nur Geometrie und Bewertung.
"""

import math

from mlg_sprache import t, tt

# Grösster sinnvoller Versatz - darüber wird die Variante verworfen
MAX_VERSATZ = 1.5

# Aufschlag in der Bewertung für 90°-Bögen (Druckverlust, Platz) in Metern
AUFSCHLAG_90 = 0.25

# Aufschlag für einen Hochpunkt in einer Gefälleleitung (Luftsack) - solche
# Varianten rutschen nach hinten, bleiben aber sichtbar
AUFSCHLAG_HOCHPUNKT = 1.0

# Aufschlag, wenn der Sicherheitsabstand unterschritten werden musste
AUFSCHLAG_ABSTAND = 0.5

# Aufschlag, wenn beide Leitungen geändert werden (doppelte Arbeit)
AUFSCHLAG_BEIDE = 0.3

# Rückfall-Abstände, wenn mit der Vorgabe nichts kollisionsfrei passt
RUECKFALL_ABSTAENDE = (0.05, 0.02)

RICHTUNGEN = {
    u"oben": (u"Oben", u"Above", u"Arriba"),
    u"unten": (u"Unten", u"Below", u"Abajo"),
    u"links": (u"Links", u"Left", u"Izquierda"),
    u"rechts": (u"Rechts", u"Right", u"Derecha"),
    u"seite1": (u"Seite 1", u"Side 1", u"Lado 1"),
    u"seite2": (u"Seite 2", u"Side 2", u"Lado 2"),
    u"seite3": (u"Seite 3", u"Side 3", u"Lado 3"),
    u"seite4": (u"Seite 4", u"Side 4", u"Lado 4"),
}


class Fehler(Exception):
    """Die Leitung lässt sich nicht automatisch umgehen."""


# --- Vektoren ----------------------------------------------------------------

def plus(a, b):
    return (a[0] + b[0], a[1] + b[1], a[2] + b[2])


def minus(a, b):
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def mal(a, faktor):
    return (a[0] * faktor, a[1] * faktor, a[2] * faktor)


def skalar(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def kreuz(a, b):
    return (a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2],
            a[0] * b[1] - a[1] * b[0])


def laenge(a):
    return math.sqrt(skalar(a, a))


def einheit(a):
    betrag = laenge(a)
    if betrag < 1e-12:
        raise Fehler(t(u"Die Leitung hat keine Länge.",
                       u"The segment has no length.",
                       u"El tramo no tiene longitud."))
    return mal(a, 1.0 / betrag)


# --- Querschnitt -------------------------------------------------------------

class Querschnitt(object):
    """Aussenmass der Leitung inkl. Dämmung.

    rund:  durchmesser
    eckig: breite entlang achse_b, hoehe entlang achse_h (Einheitsvektoren)
    """

    def __init__(self, durchmesser=None, breite=None, hoehe=None,
                 achse_b=None, achse_h=None, daemmung=0.0):
        self.durchmesser = durchmesser
        self.breite = breite
        self.hoehe = hoehe
        self.achse_b = achse_b
        self.achse_h = achse_h
        self.daemmung = daemmung or 0.0

    @property
    def rund(self):
        return self.durchmesser is not None

    def halb(self, richtung):
        """Halbe Ausdehnung quer zur Leitung in dieser Richtung."""
        if self.rund:
            return self.durchmesser / 2.0 + self.daemmung
        return (abs(skalar(richtung, self.achse_b)) * self.breite / 2.0
                + abs(skalar(richtung, self.achse_h)) * self.hoehe / 2.0
                + self.daemmung)

    @property
    def groesste(self):
        """Grösstes Aussenmass (Durchmesser bzw. längere Seite) inkl.
        Dämmung - Grundlage für Platzbedarf der Bögen."""
        if self.rund:
            return self.durchmesser + 2.0 * self.daemmung
        return max(self.breite, self.hoehe) + 2.0 * self.daemmung

    def text(self):
        if self.rund:
            return u"ø %d mm" % round(self.durchmesser * 1000.0)
        return u"%d × %d mm" % (round(self.breite * 1000.0),
                                     round(self.hoehe * 1000.0))


def hat_gefaelle(start, ende):
    """Geneigt, aber nicht senkrecht - etwa eine Abwasserleitung."""
    richtung = einheit(minus(ende, start))
    return 1e-3 < abs(richtung[2]) < 0.99


def richtungen(achse):
    """Ausweichrichtungen quer zur Leitung: {code: Einheitsvektor}."""
    if abs(achse[2]) < 0.9:
        oben = einheit(minus((0.0, 0.0, 1.0),
                             mal(achse, skalar(achse, (0.0, 0.0, 1.0)))))
        links = kreuz(oben, achse)
        return {u"oben": oben, u"unten": mal(oben, -1.0),
                u"links": links, u"rechts": mal(links, -1.0)}
    # Steigleitung: zwei waagerechte Achsen
    erste = einheit(minus((1.0, 0.0, 0.0),
                          mal(achse, skalar(achse, (1.0, 0.0, 0.0)))))
    zweite = kreuz(achse, erste)
    return {u"seite1": erste, u"seite2": mal(erste, -1.0),
            u"seite3": zweite, u"seite4": mal(zweite, -1.0)}


# --- Varianten ---------------------------------------------------------------

UMGEHUNG = u"umgehung"
VERSCHIEBEN = u"verschieben"

# Nachbar an einem Ende, der sich nicht mitziehen lässt (Abzweig, Gerät ...)
FEST = u"fest"


class Variante(object):
    """Eine Lösung.

    art UMGEHUNG     Leitung auftrennen, Mittelstück versetzen, 4 Bögen
    art VERSCHIEBEN  ganzes Stück parallel versetzen; die Bögen an den Enden
                     gehen mit, die anschliessenden Leitungen werden länger
                     oder kürzer - keine neuen Bögen
    """

    def __init__(self, richtung, winkel, art=UMGEHUNG):
        self.richtung = richtung
        self.winkel = winkel
        self.art = art
        self.versatz = 0.0
        self.punkte = []            # P1, P1o, P2o, P2
        self.zusatzlaenge = 0.0
        self.gueltig = True
        self.grund = u""
        self.kollisionen = []       # Beschreibungen, gesetzt von der Prüfung
        self.warnungen = []         # (Text, kritisch) - Hinweise zur Variante
        self.hochpunkt = False      # Gefälleleitung weicht nach oben aus
        self.skizze = {}
        self.pruefstuecke = None    # Stücke für die Prüfung, falls abweichend
        self.abstand = None         # Sicherheitsabstand dieser Variante
        self.aufschlag = 0.0        # zusätzlicher Aufschlag in der Bewertung

    @property
    def segmente(self):
        """Die Stücke, die neu entstehen bzw. neu belegt werden - für
        Vorschau und Kollisionsprüfung."""
        if self.pruefstuecke is not None:
            return list(self.pruefstuecke)
        return list(zip(self.punkte[:-1], self.punkte[1:]))

    @property
    def name(self):
        if self.art == VERSCHIEBEN:
            return t(u"Verschieben: %s", u"Shift: %s", u"Desplazar: %s") % (
                tt(RICHTUNGEN[self.richtung]))
        return u"%s %d°" % (tt(RICHTUNGEN[self.richtung]), self.winkel)

    @property
    def kosten(self):
        return (self.zusatzlaenge + self.aufschlag
                + (AUFSCHLAG_90 if self.winkel == 90 else 0.0)
                + (AUFSCHLAG_HOCHPUNKT if self.hochpunkt else 0.0))

    def sortierschluessel(self):
        return (not self.gueltig, len(self.kollisionen) > 0,
                len(self.kollisionen), self.kosten)

    def zusammenfassung(self):
        if not self.gueltig:
            return self.grund
        if self.art == VERSCHIEBEN:
            return t(u"Versatz %d cm · ganzes Stück · keine neuen Bögen",
                     u"Offset %d cm · whole segment · no new elbows",
                     u"Desvío %d cm · tramo entero · sin codos nuevos") % (
                round(self.versatz * 100.0))
        return t(u"Versatz %d cm · +%.2f m · 4 Bögen",
                 u"Offset %d cm · +%.2f m · 4 elbows",
                 u"Desvío %d cm · +%.2f m · 4 codos") % (
            round(self.versatz * 100.0), self.zusatzlaenge)


def _ecken(kasten):
    klein, gross = kasten
    return [(x, y, z) for x in (klein[0], gross[0])
            for y in (klein[1], gross[1]) for z in (klein[2], gross[2])]


def _gefaelle_hinweise(variante, richtung):
    steigt = skalar(richtung, (0.0, 0.0, 1.0)) > 0.1
    variante.hochpunkt = steigt
    if steigt:
        variante.warnungen.append((t(
            u"Hochpunkt in Gefälleleitung (Luftsack) - nur für "
            u"Druckleitungen",
            u"High point in a sloped line (air pocket) - pressure lines only",
            u"Punto alto en tubería con pendiente (bolsa de aire): solo "
            u"para tuberías a presión"), True))
    variante.warnungen.append((t(
        u"Gefälle bleibt im Mittelstück erhalten",
        u"Slope is kept in the middle section",
        u"La pendiente se mantiene en el tramo central"), False))


def _verschieben(start, ende, code, richtung, versatz, nachbarn, mindest,
                 skizze):
    """Das ganze Stück um 'versatz' in 'richtung' versetzen.

    nachbarn: für Anfang und Ende je None (freies Ende), FEST oder
              (Richtung vom Stück weg, Länge) der Leitung hinter dem Bogen.
    Geht nur, wenn jede Nachbarleitung parallel zur Verschiebung liegt - dann
    wird sie einfach länger oder kürzer.
    """
    variante = Variante(code, 0, VERSCHIEBEN)
    variante.versatz = versatz
    versetzt = mal(richtung, versatz)
    neu_a, neu_e = plus(start, versetzt), plus(ende, versetzt)
    variante.punkte = [neu_a, neu_e]
    stuecke = [(neu_a, neu_e)]
    zusatz = 0.0
    for ende_nr, (punkt, neu, nachbar) in enumerate(
            ((start, neu_a, nachbarn[0]), (ende, neu_e, nachbarn[1]))):
        if nachbar is None:
            continue
        if nachbar == FEST:
            variante.gueltig = False
            variante.grund = t(
                u"Ende hängt an Abzweig oder Gerät - Verschieben geht nicht",
                u"End is connected to a tee or equipment - cannot shift",
                u"El extremo está conectado a una derivación o equipo; no "
                u"se puede desplazar")
            break
        nachbar_richtung, nachbar_laenge = nachbar
        anteil = skalar(nachbar_richtung, richtung)
        if abs(anteil) < 0.98:
            variante.gueltig = False
            variante.grund = t(
                u"Anschlussleitung liegt nicht in Verschieberichtung",
                u"Connected segment is not aligned with the shift",
                u"El tramo conectado no está en la dirección del "
                u"desplazamiento")
            break
        # Zeigt der Nachbar in Verschieberichtung, wird er kürzer
        aenderung = -versatz * anteil
        if nachbar_laenge + aenderung < mindest:
            variante.gueltig = False
            variante.grund = t(u"Anschlussleitung würde zu kurz",
                               u"Connected segment would become too short",
                               u"El tramo conectado quedaría demasiado "
                               u"corto")
            break
        zusatz += aenderung
        if aenderung > 0:
            stuecke.append((punkt, neu) if ende_nr == 0 else (neu, punkt))
    variante.zusatzlaenge = max(zusatz, 0.0)
    variante.pruefstuecke = stuecke
    variante.skizze = dict(skizze)
    variante.skizze[u"pfad"] = [(0.0, 0.0), (0.0, versatz),
                                (skizze[u"laenge"], versatz),
                                (skizze[u"laenge"], 0.0)]
    return variante


def _markiere_abstand(variante, abstand, soll_abstand):
    variante.abstand = abstand
    if soll_abstand is None or abstand >= soll_abstand - 1e-9:
        return
    variante.aufschlag += AUFSCHLAG_ABSTAND
    variante.warnungen.insert(0, (t(
        u"Nur %d cm Sicherheitsabstand (Vorgabe %d cm)",
        u"Only %d cm clearance (target %d cm)",
        u"Solo %d cm de distancia de seguridad (objetivo %d cm)") % (
        round(abstand * 100.0), round(soll_abstand * 100.0)), True))


def abstands_runden(soll_abstand):
    """Die Abstände, mit denen nacheinander gerechnet wird: erst die
    Vorgabe, dann kleinere Rückfallwerte."""
    return [soll_abstand] + [a for a in RUECKFALL_ABSTAENDE
                             if a < soll_abstand - 1e-9]


def varianten(start, ende, querschnitt, hindernis, abstand=0.05,
              winkel=(45, 90), max_versatz=MAX_VERSATZ, nachbarn=None,
              nur_richtungen=None, fester_versatz=None, soll_abstand=None):
    """Alle Umgehungen (gültige und verworfene) für eine Leitung.

    hindernis: Hüllkasten (min, max) des Hindernisses

    Leitungen mit Gefälle werden auch umgangen: das versetzte Mittelstück
    läuft parallel zur Leitung, das Gefälle bleibt dort erhalten. Nach oben
    entstünde aber ein Hochpunkt (Luftsack, bei Abwasser unzulässig) - diese
    Varianten bekommen eine Warnung und rutschen in der Bewertung nach hinten.

    nachbarn: (anfang, ende) für die Variante "Verschieben", siehe
    _verschieben(); None lässt sie weg.
    nur_richtungen: nur diese Ausweichrichtungen (z.B. [u"unten"])
    fester_versatz: statt aus dem Hindernis berechnet - für "beide weichen
                    je zur Hälfte aus"
    soll_abstand:   Vorgabe; liegt 'abstand' darunter, wird gewarnt
    """
    gefaelle = hat_gefaelle(start, ende)
    achse = einheit(minus(ende, start))
    gesamt = laenge(minus(ende, start))
    ecken = [minus(ecke, start) for ecke in _ecken(hindernis)]
    t0 = min(skalar(e, achse) for e in ecken)
    t1 = max(skalar(e, achse) for e in ecken)

    groesse = querschnitt.groesste
    halb_axial = groesse / 2.0
    ergebnis = []
    for code, richtung in sorted(richtungen(achse).items()):
        if nur_richtungen is not None and code not in nur_richtungen:
            continue
        werte = [skalar(e, richtung) for e in ecken]
        if fester_versatz is not None:
            versatz = fester_versatz
        else:
            oberkante = max(werte)
            versatz = oberkante + abstand + querschnitt.halb(richtung)
            versatz = max(versatz, groesse + abstand)
        for grad in winkel:
            variante = Variante(code, grad)
            variante.versatz = versatz
            if gefaelle:
                _gefaelle_hinweise(variante, richtung)
            bogen = math.radians(grad)
            schraeg = versatz / math.sin(bogen)
            anlauf = versatz / math.tan(bogen) if grad < 90 else 0.0
            # Platz für die Bögen: 90° brauchen mehr als 45°
            mindest = groesse * (2.0 if grad >= 90 else 1.0) + 0.02

            t_a = t0 - abstand - halb_axial
            t_b = t1 + abstand + halb_axial
            t_p1 = t_a - anlauf
            t_p2 = t_b + anlauf
            versetzt = mal(richtung, versatz)
            p1 = plus(start, mal(achse, t_p1))
            p2 = plus(start, mal(achse, t_p2))
            variante.punkte = [p1, plus(plus(start, mal(achse, t_a)),
                                         versetzt),
                               plus(plus(start, mal(achse, t_b)), versetzt),
                               p2]
            variante.zusatzlaenge = 2.0 * (schraeg - anlauf)
            variante.skizze = {
                u"laenge": gesamt,
                u"hindernis_t": (t0, t1),
                u"hindernis_d": (min(werte), max(werte)),
                u"pfad": [(0.0, 0.0), (t_p1, 0.0), (t_a, versatz),
                          (t_b, versatz), (t_p2, 0.0), (gesamt, 0.0)],
                u"halb": querschnitt.halb(richtung),
            }

            if versatz > max_versatz:
                variante.gueltig = False
                variante.grund = t(u"Versatz %d cm zu groß",
                                   u"Offset %d cm too large",
                                   u"Desvío de %d cm demasiado grande") % (
                    round(versatz * 100.0))
            elif t_p1 < mindest or t_p2 > gesamt - mindest:
                variante.gueltig = False
                variante.grund = t(
                    u"Leitung zu kurz für die Umgehung (Stück %.2f m, "
                    u"nötig %.2f m)",
                    u"Segment too short for the bypass (%.2f m, needs "
                    u"%.2f m)",
                    u"Tramo demasiado corto para el desvío (%.2f m, "
                    u"necesita %.2f m)") % (
                    gesamt, t_p2 - t_p1 + 2.0 * mindest)
            elif schraeg < mindest or (t_b - t_a) < mindest:
                variante.gueltig = False
                variante.grund = t(u"Kein Platz für die Bögen",
                                   u"No room for the elbows",
                                   u"No hay espacio para los codos")
            _markiere_abstand(variante, abstand, soll_abstand)
            ergebnis.append(variante)

        if nachbarn is not None:
            skizze = {u"laenge": gesamt, u"hindernis_t": (t0, t1),
                      u"hindernis_d": (min(werte), max(werte)),
                      u"halb": querschnitt.halb(richtung)}
            variante = _verschieben(start, ende, code, richtung, versatz,
                                    nachbarn, groesse + 0.02, skizze)
            if gefaelle and variante.gueltig and \
                    skalar(richtung, (0.0, 0.0, 1.0)) > 0.1:
                variante.hochpunkt = True
                variante.warnungen.append((t(
                    u"Leitung wird angehoben - Gefälle der "
                    u"Anschlussleitungen prüfen",
                    u"Segment is raised - check the slope of the connected "
                    u"segments",
                    u"El tramo sube: compruebe la pendiente de los tramos "
                    u"conectados"), True))
            if versatz > max_versatz and variante.gueltig:
                variante.gueltig = False
                variante.grund = t(u"Versatz %d cm zu groß",
                                   u"Offset %d cm too large",
                                   u"Desvío de %d cm demasiado grande") % (
                    round(versatz * 100.0))
            _markiere_abstand(variante, abstand, soll_abstand)
            ergebnis.append(variante)
    return ergebnis


def voller_versatz(start, ende, querschnitt, hindernis, abstand, richtung):
    """Versatz, den die Leitung allein in dieser Richtung bräuchte."""
    code_liste = [richtung]
    liste = varianten(start, ende, querschnitt, hindernis, abstand=abstand,
                      winkel=(90,), nur_richtungen=code_liste)
    return liste[0].versatz if liste else None


class Kombination(object):
    """Beide Leitungen weichen je zur Hälfte aus - z.B. A nach unten und
    B nach oben. teile: je eine Variante für A und B, jede mit festem
    halbem Versatz berechnet."""

    art = u"beide"

    def __init__(self, teile):
        self.teile = list(teile)
        self.leitung = None          # wird von aussen gesetzt (Anzeige)

    @property
    def gueltig(self):
        return all(teil.gueltig for teil in self.teile)

    @property
    def grund(self):
        for teil in self.teile:
            if not teil.gueltig:
                return teil.grund
        return u""

    @property
    def kollisionen(self):
        gesamt = []
        for teil in self.teile:
            for name in teil.kollisionen:
                if name not in gesamt:
                    gesamt.append(name)
        return gesamt

    @property
    def warnungen(self):
        gesamt = []
        for teil in self.teile:
            for warnung in teil.warnungen:
                if warnung not in gesamt:
                    gesamt.append(warnung)
        return gesamt

    @property
    def hochpunkt(self):
        return any(teil.hochpunkt for teil in self.teile)

    @property
    def versatz(self):
        return max(teil.versatz for teil in self.teile)

    @property
    def zusatzlaenge(self):
        return sum(teil.zusatzlaenge for teil in self.teile)

    @property
    def kosten(self):
        return sum(teil.kosten for teil in self.teile) + AUFSCHLAG_BEIDE

    @property
    def skizze(self):
        return self.teile[0].skizze

    @property
    def name(self):
        namen = []
        for teil in self.teile:
            kennung = getattr(getattr(teil, "leitung", None), "kennung", u"")
            namen.append((kennung + u" " + teil.name).strip())
        return u" + ".join(namen)

    def sortierschluessel(self):
        return (not self.gueltig, len(self.kollisionen) > 0,
                len(self.kollisionen), self.kosten)

    def zusammenfassung(self):
        if not self.gueltig:
            return self.grund
        return t(u"Beide je %d cm · +%.2f m",
                 u"Both %d cm each · +%.2f m",
                 u"Ambos %d cm cada uno · +%.2f m") % (
            round(self.versatz * 100.0), self.zusatzlaenge)


# Gegenrichtungen für "beide weichen aus": nur senkrecht, dort passen die
# Richtungen beider waagerechter Leitungen zusammen
GEGENRICHTUNGEN = ((u"unten", u"oben"), (u"oben", u"unten"))


def beste(liste, anzahl=3):
    """Die besten Varianten: gültig und kollisionsfrei zuerst, dann die
    kürzesten. Verworfene füllen nur auf, wenn sonst weniger als 'anzahl'
    übrig sind - so sieht man, warum es nicht geht."""
    return sorted(liste, key=lambda v: v.sortierschluessel())[:anzahl]
