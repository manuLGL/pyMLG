# -*- coding: utf-8 -*-
"""Revit-freie Logik von FilterLegend.

Aus den Ansichten kommen Einträge (je Filter einer) mit den Grafik-
Überschreibungen als einfache Werte. Hier wird entschieden, welche Füllung
das Farbfeld bekommt, Einträge mehrerer Ansichten zusammengeführt,
sortiert und das Raster der Legende berechnet (alles offline testbar).

Muster sind für die Logik undurchsichtige Werte (in Revit ElementIds),
None heisst "nicht gesetzt", VOLL steht für <Vollton>.
Farben sind (r, g, b) oder None.
"""

import re

VOLL = u"<vollton>"
SCHWARZ = (0, 0, 0)

# welche Überschreibung das Farbfeld bestimmt
QUELLE_AUTO = 0
QUELLE_OBERFLAECHE = 1
QUELLE_SCHNITT = 2
QUELLEN = (
    (QUELLE_AUTO, (u"Automatisch", u"Automatic", u"Automático")),
    (QUELLE_OBERFLAECHE, (u"Oberfläche (Projektion)", u"Surface (projection)",
                          u"Superficie (proyección)")),
    (QUELLE_SCHNITT, (u"Schnitt", u"Cut", u"Corte")),
)

REIHENFOLGE_ANSICHT = 0
REIHENFOLGE_NAME = 1
REIHENFOLGEN = (
    (REIHENFOLGE_ANSICHT, (u"wie in der Ansicht", u"as in the view", u"como en la vista")),
    (REIHENFOLGE_NAME, (u"nach Name", u"by name", u"por nombre")),
)


class Fuellung(object):
    """Vorder- und Hintergrundmuster mit Farbe - wie ein Füllbereichstyp."""

    def __init__(self, vg_muster=None, vg_farbe=None, hg_muster=None, hg_farbe=None):
        self.vg_muster = vg_muster
        self.vg_farbe = tuple(vg_farbe) if vg_farbe else None
        self.hg_muster = hg_muster
        self.hg_farbe = tuple(hg_farbe) if hg_farbe else None

    def leer(self):
        return not (self.vg_muster or self.vg_farbe or self.hg_muster or self.hg_farbe)

    def __eq__(self, other):
        return isinstance(other, Fuellung) and self._werte() == other._werte()

    def __ne__(self, other):
        return not self == other

    def __hash__(self):
        return hash(self._werte())

    def _werte(self):
        return (self.vg_muster, self.vg_farbe, self.hg_muster, self.hg_farbe)

    def __repr__(self):
        return u"Fuellung(%r, %r, %r, %r)" % self._werte()


class Grafik(object):
    """Die für die Legende wichtigen Überschreibungen eines Filters."""

    def __init__(self, oberflaeche=None, schnitt=None, linie=None, schnittlinie=None):
        self.oberflaeche = oberflaeche      # Fuellung oder None
        self.schnitt = schnitt              # Fuellung oder None
        self.linie = tuple(linie) if linie else None              # Projektionslinien
        self.schnittlinie = tuple(schnittlinie) if schnittlinie else None


class Eintrag(object):
    """Ein Filter, wie er in einer Ansicht eingestellt ist."""

    def __init__(self, schluessel, ref, name, grafik, sichtbar=True, aktiv=True,
                 ansicht=u""):
        self.schluessel = schluessel    # Zahl der Filter-Id
        self.ref = ref                  # Filter-Id für Revit
        self.name = name or u""
        self.grafik = grafik or Grafik()
        self.sichtbar = sichtbar        # Häkchen "Sichtbarkeit" im Filter-Reiter
        self.aktiv = aktiv              # Häkchen "Aktivieren" (Revit 2021+)
        self.ansicht = ansicht          # Name der Ansicht, aus der er stammt
        self.abweichend = []            # weitere Ansichten mit anderer Grafik
        self.position = 0               # Reihenfolge in der Ansicht

    def __repr__(self):
        return u"Eintrag(%s)" % self.name


def _bereinigt(fuellung):
    """Fuellung so ergänzen, wie Revit sie zeichnet: Farbe ohne Muster ->
    Vollton, Muster ohne Farbe -> schwarz. Leer -> None."""
    if fuellung is None or fuellung.leer():
        return None
    vg_muster, vg_farbe = fuellung.vg_muster, fuellung.vg_farbe
    hg_muster, hg_farbe = fuellung.hg_muster, fuellung.hg_farbe
    if vg_farbe and not vg_muster:
        vg_muster = VOLL
    if vg_muster and not vg_farbe:
        vg_farbe = SCHWARZ
    if hg_farbe and not hg_muster:
        hg_muster = VOLL
    if hg_muster and not hg_farbe:
        hg_farbe = SCHWARZ
    # Vollton vorne verdeckt den Hintergrund
    if vg_muster == VOLL:
        hg_muster, hg_farbe = None, None
    return Fuellung(vg_muster, vg_farbe, hg_muster, hg_farbe)


def darstellung(grafik, quelle=QUELLE_AUTO):
    """Füllung für das Farbfeld eines Filters oder None (Filter ändert
    keine Farbe).

    Automatisch: Oberflächenmuster, sonst Schnittmuster, sonst die Farbe
    der Projektions- bzw. Schnittlinien als Vollton. Bei fester Quelle nur
    deren Muster, sonst deren Linienfarbe."""
    if grafik is None:
        return None
    if quelle == QUELLE_OBERFLAECHE:
        reihe = ((grafik.oberflaeche, grafik.linie),)
    elif quelle == QUELLE_SCHNITT:
        reihe = ((grafik.schnitt, grafik.schnittlinie),)
    else:
        reihe = ((grafik.oberflaeche, None), (grafik.schnitt, None),
                 (None, grafik.linie), (None, grafik.schnittlinie))
    for fuellung, linie in reihe:
        ergebnis = _bereinigt(fuellung)
        if ergebnis is not None:
            return ergebnis
        if linie:
            return Fuellung(VOLL, linie)
    return None


def vorgewaehlt(eintrag, quelle=QUELLE_AUTO):
    """Standard-Häkchen im Dialog: aktiv, sichtbar und mit Farbe."""
    return bool(eintrag.aktiv and eintrag.sichtbar
                and darstellung(eintrag.grafik, quelle) is not None)


def _gleiche_grafik(a, b):
    return all(darstellung(a, q) == darstellung(b, q)
               for q in (QUELLE_OBERFLAECHE, QUELLE_SCHNITT))


def zusammenfuehren(listen):
    """Einträge mehrerer Ansichten zu einer Liste (je Filter einer).
    Der erste Fund gilt; zeigt eine spätere Ansicht denselben Filter
    anders, steht sie in eintrag.abweichend. position = erste Fundstelle."""
    ergebnis = []
    gefunden = {}
    for liste in listen:
        for eintrag in liste:
            erster = gefunden.get(eintrag.schluessel)
            if erster is None:
                eintrag.position = len(ergebnis)
                eintrag.abweichend = []
                gefunden[eintrag.schluessel] = eintrag
                ergebnis.append(eintrag)
            elif not _gleiche_grafik(erster.grafik, eintrag.grafik) \
                    and eintrag.ansicht not in erster.abweichend:
                erster.abweichend.append(eintrag.ansicht)
    return ergebnis


def natuerlich(text):
    """Sortierschlüssel, der Zahlen als Zahlen vergleicht (F2 < F10)."""
    return [(0, int(teil), u"") if teil.isdigit() else (1, 0, teil)
            for teil in re.split(r"(\d+)", (text or u"").lower()) if teil]


def sortiere(eintraege, art=REIHENFOLGE_ANSICHT):
    if art == REIHENFOLGE_NAME:
        return sorted(eintraege, key=lambda e: natuerlich(e.name))
    return sorted(eintraege, key=lambda e: e.position)


def raster(anzahl, spalten, feld_hoehe, abstand, spaltenbreite):
    """Linke obere Ecke (x, y) jedes Farbfelds, Papier-mm, y nach oben.
    Gefüllt wird spaltenweise von oben nach unten, die Spalten sind
    möglichst gleich lang."""
    spalten = max(1, min(int(spalten), anzahl or 1))
    je_spalte = -(-anzahl // spalten)       # aufgerundet
    punkte = []
    for index in range(anzahl):
        spalte, zeile = divmod(index, je_spalte)
        punkte.append((spalte * spaltenbreite, -zeile * (feld_hoehe + abstand)))
    return punkte


def beschriftung(eintrag, eigene=None):
    """Text neben dem Farbfeld: eigene Beschriftung oder der Filtername."""
    text = (eigene or u"").strip()
    return text or eintrag.name


def zahl(text):
    try:
        return float((u"%s" % (text or u"")).strip().replace(u",", u"."))
    except ValueError:
        return None


def pruefe(text, kleinst, groesst, ganz=False):
    """Zahl im Bereich [kleinst, groesst] oder None."""
    wert = zahl(text)
    if wert is None or not kleinst <= wert <= groesst:
        return None
    if ganz:
        if wert != int(wert):
            return None
        return int(wert)
    return wert


def hexfarbe(rgb):
    return u"#%02X%02X%02X" % tuple(int(k) for k in rgb)
