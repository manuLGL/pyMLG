# -*- coding: utf-8 -*-
"""Eigener DXF-Leser für DxfLegend (ASCII-DXF, ohne Revit).

ezdxf scheidet aus: es braucht numpy, das in der pyRevit-CPython-Umgebung
nicht vorhanden ist. Eine Legende braucht nur wenige Objekttypen:

    LINE, ARC, CIRCLE, ELLIPSE, LWPOLYLINE, POLYLINE, SPLINE, LEADER
    TEXT, MTEXT, ATTRIB
    HATCH (Volltonfüllung und Muster), SOLID, TRACE
    INSERT (Blöcke, auch verschachtelt und als Raster), DIMENSION

Alles wird in Weltkoordinaten aufgelöst - Blöcke werden aufgelöst, Farbe,
Linientyp und Linienstärke sind fertig bestimmt (VONLAYER/VONBLOCK).

    zeichnung = lese_datei(pfad)
    zeichnung.kurven / .texte / .flaechen / .linientypen / .ignoriert

Längen bleiben in Zeichnungseinheiten; der Massstab kommt in logik.py.
"""

import io
import math
import re

from dxf_legende import geometrie as geo

# Farbe für ACI 7 (weiss/schwarz) - logik.py macht daraus auf weissem
# Papier schwarz
WEISS = (255, 255, 255)

MAX_TIEFE = 32

# Unterlängen ab Grundlinie in Texthöhen (für TEXT "unten")
UNTERLAENGE = 0.3


class DxfFehler(Exception):
    """Die Datei ist kein lesbares DXF."""


class Kurve(object):
    """Polylinie mit Bögen in Weltkoordinaten."""

    def __init__(self, punkte, geschlossen, farbe, linientyp, lt_faktor,
                 staerke_mm, breite, layer):
        self.punkte = punkte            # [(x, y, bulge)]
        self.geschlossen = geschlossen
        self.farbe = farbe              # (r, g, b)
        self.linientyp = linientyp      # Name in Grossbuchstaben
        self.lt_faktor = lt_faktor      # Linientypfaktor des Objekts * Block
        self.staerke_mm = staerke_mm    # None = Vorgabe
        self.breite = breite            # Polylinienbreite (Einheiten)
        self.layer = layer


class Text(object):
    """Einzeiliger oder mehrzeiliger Text, verankert wie in Revit:
    h_ausr links/mitte/rechts, v_ausr oben/mitte/unten. Bei "mitte" liegt
    der Punkt auf halber Versalhöhe der (einzigen) Zeile."""

    def __init__(self, inhalt, x, y, hoehe, drehung, breitenfaktor, schrift,
                 h_ausr, v_ausr, farbe, layer):
        self.inhalt = inhalt
        self.x = x
        self.y = y
        self.hoehe = hoehe
        self.drehung = drehung          # Bogenmass
        self.breitenfaktor = breitenfaktor
        self.schrift = schrift          # Schriftname oder None (SHX)
        self.h_ausr = h_ausr
        self.v_ausr = v_ausr
        self.farbe = farbe
        self.layer = layer


class Flaeche(object):
    """Gefüllte Fläche (Schraffur oder SOLID).

    muster None = Volltonfüllung, sonst Liste von Musterlinien
    (winkel, (bx, by), (dx, dy), [striche]) in Weltkoordinaten -
    striche wie in DXF: positiv Strich, negativ Lücke, 0 Punkt."""

    def __init__(self, schleifen, farbe, muster_name, muster, layer):
        self.schleifen = schleifen      # [[(x, y, bulge)]], geschlossen
        self.farbe = farbe
        self.muster_name = muster_name
        self.muster = muster
        self.layer = layer


class Zeichnung(object):

    def __init__(self):
        self.kurven = []
        self.texte = []
        self.flaechen = []
        self.linientypen = {}   # NAME -> [striche] (leer = durchgezogen)
        self.einheiten = 0      # $INSUNITS
        self.ltscale = 1.0      # $LTSCALE
        self.ignoriert = {}     # Objekttyp -> Anzahl
        self.version = u""

    def ist_leer(self):
        return not (self.kurven or self.texte or self.flaechen)


# ---------------------------------------------------------------------------
# AutoCAD-Farbindex
# ---------------------------------------------------------------------------

_ACI_GRUND = {1: (255, 0, 0), 2: (255, 255, 0), 3: (0, 255, 0),
              4: (0, 255, 255), 5: (0, 0, 255), 6: (255, 0, 255),
              7: WEISS, 8: (128, 128, 128), 9: (192, 192, 192)}
_ACI_GRAU = {250: (51, 51, 51), 251: (80, 80, 80), 252: (105, 105, 105),
             253: (130, 130, 130), 254: (190, 190, 190), 255: WEISS}
_ACI_HELLIGKEIT = (255, 204, 153, 127, 76)


def aci_rgb(index):
    """AutoCAD-Farbindex 1..255 -> (r, g, b)."""
    index = abs(int(index))
    if index in _ACI_GRUND:
        return _ACI_GRUND[index]
    if index in _ACI_GRAU:
        return _ACI_GRAU[index]
    if not 10 <= index <= 249:
        return WEISS
    farbton = (index // 10 - 1) * 15.0
    stufe = index % 10
    wert = float(_ACI_HELLIGKEIT[stufe // 2])
    saettigung = 0.5 if stufe % 2 else 1.0
    return _hsv(farbton, saettigung, wert)


def _hsv(farbton, saettigung, wert):
    h = (farbton % 360.0) / 60.0
    sektor = int(h)
    anteil = h - sektor
    p = wert * (1.0 - saettigung)
    q = wert * (1.0 - saettigung * anteil)
    t_ = wert * (1.0 - saettigung * (1.0 - anteil))
    r, g, b = [(wert, t_, p), (q, wert, p), (p, wert, t_),
               (p, q, wert), (t_, p, wert), (wert, p, q)][sektor % 6]
    return (int(r), int(g), int(b))


def truecolor(wert):
    wert = int(wert) & 0xFFFFFF
    return ((wert >> 16) & 255, (wert >> 8) & 255, wert & 255)


# ---------------------------------------------------------------------------
# Texte
# ---------------------------------------------------------------------------

_SONDER = {u"c": u"\u00d8", u"d": u"\u00b0", u"p": u"\u00b1",
           u"%": u"%", u"u": u"", u"o": u"", u"k": u""}


def _sonderzeichen(text):
    """%%c, %%d, %%p, %%nnn und \\U+XXXX auflösen."""
    def ersetze_prozent(treffer):
        code = treffer.group(1)
        if code.isdigit():
            try:
                return chr(int(code))
            except (ValueError, OverflowError):
                return u""
        return _SONDER.get(code.lower(), u"")
    text = re.sub(r"%%(\d{3}|.)", ersetze_prozent, text)
    return re.sub(r"\\U\+([0-9A-Fa-f]{4})",
                  lambda t_: chr(int(t_.group(1), 16)), text)


def bereinige_mtext(roh):
    """MTEXT-Formatierung entfernen.

    Rückgabe (text, format): format enthält die Angaben, die vor dem
    ersten sichtbaren Zeichen stehen und damit für den ganzen Text gelten:
    'schrift', 'hoehe' (absolut) bzw. 'hoehe_faktor', 'breite', 'aci',
    'rgb'."""
    text = _sonderzeichen(roh)
    ausgabe = []
    fmt = {}
    i = 0
    n = len(text)

    def vorne():
        return not u"".join(ausgabe).strip()

    while i < n:
        z = text[i]
        if z in u"{}":
            i += 1
            continue
        if z != u"\\" or i + 1 >= n:
            ausgabe.append(z)
            i += 1
            continue
        code = text[i + 1]
        if code in u"\\{}":
            ausgabe.append(code)
            i += 2
        elif code == u"P":
            ausgabe.append(u"\n")
            i += 2
        elif code in u"~":
            ausgabe.append(u" ")
            i += 2
        elif code in u"LlOoKkNn":
            i += 2
        elif code == u"S":
            ende = text.find(u";", i)
            ende = n if ende < 0 else ende
            teile = re.split(r"[\^/#]", text[i + 2:ende], maxsplit=1)
            ausgabe.append(u"/".join(teil.strip() for teil in teile if teil.strip()))
            i = ende + 1
        elif code in u"fFHWQTACcp":
            ende = text.find(u";", i)
            ende = n if ende < 0 else ende
            wert = text[i + 2:ende]
            if vorne():
                _merke_format(fmt, code, wert)
            i = ende + 1
        else:
            # Unbekannter Code: Zeichen behalten
            ausgabe.append(code)
            i += 2
    zeilen = [z.rstrip() for z in u"".join(ausgabe).split(u"\n")]
    while zeilen and not zeilen[-1]:
        zeilen.pop()
    return u"\n".join(zeilen), fmt


def _merke_format(fmt, code, wert):
    try:
        if code == u"f" and u"schrift" not in fmt:
            name = wert.split(u"|")[0].strip()
            if name:
                fmt[u"schrift"] = name
        elif code == u"H" and u"hoehe" not in fmt and u"hoehe_faktor" not in fmt:
            if wert.lower().endswith(u"x"):
                fmt[u"hoehe_faktor"] = float(wert[:-1])
            else:
                fmt[u"hoehe"] = float(wert)
        elif code == u"W" and u"breite" not in fmt:
            fmt[u"breite"] = float(wert.rstrip(u"xX"))
        elif code == u"C" and u"aci" not in fmt:
            fmt[u"aci"] = int(wert)
        elif code == u"c" and u"rgb" not in fmt:
            # \c: BGR-Ganzzahl
            bgr = int(wert)
            fmt[u"rgb"] = (bgr & 255, (bgr >> 8) & 255, (bgr >> 16) & 255)
    except ValueError:
        pass


_TTF_NAMEN = {
    u"arial": u"Arial", u"arialbd": u"Arial", u"ariali": u"Arial",
    u"arialn": u"Arial Narrow", u"arialnb": u"Arial Narrow",
    u"times": u"Times New Roman", u"timesbd": u"Times New Roman",
    u"calibri": u"Calibri", u"calibrib": u"Calibri", u"verdana": u"Verdana",
    u"tahoma": u"Tahoma", u"segoeui": u"Segoe UI", u"cour": u"Courier New",
    u"isocpeur": u"ISOCPEUR", u"isocteur": u"ISOCTEUR", u"gothic": u"Century Gothic",
    u"swiss": u"Arial", u"swissl": u"Arial", u"swisscl": u"Arial Narrow",
    u"romantic": u"Times New Roman",
}


def schrift_aus_datei(datei):
    """Schriftdatei aus der Stiltabelle -> Schriftname (None bei SHX)."""
    datei = (datei or u"").strip()
    if not datei:
        return None
    name = datei.replace(u"\\", u"/").split(u"/")[-1]
    stamm, _, endung = name.rpartition(u".")
    if not stamm:
        stamm, endung = name, u""
    if endung.lower() not in (u"ttf", u"otf", u"ttc"):
        return None if endung.lower() == u"shx" or not endung else stamm
    return _TTF_NAMEN.get(stamm.lower(), stamm[:1].upper() + stamm[1:])


# ---------------------------------------------------------------------------
# Einlesen
# ---------------------------------------------------------------------------

def lese_datei(pfad):
    with io.open(pfad, "rb") as datei:
        roh = datei.read()
    return lese_bytes(roh)


def lese_bytes(roh):
    if roh[:18] == b"AutoCAD Binary DXF":
        raise DxfFehler(u"binary")
    vorschau = roh[:200000].decode("latin-1")
    version = _kopfwert(vorschau, u"$ACADVER") or u""
    if version >= u"AC1021":
        text = roh.decode("utf-8", "replace")
    else:
        seite = (_kopfwert(vorschau, u"$DWGCODEPAGE") or u"").upper()
        kodierung = u"cp1252"
        if seite.startswith(u"ANSI_"):
            kodierung = u"cp" + seite[5:]
        try:
            text = roh.decode(kodierung, "replace")
        except LookupError:
            text = roh.decode("cp1252", "replace")
    if text.startswith(u"\ufeff"):
        text = text[1:]
    return lese_text(text)


def _kopfwert(text, name):
    treffer = re.search(r"\n\s*9\s*\r?\n\s*" + re.escape(name)
                        + r"\s*\r?\n\s*\d+\s*\r?\n([^\r\n]*)", text)
    return treffer.group(1).strip() if treffer else None


def _paare(text):
    zeilen = text.splitlines()
    if len(zeilen) % 2:
        zeilen = zeilen[:-1]
    paare = []
    for i in range(0, len(zeilen), 2):
        try:
            code = int(zeilen[i].strip())
        except ValueError:
            raise DxfFehler(u"line %d" % (i + 1))
        wert = zeilen[i + 1]
        if code not in (1, 3):
            wert = wert.strip()
        paare.append((code, wert))
    return paare


def _abschnitte(paare):
    abschnitte = {}
    i = 0
    n = len(paare)
    while i < n:
        if paare[i] == (0, u"SECTION") and i + 1 < n and paare[i + 1][0] == 2:
            name = paare[i + 1][1].upper()
            j = i + 2
            while j < n and paare[j] != (0, u"ENDSEC"):
                j += 1
            abschnitte[name] = paare[i + 2:j]
            i = j + 1
        else:
            i += 1
    return abschnitte


def _objekte(paare):
    """Paare -> [(TYP, [(code, wert)])] je Objekt (Code-0-Gruppe)."""
    objekte = []
    aktuell = None
    for code, wert in paare:
        if code == 0:
            aktuell = (wert.strip().upper(), [])
            objekte.append(aktuell)
        elif aktuell is not None:
            aktuell[1].append((code, wert))
    return objekte


def _gruppiere(objekte):
    """POLYLINE + VERTEX…SEQEND und INSERT + ATTRIB…SEQEND zusammenfassen.
    Rückgabe [(typ, tags, unterobjekte)]."""
    ergebnis = []
    i = 0
    n = len(objekte)
    while i < n:
        typ, tags = objekte[i]
        unter = []
        if typ == u"POLYLINE" or (typ == u"INSERT" and _wert(tags, 66, 0, int) == 1):
            j = i + 1
            while j < n and objekte[j][0] in (u"VERTEX", u"ATTRIB"):
                unter.append(objekte[j])
                j += 1
            if j < n and objekte[j][0] == u"SEQEND":
                j += 1
            i = j
        else:
            i += 1
        ergebnis.append((typ, tags, unter))
    return ergebnis


def _wert(tags, code, vorgabe=None, art=None):
    for c, w in tags:
        if c == code:
            if art is None:
                return w
            try:
                return art(w)
            except ValueError:
                return vorgabe
    return vorgabe


def _zahl(tags, code, vorgabe=0.0):
    return _wert(tags, code, vorgabe, float)


def lese_text(text):
    paare = _paare(text)
    abschnitte = _abschnitte(paare)
    if u"ENTITIES" not in abschnitte and u"BLOCKS" not in abschnitte:
        raise DxfFehler(u"no entities")
    leser = _Leser()
    leser.kopf(abschnitte.get(u"HEADER", []))
    leser.tabellen(abschnitte.get(u"TABLES", []))
    leser.bloecke(abschnitte.get(u"BLOCKS", []))
    leser.zeichne_alles(_gruppiere(_objekte(abschnitte.get(u"ENTITIES", []))))
    return leser.zeichnung


class _Kontext(object):
    """Eigenschaften der Blockreferenz für VONBLOCK und Layer 0."""

    def __init__(self, layer=None, farbe=WEISS, linientyp=u"CONTINUOUS",
                 staerke=None, lt_faktor=1.0):
        self.layer = layer
        self.farbe = farbe
        self.linientyp = linientyp
        self.staerke = staerke
        self.lt_faktor = lt_faktor


class _Layer(object):

    def __init__(self, farbe=WEISS, linientyp=u"CONTINUOUS", staerke=None,
                 sichtbar=True):
        self.farbe = farbe
        self.linientyp = linientyp
        self.staerke = staerke
        self.sichtbar = sichtbar


class _Leser(object):

    def __init__(self):
        self.zeichnung = Zeichnung()
        self.layer = {}
        self.stile = {}         # TEXTSTIL -> (schrift, breite, hoehe)
        self.blockdef = {}      # NAME -> (basis, objekte)

    # -- Kopf und Tabellen ---------------------------------------------------

    def kopf(self, paare):
        variable = None
        for code, wert in paare:
            if code == 9:
                variable = wert.upper()
            elif variable == u"$INSUNITS" and code == 70:
                self.zeichnung.einheiten = int(wert)
            elif variable == u"$LTSCALE" and code == 40:
                try:
                    self.zeichnung.ltscale = float(wert) or 1.0
                except ValueError:
                    pass
            elif variable == u"$ACADVER" and code == 1:
                self.zeichnung.version = wert

    def tabellen(self, paare):
        for typ, tags in _objekte(paare):
            name = (_wert(tags, 2) or u"").strip()
            if typ == u"LAYER":
                self._layer(name.upper(), tags)
            elif typ == u"LTYPE":
                striche = [float(w) for c, w in tags if c == 49]
                self.zeichnung.linientypen[name.upper()] = striche
            elif typ == u"STYLE":
                schrift = None
                for c, w in tags:
                    if c == 1000 and w.strip():
                        schrift = w.strip()
                        break
                if schrift is None:
                    schrift = schrift_aus_datei(_wert(tags, 3))
                self.stile[name.upper()] = (schrift, _zahl(tags, 41, 1.0) or 1.0,
                                            _zahl(tags, 40, 0.0))

    def _layer(self, name, tags):
        aci = _wert(tags, 62, 7, int)
        tc = _wert(tags, 420, None, int)
        farbe = truecolor(tc) if tc is not None else aci_rgb(aci or 7)
        flags = _wert(tags, 70, 0, int)
        staerke = _wert(tags, 370, -3, int)
        self.layer[name] = _Layer(
            farbe=farbe,
            linientyp=(_wert(tags, 6) or u"CONTINUOUS").upper(),
            staerke=staerke / 100.0 if staerke >= 0 else None,
            sichtbar=aci >= 0 and not flags & 1)

    def bloecke(self, paare):
        name = None
        basis = (0.0, 0.0)
        inhalt = []
        for typ, tags in _objekte(paare):
            if typ == u"BLOCK":
                name = (_wert(tags, 2) or u"").upper()
                basis = (_zahl(tags, 10), _zahl(tags, 20))
                inhalt = []
            elif typ == u"ENDBLK":
                if name is not None:
                    self.blockdef[name] = (basis, _gruppiere(inhalt))
                name = None
            elif name is not None:
                inhalt.append((typ, tags))

    # -- Eigenschaften ---------------------------------------------------------

    def _layer_von(self, tags, kontext):
        name = (_wert(tags, 8) or u"0").upper()
        if name == u"0" and kontext.layer is not None:
            name = kontext.layer
        return name

    def _eigenschaften(self, tags, kontext):
        """(sichtbar, layer, farbe, linientyp, stärke, lt_faktor)."""
        layername = self._layer_von(tags, kontext)
        layer = self.layer.get(layername, _Layer())
        tc = _wert(tags, 420, None, int)
        aci = _wert(tags, 62, 256, int)
        if tc is not None:
            farbe = truecolor(tc)
        elif aci == 0:
            farbe = kontext.farbe
        elif aci == 256 or aci is None:
            farbe = layer.farbe
        else:
            farbe = aci_rgb(aci)
        lt = (_wert(tags, 6) or u"BYLAYER").upper()
        if lt == u"BYLAYER":
            lt = layer.linientyp
        elif lt == u"BYBLOCK":
            lt = kontext.linientyp
        st = _wert(tags, 370, -1, int)
        if st == -1:
            staerke = layer.staerke
        elif st == -2:
            staerke = kontext.staerke
        elif st < 0:
            staerke = None
        else:
            staerke = st / 100.0
        sichtbar = layer.sichtbar and _wert(tags, 60, 0, int) != 1
        lt_faktor = (_zahl(tags, 48, 1.0) or 1.0) * kontext.lt_faktor
        return sichtbar, layername, farbe, lt, staerke, lt_faktor

    def _zaehle(self, typ):
        self.zeichnung.ignoriert[typ] = self.zeichnung.ignoriert.get(typ, 0) + 1

    # -- Zeichnen --------------------------------------------------------------

    def zeichne_alles(self, objekte):
        # Nur Modellbereich - es sei denn, es gibt dort nichts
        modell = [o for o in objekte if _wert(o[1], 67, 0, int) != 1]
        self.zeichne(modell or objekte, geo.EINHEIT, _Kontext(), 0)

    def zeichne(self, objekte, m, kontext, tiefe):
        for typ, tags, unter in objekte:
            try:
                self._objekt(typ, tags, unter, m, kontext, tiefe)
            except (ValueError, TypeError, ZeroDivisionError, IndexError,
                    KeyError):
                self._zaehle(typ + u" (?)")

    def _objekt(self, typ, tags, unter, m, kontext, tiefe):
        if typ in (u"INSERT", u"DIMENSION"):
            self._einfuegen(typ, tags, unter, m, kontext, tiefe)
            return
        eig = self._eigenschaften(tags, kontext)
        if not eig[0]:
            return
        if typ == u"LINE":
            punkte = [(_zahl(tags, 10), _zahl(tags, 20), 0.0),
                      (_zahl(tags, 11), _zahl(tags, 21), 0.0)]
            self._kurve(punkte, False, m, eig, 0.0)
        elif typ in (u"ARC", u"CIRCLE"):
            self._bogen(typ, tags, m, eig)
        elif typ == u"ELLIPSE":
            self._ellipse(tags, m, eig)
        elif typ == u"LWPOLYLINE":
            self._lwpolyline(tags, m, eig)
        elif typ == u"POLYLINE":
            self._polyline(tags, unter, m, eig)
        elif typ == u"SPLINE":
            self._spline(tags, m, eig)
        elif typ == u"LEADER":
            xs = [float(w) for c, w in tags if c == 10]
            ys = [float(w) for c, w in tags if c == 20]
            self._kurve([(x, y, 0.0) for x, y in zip(xs, ys)], False, m, eig, 0.0)
        elif typ in (u"TEXT", u"ATTRIB"):
            if typ == u"ATTRIB" and _wert(tags, 70, 0, int) & 1:
                return
            self._text(typ, tags, m, eig)
        elif typ == u"MTEXT":
            self._mtext(tags, m, eig)
        elif typ == u"HATCH":
            self._schraffur(tags, m, eig)
        elif typ in (u"SOLID", u"TRACE"):
            ecken = [(_zahl(tags, 10), _zahl(tags, 20)), (_zahl(tags, 11), _zahl(tags, 21)),
                     (_zahl(tags, 13, None), _zahl(tags, 23, None)),
                     (_zahl(tags, 12), _zahl(tags, 22))]
            if ecken[2][0] is None or (abs(ecken[2][0] - ecken[3][0]) < 1e-12
                                       and abs(ecken[2][1] - ecken[3][1]) < 1e-12):
                del ecken[2]
            schleife = [(x, y, 0.0) for x, y in ecken]
            schleife = self._ocs(tags, schleife)
            self._flaeche([geo.transformiere(schleife, True, m)], None, u"SOLID", eig)
        elif typ in (u"ATTDEF", u"POINT", u"VIEWPORT", u"SEQEND"):
            return
        else:
            self._zaehle(typ)

    # -- Hilfen ----------------------------------------------------------------

    @staticmethod
    def _gespiegelt(tags):
        """Objekt-Koordinatensystem mit Extrusion (0, 0, -1): x gespiegelt."""
        return _zahl(tags, 230, 1.0) < 0

    def _ocs(self, tags, punkte):
        if not self._gespiegelt(tags):
            return punkte
        return [(-p[0], p[1], -p[2]) for p in punkte]

    def _kurve(self, punkte, geschlossen, m, eig, breite):
        if len(punkte) < 2:
            return
        _s, layer, farbe, lt, staerke, ltf = eig
        punkte = geo.transformiere(punkte, geschlossen, m)
        faktor = geo.mittlerer_faktor(m)
        self.zeichnung.kurven.append(Kurve(
            punkte, geschlossen, farbe, lt, ltf * faktor, staerke,
            breite * faktor, layer))

    def _flaeche(self, schleifen, muster, muster_name, eig):
        schleifen = [s for s in schleifen if len(s) >= 2]
        if not schleifen:
            return
        self.zeichnung.flaechen.append(
            Flaeche(schleifen, eig[2], muster_name, muster, eig[1]))

    # -- Objekte ---------------------------------------------------------------

    def _bogen(self, typ, tags, m, eig):
        cx, cy, r = _zahl(tags, 10), _zahl(tags, 20), _zahl(tags, 40)
        if r <= 0:
            return
        if typ == u"CIRCLE":
            punkte = [(cx - r, cy, 1.0), (cx + r, cy, 1.0)]
            self._kurve(self._ocs(tags, punkte), True, m, eig, 0.0)
            return
        start = math.radians(_zahl(tags, 50))
        ende = math.radians(_zahl(tags, 51))
        oeffnung = (ende - start) % (2.0 * math.pi)
        if oeffnung < 1e-12:
            oeffnung = 2.0 * math.pi
        if oeffnung > math.pi * 1.999:
            mitte = start + oeffnung / 2.0
            punkte = [(cx + r * math.cos(start), cy + r * math.sin(start), geo.bulge_aus_winkel(oeffnung / 2.0)),
                      (cx + r * math.cos(mitte), cy + r * math.sin(mitte), geo.bulge_aus_winkel(oeffnung / 2.0)),
                      (cx + r * math.cos(start + oeffnung), cy + r * math.sin(start + oeffnung), 0.0)]
        else:
            punkte = [(cx + r * math.cos(start), cy + r * math.sin(start), geo.bulge_aus_winkel(oeffnung)),
                      (cx + r * math.cos(ende), cy + r * math.sin(ende), 0.0)]
        self._kurve(self._ocs(tags, punkte), False, m, eig, 0.0)

    def _ellipse(self, tags, m, eig):
        mitte = (_zahl(tags, 10), _zahl(tags, 20))
        achse = (_zahl(tags, 11), _zahl(tags, 21))
        start = _zahl(tags, 41, 0.0)
        ende = _zahl(tags, 42, 2.0 * math.pi)
        voll = abs((ende - start) - 2.0 * math.pi) < 1e-9
        # Mittelpunkt und Achse stehen in Weltkoordinaten (kein OCS)
        zug = geo.ellipse_punkte(mitte, achse, _zahl(tags, 40, 1.0), start, ende)
        if voll:
            zug = zug[:-1]
        self._kurve([(x, y, 0.0) for x, y in zug], voll, m, eig, 0.0)

    def _lwpolyline(self, tags, m, eig):
        punkte = []
        breite = _zahl(tags, 43, 0.0)
        for code, wert in tags:
            if code == 10:
                punkte.append([float(wert), 0.0, 0.0])
            elif code == 20 and punkte:
                punkte[-1][1] = float(wert)
            elif code == 42 and punkte:
                punkte[-1][2] = float(wert)
            elif code in (40, 41) and punkte:
                breite = max(breite, float(wert))
        geschlossen = bool(_wert(tags, 70, 0, int) & 1)
        punkte = self._ocs(tags, [tuple(p) for p in punkte])
        self._kurve(punkte, geschlossen, m, eig, breite)

    def _polyline(self, tags, unter, m, eig):
        flags = _wert(tags, 70, 0, int)
        if flags & (16 | 64):
            self._zaehle(u"POLYLINE (3D/mesh)")
            return
        breite = max(_zahl(tags, 40, 0.0), _zahl(tags, 41, 0.0))
        punkte = []
        for typ, vtags in unter:
            if typ != u"VERTEX" or _wert(vtags, 70, 0, int) & 16:
                continue
            punkte.append((_zahl(vtags, 10), _zahl(vtags, 20), _zahl(vtags, 42, 0.0)))
            breite = max(breite, _zahl(vtags, 40, 0.0), _zahl(vtags, 41, 0.0))
        self._kurve(self._ocs(tags, punkte), bool(flags & 1), m, eig, breite)

    def _spline(self, tags, m, eig):
        grad = _wert(tags, 71, 3, int)
        knoten = [float(w) for c, w in tags if c == 40]
        gewichte = [float(w) for c, w in tags if c == 41]
        kx = [float(w) for c, w in tags if c == 10]
        ky = [float(w) for c, w in tags if c == 20]
        fx = [float(w) for c, w in tags if c == 11]
        fy = [float(w) for c, w in tags if c == 21]
        if kx:
            zug = geo.spline_punkte(grad, knoten, list(zip(kx, ky)), gewichte)
        else:
            zug = list(zip(fx, fy))
        geschlossen = bool(_wert(tags, 70, 0, int) & 1)
        if geschlossen and len(zug) > 2 and zug[0] == zug[-1]:
            zug = zug[:-1]
        self._kurve([(x, y, 0.0) for x, y in zug], geschlossen, m, eig, 0.0)

    # -- Texte -----------------------------------------------------------------

    def _stil(self, tags):
        name = (_wert(tags, 7) or u"STANDARD").upper()
        return self.stile.get(name, (None, 1.0, 0.0))

    def _text(self, typ, tags, m, eig):
        inhalt = _sonderzeichen(_wert(tags, 1) or u"").strip()
        if not inhalt:
            return
        schrift, stil_breite, stil_hoehe = self._stil(tags)
        h = _zahl(tags, 40, 0.0) or stil_hoehe or 1.0
        breite = _zahl(tags, 41, 0.0) or stil_breite
        drehung = math.radians(_zahl(tags, 50, 0.0))
        hz = _wert(tags, 72, 0, int)
        # Vertikale Ausrichtung: TEXT Code 73, ATTRIB Code 74
        vt = _wert(tags, 74 if typ == u"ATTRIB" else 73, 0, int)
        p10 = (_zahl(tags, 10), _zahl(tags, 20))
        p11 = (_zahl(tags, 11, None), _zahl(tags, 21, None))
        if (hz == 0 and vt == 0) or hz in (3, 5) or p11[0] is None:
            punkt = p10
        else:
            punkt = p11
        if hz in (3, 5) and p11[0] is not None:
            # Ausgerichtet/Einpassen: Richtung aus den beiden Punkten
            drehung = math.atan2(p11[1] - p10[1], p11[0] - p10[0])
        h_ausr = {1: u"mitte", 2: u"rechts", 4: u"mitte"}.get(hz, u"links")
        if hz == 4:
            hoch = 0.0
        else:
            hoch = {0: 0.5, 1: 0.5 + UNTERLAENGE, 2: 0.0, 3: -0.5}.get(vt, 0.5)
        if self._gespiegelt(tags):
            punkt = (-punkt[0], punkt[1])
            drehung = math.pi - drehung
        vx, vy = -math.sin(drehung), math.cos(drehung)
        anker = (punkt[0] + vx * h * hoch, punkt[1] + vy * h * hoch)
        self._text_ausgeben(inhalt, anker, h, drehung, breite, schrift,
                            h_ausr, u"mitte", m, eig)

    def _mtext(self, tags, m, eig):
        roh = u"".join(w for c, w in tags if c == 3) + (_wert(tags, 1) or u"")
        inhalt, fmt = bereinige_mtext(roh)
        if not inhalt.strip():
            return
        schrift, stil_breite, stil_hoehe = self._stil(tags)
        h = _zahl(tags, 40, 0.0) or stil_hoehe or 1.0
        if u"hoehe" in fmt:
            h = fmt[u"hoehe"]
        elif u"hoehe_faktor" in fmt:
            h *= fmt[u"hoehe_faktor"]
        breite = fmt.get(u"breite", stil_breite)
        schrift = fmt.get(u"schrift", schrift)
        if _wert(tags, 11) is not None:
            drehung = math.atan2(_zahl(tags, 21), _zahl(tags, 11))
        else:
            drehung = _zahl(tags, 50, 0.0)
        anhang = _wert(tags, 71, 1, int)
        zeile, spalte = divmod(max(1, min(9, anhang)) - 1, 3)
        h_ausr = (u"links", u"mitte", u"rechts")[spalte]
        punkt = (_zahl(tags, 10), _zahl(tags, 20))
        if self._gespiegelt(tags):
            punkt = (-punkt[0], punkt[1])
            drehung = math.pi - drehung
        if u"\n" in inhalt:
            v_ausr = (u"oben", u"mitte", u"unten")[zeile]
            anker = punkt
        else:
            hoch = (-0.5, 0.0, 0.5)[zeile]
            vx, vy = -math.sin(drehung), math.cos(drehung)
            anker = (punkt[0] + vx * h * hoch, punkt[1] + vy * h * hoch)
            v_ausr = u"mitte"
        if u"rgb" in fmt:
            eig = eig[:2] + (fmt[u"rgb"],) + eig[3:]
        elif u"aci" in fmt and fmt[u"aci"] not in (0, 256):
            eig = eig[:2] + (aci_rgb(fmt[u"aci"]),) + eig[3:]
        self._text_ausgeben(inhalt, anker, h, drehung, breite, schrift,
                            h_ausr, v_ausr, m, eig)

    def _text_ausgeben(self, inhalt, anker, h, drehung, breite, schrift,
                       h_ausr, v_ausr, m, eig):
        x, y = geo.anwenden(m, anker[0], anker[1])
        ux, uy = geo.richtung(m, math.cos(drehung), math.sin(drehung))
        vx, vy = geo.richtung(m, -math.sin(drehung), math.cos(drehung))
        laenge_u = math.hypot(ux, uy) or 1.0
        laenge_v = math.hypot(vx, vy) or 1.0
        self.zeichnung.texte.append(Text(
            inhalt, x, y, h * laenge_v, math.atan2(uy, ux),
            breite * laenge_u / laenge_v, schrift, h_ausr, v_ausr, eig[2],
            eig[1]))

    # -- Blöcke ----------------------------------------------------------------

    def _einfuegen(self, typ, tags, unter, m, kontext, tiefe):
        if tiefe >= MAX_TIEFE:
            return
        name = (_wert(tags, 2) or u"").upper()
        block = self.blockdef.get(name)
        sichtbar, layer, farbe, lt, staerke, ltf = self._eigenschaften(tags, kontext)
        if not sichtbar:
            return
        neu = _Kontext(layer, farbe, lt, staerke, kontext.lt_faktor)
        if typ == u"DIMENSION":
            # Bemaßungsblock liegt schon in den Koordinaten des Behälters
            if block is not None:
                self.zeichne(block[1], m, neu, tiefe + 1)
            return
        if block is None:
            self._zaehle(u"INSERT ?")
        else:
            basis, objekte = block
            sx = _zahl(tags, 41, 1.0) or 1.0
            sy = _zahl(tags, 42, 1.0) or 1.0
            drehung = math.radians(_zahl(tags, 50, 0.0))
            ix, iy = _zahl(tags, 10), _zahl(tags, 20)
            spalten = max(1, _wert(tags, 70, 1, int) or 1)
            reihen = max(1, _wert(tags, 71, 1, int) or 1)
            dx, dy = _zahl(tags, 44, 0.0), _zahl(tags, 45, 0.0)
            ocs = (geo.skalierung(-1.0, 1.0) if self._gespiegelt(tags)
                   else geo.EINHEIT)
            for reihe in range(min(reihen, 100)):
                for spalte in range(min(spalten, 100)):
                    lokal = geo.verkette(
                        geo.drehung(drehung),
                        geo.verschiebung(spalte * dx, reihe * dy))
                    lokal = geo.verkette(lokal, geo.skalierung(sx, sy))
                    lokal = geo.verkette(lokal, geo.verschiebung(-basis[0], -basis[1]))
                    lokal = geo.verkette(geo.verschiebung(ix, iy), lokal)
                    gesamt = geo.verkette(m, geo.verkette(ocs, lokal))
                    self.zeichne(objekte, gesamt, neu, tiefe + 1)
        # Attribute liegen im Koordinatensystem des Behälters
        for atyp, atags in unter:
            if atyp == u"ATTRIB":
                self._objekt(atyp, atags, [], m, kontext, tiefe)

    # -- Schraffuren -----------------------------------------------------------

    def _schraffur(self, tags, m, eig):
        leser = _Cursor(tags)
        muster_name = (leser.suche(2) or u"SOLID").strip().upper()
        voll = leser.suche(70, int) == 1
        anzahl = leser.suche(91, int) or 0
        schleifen = []
        for _ in range(anzahl):
            art = leser.suche(92, int)
            if art is None:
                break
            if art & 2:
                schleife = self._hatch_polylinie(leser)
            else:
                schleife = self._hatch_kanten(leser)
            if schleife and len(schleife) >= 2:
                schleifen.append(geo.transformiere(self._ocs(tags, schleife), True, m))
        muster = None
        if not voll:
            muster = self._hatch_muster(leser, tags, m)
        farbe_eig = eig
        if _wert(tags, 450, 0, int) == 1:
            # Verlauf: erste Verlaufsfarbe als Vollton
            tc = _wert(tags, 421, None, int)
            if tc is not None:
                farbe_eig = eig[:2] + (truecolor(tc),) + eig[3:]
            voll, muster = True, None
        if not voll and not muster:
            # Muster ohne Musterlinien: wie Vollton behandeln
            voll = True
        self._flaeche(schleifen, muster, u"SOLID" if voll else muster_name, farbe_eig)

    def _hatch_polylinie(self, leser):
        mit_bulge = leser.lies(72, int)
        leser.lies(73, int)
        anzahl = leser.lies(93, int) or 0
        punkte = []
        for _ in range(anzahl):
            x = leser.lies(10, float)
            y = leser.lies(20, float)
            b = leser.lies(42, float) if mit_bulge else None
            if x is None or y is None:
                break
            punkte.append((x, y, b or 0.0))
        if len(punkte) > 2 and abs(punkte[0][0] - punkte[-1][0]) < 1e-12 \
                and abs(punkte[0][1] - punkte[-1][1]) < 1e-12:
            punkte.pop()
        return punkte

    def _hatch_kanten(self, leser):
        anzahl = leser.lies(93, int) or 0
        teile = []      # je Kante [(x, y, bulge)...] inkl. Endpunkt
        for _ in range(anzahl):
            art = leser.lies(72, int)
            if art == 1:
                x1, y1 = leser.lies(10, float), leser.lies(20, float)
                x2, y2 = leser.lies(11, float), leser.lies(21, float)
                teile.append([(x1, y1, 0.0), (x2, y2, 0.0)])
            elif art == 2:
                cx, cy = leser.lies(10, float), leser.lies(20, float)
                r = leser.lies(40, float)
                s, e = leser.lies(50, float), leser.lies(51, float)
                ccw = leser.lies(73, int)
                teile.append(_kantenbogen(cx, cy, r, s, e, ccw != 0))
            elif art == 3:
                cx, cy = leser.lies(10, float), leser.lies(20, float)
                ax, ay = leser.lies(11, float), leser.lies(21, float)
                verh = leser.lies(40, float)
                s, e = leser.lies(50, float), leser.lies(51, float)
                ccw = leser.lies(73, int)
                if ccw == 0:
                    s, e = 360.0 - s, 360.0 - e
                    zug = geo.ellipse_punkte((cx, cy), (ax, ay), verh,
                                             math.radians(e), math.radians(s))
                    zug.reverse()
                else:
                    zug = geo.ellipse_punkte((cx, cy), (ax, ay), verh,
                                             math.radians(s), math.radians(e))
                teile.append([(x, y, 0.0) for x, y in zug])
            elif art == 4:
                grad = leser.lies(94, int)
                leser.lies(73, int)
                leser.lies(74, int)
                nk = leser.lies(95, int) or 0
                nkp = leser.lies(96, int) or 0
                knoten = [leser.lies(40, float) for _k in range(nk)]
                kp = [(leser.lies(10, float), leser.lies(20, float)) for _k in range(nkp)]
                gewichte = []
                while leser.naechster() == 42:
                    gewichte.append(leser.lies(42, float))
                if leser.naechster() == 97:
                    nf = leser.lies(97, int) or 0
                    for _k in range(nf):
                        leser.lies(11, float)
                        leser.lies(21, float)
                    for code in (12, 22, 13, 23):
                        if leser.naechster() == code:
                            leser.lies(code, float)
                zug = geo.spline_punkte(grad, knoten, kp, gewichte)
                teile.append([(x, y, 0.0) for x, y in zug])
            else:
                break
        return _verkette_kanten(teile)

    def _hatch_muster(self, leser, tags, m):
        anzahl = leser.suche(78, int) or 0
        linien = []
        for _ in range(anzahl):
            winkel = leser.lies(53, float)
            bx, by = leser.lies(43, float), leser.lies(44, float)
            dx, dy = leser.lies(45, float), leser.lies(46, float)
            n = leser.lies(79, int) or 0
            striche = [leser.lies(49, float) for _k in range(n)]
            if winkel is None or dx is None:
                break
            linien.append(_muster_abbilden(math.radians(winkel), (bx or 0.0, by or 0.0),
                                           (dx, dy or 0.0), [s or 0.0 for s in striche],
                                           m, self._gespiegelt(tags)))
        return linien


def _muster_abbilden(winkel, basis, versatz, striche, m, gespiegelt):
    if gespiegelt:
        winkel = math.pi - winkel
        basis = (-basis[0], basis[1])
        versatz = (-versatz[0], versatz[1])
    ux, uy = geo.richtung(m, math.cos(winkel), math.sin(winkel))
    faktor = math.hypot(ux, uy) or 1.0
    bx, by = geo.anwenden(m, basis[0], basis[1])
    dx, dy = geo.richtung(m, versatz[0], versatz[1])
    return (math.atan2(uy, ux), (bx, by), (dx, dy), [s * faktor for s in striche])


def _kantenbogen(cx, cy, r, start, ende, ccw):
    """Bogenkante einer Schraffur -> [(x, y, bulge), (x, y, 0)]."""
    if not ccw:
        # Im Uhrzeigersinn speichert DXF die Winkel gespiegelt
        start, ende = 360.0 - start, 360.0 - ende
    s, e = math.radians(start), math.radians(ende)
    if ccw:
        oeffnung = (e - s) % (2.0 * math.pi) or 2.0 * math.pi
    else:
        oeffnung = -((s - e) % (2.0 * math.pi) or 2.0 * math.pi)
    p1 = (cx + r * math.cos(s), cy + r * math.sin(s))
    if abs(oeffnung) > math.pi * 1.999:
        mitte = s + oeffnung / 2.0
        b = geo.bulge_aus_winkel(oeffnung / 2.0)
        return [(p1[0], p1[1], b),
                (cx + r * math.cos(mitte), cy + r * math.sin(mitte), b),
                (p1[0], p1[1], 0.0)]
    p2 = (cx + r * math.cos(s + oeffnung), cy + r * math.sin(s + oeffnung))
    return [(p1[0], p1[1], geo.bulge_aus_winkel(oeffnung)), (p2[0], p2[1], 0.0)]


def _umkehren(teil):
    """Kante umdrehen: Bulge gehört dann zum Vorgängerpunkt, Vorzeichen
    wechselt."""
    punkte = list(reversed(teil))
    ergebnis = []
    for i, p in enumerate(punkte):
        b = -punkte[i + 1][2] if i + 1 < len(punkte) else 0.0
        ergebnis.append((p[0], p[1], b))
    return ergebnis


def _verkette_kanten(teile):
    """Kanten zu einer Schleife verbinden (Endpunkte jeweils weglassen)."""
    if not teile:
        return []
    tol = 1e-6

    def gleich(a, b):
        return abs(a[0] - b[0]) <= tol * max(1.0, abs(a[0])) and \
            abs(a[1] - b[1]) <= tol * max(1.0, abs(a[1]))

    schleife = []
    vorher = None
    for teil in teile:
        if vorher is not None and not gleich(teil[0], vorher) and gleich(teil[-1], vorher):
            teil = _umkehren(teil)
        schleife.extend(teil[:-1])
        vorher = teil[-1]
    if not schleife:
        return []
    if vorher is not None and not gleich(vorher, schleife[0]):
        schleife.append((vorher[0], vorher[1], 0.0))
    return schleife


class _Cursor(object):
    """Liest die Tags einer Schraffur der Reihe nach."""

    def __init__(self, tags):
        self.tags = tags
        self.i = 0

    def naechster(self):
        return self.tags[self.i][0] if self.i < len(self.tags) else None

    def lies(self, code, art):
        """Nächsten Tag lesen, wenn er den Code hat (sonst None)."""
        if self.i < len(self.tags) and self.tags[self.i][0] == code:
            wert = self.tags[self.i][1]
            self.i += 1
            try:
                return art(wert)
            except ValueError:
                return None
        return None

    def suche(self, code, art=None):
        """Vorspulen bis zum nächsten Tag mit dem Code."""
        while self.i < len(self.tags):
            c, w = self.tags[self.i]
            self.i += 1
            if c == code:
                if art is None:
                    return w
                try:
                    return art(w)
                except ValueError:
                    return None
        return None
