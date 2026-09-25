# -*- coding: utf-8 -*-
"""Kollisionsberichte aus Navisworks lesen - XML und CSV. Ohne Revit.

XML (Clash Detective > Berichte > Format "XML") - Aufbau, soweit genutzt:

    <exchange units="m">
      <batchtest>
        <clashtests>
          <clashtest name="TGA vs TWP">
            <clashresults>
              <clashresult name="Kollision1" guid="..." status="new"
                           distance="-0.052">
                <gridlocation>C-4 : Ebene 01</gridlocation>
                <clashpoint><pos3f x="12.3" y="4.5" z="6.7"/></clashpoint>
                <clashobjects>
                  <clashobject>
                    <objectattribute><name>Element ID</name>
                                     <value>398254</value></objectattribute>
                    <pathlink><node>Datei</node><node>TGA.nwc</node>...
                    <smarttags><smarttag><name>Item Name</name>...
              <clashgroup name="Gruppe 1" ...>
                <clashresults><clashresult .../></clashresults>

CSV: Spalten werden an ihren Überschriften erkannt (Deutsch, Englisch,
Spanisch), Trennzeichen ; , oder Tab. Die Spalten der zwei Elemente tragen
eine 1 bzw. 2 ("Item 1 Element ID", "Elemento 2 ID", "Datei 1" ...).

Längen und Punkte werden in Meter umgerechnet. Die Koordinaten sind die von
Navisworks - meist die gemeinsamen Koordinaten des Revit-Exports.
"""

import io
import os
import re
import xml.etree.ElementTree as ET

from mlg_sprache import t

# Status in der Schreibweise dieses Werkzeugs
NEU = u"neu"
AKTIV = u"aktiv"
GEPRUEFT = u"geprueft"
GENEHMIGT = u"genehmigt"
BEHOBEN = u"behoben"
STATUS = (NEU, AKTIV, GEPRUEFT, GENEHMIGT, BEHOBEN)

_STATUS_NAMEN = {
    NEU: (u"new", u"neu", u"nuevo", u"nueva", u"open", u"offen", u"abierto",
          u"reopened"),
    AKTIV: (u"active", u"aktiv", u"activo", u"activa"),
    GEPRUEFT: (u"reviewed", u"geprüft", u"geprueft", u"revisado",
               u"revisada"),
    GENEHMIGT: (u"approved", u"genehmigt", u"aprobado", u"aprobada"),
    BEHOBEN: (u"resolved", u"behoben", u"gelöst", u"geloest", u"resuelto",
              u"resuelta", u"closed", u"geschlossen", u"cerrado"),
}

_EINHEITEN = {
    u"m": 1.0, u"meter": 1.0, u"meters": 1.0, u"metres": 1.0,
    u"cm": 0.01, u"mm": 0.001, u"km": 1000.0,
    u"ft": 0.3048, u"feet": 0.3048, u"foot": 0.3048,
    u"in": 0.0254, u"inch": 0.0254, u"inches": 0.0254,
    u"yd": 0.9144,
}

_DATEIENDUNGEN = (u".nwc", u".nwd", u".nwf", u".rvt", u".ifc", u".dwg",
                  u".dgn", u".fbx")


class Fehler(Exception):
    """Die Datei ist kein lesbarer Kollisionsbericht."""


class Objekt(object):
    """Ein an der Kollision beteiligtes Element."""

    def __init__(self, element_id=None, datei=u"", name=u"", typ=u"",
                 ebene=u""):
        self.element_id = element_id
        self.datei = datei or u""
        self.name = name or u""
        self.typ = typ or u""
        self.ebene = ebene or u""

    @property
    def beschreibung(self):
        teile = [teil for teil in (self.name, self.typ) if teil]
        if len(teile) == 2 and teile[0] == teile[1]:
            teile = teile[:1]
        return u" · ".join(teile) or t(u"(ohne Namen)", u"(no name)",
                                       u"(sin nombre)")


class Clash(object):
    """Eine Kollision. Längen in Metern, Punkt als (x, y, z) in Metern."""

    def __init__(self):
        self.schluessel = u""
        self.name = u""
        self.guid = u""
        self.test = u""
        self.gruppe = u""
        self.status_navis = NEU
        self.abstand = None
        self.punkt = None
        self.raster = u""
        self.beschreibung = u""
        self.datum = u""
        self.zugewiesen = u""
        self.kommentare_navis = []
        self.objekte = []

    @property
    def ebene(self):
        """Der Teil der Rasterposition nach dem Doppelpunkt: 'C-4 : EG'."""
        if u":" in self.raster:
            return self.raster.split(u":", 1)[1].strip()
        for objekt in self.objekte:
            if objekt.ebene:
                return objekt.ebene
        return u""

    def __repr__(self):
        return "<Clash %r>" % self.schluessel


class Bericht(object):
    def __init__(self, pfad, clashes, format_name):
        self.pfad = pfad
        self.clashes = clashes
        self.format = format_name

    @property
    def name(self):
        return os.path.basename(self.pfad)

    @property
    def tests(self):
        gesehen = []
        for clash in self.clashes:
            if clash.test not in gesehen:
                gesehen.append(clash.test)
        return gesehen


# ---------------------------------------------------------------------------
# Hilfen
# ---------------------------------------------------------------------------

def status_aus_text(text):
    """Navisworks-Status (beliebige Sprache) -> NEU, AKTIV ... ; sonst None."""
    wert = (text or u"").strip().lower()
    for status, namen in _STATUS_NAMEN.items():
        if wert in namen:
            return status
    return None


def einheit_faktor(text, vorgabe=1.0):
    return _EINHEITEN.get((text or u"").strip().lower(), vorgabe)


_ZAHL = re.compile(r"-?\d+(?:[.,]\d+)?(?:[eE][-+]?\d+)?")


def zahl(text):
    """Erste Zahl im Text, Dezimalkomma erlaubt. None, wenn keine."""
    treffer = _ZAHL.search(text or u"")
    if not treffer:
        return None
    try:
        return float(treffer.group(0).replace(u",", u"."))
    except ValueError:
        return None


def ganzzahl(text):
    treffer = re.search(r"\d+", text or u"")
    if not treffer:
        return None
    try:
        return int(treffer.group(0))
    except ValueError:
        return None


def punkt_aus_text(text):
    """'x:12.3 m, y:4.5 m, z:6.7 m' oder '12,3; 4,5; 6,7' -> (x, y, z)."""
    text = text or u""
    benannt = dict((achse.lower(), wert) for achse, wert in re.findall(
        r"([xyzXYZ])\s*[:=]\s*(-?\d+(?:[.,]\d+)?)", text))
    if len(benannt) == 3:
        return tuple(float(benannt[a].replace(u",", u".")) for a in u"xyz")
    if u";" in text:
        teile = text.split(u";")
    elif text.count(u",") == 2:
        teile = text.split(u",")
    else:
        teile = text.split()
    werte = [zahl(teil) for teil in teile]
    werte = [w for w in werte if w is not None]
    if len(werte) == 3:
        return tuple(werte)
    return None


def _normiert(text):
    """Klein, ohne Leer- und Sonderzeichen: 'Element ID' -> 'elementid'."""
    text = (text or u"").strip().lower()
    for alt, neu in ((u"ä", u"a"), (u"ö", u"o"), (u"ü", u"u"), (u"ß", u"ss"),
                     (u"á", u"a"), (u"é", u"e"), (u"í", u"i"), (u"ó", u"o"),
                     (u"ú", u"u"), (u"ñ", u"n")):
        text = text.replace(alt, neu)
    return re.sub(r"[^a-z0-9]", u"", text)


_ID_NAMEN = (u"elementid", u"iddeelemento", u"idelemento", u"elementoid",
             u"elementids", u"revitid", u"revitelementid", u"id")


def ist_id_name(name):
    return _normiert(name) in _ID_NAMEN


def _dateiname(text):
    """Dateiname aus einem Pfadknoten, wenn er eine Modelldatei ist."""
    text = (text or u"").strip()
    klein = text.lower()
    for endung in _DATEIENDUNGEN:
        if klein.endswith(endung):
            return os.path.basename(text.replace(u"\\", u"/"))
    return u""


def _eindeutig(clashes):
    """Schlüssel: GUID, sonst Test|Name; Dubletten bekommen #2, #3 ..."""
    vergeben = set()
    for clash in clashes:
        basis = clash.guid or u"%s|%s" % (clash.test, clash.name)
        schluessel = basis
        zaehler = 1
        while schluessel in vergeben:
            zaehler += 1
            schluessel = u"%s#%d" % (basis, zaehler)
        vergeben.add(schluessel)
        clash.schluessel = schluessel
    return clashes


# ---------------------------------------------------------------------------
# XML
# ---------------------------------------------------------------------------

def _lokal(tag):
    """Tag ohne Namensraum; Kommentare haben keinen Text-Tag."""
    try:
        return tag.split(u"}")[-1]
    except AttributeError:
        return u""


def _kinder(knoten, name):
    return [kind for kind in list(knoten) if _lokal(kind.tag) == name]


def _kind(knoten, name):
    for kind in list(knoten):
        if _lokal(kind.tag) == name:
            return kind
    return None


def _text(knoten, name):
    kind = _kind(knoten, name) if knoten is not None else None
    if kind is None or kind.text is None:
        return u""
    return kind.text.strip()


def _datum(knoten):
    if knoten is None:
        return u""
    datum = _kind(knoten, u"date")
    if datum is None:
        return u""
    try:
        return u"%04d-%02d-%02d %02d:%02d" % tuple(
            int(datum.get(teil) or 0)
            for teil in (u"year", u"month", u"day", u"hour", u"minute"))
    except ValueError:
        return u""


def _eigenschaften(clashobjekt):
    """(Name, Wert) aus objectattribute und smarttags."""
    paare = []
    for attribut in _kinder(clashobjekt, u"objectattribute"):
        paare.append((_text(attribut, u"name"), _text(attribut, u"value")))
    tags = _kind(clashobjekt, u"smarttags")
    if tags is not None:
        for tag in _kinder(tags, u"smarttag"):
            paare.append((_text(tag, u"name"), _text(tag, u"value")))
    return paare


def _wert(paare, namen):
    gesucht = [_normiert(name) for name in namen]
    for name, wert in paare:
        if _normiert(name) in gesucht and wert:
            return wert
    return u""


def _objekt_aus_xml(knoten):
    paare = _eigenschaften(knoten)
    element_id = None
    for name, wert in paare:
        if ist_id_name(name):
            element_id = ganzzahl(wert)
            if element_id is not None:
                break

    datei = _wert(paare, (u"Item Source File", u"Source File",
                          u"Quelldatei", u"Element Quelldatei",
                          u"Archivo de origen", u"Elemento Archivo de origen"))
    datei = _dateiname(datei) or datei
    pfad = _kind(knoten, u"pathlink")
    if pfad is not None and not datei:
        for teil in _kinder(pfad, u"node"):
            datei = _dateiname(teil.text)
            if datei:
                break

    return Objekt(
        element_id=element_id,
        datei=datei,
        name=_wert(paare, (u"Item Name", u"Element Name", u"Name",
                           u"Nombre", u"Elemento Nombre")),
        typ=_wert(paare, (u"Item Type", u"Element Typ", u"Type", u"Typ",
                          u"Tipo", u"Elemento Tipo", u"Category",
                          u"Kategorie", u"Categoría")),
        ebene=_text(knoten, u"layer"))


def _clash_aus_xml(knoten, test, gruppe):
    clash = Clash()
    clash.name = knoten.get(u"name") or u""
    clash.guid = knoten.get(u"guid") or u""
    clash.test = test
    clash.gruppe = gruppe
    clash.status_navis = (status_aus_text(knoten.get(u"status"))
                          or status_aus_text(_text(knoten, u"resultstatus"))
                          or NEU)
    abstand = knoten.get(u"distance")
    clash.abstand = zahl(abstand) if abstand else None
    clash.raster = _text(knoten, u"gridlocation")
    clash.beschreibung = _text(knoten, u"description")
    clash.zugewiesen = _text(knoten, u"assignedto")
    clash.datum = _datum(_kind(knoten, u"createddate"))

    punkt = _kind(knoten, u"clashpoint")
    pos = _kind(punkt, u"pos3f") if punkt is not None else None
    if pos is not None:
        try:
            clash.punkt = tuple(float(pos.get(a)) for a in (u"x", u"y", u"z"))
        except (TypeError, ValueError):
            clash.punkt = None

    kommentare = _kind(knoten, u"comments")
    if kommentare is not None:
        for kommentar in _kinder(kommentare, u"comment"):
            text = _text(kommentar, u"body")
            if text:
                benutzer = _text(kommentar, u"user")
                clash.kommentare_navis.append(
                    u"%s: %s" % (benutzer, text) if benutzer else text)

    objekte = _kind(knoten, u"clashobjects")
    if objekte is not None:
        clash.objekte = [_objekt_aus_xml(objekt) for objekt
                         in _kinder(objekte, u"clashobject")]
    return clash


def _skaliere(clash, faktor):
    if faktor == 1.0:
        return
    if clash.abstand is not None:
        clash.abstand *= faktor
    if clash.punkt is not None:
        clash.punkt = tuple(wert * faktor for wert in clash.punkt)


def lies_xml(pfad):
    try:
        wurzel = ET.parse(pfad).getroot()
    except Exception as fehler:
        raise Fehler(t(u"Die XML-Datei lässt sich nicht lesen:\n%s",
                       u"The XML file cannot be read:\n%s",
                       u"No se puede leer el archivo XML:\n%s") % fehler)

    faktor = einheit_faktor(wurzel.get(u"units"))
    clashes = []
    tests = [k for k in wurzel.iter() if _lokal(k.tag) == u"clashtest"]
    for test in tests:
        testname = test.get(u"name") or u""
        stapel = _kind(test, u"clashresults")
        if stapel is None:
            continue
        testfaktor = einheit_faktor(test.get(u"units"), faktor)
        gefunden = []
        for eintrag in list(stapel):
            art = _lokal(eintrag.tag)
            if art == u"clashresult":
                gefunden.append(_clash_aus_xml(eintrag, testname, u""))
            elif art == u"clashgroup":
                gruppe = _clash_aus_xml(eintrag, testname, u"")
                innen = _kind(eintrag, u"clashresults")
                for kind in (_kinder(innen, u"clashresult")
                             if innen is not None else []):
                    clash = _clash_aus_xml(kind, testname, gruppe.name)
                    if clash.punkt is None:
                        clash.punkt = gruppe.punkt
                    gefunden.append(clash)
        for clash in gefunden:
            _skaliere(clash, testfaktor)
        clashes.extend(gefunden)

    if not tests:
        raise Fehler(t(u"Die XML-Datei enthält keine Kollisionstests.\n"
                       u"Erwartet wird ein Bericht aus dem Clash Detective "
                       u"im Format \"XML\".",
                       u"The XML file contains no clash tests.\n"
                       u"Expected a Clash Detective report in \"XML\" "
                       u"format.",
                       u"El archivo XML no contiene pruebas de conflictos.\n"
                       u"Se espera un informe de Clash Detective en formato "
                       u"\"XML\"."))
    return Bericht(pfad, _eindeutig(clashes), u"XML")


# ---------------------------------------------------------------------------
# CSV
# ---------------------------------------------------------------------------

def _lies_text(pfad):
    with io.open(pfad, "rb") as datei:
        roh = datei.read()
    if roh[:2] in (b"\xff\xfe", b"\xfe\xff"):
        return roh.decode("utf-16")
    try:
        return roh.decode("utf-8-sig")
    except UnicodeError:
        return roh.decode("cp1252", "replace")


def _trennzeichen(kopfzeile):
    zaehler = {}
    in_zitat = False
    for zeichen in kopfzeile:
        if zeichen == u'"':
            in_zitat = not in_zitat
        elif not in_zitat and zeichen in (u";", u",", u"\t"):
            zaehler[zeichen] = zaehler.get(zeichen, 0) + 1
    if not zaehler:
        return u","
    return max(zaehler, key=lambda z: zaehler[z])


def csv_zeilen(text, trenner):
    """Einfacher CSV-Leser mit Anführungszeichen und Zeilenumbrüchen in
    Feldern (das csv-Modul kann unter IronPython 2 kein Unicode)."""
    zeilen = []
    feld = []
    zeile = []
    in_zitat = False
    i = 0
    laenge = len(text)
    while i < laenge:
        zeichen = text[i]
        if in_zitat:
            if zeichen == u'"':
                if i + 1 < laenge and text[i + 1] == u'"':
                    feld.append(u'"')
                    i += 1
                else:
                    in_zitat = False
            else:
                feld.append(zeichen)
        elif zeichen == u'"':
            in_zitat = True
        elif zeichen == trenner:
            zeile.append(u"".join(feld))
            feld = []
        elif zeichen in (u"\r", u"\n"):
            if zeichen == u"\r" and i + 1 < laenge and text[i + 1] == u"\n":
                i += 1
            zeile.append(u"".join(feld))
            feld = []
            if any(wert.strip() for wert in zeile):
                zeilen.append(zeile)
            zeile = []
        else:
            feld.append(zeichen)
        i += 1
    if feld or zeile:
        zeile.append(u"".join(feld))
        if any(wert.strip() for wert in zeile):
            zeilen.append(zeile)
    return zeilen


# Felder der Kollision: normierte Überschriften
_CLASH_SPALTEN = {
    u"name": (u"name", u"clashname", u"clash", u"kollision",
              u"kollisionsname", u"nombre", u"conflicto", u"interferencia",
              u"nombredeconflicto", u"nombredelconflicto"),
    u"guid": (u"guid", u"clashguid", u"clashid", u"id", u"kollisionsid",
              u"iddeconflicto"),
    u"test": (u"test", u"testname", u"clashtest", u"prueba",
              u"nombredeprueba", u"kollisionstest"),
    u"status": (u"status", u"clashstatus", u"estado", u"resultstatus"),
    u"abstand": (u"distance", u"abstand", u"distancia", u"penetration"),
    u"punkt": (u"clashpoint", u"punkt", u"point", u"kollisionspunkt",
               u"puntodeconflicto", u"punto", u"position", u"location",
               u"clashposition"),
    u"x": (u"x", u"clashpointx", u"pointx", u"punktx", u"puntox", u"posx"),
    u"y": (u"y", u"clashpointy", u"pointy", u"punkty", u"puntoy", u"posy"),
    u"z": (u"z", u"clashpointz", u"pointz", u"punktz", u"puntoz", u"posz"),
    u"raster": (u"gridlocation", u"grid", u"raster", u"rasterposition",
                u"rasterschnittpunkt", u"ubicaciondelarejilla",
                u"ubicacionderejilla", u"rejilla"),
    u"gruppe": (u"group", u"gruppe", u"grupo", u"clashgroup"),
    u"beschreibung": (u"description", u"beschreibung", u"descripcion"),
    u"kommentar": (u"comment", u"comments", u"kommentar", u"kommentare",
                   u"comentario", u"comentarios"),
    u"zugewiesen": (u"assignedto", u"zugewiesenan", u"asignadoa",
                    u"assigned"),
    u"datum": (u"date", u"datum", u"fecha", u"datefound", u"found",
               u"createddate", u"gefunden", u"encontrado"),
}

_OBJEKT_WOERTER = (u"item", u"element", u"elemento", u"objekt", u"object",
                   u"objeto", u"selection", u"auswahl", u"seleccion")


def _objekt_feld(rest):
    if rest in _ID_NAMEN or rest in (u"revitid",):
        return u"id"
    if not rest or rest in (u"name", u"nombre", u"itemname"):
        return u"name"
    for wort in (u"file", u"datei", u"archivo", u"source", u"quelle",
                 u"model", u"modell", u"modelo"):
        if wort in rest:
            return u"datei"
    for wort in (u"layer", u"level", u"ebene", u"nivel", u"capa"):
        if wort in rest:
            return u"ebene"
    for wort in (u"type", u"typ", u"tipo", u"category", u"kategorie",
                 u"categoria"):
        if wort in rest:
            return u"typ"
    if u"name" in rest or u"nombre" in rest:
        return u"name"
    if rest.endswith(u"id"):
        return u"id"
    return None


def ordne_spalten(kopf):
    """Überschriften -> {feld: index} und {(nr, feld): index}.

    Zusätzlich die Einheit, falls eine Überschrift sie nennt: 'Abstand (mm)'.
    """
    clash_spalten = {}
    objekt_spalten = {}
    einheit = None
    for index, ueberschrift in enumerate(kopf):
        roh = (ueberschrift or u"").strip()
        klammer = re.search(r"\(\s*(mm|cm|m|ft|in)\s*\)", roh.lower())
        if klammer and einheit is None:
            einheit = klammer.group(1)
            roh = roh[:klammer.start()] + roh[klammer.end():]
        norm = _normiert(roh)

        nummer = None
        rest = norm
        for wort in _OBJEKT_WOERTER:
            treffer = re.search(wort + r"0?([12])", norm)
            if treffer:
                nummer = int(treffer.group(1))
                rest = norm[:treffer.start()] + norm[treffer.end():]
                break
        if nummer is None:
            treffer = re.match(r"^(.*[a-z])([12])$", norm) or \
                re.match(r"^([12])([a-z].*)$", norm)
            if treffer:
                if treffer.group(2).isdigit():
                    nummer, rest = int(treffer.group(2)), treffer.group(1)
                else:
                    nummer, rest = int(treffer.group(1)), treffer.group(2)

        if nummer is not None:
            feld = _objekt_feld(rest)
            if feld and (nummer, feld) not in objekt_spalten:
                objekt_spalten[(nummer, feld)] = index
            continue

        for feld, namen in _CLASH_SPALTEN.items():
            if norm in namen and feld not in clash_spalten:
                clash_spalten[feld] = index
                break
    return clash_spalten, objekt_spalten, einheit


def lies_csv(pfad):
    text = _lies_text(pfad)
    kopfzeile = text.lstrip(u"﻿").split(u"\n", 1)[0]
    zeilen = csv_zeilen(text.lstrip(u"﻿"), _trennzeichen(kopfzeile))
    if len(zeilen) < 2:
        raise Fehler(t(u"Die CSV-Datei enthält keine Kollisionen.",
                       u"The CSV file contains no clashes.",
                       u"El archivo CSV no contiene conflictos."))

    clash_spalten, objekt_spalten, einheit = ordne_spalten(zeilen[0])
    if u"name" not in clash_spalten and u"guid" not in clash_spalten:
        raise Fehler(t(u"In der CSV-Datei fehlt eine Spalte für den Namen "
                       u"der Kollision.\nGefundene Spalten: %s",
                       u"The CSV file has no column for the clash name.\n"
                       u"Columns found: %s",
                       u"Al archivo CSV le falta una columna con el nombre "
                       u"del conflicto.\nColumnas encontradas: %s")
                     % u", ".join(zeilen[0]))
    faktor = einheit_faktor(einheit)

    def wert(zeile, index):
        if index is None or index >= len(zeile):
            return u""
        return zeile[index].strip()

    clashes = []
    for zeile in zeilen[1:]:
        feld = dict((name, wert(zeile, index))
                    for name, index in clash_spalten.items())
        clash = Clash()
        clash.name = feld.get(u"name", u"")
        clash.guid = feld.get(u"guid", u"")
        if not clash.name and not clash.guid:
            continue
        clash.test = feld.get(u"test", u"")
        clash.gruppe = feld.get(u"gruppe", u"")
        clash.status_navis = status_aus_text(feld.get(u"status")) or NEU
        clash.abstand = zahl(feld.get(u"abstand"))
        clash.raster = feld.get(u"raster", u"")
        clash.beschreibung = feld.get(u"beschreibung", u"")
        clash.zugewiesen = feld.get(u"zugewiesen", u"")
        clash.datum = feld.get(u"datum", u"")
        if feld.get(u"kommentar"):
            clash.kommentare_navis = [feld[u"kommentar"]]
        if all(feld.get(achse) for achse in (u"x", u"y", u"z")):
            werte = [zahl(feld[achse]) for achse in (u"x", u"y", u"z")]
            clash.punkt = tuple(werte) if None not in werte else None
        elif feld.get(u"punkt"):
            clash.punkt = punkt_aus_text(feld[u"punkt"])

        for nummer in (1, 2):
            spalten = dict((name, index) for (nr, name), index
                           in objekt_spalten.items() if nr == nummer)
            if not spalten:
                continue
            objekt = Objekt(
                element_id=ganzzahl(wert(zeile, spalten.get(u"id"))),
                datei=_dateiname(wert(zeile, spalten.get(u"datei")))
                or wert(zeile, spalten.get(u"datei")),
                name=wert(zeile, spalten.get(u"name")),
                typ=wert(zeile, spalten.get(u"typ")),
                ebene=wert(zeile, spalten.get(u"ebene")))
            clash.objekte.append(objekt)

        _skaliere(clash, faktor)
        clashes.append(clash)

    return Bericht(pfad, _eindeutig(clashes), u"CSV")


def lies(pfad):
    """Bericht nach Dateiendung lesen. Wirft Fehler mit Klartext."""
    endung = os.path.splitext(pfad)[1].lower()
    if endung == u".xml":
        return lies_xml(pfad)
    if endung in (u".csv", u".txt"):
        return lies_csv(pfad)
    if endung in (u".nwd", u".nwf", u".nwc"):
        raise Fehler(t(
            u"NWD-, NWF- und NWC-Dateien lassen sich ohne Navisworks nicht "
            u"lesen.\n\nIn Navisworks: Clash Detective > Berichte > Format "
            u"\"XML\" und bei den Inhalten die Element-ID einschliessen. Die "
            u"XML-Datei hier laden.",
            u"NWD, NWF and NWC files cannot be read without Navisworks.\n\n"
            u"In Navisworks: Clash Detective > Report > format \"XML\" and "
            u"include the element id in the contents. Load the XML file "
            u"here.",
            u"Los archivos NWD, NWF y NWC no se pueden leer sin "
            u"Navisworks.\n\nEn Navisworks: Clash Detective > Informe > "
            u"formato \"XML\" e incluya el ID de elemento en el contenido. "
            u"Cargue aquí el archivo XML."))
    raise Fehler(t(u"Unbekanntes Dateiformat: %s",
                   u"Unknown file format: %s",
                   u"Formato de archivo desconocido: %s") % endung)
