# -*- coding: utf-8 -*-
"""Revit-Zugriffe der Raum-Automatik.

Ablauf (in einer offenen Transaktion):
    1. NewRooms2 legt in jedem freien geschlossenen Bereich des Geschosses
       einen Raum an (Phase der Ansicht).
    2. Einrichtung (Familienexemplare, auch Detailbauteile im Grundriss
       des Geschosses) und Türen werden per GetRoomAtPoint den Räumen
       zugeordnet. Türen: je ein Punkt vor und hinter der Wand.
    3. Räume ohne Regel werden wieder gelöscht, die übrigen bekommen Name,
       Name ER und Raumkategorie (Schlüssel) und lückenlose Nummern.

Nur das aktive Modell - Einrichtung aus Verknüpfungen zählt nicht.
"""

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    StorageType,
    ViewPlan,
    ViewSchedule,
    XYZ,
)
from System.Collections.Generic import List  # noqa: E402

from raum_automatik import logik as lg  # noqa: E402

FUSS = 0.3048

KATEGORIE_PARAMETER = u"Raumkategorie"
NAME_ER_PARAMETER = u"Name ER"
BODEN_PARAMETER = u"Fußboden"

# Prüfpunkte so hoch über der Ebene
PRUEF_HOEHE = 1.0 / FUSS
# Einrichtung ohne Ebene zählt, wenn sie so hoch über/unter der Ebene steht
HOEHE_UNTEN = 0.5 / FUSS
HOEHE_OBEN = 2.5 / FUSS
# Türen: so weit vor/hinter der Wandfläche prüfen
TUER_ABSTAND = 0.3 / FUSS


def id_wert(element_id):
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def _kat_id(bic):
    return id_wert(ElementId(bic))


def _text(funktion):
    try:
        return funktion() or u""
    except Exception:
        return u""


def ebene_und_phase(doc, ansicht):
    """(Ebene, Phase) eines Grundrisses oder (None, None)."""
    if not isinstance(ansicht, ViewPlan) or ansicht.GenLevel is None:
        return None, None
    phase = None
    param = ansicht.get_Parameter(BuiltInParameter.VIEW_PHASE)
    if param is not None:
        phase = doc.GetElement(param.AsElementId())
    if phase is None:
        phasen = list(doc.Phases)
        phase = phasen[-1] if phasen else None
    return ansicht.GenLevel, phase


# ---------------------------------------------------------------------------
# Raumkategorie (Schlüssel)
# ---------------------------------------------------------------------------

def schluessel(doc, parameter=KATEGORIE_PARAMETER):
    """{Vergleichsname: ElementId} der Zeilen der Schlüsselbauteilliste
    zum Parameter Raumkategorie."""
    gesucht = lg.schluessel_name(parameter)
    raum_kat = _kat_id(BuiltInCategory.OST_Rooms)
    ergebnis = {}
    for liste in FilteredElementCollector(doc).OfClass(ViewSchedule):
        try:
            definition = liste.Definition
            if not definition.IsKeySchedule:
                continue
            if id_wert(definition.CategoryId) != raum_kat:
                continue
            if lg.schluessel_name(liste.KeyScheduleParameterName) != gesucht:
                continue
        except Exception:
            continue
        for zeile in FilteredElementCollector(doc).OwnedByView(liste.Id):
            name = _text(lambda: zeile.Name)
            if name:
                ergebnis[lg.schluessel_name(name)] = zeile.Id
    return ergebnis


# ---------------------------------------------------------------------------
# Einrichtung und Türen
# ---------------------------------------------------------------------------

def _name(inst):
    symbol = inst.Symbol
    if symbol is None:
        return u""
    typ = _text(lambda: symbol.Name)
    if not typ:
        p = symbol.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        typ = _text(lambda: p.AsString()) if p is not None else u""
    return u"%s %s" % (_text(lambda: symbol.FamilyName), typ)


class _Ebenenpruefung(object):
    """Steht ein Exemplar auf dem Geschoss? Ansichten werden gemerkt."""

    def __init__(self, doc, ebene):
        self.doc = doc
        self.ebene = id_wert(ebene.Id)
        self.hoehe = ebene.ProjectElevation
        self._ansichten = {}

    def ansicht(self, inst):
        """Besitzer-Ansicht (Detailbauteile) oder None."""
        besitzer = inst.OwnerViewId
        if besitzer is None:
            return None
        kennung = id_wert(besitzer)
        if kennung <= 0:
            return None
        if kennung not in self._ansichten:
            self._ansichten[kennung] = self.doc.GetElement(besitzer)
        return self._ansichten[kennung]

    def schnell(self, inst):
        """Detailbauteile: True/False nach der Ebene ihres Grundrisses.
        Modellelemente: None = die Höhe entscheidet, nicht die Ebene, an
        der sie hängen (Möbel auf FBOK, Grundriss auf RDOK ...)."""
        ansicht = self.ansicht(inst)
        if ansicht is not None:
            gen = getattr(ansicht, "GenLevel", None)
            return gen is not None and id_wert(gen.Id) == self.ebene
        return None

    def hoehe_passt(self, z):
        return self.hoehe - HOEHE_UNTEN <= z <= self.hoehe + HOEHE_OBEN


def _box(pruefung, inst):
    try:
        return inst.get_BoundingBox(pruefung.ansicht(inst))
    except Exception:
        return None


class Gegenstand(object):
    def __init__(self, punkt, merkmale, beschreibung):
        self.punkt = punkt
        self.merkmale = merkmale
        self.beschreibung = beschreibung


def einrichtung(doc, ebene):
    """Gegenstände des Geschosses mit mindestens einem Merkmal."""
    tueren = _kat_id(BuiltInCategory.OST_Doors)
    fenster = _kat_id(BuiltInCategory.OST_Windows)
    sanitaer = _kat_id(BuiltInCategory.OST_PlumbingFixtures)
    pruefung = _Ebenenpruefung(doc, ebene)
    # Merkmale je Typ - Name und Regeln nur einmal je Typ auswerten
    je_typ = {}
    ergebnis = []
    for inst in FilteredElementCollector(doc).OfClass(FamilyInstance):
        kat = inst.Category
        if kat is None:
            continue
        kat_id = id_wert(kat.Id)
        if kat_id in (tueren, fenster):
            continue
        auf_ebene = pruefung.schnell(inst)
        if auf_ebene is False:
            continue
        typ_id = id_wert(inst.GetTypeId())
        if typ_id not in je_typ:
            name = _name(inst)
            je_typ[typ_id] = (name.strip(),
                              lg.merkmale(name, kat_id == sanitaer))
        name, merkmale = je_typ[typ_id]
        if not merkmale:
            continue
        box = None
        # unter CPython ist Location nicht als LocationPoint erkennbar
        punkt = getattr(inst.Location, "Point", None)
        if punkt is None:
            box = _box(pruefung, inst)
            if box is None:
                continue
            punkt = box.Min.Add(box.Max).Multiply(0.5)
        if auf_ebene is None and not pruefung.hoehe_passt(punkt.Z):
            continue
        if lg.BETT in merkmale and lg.DOPPELBETT not in merkmale:
            # Doppelbett an der Breite erkennen - Geometrie nur für Betten
            box = box or _box(pruefung, inst)
            if box is not None:
                breite = min(box.Max.X - box.Min.X,
                             box.Max.Y - box.Min.Y) * FUSS
                merkmale = lg.merkmale(name, kat_id == sanitaer, breite)
        ergebnis.append(Gegenstand(punkt, merkmale, name))
    return ergebnis


def tuer_punkte(doc, ebene):
    """[(Punkt vorne, Punkt hinten)] der Türen des Geschosses."""
    pruefung = _Ebenenpruefung(doc, ebene)
    ergebnis = []
    for inst in (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_Doors)
                 .WhereElementIsNotElementType()):
        punkt = getattr(inst.Location, "Point", None)
        if punkt is None:
            continue
        auf_ebene = pruefung.schnell(inst)
        if auf_ebene is False or (auf_ebene is None
                                  and not pruefung.hoehe_passt(punkt.Z)):
            continue
        richtung = getattr(inst, "FacingOrientation", None)
        if richtung is None or richtung.GetLength() < 1e-6:
            continue
        dicke = 0.0
        try:
            wand = inst.Host
            dicke = wand.Width if wand is not None else 0.0
        except Exception:
            dicke = 0.0
        weg = richtung.Normalize().Multiply(dicke / 2.0 + TUER_ABSTAND)
        ergebnis.append((punkt.Add(weg), punkt.Subtract(weg)))
    return ergebnis


# ---------------------------------------------------------------------------
# Räume
# ---------------------------------------------------------------------------

def vorhandene_raeume(doc, ebene, phase):
    """Platzierte Räume des Geschosses in der Phase."""
    ergebnis = []
    for raum in (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_Rooms)
                 .WhereElementIsNotElementType()):
        if raum.Location is None:
            continue
        if raum.Area <= 0 or id_wert(raum.LevelId) != id_wert(ebene.Id):
            continue
        if phase is not None:
            p = raum.get_Parameter(BuiltInParameter.ROOM_PHASE)
            if p is not None and id_wert(p.AsElementId()) != id_wert(phase.Id):
                continue
        ergebnis.append(raum)
    return ergebnis


def _raumname(raum):
    p = raum.get_Parameter(BuiltInParameter.ROOM_NAME)
    return _text(lambda: p.AsString()) if p is not None else u""


class Ergebnis(object):
    def __init__(self):
        self.angelegt = []       # (Raum, Bereich)
        self.verworfen = 0
        self.bereiche = []       # alle neuen Bereiche, für das Protokoll
        self.ohne_schluessel = set()
        self.fehlende_parameter = set()
        self.fehler = []


def lege_an(doc, ebene, phase, gegenstaende, tueren, schluessel_ids):
    """Räume anlegen, bewerten, Regel-lose löschen, Parameter setzen.

    Läuft in einer offenen Transaktion. Rückgabe: Ergebnis.
    """
    ergebnis = Ergebnis()
    vorhanden = vorhandene_raeume(doc, ebene, phase)
    neue_ids = list(doc.Create.NewRooms2(ebene, phase))
    if not neue_ids:
        return ergebnis
    doc.Regenerate()

    neu = dict((id_wert(i), doc.GetElement(i)) for i in neue_ids)
    bereiche = {}
    for kennung, raum in neu.items():
        bereiche[kennung] = lg.Bereich(kennung, True, raum.Area * FUSS * FUSS)
    for raum in vorhanden:
        kennung = id_wert(raum.Id)
        bereiche[kennung] = lg.Bereich(kennung, False, raum.Area * FUSS * FUSS,
                                       _raumname(raum))

    hoehe = ebene.ProjectElevation + PRUEF_HOEHE

    def raum_bei(punkt):
        raum = doc.GetRoomAtPoint(XYZ(punkt.X, punkt.Y, hoehe), phase)
        if raum is None:
            return None
        kennung = id_wert(raum.Id)
        return kennung if kennung in bereiche else None

    for g in gegenstaende:
        kennung = raum_bei(g.punkt)
        if kennung is not None and kennung in neu:
            bereiche[kennung].merkmale |= g.merkmale
            bereiche[kennung].inhalt.append(g.beschreibung)

    verbindungen = []
    for vorne, hinten in tueren:
        a, b = raum_bei(vorne), raum_bei(hinten)
        if a == b:
            continue
        for k in (a, b):
            if k is not None:
                bereiche[k].tueren += 1
        if a is not None and b is not None:
            verbindungen.append((a, b))

    alle = list(bereiche.values())
    lg.bewerte(alle)
    lg.nummeriere(alle, verbindungen)

    nummern = []
    behalten = []
    weg = List[ElementId]()
    for i in neue_ids:
        raum = neu[id_wert(i)]
        p = raum.get_Parameter(BuiltInParameter.ROOM_NUMBER)
        nummern.append(_text(lambda: p.AsString()))
        bereich = bereiche[id_wert(i)]
        if bereich.art is None:
            weg.Add(i)
            ergebnis.verworfen += 1
        else:
            behalten.append((raum, bereich))

    if weg.Count:
        doc.Delete(weg)

    # Revit zählt die gelöschten Räume mit - Nummern lückenlos nachziehen
    if nummern and all(n.isdigit() for n in nummern):
        frei = sorted(nummern, key=int)
        for (raum, _), nummer in zip(behalten, frei):
            p = raum.get_Parameter(BuiltInParameter.ROOM_NUMBER)
            if p is not None and not p.IsReadOnly:
                p.Set(nummer)

    ergebnis.bereiche = [bereiche[id_wert(i)] for i in neue_ids]
    for raum, bereich in behalten:
        try:
            _fuelle(raum, bereich, schluessel_ids, ergebnis)
        except Exception as ausnahme:
            ergebnis.fehler.append(u"%s: %s" % (bereich.raumname(), ausnahme))
        ergebnis.angelegt.append((raum, bereich))
    return ergebnis


def _fuelle(raum, bereich, schluessel_ids, ergebnis):
    name = bereich.raumname()
    raum.get_Parameter(BuiltInParameter.ROOM_NAME).Set(name)

    p = raum.LookupParameter(NAME_ER_PARAMETER)
    if p is None or p.IsReadOnly:
        ergebnis.fehlende_parameter.add(NAME_ER_PARAMETER)
    else:
        p.Set(name)

    kategorie = bereich.kategorie()
    p = raum.LookupParameter(KATEGORIE_PARAMETER)
    if p is None or p.IsReadOnly:
        ergebnis.fehlende_parameter.add(KATEGORIE_PARAMETER)
    elif p.StorageType == StorageType.ElementId:
        schluessel_id = schluessel_ids.get(lg.schluessel_name(kategorie))
        if schluessel_id is None:
            ergebnis.ohne_schluessel.add(kategorie)
        else:
            p.Set(schluessel_id)
    elif p.StorageType == StorageType.String:
        p.Set(kategorie)

    # eigener Parameter "Fußboden", sonst der eingebaute Bodenbelag
    p = raum.LookupParameter(BODEN_PARAMETER)
    if p is None:
        p = raum.get_Parameter(BuiltInParameter.ROOM_FINISH_FLOOR)
    if p is None or p.IsReadOnly or p.StorageType != StorageType.String:
        ergebnis.fehlende_parameter.add(BODEN_PARAMETER)
    else:
        p.Set(bereich.boden())
