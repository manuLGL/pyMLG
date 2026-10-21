# -*- coding: utf-8 -*-
"""Revit-Zugriffe von ColumnNumbering.

Ein Lauf ist eine Transaktionsgruppe:

    je Klick (bzw. einmal für den ganzen Pfad) eine Transaktion:
        Nummer in den Parameter schreiben
        Vorschau in der aktiven Ansicht: Textnotiz mit der Nummer, Stütze grün
    am Ende "Übernehmen?"
        ja:   Vorschau wieder entfernen, Gruppe zusammenfassen
              -> ein Rückgängig-Schritt
        nein: Gruppe zurückrollen -> nichts geändert

PickObject ist erlaubt, solange nur die Gruppe offen ist (keine
Transaktion).
"""

import uuid

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    BuiltInParameter,
    CategoryType,
    Color,
    CurveElement,
    ElementId,
    ElementMulticategoryFilter,
    ElementTypeGroup,
    FailureProcessingResult,
    FailureSeverity,
    FilteredElementCollector,
    FillPatternElement,
    IFailuresPreprocessor,
    OverrideGraphicSettings,
    StorageType,
    TextNote,
    Transaction,
    TransactionGroup,
    XYZ,
)
from Autodesk.Revit.Exceptions import OperationCanceledException  # noqa: E402
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType  # noqa: E402

from mlg_sprache import t  # noqa: E402
from stuetzen_nummer import logik as lg  # noqa: E402

FUSS = 0.3048

# Stützen höchstens so weit auseinander (Grundriss) gelten als übereinander
STAPEL_TOLERANZ = 0.05 / FUSS

# Kategorien werden über den Wert ihrer Id geführt (BuiltInCategory als
# Zahl) - gleich in jedem Projekt und jeder Sprache
TRAGWERK = int(BuiltInCategory.OST_StructuralColumns)
ARCHITEKTUR = int(BuiltInCategory.OST_Columns)
WAND = int(BuiltInCategory.OST_Walls)
STANDARD = (TRAGWERK, ARCHITEKTUR)

# Diese Kategorien gelten beim Stapeln als eine (Tragwerks- und
# architektonische Stütze übereinander bekommen dieselbe Nummer)
STAPEL_GRUPPEN = ((TRAGWERK, ARCHITEKTUR),)

# Linienelemente, deren Enden höchstens so weit auseinander liegen (Höhe),
# gelten als waagrecht (Wände, Träger, Rohre) -> Mitte statt unterer
# Endpunkt
WAAGRECHT_TOLERANZ = 0.01 / FUSS

# Diese eingebauten Parameter stehen in der Liste oben
BEVORZUGT = ("ALL_MODEL_MARK", "ALL_MODEL_INSTANCE_COMMENTS")

VORSCHAU_FARBE = Color(0, 150, 60)

_klassen = {}
_pruefungen = {}


class Abbruch(Exception):
    """Der Nutzer hat die Auswahl abgebrochen."""


def id_wert(element_id):
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def beschreibe(element):
    try:
        return u"%s [%d]" % (element.Name, id_wert(element.Id))
    except Exception:
        return u"[%d]" % id_wert(element.Id)


# ---------------------------------------------------------------- Kategorien

def kategorie_wert(element):
    kategorie = getattr(element, "Category", None)
    return id_wert(kategorie.Id) if kategorie is not None else None


def _modellkategorien(doc):
    """{Id-Wert: Kategorie} aller Modellkategorien, die Parameter tragen
    können (keine Beschriftungen, keine internen)."""
    ergebnis = {}
    for kategorie in doc.Settings.Categories:
        try:
            if (kategorie.CategoryType != CategoryType.Model
                    or kategorie.IsTagCategory
                    or not kategorie.AllowsBoundParameters
                    or not kategorie.IsVisibleInUI):
                continue
        except Exception:
            continue
        ergebnis[id_wert(kategorie.Id)] = kategorie
    return ergebnis


def kategorie_optionen(doc):
    """[(Id-Wert, Name)] der Modellkategorien, von denen es im Projekt
    Elemente gibt - nach Namen sortiert."""
    ergebnis = []
    for wert_, kategorie in _modellkategorien(doc).items():
        vorhanden = FilteredElementCollector(doc) \
            .OfCategoryId(kategorie.Id).WhereElementIsNotElementType() \
            .FirstElementId()
        if vorhanden != ElementId.InvalidElementId:
            ergebnis.append((wert_, u"%s" % kategorie.Name))
    return sorted(ergebnis, key=lambda e: e[1].lower())


# ---------------------------------------------------------------- Elemente

def ist_stuetze(element, kategorien):
    """Element einer der gewählten Kategorien mit bestimmbarer Lage."""
    return (kategorie_wert(element) in kategorien
            and punkt(element) is not None)


def ist_linienelement(element):
    """Waagrechtes Linienelement (Wand, Träger, Rohr ...)."""
    kurve = getattr(element.Location, "Curve", None)
    if kurve is None:
        return False
    a, b = kurve.GetEndPoint(0), kurve.GetEndPoint(1)
    return abs(a.Z - b.Z) <= WAAGRECHT_TOLERANZ


def _sammler(doc, kategorien, ansicht=None):
    from System.Collections.Generic import List
    vorhanden = _modellkategorien(doc)
    liste = List[ElementId]()
    for k in kategorien:
        if k in vorhanden:
            liste.Add(vorhanden[k].Id)
    if liste.Count == 0:
        return []
    sammler = (FilteredElementCollector(doc, ansicht.Id) if ansicht
               else FilteredElementCollector(doc))
    return sammler.WherePasses(ElementMulticategoryFilter(liste)) \
        .WhereElementIsNotElementType()


def stuetzen(doc, kategorien, ansicht=None):
    """Alle Elemente der Kategorien (oder die in der Ansicht sichtbaren)."""
    return [e for e in _sammler(doc, kategorien, ansicht)
            if ist_stuetze(e, kategorien)]


def punkt(stuetze):
    """Einfügepunkt (XYZ); bei geneigten Stützen der untere Endpunkt, bei
    waagrechten Linienelementen (Wände) die Mitte - sonst hinge die
    Reihenfolge im Pfad-Modus von der Zeichenrichtung der Wand ab.
    Elemente ohne Punkt oder Linie (Decken, Dächer ...): Mitte des
    Umgrenzungsrahmens auf seiner Unterkante."""
    ort = stuetze.Location
    p = getattr(ort, "Point", None)
    if p is not None:
        return p
    kurve = getattr(ort, "Curve", None)
    if kurve is None:
        try:
            rahmen = stuetze.get_BoundingBox(None)
        except Exception:
            rahmen = None
        if rahmen is None:
            return None
        return XYZ((rahmen.Min.X + rahmen.Max.X) / 2.0,
                   (rahmen.Min.Y + rahmen.Max.Y) / 2.0, rahmen.Min.Z)
    a, b = kurve.GetEndPoint(0), kurve.GetEndPoint(1)
    if abs(a.Z - b.Z) <= WAAGRECHT_TOLERANZ:
        return kurve.Evaluate(0.5, True)
    return a if a.Z <= b.Z else b


def stapel_gruppe(kategorie):
    """Die Kategorien, mit denen gestapelt wird."""
    for gruppe in STAPEL_GRUPPEN:
        if kategorie in gruppe:
            return gruppe
    return (kategorie,)


def _xy(p):
    return (p.X, p.Y)


def grundriss(stuetze):
    """(x, y) des Einfügepunkts."""
    return _xy(punkt(stuetze))


# ---------------------------------------------------------------- Parameter

def _stichprobe(doc, kategorien, je_kategorie):
    for k in kategorien:
        for nummer, element in enumerate(_sammler(doc, [k])):
            if nummer >= je_kategorie:
                break
            yield element


def parameter_optionen(doc, kategorien, stichprobe=50):
    """[(Schlüssel, Anzeigename)] der beschreibbaren Text-Parameter der
    Elemente (Stichprobe je Kategorie).
    Schlüssel: 'bip:ALL_MODEL_MARK' oder 'name:<Name>'."""
    gefunden = {}
    for element in _stichprobe(doc, kategorien, stichprobe):
        for parameter in element.Parameters:
            if (parameter.StorageType != StorageType.String
                    or parameter.IsReadOnly):
                continue
            definition = parameter.Definition
            eingebaut = getattr(definition, "BuiltInParameter", None)
            if (eingebaut is not None
                    and eingebaut != BuiltInParameter.INVALID):
                schluessel = u"bip:%s" % eingebaut.ToString()
            else:
                schluessel = u"name:%s" % definition.Name
            gefunden.setdefault(schluessel, definition.Name)

    def rang(eintrag):
        schluessel, name = eintrag
        for i, bip in enumerate(BEVORZUGT):
            if schluessel == u"bip:" + bip:
                return (i, u"")
        return (len(BEVORZUGT), name.lower())
    return sorted(gefunden.items(), key=rang)


def standard_optionen():
    """Falls es noch keine Stützen gibt: Kennzeichen und Kommentare."""
    return [(u"bip:ALL_MODEL_MARK", t(u"Kennzeichen", u"Mark", u"Marca")),
            (u"bip:ALL_MODEL_INSTANCE_COMMENTS",
             t(u"Kommentare", u"Comments", u"Comentarios"))]


def parameter(element, schluessel):
    art, _, name = schluessel.partition(u":")
    if art == u"bip":
        try:
            return element.get_Parameter(getattr(BuiltInParameter, name))
        except Exception:
            return None
    return element.LookupParameter(name)


def wert(element, schluessel):
    p = parameter(element, schluessel)
    return (p.AsString() or u"") if p is not None else u""


# ---------------------------------------------------------------- Auswahl

def _filter(name, erlaubt):
    """ISelectionFilter als .NET-Klasse. Der Namensraum ist je Sitzung
    eindeutig, weil die CPython-Engine im Revit-Prozess weiterlebt."""
    if name not in _klassen:
        namensraum = "pyMLG.StuetzenNummer_%s_%s" % (name, uuid.uuid4().hex)

        class Filter(ISelectionFilter):
            __namespace__ = namensraum

            def AllowElement(self, element):
                try:
                    return bool(_pruefungen[name](element))
                except Exception:
                    return False

            def AllowReference(self, referenz, punkt_):
                return False

        _klassen[name] = Filter
    _pruefungen[name] = erlaubt
    return _klassen[name]()


def waehle_stuetze(uidoc, kategorien, hinweis):
    """Eine Stütze anklicken - None bei ESC."""
    try:
        referenz = uidoc.Selection.PickObject(
            ObjectType.Element,
            _filter("stuetze", lambda e: ist_stuetze(e, kategorien)),
            hinweis)
    except OperationCanceledException:
        return None
    return uidoc.Document.GetElement(referenz.ElementId)


def _ist_linie(element):
    return (isinstance(element, CurveElement)
            and element.GeometryCurve is not None)


def gewaehlte_linien(doc, auswahl):
    linien = [doc.GetElement(i) for i in auswahl]
    return [e for e in linien if _ist_linie(e)]


def waehle_linie(uidoc, hinweis):
    """Eine Linie anklicken - None bei ESC."""
    try:
        referenz = uidoc.Selection.PickObject(
            ObjectType.Element, _filter("linie", _ist_linie), hinweis)
    except OperationCanceledException:
        return None
    return uidoc.Document.GetElement(referenz.ElementId)


def waehle_linien(uidoc):
    """Linien nacheinander anklicken (ESC beendet). Die gewählten bleiben
    markiert, damit man sieht, was schon dabei ist."""
    gewaehlt = []
    while True:
        hinweis = t(u"Linie %d des Pfads anklicken - ESC = fertig",
                    u"Pick line %d of the path - ESC = done",
                    u"Seleccione la línea %d del recorrido - ESC = terminar") \
            % (len(gewaehlt) + 1)
        try:
            referenz = uidoc.Selection.PickObject(
                ObjectType.Element, _filter("linie", _ist_linie), hinweis)
        except OperationCanceledException:
            break
        element = uidoc.Document.GetElement(referenz.ElementId)
        if all(e.Id != element.Id for e in gewaehlt):
            gewaehlt.append(element)
            _markiere(uidoc, gewaehlt)
    return gewaehlt


def _markiere(uidoc, elemente):
    from System.Collections.Generic import List
    ids = List[ElementId]()
    for e in elemente:
        ids.Add(e.Id)
    try:
        uidoc.Selection.SetElementIds(ids)
    except Exception:
        pass


def pfad(linien):
    """Linienelemente -> Segmente für logik.entlang (Grundriss)."""
    stuecke = [[_xy(p) for p in e.GeometryCurve.Tessellate()]
               for e in linien]
    return lg.ketten(stuecke)


# ---------------------------------------------------------------- Lauf

def _vorverarbeiter():
    """Warnungen (z.B. doppeltes Kennzeichen bei übereinanderstehenden
    Stützen) still bestätigen."""
    if "fehler" not in _klassen:
        namensraum = "pyMLG.StuetzenNummer_%s" % uuid.uuid4().hex

        class Vorverarbeiter(IFailuresPreprocessor):
            __namespace__ = namensraum

            def PreprocessFailures(self, accessor):
                for meldung in list(accessor.GetFailureMessages()):
                    if meldung.GetSeverity() == FailureSeverity.Warning:
                        accessor.DeleteWarning(meldung)
                return FailureProcessingResult.Continue

        _klassen["fehler"] = Vorverarbeiter
    return _klassen["fehler"]()


def _vollfuellung(doc):
    for muster in FilteredElementCollector(doc).OfClass(FillPatternElement):
        try:
            if muster.GetFillPattern().IsSolidFill:
                return muster.Id
        except Exception:
            pass
    return None


class Lauf(object):
    """Nummern vergeben mit Vorschau in der aktiven Ansicht."""

    def __init__(self, uidoc, schluessel, kategorien, stapeln):
        self.uidoc = uidoc
        self.doc = uidoc.Document
        self.ansicht = self.doc.ActiveView
        self.schluessel = schluessel
        self.kategorien = kategorien
        self.stapeln = stapeln
        self.nummern = {}          # Id-Wert -> Nummer
        self.reihenfolge = []      # [(Nummer, angeklickte Stütze, Anzahl)]
        self.fehler = []           # [(Stütze, Grund)]
        self._notizen = []
        self._alte_ueberschreibung = {}
        self._notiztyp = self.doc.GetDefaultElementTypeId(
            ElementTypeGroup.TextNoteType)
        self._notizen_moeglich = self._notiztyp != ElementId.InvalidElementId
        self._fuellung = _vollfuellung(self.doc)
        self._alle = {}            # Stapelgruppe -> ([Elemente], [xy])
        self.gruppe = TransactionGroup(
            self.doc, t(u"Elemente nummerieren", u"Number elements",
                        u"Numerar elementos"))
        self.gruppe.Start()

    # --- Zuordnung --------------------------------------------------------
    def vergeben(self, stuetze):
        return self.nummern.get(id_wert(stuetze.Id))

    def _stapel(self, stuetze):
        """Das Element und - falls gewünscht - alle Elemente derselben
        Kategorie über und unter ihm, die noch keine Nummer haben.
        Waagrechte Linienelemente (Wände, Träger) werden nie gestapelt: ein
        einzelner Punkt je Wand trifft die Wand darüber nur, wenn sie exakt
        gleich gezeichnet ist (siehe punkt)."""
        if not self.stapeln or ist_linienelement(stuetze):
            return [stuetze]
        gruppe = tuple(k for k in stapel_gruppe(kategorie_wert(stuetze))
                       if k in self.kategorien)
        if gruppe not in self._alle:
            alle = [e for e in stuetzen(self.doc, gruppe)
                    if not ist_linienelement(e)]
            self._alle[gruppe] = (alle, [grundriss(e) for e in alle])
        alle, alle_xy = self._alle[gruppe]
        treffer = lg.uebereinander(grundriss(stuetze), alle_xy,
                                   STAPEL_TOLERANZ)
        ergebnis = [stuetze]
        for i in treffer:
            e = alle[i]
            if e.Id != stuetze.Id and self.vergeben(e) is None:
                ergebnis.append(e)
        return ergebnis

    def nummeriere(self, liste, zaehler):
        """Stützen der Reihe nach nummerieren - eine Transaktion für alle.
        Stützen mit schon vergebener Nummer (z.B. als Teil eines Stapels)
        werden übersprungen und verbrauchen keine Nummer."""
        auftraege = []
        for stuetze in liste:
            if self.vergeben(stuetze) is not None:
                continue
            nummer = zaehler.weiter()
            gruppe = self._stapel(stuetze)
            for e in gruppe:
                self.nummern[id_wert(e.Id)] = nummer
            auftraege.append((stuetze, nummer, gruppe))
        if not auftraege:
            return
        transaktion = Transaction(self.doc, t(u"Element nummerieren",
                                              u"Number element",
                                              u"Numerar elemento"))
        transaktion.Start()
        try:
            optionen = transaktion.GetFailureHandlingOptions()
            optionen.SetFailuresPreprocessor(_vorverarbeiter())
            transaktion.SetFailureHandlingOptions(optionen)
            for stuetze, nummer, gruppe in auftraege:
                for e in gruppe:
                    self._schreibe(e, nummer)
                self._vorschau(stuetze, nummer)
                self.reihenfolge.append((nummer, stuetze, len(gruppe)))
        except Exception:
            transaktion.RollBack()
            raise
        transaktion.Commit()
        try:
            self.uidoc.RefreshActiveView()
        except Exception:
            pass

    def _schreibe(self, element, nummer):
        p = parameter(element, self.schluessel)
        if p is None:
            self.fehler.append((element, t(u"Parameter fehlt",
                                           u"parameter missing",
                                           u"falta el parámetro")))
        elif p.IsReadOnly:
            self.fehler.append((element, t(u"Parameter schreibgeschützt",
                                           u"parameter is read-only",
                                           u"parámetro de solo lectura")))
        else:
            try:
                p.Set(nummer)
            except Exception as fehler:
                self.fehler.append((element, u"%s" % fehler))

    # --- Vorschau ---------------------------------------------------------
    def _vorschau(self, stuetze, nummer):
        ansicht = self.ansicht
        if self._notizen_moeglich:
            try:
                p = punkt(stuetze)
                notiz = TextNote.Create(self.doc, ansicht.Id,
                                        XYZ(p.X, p.Y, p.Z), nummer,
                                        self._notiztyp)
                self._notizen.append(notiz.Id)
            except Exception:
                # z.B. 3D-Ansicht: dort gibt es keine Textnotizen
                self._notizen_moeglich = False
        schluessel = id_wert(stuetze.Id)
        if schluessel in self._alte_ueberschreibung:
            return
        try:
            alt = ansicht.GetElementOverrides(stuetze.Id)
            neu = OverrideGraphicSettings(alt)
            neu.SetProjectionLineColor(VORSCHAU_FARBE)
            neu.SetCutLineColor(VORSCHAU_FARBE)
            if self._fuellung is not None:
                neu.SetSurfaceForegroundPatternId(self._fuellung)
                neu.SetSurfaceForegroundPatternColor(VORSCHAU_FARBE)
                neu.SetCutForegroundPatternId(self._fuellung)
                neu.SetCutForegroundPatternColor(VORSCHAU_FARBE)
            ansicht.SetElementOverrides(stuetze.Id, neu)
            self._alte_ueberschreibung[schluessel] = (stuetze.Id, alt)
        except Exception:
            pass

    def _vorschau_entfernen(self):
        transaktion = Transaction(self.doc, t(u"Vorschau entfernen",
                                              u"Remove preview",
                                              u"Quitar vista previa"))
        transaktion.Start()
        for element_id in self._notizen:
            try:
                self.doc.Delete(element_id)
            except Exception:
                pass
        for element_id, alt in self._alte_ueberschreibung.values():
            try:
                self.ansicht.SetElementOverrides(element_id, alt)
            except Exception:
                pass
        transaktion.Commit()

    # --- Ende -------------------------------------------------------------
    def doppelte(self):
        """{Nummer: [Stützen]} - Stützen außerhalb dieses Laufs, die schon
        eine der neuen Nummern tragen."""
        neu = set(self.nummern.values())
        ergebnis = {}
        for e in stuetzen(self.doc, self.kategorien):
            if self.vergeben(e) is None:
                w = wert(e, self.schluessel)
                if w in neu:
                    ergebnis.setdefault(w, []).append(e)
        return ergebnis

    def uebernehmen(self):
        try:
            self._vorschau_entfernen()
        except Exception:
            self.gruppe.RollBack()
            raise
        self.gruppe.Assimilate()

    def verwerfen(self):
        if self.gruppe.HasStarted():
            self.gruppe.RollBack()
