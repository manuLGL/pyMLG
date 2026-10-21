# -*- coding: utf-8 -*-
"""Revit-Zugriffe von ApartmentNumbering.

Eine Wohnung ist eine Modellgruppe, deren Typname mit WHG_TYP beginnt
(nur oberste Gruppen - verschachtelte gehören zur äusseren). Ihre Räume
(auch in verschachtelten Gruppen) bekommen die Werte.

Räume in Gruppen dürfen nur dann je Instanz verschiedene Werte haben, wenn
der Parameter "Werte je nach Gruppeninstanz unterschiedlich" erlaubt
(InternalDefinition.VariesAcrossGroups). Sonst meldet Revit beim Schreiben
"Gruppe ausserhalb des Gruppenbearbeitungsmodus geändert" und bietet an,
die Gruppe aufzulösen - das passiert hier nie: das Werkzeug prüft vorher
und rollt bei solchen Fehlern zurück.

Ein Lauf ist eine Transaktionsgruppe:

    Varianz einschalten (falls bestätigt), Werte schreiben,
    Vorschau in der aktiven Ansicht (Textnotiz mit der TOP-Nummer)
    am Ende "Übernehmen?"
        ja:   Vorschau entfernen, Gruppe zusammenfassen
              -> ein Rückgängig-Schritt
        nein: Gruppe zurückrollen -> nichts geändert
"""

import uuid

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    CurveElement,
    ElementId,
    ElementTypeGroup,
    FailureProcessingResult,
    FailureSeverity,
    FilteredElementCollector,
    Group,
    IFailuresPreprocessor,
    Level,
    SpatialElementBoundaryOptions,
    StorageType,
    TextNote,
    Transaction,
    TransactionGroup,
    TransactionStatus,
    XYZ,
)
from Autodesk.Revit.DB.Architecture import Room  # noqa: E402

from mlg_sprache import t  # noqa: E402
from wohnung_top import logik as lg  # noqa: E402

_klassen = {}

# Ein Ring mit weniger Punkten ist keine Fläche
MIN_PUNKTE = 3


def id_wert(element_id):
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def _xy(p):
    return (p.X, p.Y)


# ---------------------------------------------------------------- Wohnungen

class Wohnung(object):
    """Eine Wohnungsgruppe mit ihren Räumen."""

    def __init__(self, gruppe, name, raeume, ebene_id):
        self.gruppe = gruppe
        self.name = name
        self.raeume = raeume
        self.ebene_id = ebene_id
        self._flaechen = None

    @property
    def schluessel(self):
        return id_wert(self.gruppe.Id)

    def beschreibung(self):
        return u"%s [%d]" % (self.name, self.schluessel)

    def flaechen(self):
        """[[Ring, ...] je Raum] im Grundriss. Räume ohne Umgrenzung
        (nicht platziert, nicht umschlossen) fehlen; hat keiner eine, gilt
        der Umgrenzungsrahmen der Gruppe."""
        if self._flaechen is None:
            optionen = SpatialElementBoundaryOptions()
            flaechen = []
            for raum in self.raeume:
                ringe = []
                try:
                    schleifen = raum.GetBoundarySegments(optionen) or []
                except Exception:
                    schleifen = []
                for schleife in schleifen:
                    ring = []
                    for segment in schleife:
                        punkte = [_xy(p) for p in
                                  segment.GetCurve().Tessellate()]
                        ring.extend(punkte[1:] if ring else punkte)
                    if len(ring) >= MIN_PUNKTE:
                        ringe.append(ring)
                if ringe:
                    flaechen.append(ringe)
            if not flaechen:
                rahmen = self.gruppe.get_BoundingBox(None)
                if rahmen is not None:
                    x0, y0 = rahmen.Min.X, rahmen.Min.Y
                    x1, y1 = rahmen.Max.X, rahmen.Max.Y
                    flaechen.append([[(x0, y0), (x1, y0), (x1, y1),
                                      (x0, y1)]])
            self._flaechen = flaechen
        return self._flaechen

    def mitte(self, hoehe):
        """Punkt für die Vorschau-Notiz: Mitte des Umgrenzungsrahmens."""
        rahmen = self.gruppe.get_BoundingBox(None)
        if rahmen is None:
            return None
        return XYZ((rahmen.Min.X + rahmen.Max.X) / 2.0,
                   (rahmen.Min.Y + rahmen.Max.Y) / 2.0, hoehe)


def _raeume(doc, gruppe):
    """Räume der Gruppe, auch aus verschachtelten Gruppen."""
    ergebnis = []
    for i in gruppe.GetMemberIds():
        element = doc.GetElement(i)
        if isinstance(element, Room):
            ergebnis.append(element)
        elif isinstance(element, Group):
            ergebnis.extend(_raeume(doc, element))
    return ergebnis


def wohnungen(doc):
    """Alle obersten Modellgruppen mit WHG_TYP-Typnamen."""
    ergebnis = []
    sammler = FilteredElementCollector(doc).OfClass(Group) \
        .OfCategory(BuiltInCategory.OST_IOSModelGroups)
    for gruppe in sammler:
        if gruppe.GroupId != ElementId.InvalidElementId:
            continue                    # verschachtelt
        typ = gruppe.GroupType
        name = typ.Name if typ is not None else u""
        if not lg.ist_wohnung(name):
            continue
        raeume = _raeume(doc, gruppe)
        ebene_id = raeume[0].LevelId if raeume else gruppe.LevelId
        ergebnis.append(Wohnung(gruppe, name, raeume, ebene_id))
    return ergebnis


class Ebene(object):
    def __init__(self, ebene, anzahl):
        self.ebene = ebene
        self.name = ebene.Name
        self.hoehe = ebene.Elevation
        self.anzahl = anzahl

    @property
    def schluessel(self):
        return id_wert(self.ebene.Id)


def ebenen(doc, liste):
    """Ebenen mit Wohnungen, von oben nach unten."""
    anzahl = {}
    for w in liste:
        anzahl[id_wert(w.ebene_id)] = anzahl.get(id_wert(w.ebene_id), 0) + 1
    ergebnis = [Ebene(e, anzahl[id_wert(e.Id)])
                for e in FilteredElementCollector(doc).OfClass(Level)
                if id_wert(e.Id) in anzahl]
    return sorted(ergebnis, key=lambda e: -e.hoehe)


# ---------------------------------------------------------------- Linie

# Linienenden höchstens so weit auseinander gelten als verbunden (Fuss) -
# wie stuetzen_nummer.logik.VERBINDUNG
VERBINDUNG = 0.01


def linienzug(doc, linien):
    """Die angeklickten Linien samt allen, die über ihre Endpunkte daran
    hängen (gleiche Ansicht bei Detaillinien). Revit legt eine in einem Zug
    gezeichnete Linie als einzelne Stücke an."""
    eigner = set(id_wert(e.OwnerViewId) for e in linien)
    kandidaten = [e for e in FilteredElementCollector(doc)
                  .OfCategory(BuiltInCategory.OST_Lines)
                  .WhereElementIsNotElementType()
                  if isinstance(e, CurveElement)
                  and e.GeometryCurve is not None
                  and id_wert(e.OwnerViewId) in eigner]
    gewaehlt = set(id_wert(e.Id) for e in linien)
    enden = []
    for e in kandidaten:
        kurve = e.GeometryCurve
        enden.append(tuple((p.X, p.Y, p.Z) for p in
                           (kurve.GetEndPoint(0), kurve.GetEndPoint(1))))
    start = [i for i, e in enumerate(kandidaten)
             if id_wert(e.Id) in gewaehlt]
    return [kandidaten[i] for i in lg.verbunden(enden, start, VERBINDUNG)]


# ---------------------------------------------------------------- Parameter

def parameter(element, name):
    """Parameter nach Namen - Gross-/Kleinschreibung egal."""
    p = element.LookupParameter(name)
    if p is not None:
        return p
    klein = name.lower()
    for kandidat in element.Parameters:
        if kandidat.Definition.Name.lower() == klein:
            return kandidat
    return None


def pruefe_parameter(raeume, namen):
    """(fehlend, fest): Namen, die keinem Raum gehören, und
    {Name: InternalDefinition} der Parameter, die nicht je Gruppeninstanz
    variieren dürfen."""
    gefunden, fest = set(), {}
    for raum in raeume:
        for name in namen:
            if name in gefunden:
                continue
            p = parameter(raum, name)
            if p is None:
                continue
            gefunden.add(name)
            definition = p.Definition
            if not getattr(definition, "VariesAcrossGroups", True):
                fest[name] = definition
    return [n for n in namen if n not in gefunden], fest


def _setze(p, wert):
    """Wert passend zum Speichertyp setzen - Fehlertext oder None."""
    if p.IsReadOnly:
        return t(u"schreibgeschützt", u"read-only", u"de solo lectura")
    art = p.StorageType
    try:
        if art == StorageType.String:
            p.Set(wert)
        elif art == StorageType.Integer:
            p.Set(int(wert))
        elif art == StorageType.Double:
            p.Set(float(wert.replace(u",", u".")))
        else:
            return t(u"falscher Parametertyp", u"wrong parameter type",
                     u"tipo de parámetro incorrecto")
    except ValueError:
        return t(u"'%s' ist keine Zahl", u"'%s' is not a number",
                 u"'%s' no es un número") % wert
    except Exception as fehler:
        return u"%s" % fehler
    return None


# ---------------------------------------------------------------- Lauf

class GruppenFehler(Exception):
    """Revit hat das Schreiben abgelehnt (z.B. Gruppe ausserhalb des
    Bearbeitungsmodus geändert) - alles wurde zurückgerollt."""


def _vorverarbeiter(fehler):
    """Warnungen still bestätigen. Fehler (etwa die Gruppen-Meldung mit
    "Gruppierung aufheben") nie auflösen lassen: zurückrollen und den Text
    in fehler sammeln."""
    if "fehler" not in _klassen:
        namensraum = "pyMLG.WohnungTop_%s" % uuid.uuid4().hex

        class Vorverarbeiter(IFailuresPreprocessor):
            __namespace__ = namensraum

            def PreprocessFailures(self, accessor):
                abbruch = False
                for meldung in list(accessor.GetFailureMessages()):
                    if meldung.GetSeverity() == FailureSeverity.Warning:
                        accessor.DeleteWarning(meldung)
                    else:
                        abbruch = True
                        try:
                            _klassen["sammlung"].append(
                                meldung.GetDescriptionText())
                        except Exception:
                            pass
                if abbruch:
                    return FailureProcessingResult.ProceedWithRollBack
                return FailureProcessingResult.Continue

        _klassen["fehler"] = Vorverarbeiter
    _klassen["sammlung"] = fehler
    return _klassen["fehler"]()


class Lauf(object):
    """Werte schreiben mit Vorschau in der aktiven Ansicht."""

    def __init__(self, uidoc):
        self.uidoc = uidoc
        self.doc = uidoc.Document
        self.ansicht = self.doc.ActiveView
        self.raeume = 0
        self.fehler = []           # [(Wohnung, Raum, Parameter, Grund)]
        self._notizen = []
        self._notiztyp = self.doc.GetDefaultElementTypeId(
            ElementTypeGroup.TextNoteType)
        self.gruppe = TransactionGroup(
            self.doc, t(u"Wohnungen nummerieren", u"Number apartments",
                        u"Numerar viviendas"))
        self.gruppe.Start()

    def _transaktion(self, name, arbeit):
        gesammelt = []
        transaktion = Transaction(self.doc, name)
        transaktion.Start()
        optionen = transaktion.GetFailureHandlingOptions()
        optionen.SetFailuresPreprocessor(_vorverarbeiter(gesammelt))
        transaktion.SetFailureHandlingOptions(optionen)
        try:
            arbeit()
        except Exception:
            transaktion.RollBack()
            raise
        if transaktion.Commit() != TransactionStatus.Committed:
            raise GruppenFehler(u"\n".join(gesammelt) or t(
                u"Revit hat die Änderung abgelehnt.",
                u"Revit rejected the change.",
                u"Revit ha rechazado el cambio."))

    def erlaube_varianz(self, definitionen):
        """'Werte je nach Gruppeninstanz unterschiedlich' einschalten."""
        def arbeit():
            for definition in definitionen:
                definition.SetAllowVaryBetweenGroups(self.doc, True)
        self._transaktion(t(u"Werte je Gruppeninstanz erlauben",
                            u"Allow values per group instance",
                            u"Permitir valores por instancia de grupo"),
                          arbeit)

    def schreibe(self, auftraege, vorschau_ebene):
        """auftraege: [(Wohnung, {Parameter: Wert})]. Für Wohnungen auf
        vorschau_ebene (Id-Wert oder None) kommt die TOP-Nummer als
        Textnotiz in die aktive Ansicht."""
        def arbeit():
            for wohnung, werte in auftraege:
                for raum in wohnung.raeume:
                    self.raeume += 1
                    for name in lg.ALLE:
                        if name not in werte:
                            continue
                        p = parameter(raum, name)
                        if p is None:
                            continue    # gibt es nicht -> nicht füllen
                        grund = _setze(p, werte[name])
                        if grund:
                            self.fehler.append((wohnung, raum, name, grund))
                if id_wert(wohnung.ebene_id) == vorschau_ebene:
                    self._vorschau(wohnung, werte.get(lg.TOP, u""))
        self._transaktion(t(u"Wohnungen nummerieren", u"Number apartments",
                            u"Numerar viviendas"), arbeit)
        try:
            self.uidoc.RefreshActiveView()
        except Exception:
            pass

    # --- Vorschau ---------------------------------------------------------
    def _vorschau(self, wohnung, nummer):
        if self._notiztyp == ElementId.InvalidElementId or not nummer:
            return
        ebene = self.doc.GetElement(wohnung.ebene_id)
        p = wohnung.mitte(ebene.Elevation if ebene is not None else 0.0)
        if p is None:
            return
        try:
            notiz = TextNote.Create(self.doc, self.ansicht.Id, p,
                                    u"TOP %s" % nummer, self._notiztyp)
            self._notizen.append(notiz.Id)
        except Exception:
            # z.B. 3D-Ansicht oder Plan: dort gibt es keine Notiz
            self._notiztyp = ElementId.InvalidElementId

    def _vorschau_entfernen(self):
        if not self._notizen:
            return

        def arbeit():
            for element_id in self._notizen:
                try:
                    self.doc.Delete(element_id)
                except Exception:
                    pass
        self._transaktion(t(u"Vorschau entfernen", u"Remove preview",
                            u"Quitar vista previa"), arbeit)

    # --- Ende -------------------------------------------------------------
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


def vorschau_ebene(ansicht):
    """Id-Wert der Ebene des aktiven Grundrisses (sonst None)."""
    ebene = getattr(ansicht, "GenLevel", None)
    return id_wert(ebene.Id) if ebene is not None else None
