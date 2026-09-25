# -*- coding: utf-8 -*-
"""Kollisionen automatisch lösen: Umgehung für Rohre und Luftkanäle.

Nur die Leitung der gewählten Kollision wird verändert - und nur, wenn sie
im aktiven Modell liegt. Das Hindernis (auch aus einer Verknüpfung) und alle
anderen Elemente werden nur gelesen.

Ablauf:
  analysiere()   Leitung und Hindernis bestimmen, Varianten berechnen
                 (umgehung.py) und gegen die Umgebung prüfen
  zeige_vorschau() die Variante als farbige DirectShape in die Ansicht legen
  uebernehme()   Leitung zweimal auftrennen, Mittelstück ersetzen, Bögen
                 setzen (Routing-Einstellungen des Typs), Dämmung übertragen
                 - eine Transaktion, Strg+Z in Revit nimmt alles zurück

Welche Leitung weicht aus? Nur Rohre und Kanäle des aktiven Modells. Sind
beide Elemente solche Leitungen, werden beide durchgerechnet (A = kleinere,
B = grössere) und die Karten zeigen die besten Varianten beider.
"""

import math
import time

import clr

from Autodesk.Revit.DB import (BuiltInCategory, BuiltInParameter, Color,
                               ConnectorType,
                               CurveLoop, DirectShape, ElementId,
                               ElementIntersectsSolidFilter,
                               ElementMulticategoryFilter,
                               ElementTransformUtils,
                               FailureProcessingResult, FailureSeverity,
                               FilteredElementCollector, FillPatternElement,
                               GeometryCreationUtilities, GeometryObject,
                               IFailuresPreprocessor, InsulationLiningBase,
                               Line, Arc, Outline, OverrideGraphicSettings,
                               BoundingBoxIntersectsFilter, SolidUtils,
                               Transaction, TransactionGroup, XYZ)
from Autodesk.Revit.DB.Mechanical import Duct, DuctInsulation, MechanicalUtils
from Autodesk.Revit.DB.Plumbing import Pipe, PipeInsulation, PlumbingUtils

from mlg_sprache import t

from clash_navigator import logik as lg
from clash_navigator import revit as rv
from clash_navigator import umgehung as ug

clr.AddReference("System")
from System.Collections.Generic import List  # noqa: E402

KENNUNG = u"pyMLG.ClashNavigator.Vorschau"
FUSS = 1.0 / lg.M_PRO_FUSS

# Diese Kategorien zählen bei der Prüfung als Hindernis
_KATEGORIEN = (
    BuiltInCategory.OST_Walls, BuiltInCategory.OST_Floors,
    BuiltInCategory.OST_Roofs, BuiltInCategory.OST_Ceilings,
    BuiltInCategory.OST_StructuralFraming,
    BuiltInCategory.OST_StructuralColumns, BuiltInCategory.OST_Columns,
    BuiltInCategory.OST_StructuralFoundation,
    BuiltInCategory.OST_PipeCurves, BuiltInCategory.OST_PipeFitting,
    BuiltInCategory.OST_PipeAccessory, BuiltInCategory.OST_FlexPipeCurves,
    BuiltInCategory.OST_DuctCurves, BuiltInCategory.OST_DuctFitting,
    BuiltInCategory.OST_DuctAccessory, BuiltInCategory.OST_FlexDuctCurves,
    BuiltInCategory.OST_DuctTerminal, BuiltInCategory.OST_CableTray,
    BuiltInCategory.OST_CableTrayFitting, BuiltInCategory.OST_Conduit,
    BuiltInCategory.OST_ConduitFitting,
    BuiltInCategory.OST_MechanicalEquipment,
    BuiltInCategory.OST_PlumbingFixtures, BuiltInCategory.OST_Sprinklers,
    BuiltInCategory.OST_GenericModel, BuiltInCategory.OST_Stairs,
)


class Analyse(object):
    """Ergebnis von analysiere() - Grundlage für Vorschau und Übernehmen.

    Sind beide Elemente der Kollision Leitungen des aktiven Modells, stehen
    beide in 'leitungen'; jede Variante weiss über variante.leitung, welche
    sich bewegt.
    """

    def __init__(self):
        self.doc_titel = u""
        self.leitungen = []
        self.varianten = []
        self.hinweise = []


class Leitungsdaten(object):
    """Eine Leitung, die ausweichen kann, mit allem für Vorschau und
    Übernehmen."""

    def __init__(self, kennung, element_id, text, start, ende, querschnitt,
                 anschluesse_, hindernis_text):
        self.kennung = kennung              # "A" oder "B"
        self.element_id = element_id
        self.text = text
        self.start = start
        self.ende = ende
        self.querschnitt = querschnitt
        self.anschluesse = anschluesse_     # je Ende: Anschluss oder None
        self.hindernis_text = hindernis_text


class Anschluss(object):
    """Bogen an einem Ende der Leitung und die Leitung dahinter."""

    def __init__(self, formteil_id, nachbar_id=None, punkt=None):
        self.formteil_id = formteil_id
        self.nachbar_id = nachbar_id
        self.punkt = punkt          # Verbindung Bogen - Nachbar (XYZ, Fuss)


# --- Hilfen ------------------------------------------------------------------

def _fuss(punkt):
    return XYZ(punkt[0] * FUSS, punkt[1] * FUSS, punkt[2] * FUSS)


def _m(xyz):
    return (xyz.X / FUSS, xyz.Y / FUSS, xyz.Z / FUSS)


def _id_liste(ids):
    liste = List[ElementId]()
    for element_id in ids:
        liste.Add(element_id)
    return liste


def _parameter(element, eingebaut):
    try:
        parameter = element.get_Parameter(eingebaut)
    except Exception:
        return None
    if parameter is None or not parameter.HasValue:
        return None
    return parameter.AsDouble()


def _ist_leitung(element):
    return isinstance(element, (Pipe, Duct))


def _daemmung(doc, element):
    """(Dämmelement oder None, Dicke in Metern)."""
    try:
        ids = InsulationLiningBase.GetInsulationIds(doc, element.Id)
    except Exception:
        return None, 0.0
    for element_id in ids:
        daemmung = doc.GetElement(element_id)
        if isinstance(daemmung, (PipeInsulation, DuctInsulation)):
            return daemmung, daemmung.Thickness / FUSS
    return None, 0.0


def _verbinder(element):
    """Verbinder einer Leitung oder eines Formteils (dort über MEPModel)."""
    for weg in (lambda e: e.ConnectorManager,
                lambda e: e.MEPModel.ConnectorManager):
        try:
            return list(weg(element).Connectors)
        except Exception:
            continue
    return []


def _gegenstuecke(verbinder, eigene_id):
    """Physisch verbundene Verbinder anderer Elemente (ohne Systemlogik)."""
    ergebnis = []
    try:
        for gegenstueck in verbinder.AllRefs:
            if gegenstueck.Owner is None or gegenstueck.Owner.Id == eigene_id:
                continue
            if gegenstueck.ConnectorType == ConnectorType.Logical:
                continue
            ergebnis.append(gegenstueck)
    except Exception:
        pass
    return ergebnis


def _anschluss_an(element, punkt):
    """Was hängt am Ende 'punkt' (Fuss)? Rückgabe (für umgehung, Anschluss):

    (None, None)                     freies Ende
    (ug.FEST, None)                  Abzweig, Gerät, Armatur ...
    ((Richtung, Länge), Anschluss)   Bogen und die Leitung dahinter
    """
    eigener = _verbinder_bei(element, punkt)
    if eigener is None:
        return None, None
    gegen = _gegenstuecke(eigener, element.Id)
    if not gegen:
        return None, None
    nachbar_el = gegen[0].Owner
    if _ist_leitung(nachbar_el):
        # direkt verbundene Leitung (ohne Bogen) - liegt in Flucht
        return _richtung_weg(nachbar_el, punkt), None
    formteil = nachbar_el
    verbinder = _verbinder(formteil)
    if len(verbinder) != 2:
        return ug.FEST, None
    anderer = max(verbinder, key=lambda v: v.Origin.DistanceTo(punkt))
    weiter = _gegenstuecke(anderer, formteil.Id)
    if not weiter:
        return None, Anschluss(formteil.Id)
    leitung = weiter[0].Owner
    if not _ist_leitung(leitung):
        return ug.FEST, None
    return (_richtung_weg(leitung, anderer.Origin),
            Anschluss(formteil.Id, leitung.Id, anderer.Origin))


def _richtung_weg(leitung, punkt):
    """(Richtung vom Punkt weg, Länge) der Leitung in Metern."""
    kurve = leitung.Location.Curve
    a, b = kurve.GetEndPoint(0), kurve.GetEndPoint(1)
    nah, fern = (a, b) if a.DistanceTo(punkt) < b.DistanceTo(punkt) else (b, a)
    return (ug.einheit(ug.minus(_m(fern), _m(nah))), kurve.Length / FUSS)


def anschluesse(element):
    kurve = element.Location.Curve
    ergebnis = [_anschluss_an(element, kurve.GetEndPoint(i)) for i in (0, 1)]
    return [e[0] for e in ergebnis], [e[1] for e in ergebnis]


def querschnitt(doc, element):
    _daemm, dicke = _daemmung(doc, element)
    if isinstance(element, Pipe):
        aussen = (_parameter(element, BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
                  or _parameter(element,
                                BuiltInParameter.RBS_PIPE_DIAMETER_PARAM))
        return ug.Querschnitt(durchmesser=aussen / FUSS, daemmung=dicke)
    durchmesser = _parameter(element, BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
    if durchmesser:
        return ug.Querschnitt(durchmesser=durchmesser / FUSS, daemmung=dicke)
    breite = _parameter(element, BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
    hoehe = _parameter(element, BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
    achse_b, achse_h = (1.0, 0.0, 0.0), (0.0, 0.0, 1.0)
    for verbinder in _verbinder(element):
        system = verbinder.CoordinateSystem
        achse_b = (system.BasisX.X, system.BasisX.Y, system.BasisX.Z)
        achse_h = (system.BasisY.X, system.BasisY.Y, system.BasisY.Z)
        break
    return ug.Querschnitt(breite=breite / FUSS, hoehe=hoehe / FUSS,
                          achse_b=achse_b, achse_h=achse_h, daemmung=dicke)


def _linie(element):
    kurve = element.Location.Curve
    return _m(kurve.GetEndPoint(0)), _m(kurve.GetEndPoint(1))


def _beschreibung(fund):
    text = lg.objekt_text(fund.objekt)
    if not fund.modell.ist_wirt:
        text += u" (%s)" % (fund.modell.namen[0] if fund.modell.namen
                            else u"Link")
    return text


# --- Körper für Vorschau und Prüfung ------------------------------------------

def _profil(mitte, e1, e2, quer, richtung):
    """Geschlossene Kontur quer zur Leitung, Mittelpunkt 'mitte' (Fuss)."""
    schleife = CurveLoop()
    if quer.rund:
        radius = (quer.durchmesser / 2.0 + quer.daemmung) * FUSS
        achse_x = XYZ(*e1)
        achse_y = XYZ(*e2)
        schleife.Append(Arc.Create(mitte, radius, 0.0, math.pi, achse_x,
                                   achse_y))
        schleife.Append(Arc.Create(mitte, radius, math.pi, 2.0 * math.pi,
                                   achse_x, achse_y))
        return schleife
    halb1 = quer.halb(e1) * FUSS
    halb2 = quer.halb(richtung) * FUSS
    v1, v2 = XYZ(*e1), XYZ(*e2)
    ecken = [mitte + v1 * (a * halb1) + v2 * (b * halb2)
             for a, b in ((-1.0, -1.0), (1.0, -1.0), (1.0, 1.0),
                          (-1.0, 1.0))]
    for i in range(4):
        schleife.Append(Line.CreateBound(ecken[i], ecken[(i + 1) % 4]))
    return schleife


def koerper(variante, quer, achse):
    """Volumenkörper der drei neuen Stücke (Fuss, Wirtssystem)."""
    richtung = ug.richtungen(achse)[variante.richtung]
    senkrecht = ug.einheit(ug.kreuz(achse, richtung))
    liste = []
    for von, bis in variante.segmente:
        s = ug.einheit(ug.minus(bis, von))
        e2 = ug.einheit(ug.kreuz(s, senkrecht))
        schleifen = List[CurveLoop]()
        schleifen.Add(_profil(_fuss(von), senkrecht, e2, quer, richtung))
        liste.append(GeometryCreationUtilities.CreateExtrusionGeometry(
            schleifen, XYZ(*s), ug.laenge(ug.minus(bis, von)) * FUSS))
    return liste


def _umriss(koerper_, transform=None):
    kasten = koerper_.GetBoundingBox()
    punkte = []
    for x in (kasten.Min.X, kasten.Max.X):
        for y in (kasten.Min.Y, kasten.Max.Y):
            for z in (kasten.Min.Z, kasten.Max.Z):
                punkt = kasten.Transform.OfPoint(XYZ(x, y, z))
                if transform is not None:
                    punkt = transform.OfPoint(punkt)
                punkte.append(punkt)
    return Outline(XYZ(min(p.X for p in punkte), min(p.Y for p in punkte),
                       min(p.Z for p in punkte)),
                   XYZ(max(p.X for p in punkte), max(p.Y for p in punkte),
                       max(p.Z for p in punkte)))


def _gesamtumriss(umrisse, transform=None):
    """Ein Umriss um mehrere, bei Bedarf in ein anderes System gerechnet."""
    punkte = []
    for umriss in umrisse:
        for x in (umriss.MinimumPoint.X, umriss.MaximumPoint.X):
            for y in (umriss.MinimumPoint.Y, umriss.MaximumPoint.Y):
                for z in (umriss.MinimumPoint.Z, umriss.MaximumPoint.Z):
                    punkt = XYZ(x, y, z)
                    if transform is not None:
                        punkt = transform.OfPoint(punkt)
                    punkte.append(punkt)
    return Outline(XYZ(min(p.X for p in punkte), min(p.Y for p in punkte),
                       min(p.Z for p in punkte)),
                   XYZ(max(p.X for p in punkte), max(p.Y for p in punkte),
                       max(p.Z for p in punkte)))


def _kandidaten(dokument, umriss, ausschluss):
    """Ein einziger Durchlauf je Modell: alle Bauteile im Bereich aller
    Varianten. Die genaue Körperprüfung läuft danach nur noch auf diesen."""
    kategorien = List[BuiltInCategory]()
    for kategorie in _KATEGORIEN:
        kategorien.Add(kategorie)
    sammler = (FilteredElementCollector(dokument)
               .WhereElementIsNotElementType()
               .WherePasses(ElementMulticategoryFilter(kategorien))
               .WherePasses(BoundingBoxIntersectsFilter(umriss)))
    if ausschluss:
        sammler = sammler.Excluding(_id_liste(ausschluss))
    return sammler.ToElementIds()


def _name(element, modell):
    name = u"%s %s" % (element.Category.Name if element.Category else u"",
                       rv.id_wert_text(element.Id))
    if not modell.ist_wirt and modell.namen:
        name += u" (%s)" % modell.namen[0]
    return name


class Vorrat(object):
    """Alle Bauteile rund um die Kollision - einmal je Berechnung gesammelt
    (ein Durchlauf je Modell) und für alle Varianten und Runden benutzt."""

    def __init__(self, liste_modelle, kaesten_m, rand_m):
        punkte = [ecke for kasten in kaesten_m for ecke in kasten]
        klein = XYZ(*[(min(p[a] for p in punkte) - rand_m) * FUSS
                      for a in range(3)])
        gross = XYZ(*[(max(p[a] for p in punkte) + rand_m) * FUSS
                      for a in range(3)])
        umriss_wirt = Outline(klein, gross)
        self.je_modell = []
        for modell in liste_modelle:
            umkehr = None if modell.ist_wirt else modell.transform.Inverse
            try:
                ids = _kandidaten(modell.dokument,
                                  _gesamtumriss([umriss_wirt], umkehr), [])
            except Exception:
                continue
            if ids.Count:
                self.je_modell.append((modell, umkehr, list(ids)))


def pruefe_varianten(vorrat, paare, ausschluss, frist):
    """Setzt variante.kollisionen für alle (variante, körper) in 'paare'.

    Früher lief je Körper und Modell ein Durchlauf über das ganze Modell -
    bei vielen Verknüpfungen Hunderte und Revit stand minutenlang. Jetzt
    kommen die Kandidaten aus dem Vorrat, geprüft werden nur noch Körper.
    frist: Zeitpunkt (time.time()), ab dem abgebrochen wird.
    Rückgabe: False, wenn die Frist die Prüfung abgebrochen hat.
    """
    weg = set(rv.id_wert_text(i) for i in ausschluss)
    for modell, umkehr, alle_ids in vorrat.je_modell:
        if modell.ist_wirt:
            ids = _id_liste([i for i in alle_ids
                             if rv.id_wert_text(i) not in weg])
        else:
            ids = _id_liste(alle_ids)
        if not ids.Count:
            continue
        for variante, koerper_liste in paare:
            for koerper_ in koerper_liste:
                if time.time() > frist:
                    return False
                pruefkoerper = (koerper_ if umkehr is None else
                                SolidUtils.CreateTransformed(koerper_,
                                                             umkehr))
                try:
                    treffer = (FilteredElementCollector(modell.dokument, ids)
                               .WherePasses(BoundingBoxIntersectsFilter(
                                   _umriss(pruefkoerper)))
                               .WherePasses(ElementIntersectsSolidFilter(
                                   pruefkoerper)))
                    for element in treffer:
                        name = _name(element, modell)
                        if name not in variante.kollisionen:
                            variante.kollisionen.append(name)
                except Exception:
                    continue
    return True


def _ausschluss(doc, element):
    """Die Leitung selbst, ihre Dämmung und was an ihr hängt."""
    ids = [element.Id]
    try:
        ids.extend(InsulationLiningBase.GetInsulationIds(doc, element.Id))
    except Exception:
        pass
    for verbinder in _verbinder(element):
        try:
            for gegenstueck in verbinder.AllRefs:
                if gegenstueck.Owner is not None:
                    ids.append(gegenstueck.Owner.Id)
        except Exception:
            pass
    return ids


# --- Analyse -----------------------------------------------------------------

def analysiere(uiapp, clash, optionen):
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        raise rv.Fehler(t(u"Kein Projekt geöffnet.", u"No project open.",
                          u"No hay ningún proyecto abierto."))

    liste = rv.modelle(doc)
    vorab, _ = rv.punkt_intern(doc, clash.punkt,
                               optionen.get(u"koordinaten"), [])
    funde = [rv.finde(o, liste, vorab) for o in clash.objekte
             if o.element_id is not None]
    funde = [fund for fund in funde if fund is not None]

    beweglich = [f for f in funde
                 if f.modell.ist_wirt and _ist_leitung(f.element)]
    if not beweglich:
        raise rv.Fehler(t(
            u"An dieser Kollision ist kein Rohr und kein Luftkanal des "
            u"aktiven Modells beteiligt - nichts zu verlegen.",
            u"No pipe or duct of the active model is involved in this "
            u"clash - nothing to reroute.",
            u"En este conflicto no participa ninguna tubería ni conducto "
            u"del modelo activo; no hay nada que desviar."))

    if len(funde) < 2:
        raise rv.Fehler(t(u"Das zweite Element der Kollision wurde nicht "
                          u"gefunden - ohne Hindernis keine Umgehung.",
                          u"The second element of the clash was not found - "
                          u"no obstacle, no bypass.",
                          u"No se encontró el segundo elemento del "
                          u"conflicto; sin obstáculo no hay desvío."))

    # Kleinere Leitung zuerst - sie heisst "A"
    beweglich.sort(key=lambda f: querschnitt(doc, f.element).groesste)
    analyse = Analyse()
    analyse.doc_titel = doc.Title
    soll = float(optionen.get(u"abstand_cm", 5)) / 100.0
    winkel = optionen.get(u"winkel") or (45, 90)
    # Gesamtfrist für alle Prüfungen - Revit darf nicht hängen
    frist = time.time() + ZEITLIMIT
    vorrat = Vorrat(liste, [f.kasten for f in funde], ug.MAX_VERSATZ + 1.0)

    rechner = []
    for kennung, fund in zip((u"A", u"B"), beweglich):
        hindernis = [f for f in funde if f is not fund][0]
        rechner.append(_Rechner(doc, kennung, fund, hindernis, vorrat,
                                winkel, soll))
    analyse.leitungen = [r.daten for r in rechner]

    je_leitung = []
    unvollstaendig = False
    for teil in rechner:
        alle = []
        # Erst mit der Vorgabe, dann mit kleineren Abständen - aber nur so
        # lange, bis etwas kollisionsfrei passt
        for abstand in ug.abstands_runden(soll):
            runde, fertig = teil.rechne(abstand, frist)
            alle.extend(runde)
            unvollstaendig = unvollstaendig or not fertig
            if any(v.gueltig and not v.kollisionen for v in runde):
                break
        je_leitung.append(alle)

    kombinationen = []
    if len(rechner) == 2:
        kombinationen, fertig = _beide(rechner, soll, frist)
        unvollstaendig = unvollstaendig or not fertig

    if unvollstaendig:
        analyse.hinweise.append(t(
            u"Prüfung nach %d s abgebrochen - Kollisionsfreiheit nicht "
            u"sicher.",
            u"Check stopped after %d s - clash-free status not certain.",
            u"Comprobación detenida tras %d s; no es seguro que no haya "
            u"conflictos.") % ZEITLIMIT)
    analyse.varianten = auswahl(je_leitung, kombinationen)
    return analyse


# Gesamtzeit für alle Kollisionsprüfungen einer Berechnung (Sekunden)
ZEITLIMIT = 25


class _Rechner(object):
    """Varianten für eine Leitung rechnen und prüfen."""

    def __init__(self, doc, kennung, fund, hindernis, vorrat, winkel, soll):
        element = fund.element
        self.hindernis = hindernis
        self.vorrat = vorrat
        self.winkel = winkel
        self.soll = soll
        quer = querschnitt(doc, element)
        start, ende = _linie(element)
        self.nachbarn, anschluesse_ = anschluesse(element)
        self.daten = Leitungsdaten(
            kennung, element.Id,
            u"%s · %s" % (_beschreibung(fund), quer.text()),
            start, ende, quer, anschluesse_, _beschreibung(hindernis))
        self.achse = ug.einheit(ug.minus(ende, start))
        self.ausschluss = _ausschluss(doc, element)
        for anschluss in anschluesse_:
            if anschluss is not None and anschluss.nachbar_id is not None:
                self.ausschluss.append(anschluss.nachbar_id)

    def rechne(self, abstand, frist, nur_richtungen=None,
               fester_versatz=None, zusatz_ausschluss=()):
        d = self.daten
        runde = ug.varianten(d.start, d.ende, d.querschnitt,
                             self.hindernis.kasten, abstand=abstand,
                             winkel=self.winkel, nachbarn=self.nachbarn,
                             nur_richtungen=nur_richtungen,
                             fester_versatz=fester_versatz,
                             soll_abstand=self.soll)
        for variante in runde:
            variante.leitung = d
        paare = [(v, koerper(v, d.querschnitt, self.achse))
                 for v in runde if v.gueltig]
        fertig = pruefe_varianten(self.vorrat, paare,
                                  self.ausschluss + list(zusatz_ausschluss),
                                  frist)
        return runde, fertig


class _Paar(object):
    """Anzeige-Angaben für eine Kombination (wie Leitungsdaten)."""

    def __init__(self, rechner):
        self.kennung = u"A+B"
        self.text = u" / ".join(r.daten.text for r in rechner)


def _beide(rechner, soll, frist):
    """Beide weichen je zur Hälfte aus: A nach unten und B nach oben bzw.
    umgekehrt. Jede Hälfte wird ohne den Partner geprüft - der ist ja
    selbst unterwegs; auseinander liegen sie durch die Summe der Versätze.
    """
    a, b = rechner
    ergebnis = []
    fertig = True
    if abs(a.achse[2]) >= 0.9 or abs(b.achse[2]) >= 0.9:
        return ergebnis, fertig            # Steigleitungen: nicht sinnvoll
    for richtung_a, richtung_b in ug.GEGENRICHTUNGEN:
        voll = ug.voller_versatz(a.daten.start, a.daten.ende,
                                 a.daten.querschnitt, a.hindernis.kasten,
                                 soll, richtung_a)
        if voll is None:
            continue
        halb = voll / 2.0
        teile = []
        for teil, richtung, partner in ((a, richtung_a, b),
                                        (b, richtung_b, a)):
            runde, ok = teil.rechne(soll, frist, nur_richtungen=[richtung],
                                    fester_versatz=halb,
                                    zusatz_ausschluss=partner.ausschluss)
            fertig = fertig and ok
            if not runde:
                break
            teile.append(ug.beste(runde, 1)[0])
        if len(teile) == 2:
            kombination = ug.Kombination(teile)
            kombination.leitung = _Paar(rechner)
            ergebnis.append(kombination)
    return ergebnis, fertig


def auswahl(je_leitung, kombinationen=()):
    """Die Karten: bei einer Leitung die besten 3, bei zweien die besten 2
    je Leitung - damit man immer auch das andere Element verändern kann -
    und dazu die beste Variante "beide weichen aus"."""
    if len(je_leitung) == 1:
        return ug.beste(je_leitung[0], 3)
    gewaehlt = []
    for alle in je_leitung:
        gewaehlt.extend(ug.beste(alle, 2))
    if kombinationen:
        gewaehlt.extend(ug.beste(list(kombinationen), 1))
    return sorted(gewaehlt, key=lambda v: v.sortierschluessel())


# --- Vorschau ----------------------------------------------------------------

def _vollfuellung(doc):
    for muster in FilteredElementCollector(doc).OfClass(FillPatternElement):
        try:
            if muster.GetFillPattern().IsSolidFill:
                return muster.Id
        except Exception:
            continue
    return None


def _vorschau_ids(doc):
    ids = []
    for form in FilteredElementCollector(doc).OfClass(DirectShape):
        try:
            if form.ApplicationId == KENNUNG:
                ids.append(form.Id)
        except Exception:
            continue
    return ids


def _loesche_vorschau(doc):
    ids = _vorschau_ids(doc)
    if ids:
        doc.Delete(_id_liste(ids))
    return len(ids)


def entferne_vorschau(uiapp):
    uidoc = uiapp.ActiveUIDocument
    if uidoc is None:
        return 0
    doc = uidoc.Document
    if not _vorschau_ids(doc):
        return 0
    transaktion = Transaction(doc, t(u"Clash-Vorschau entfernen",
                                     u"Remove clash preview",
                                     u"Quitar vista previa"))
    transaktion.Start()
    try:
        anzahl = _loesche_vorschau(doc)
        transaktion.Commit()
        return anzahl
    except Exception:
        if transaktion.HasStarted():
            transaktion.RollBack()
        raise


def zeige_vorschau(uiapp, analyse, variante):
    """Die Variante als halbdurchsichtige DirectShape - grün, wenn sie
    kollisionsfrei ist, sonst orange."""
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
    _pruefe_dokument(doc, analyse)
    koerper_liste = []
    for teil in getattr(variante, "teile", [variante]):
        leitung = teil.leitung
        achse = ug.einheit(ug.minus(leitung.ende, leitung.start))
        koerper_liste.extend(koerper(teil, leitung.querschnitt, achse))
    ansicht = uidoc.ActiveView

    transaktion = Transaction(doc, t(u"Clash-Vorschau", u"Clash preview",
                                     u"Vista previa de conflicto"))
    transaktion.Start()
    try:
        _loesche_vorschau(doc)
        form = DirectShape.CreateElement(
            doc, ElementId(BuiltInCategory.OST_GenericModel))
        form.ApplicationId = KENNUNG
        form.ApplicationDataId = variante.name
        geometrie = List[GeometryObject]()
        for koerper_ in koerper_liste:
            geometrie.Add(koerper_)
        form.SetShape(geometrie)
        try:
            form.Name = u"pyMLG Vorschau %s" % variante.name
        except Exception:
            pass

        einstellung = OverrideGraphicSettings()
        farbe = (Color(39, 174, 96) if not variante.kollisionen
                 else Color(230, 126, 34))
        muster = _vollfuellung(doc)
        if muster is not None:
            einstellung.SetSurfaceForegroundPatternId(muster)
            einstellung.SetSurfaceForegroundPatternColor(farbe)
        einstellung.SetProjectionLineColor(farbe)
        einstellung.SetSurfaceTransparency(35)
        try:
            ansicht.SetElementOverrides(form.Id, einstellung)
        except Exception:
            pass
        transaktion.Commit()
    except Exception:
        if transaktion.HasStarted():
            transaktion.RollBack()
        raise


# --- Übernehmen --------------------------------------------------------------

class _Warnungen(IFailuresPreprocessor):
    """Warnungen (z.B. "leicht versetzt") still übergehen - Fehler bleiben
    und führen zum Zurückrollen."""

    def PreprocessFailures(self, zugriff):
        for meldung in zugriff.GetFailureMessages():
            if meldung.GetSeverity() == FailureSeverity.Warning:
                zugriff.DeleteWarning(meldung)
        return FailureProcessingResult.Continue


def _pruefe_dokument(doc, analyse):
    if doc.Title != analyse.doc_titel:
        raise rv.Fehler(t(u"Das aktive Modell hat gewechselt - bitte neu "
                          u"berechnen.",
                          u"The active model has changed - please "
                          u"recalculate.",
                          u"El modelo activo ha cambiado; vuelva a "
                          u"calcular."))


def _brich(doc, element_id, punkt):
    element = doc.GetElement(element_id)
    if isinstance(element, Pipe):
        return PlumbingUtils.BreakCurve(doc, element_id, punkt)
    return MechanicalUtils.BreakCurve(doc, element_id, punkt)


def _enthaelt(element, punkt):
    """Liegt der Punkt (Fuss) auf dem Stück (mit etwas Spiel)?"""
    kurve = element.Location.Curve
    return kurve.Distance(punkt) < 1e-3


def _verbinder_bei(element, punkt):
    beste = None
    for verbinder in _verbinder(element):
        abstand = verbinder.Origin.DistanceTo(punkt)
        if beste is None or abstand < beste[0]:
            beste = (abstand, verbinder)
    return beste[1] if beste else None


def _kopie(doc, vorlage_id, von, bis):
    """Kopie der Leitung mit neuer Lage - behält Typ, System, Grösse und
    Parameter."""
    kopien = ElementTransformUtils.CopyElement(doc, vorlage_id,
                                               XYZ(0.0, 0.0, 10.0))
    for kopie_id in kopien:
        kopie = doc.GetElement(kopie_id)
        if _ist_leitung(kopie):
            kopie.Location.Curve = Line.CreateBound(von, bis)
            return kopie
    raise rv.Fehler(t(u"Die Leitung ließ sich nicht kopieren.",
                      u"The segment could not be copied.",
                      u"No se pudo copiar el tramo."))


def _uebertrage_daemmung(doc, vorlage_daemmung, ziele):
    if vorlage_daemmung is None:
        return
    for ziel in ziele:
        try:
            if list(InsulationLiningBase.GetInsulationIds(doc, ziel.Id)):
                continue
            if isinstance(vorlage_daemmung, PipeInsulation):
                PipeInsulation.Create(doc, ziel.Id,
                                      vorlage_daemmung.GetTypeId(),
                                      vorlage_daemmung.Thickness)
            else:
                DuctInsulation.Create(doc, ziel.Id,
                                      vorlage_daemmung.GetTypeId(),
                                      vorlage_daemmung.Thickness)
        except Exception:
            continue


def uebernehme(uiapp, analyse, variante):
    """Die Variante modellieren. Rückgabe: Zahl der neuen bzw. bewegten
    Elemente. Eine Kombination ändert beide Leitungen - in einer
    Transaktionsgruppe, Strg+Z nimmt beide zurück."""
    teile = getattr(variante, "teile", None)
    if not teile:
        return _uebernehme_einzeln(uiapp, analyse, variante)
    doc = uiapp.ActiveUIDocument.Document
    gruppe = TransactionGroup(doc, t(u"Kollision lösen: %s",
                                     u"Resolve clash: %s",
                                     u"Resolver conflicto: %s")
                              % variante.name)
    gruppe.Start()
    try:
        anzahl = 0
        for teil in teile:
            anzahl += _uebernehme_einzeln(uiapp, analyse, teil)
        gruppe.Assimilate()
        return anzahl
    except Exception:
        if gruppe.HasStarted():
            gruppe.RollBack()
        raise


def _uebernehme_einzeln(uiapp, analyse, variante):
    """Eine Variante für eine Leitung modellieren."""
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document
    _pruefe_dokument(doc, analyse)
    daten = variante.leitung
    leitung = doc.GetElement(daten.element_id)
    if leitung is None or not _ist_leitung(leitung):
        raise rv.Fehler(t(u"Die Leitung gibt es nicht mehr - bitte neu "
                          u"berechnen.",
                          u"The segment no longer exists - please "
                          u"recalculate.",
                          u"El tramo ya no existe; vuelva a calcular."))
    start, ende = _linie(leitung)
    if max(ug.laenge(ug.minus(start, daten.start)),
           ug.laenge(ug.minus(ende, daten.ende))) > 1e-4:
        raise rv.Fehler(t(u"Die Leitung wurde inzwischen verändert - bitte "
                          u"neu berechnen.",
                          u"The segment has been changed - please "
                          u"recalculate.",
                          u"El tramo ha cambiado; vuelva a calcular."))

    if variante.art == ug.VERSCHIEBEN:
        return _verschiebe(uidoc, doc, leitung, analyse, variante)

    p1, p1o, p2o, p2 = [_fuss(p) for p in variante.punkte]
    daemmung, _dicke = _daemmung(doc, leitung)

    transaktion = Transaction(doc, t(u"Kollision lösen: %s",
                                     u"Resolve clash: %s",
                                     u"Resolver conflicto: %s")
                              % variante.name)
    optionen = transaktion.GetFailureHandlingOptions()
    optionen.SetFailuresPreprocessor(_Warnungen())
    transaktion.SetFailureHandlingOptions(optionen)
    transaktion.Start()
    try:
        _loesche_vorschau(doc)

        # Zweimal auftrennen: erst bei P2, dann das Stück, das P1 enthält
        teile = [leitung.Id, _brich(doc, leitung.Id, p2)]
        mit_p1 = [i for i in teile if _enthaelt(doc.GetElement(i), p1)]
        teile.append(_brich(doc, mit_p1[0], p1))

        # Das Mittelstück liegt zwischen P1 und P2 - es wird ersetzt
        mitte = (p1 + p2) / 2.0
        mittelstueck = [i for i in teile
                        if _enthaelt(doc.GetElement(i), mitte)][0]
        vorne = [i for i in teile if i != mittelstueck
                 and _enthaelt(doc.GetElement(i), p1)][0]
        hinten = [i for i in teile if i != mittelstueck
                  and _enthaelt(doc.GetElement(i), p2)][0]

        neu = [_kopie(doc, mittelstueck, von, bis)
               for von, bis in ((p1, p1o), (p1o, p2o), (p2o, p2))]
        doc.Delete(mittelstueck)
        doc.Regenerate()

        kette = [doc.GetElement(vorne)] + neu + [doc.GetElement(hinten)]
        knoten = [p1, p1o, p2o, p2]
        boegen = []
        for nummer, punkt in enumerate(knoten):
            a = _verbinder_bei(kette[nummer], punkt)
            b = _verbinder_bei(kette[nummer + 1], punkt)
            try:
                boegen.append(doc.Create.NewElbowFitting(a, b))
            except Exception as fehler:
                raise rv.Fehler(t(
                    u"Bogen %d ließ sich nicht setzen: %s\nSind im "
                    u"Rohr-/Kanaltyp Bögen in den Routing-Einstellungen "
                    u"hinterlegt?",
                    u"Elbow %d could not be placed: %s\nDoes the pipe/duct "
                    u"type have elbows in its routing preferences?",
                    u"No se pudo colocar el codo %d: %s\n¿Tiene el tipo de "
                    u"tubería/conducto codos en sus preferencias de "
                    u"trazado?") % (nummer + 1, fehler))

        _uebertrage_daemmung(doc, daemmung, neu + boegen)
        transaktion.Commit()
    except Exception:
        if transaktion.HasStarted():
            transaktion.RollBack()
        raise

    try:
        ids = List[ElementId]()
        for element in neu + boegen:
            ids.Add(element.Id)
        uidoc.Selection.SetElementIds(ids)
    except Exception:
        pass
    return len(neu) + len(boegen)


def _verschiebe(uidoc, doc, leitung, analyse, variante):
    """Ganzes Stück samt Bögen an den Enden versetzen; die Leitungen hinter
    den Bögen werden bis an den versetzten Bogen verlängert bzw. gekürzt."""
    vektor = _fuss(variante.punkte[0]) - leitung.Location.Curve.GetEndPoint(0)
    anschluesse_ = [a for a in variante.leitung.anschluesse
                    if a is not None]
    ids = [leitung.Id] + [a.formteil_id for a in anschluesse_]

    transaktion = Transaction(doc, t(u"Kollision lösen: %s",
                                     u"Resolve clash: %s",
                                     u"Resolver conflicto: %s")
                              % variante.name)
    optionen = transaktion.GetFailureHandlingOptions()
    optionen.SetFailuresPreprocessor(_Warnungen())
    transaktion.SetFailureHandlingOptions(optionen)
    transaktion.Start()
    try:
        _loesche_vorschau(doc)
        ElementTransformUtils.MoveElements(doc, _id_liste(ids), vektor)
        doc.Regenerate()
        for anschluss in anschluesse_:
            if anschluss.nachbar_id is None:
                continue
            formteil = doc.GetElement(anschluss.formteil_id)
            nachbar = doc.GetElement(anschluss.nachbar_id)
            ziel = anschluss.punkt + vektor
            bogen_verbinder = _verbinder_bei(formteil, ziel)
            ende_neu = bogen_verbinder.Origin
            eigener = min(_verbinder(nachbar), key=lambda v: min(
                v.Origin.DistanceTo(anschluss.punkt),
                v.Origin.DistanceTo(ende_neu)))
            if eigener.Origin.DistanceTo(ende_neu) > 1e-6:
                # Revit hat den Nachbarn nicht mitgezogen: Ende nachführen
                kurve = nachbar.Location.Curve
                a, b = kurve.GetEndPoint(0), kurve.GetEndPoint(1)
                if a.DistanceTo(anschluss.punkt) < b.DistanceTo(
                        anschluss.punkt):
                    nachbar.Location.Curve = Line.CreateBound(ende_neu, b)
                else:
                    nachbar.Location.Curve = Line.CreateBound(a, ende_neu)
                doc.Regenerate()
                eigener = _verbinder_bei(nachbar, ende_neu)
            if not eigener.IsConnectedTo(bogen_verbinder):
                eigener.ConnectTo(bogen_verbinder)
        transaktion.Commit()
    except Exception:
        if transaktion.HasStarted():
            transaktion.RollBack()
        raise

    try:
        uidoc.Selection.SetElementIds(_id_liste(ids))
    except Exception:
        pass
    return len(ids)
