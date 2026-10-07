# -*- coding: utf-8 -*-
"""Revit-Teil von DxfLegend: Legendenplan (Papier-mm) -> Revit-Ansicht.

Eine Legendenansicht lässt sich per API nicht neu anlegen - eine
vorhandene Legende wird ohne Detaillierung dupliziert. Gibt es keine,
wird eine Zeichnungsansicht erzeugt.

Angelegt (bzw. wiederverwendet, wenn Name und Eigenschaften passen):
    Linienmuster       LinePatternElement (Längen in Papiermass)
    Linienstile        Unterkategorien von <Linien> mit Farbe/Stift/Muster
    Texttypen          Kopie des Vorgabe-Texttyps: Höhe, Schrift, Farbe,
                       Breitenfaktor, transparent, ohne Rahmenabstand
    Füllmuster         Zeichnungs-Füllmuster aus den Musterlinien
    Füllbereichstypen  Vordergrund Muster + Farbe, kein Hintergrund

Gleicher Name mit anderen Eigenschaften:
    Vorgabe            neuer Name mit _2, _3 … - Vorhandenes bleibt, wie es ist
    ueberschreiben     das vorhandene Element bekommt die neuen Eigenschaften
                       (wirkt überall, wo es benutzt wird)

ersetzen: Gibt es schon eine Legende/Zeichnungsansicht mit dem Namen, wird
ihr Inhalt gelöscht und neu gezeichnet - auf Plänen bleibt sie stehen.
Sonst entsteht eine neue Ansicht "Name (2)".

Alles in einer Transaktion -> ein Rückgängig-Schritt.
"""

import io
import math
import os
import traceback

from Autodesk.Revit.DB import (
    Arc,
    BuiltInCategory,
    BuiltInParameter,
    Color,
    CurveLoop,
    ElementId,
    ElementTypeGroup,
    FillGrid,
    FillPattern,
    FillPatternElement,
    FillPatternHostOrientation,
    FillPatternTarget,
    FilledRegion,
    FilledRegionType,
    FilteredElementCollector,
    GraphicsStyleType,
    HorizontalTextAlignment,
    Line,
    LinePattern,
    LinePatternElement,
    LinePatternSegment,
    LinePatternSegmentType,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    Transaction,
    TransactionStatus,
    UV,
    VerticalTextAlignment,
    View,
    ViewDrafting,
    ViewDuplicateOption,
    ViewFamily,
    ViewFamilyType,
    ViewType,
    XYZ,
)
from System import Byte, Double
from System.Collections.Generic import List

from dxf_legende import geometrie as geo
from dxf_legende import logik as lg
from mlg_sprache import t

MM = 1.0 / 304.8

# OST_InvisibleLines, falls der Name in der API fehlt
_UNSICHTBAR_ID = -2000064

PROTOKOLL = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "pyMLG",
    "DxfLegend_Protokoll.txt")


class LegendenFehler(Exception):
    """Fachlicher Fehler - nur als Text anzeigen."""


def id_wert(element_id):
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def _name(element):
    try:
        return element.Name
    except Exception:
        parameter = element.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        return parameter.AsString() if parameter else u""


def legenden(doc):
    """Legendenansichten des Projekts (ohne Vorlagen), nach Name."""
    ansichten = [v for v in FilteredElementCollector(doc).OfClass(View)
                 if v.ViewType == ViewType.Legend and not v.IsTemplate]
    return sorted(ansichten, key=lambda v: _name(v).lower())


def vorhandene_ansicht(doc, name):
    """Legende oder Zeichnungsansicht mit diesem Namen (ohne Vorlagen)."""
    gesucht = lg.revit_name(name, 200).lower()
    for ansicht in FilteredElementCollector(doc).OfClass(View):
        if ansicht.IsTemplate or ansicht.ViewType not in (ViewType.Legend,
                                                          ViewType.DraftingView):
            continue
        if _name(ansicht).lower() == gesucht:
            return ansicht
    return None


def _revit_farbe(rgb):
    return Color(Byte(int(rgb[0])), Byte(int(rgb[1])), Byte(int(rgb[2])))


def _gleiche_farbe(farbe, rgb):
    try:
        return farbe.IsValid and (farbe.Red, farbe.Green, farbe.Blue) == tuple(rgb)
    except Exception:
        return False


def _farbwert(rgb):
    """RGB -> Ganzzahl für Parameter LINE_COLOR (Rot im niedrigsten Byte)."""
    return int(rgb[0]) + int(rgb[1]) * 256 + int(rgb[2]) * 65536


def _nah(a, b, tol=1e-6):
    return abs(a - b) <= tol


class Ergebnis(object):

    def __init__(self):
        self.ansicht = None
        self.linien = 0
        self.texte = 0
        self.flaechen = 0
        self.neu = {u"muster": 0, u"stile": 0, u"texttypen": 0,
                    u"fuellmuster": 0, u"fuelltypen": 0}
        self.geaendert = dict((art, 0) for art in self.neu)
        self.ersetzt = False    # vorhandene Ansicht neu gezeichnet
        self.geloescht = 0      # dabei gelöschte Elemente
        self.fehler = {}        # Art -> Anzahl
        self.meldungen = []     # erste Ausnahmen für das Protokoll

    def fehlschlag(self, art, ausnahme=None):
        self.fehler[art] = self.fehler.get(art, 0) + 1
        if ausnahme is not None and len(self.meldungen) < 50:
            self.meldungen.append(u"%s: %s" % (art, traceback.format_exc()))

    def schreibe_protokoll(self):
        if not self.meldungen:
            return None
        try:
            if not os.path.isdir(os.path.dirname(PROTOKOLL)):
                os.makedirs(os.path.dirname(PROTOKOLL))
            with io.open(PROTOKOLL, "w", encoding="utf-8") as datei:
                datei.write(u"\n".join(self.meldungen) + u"\n")
            return PROTOKOLL
        except Exception:
            return None


def erzeuge(doc, plan, name, massstab, vorlage=None, ueberschreiben=False,
            ersetzen=False):
    """Legende (Duplikat von vorlage) bzw. Zeichnungsansicht mit dem Plan
    füllen - oder mit ersetzen die gleichnamige vorhandene Ansicht.
    Rückgabe Ergebnis."""
    ergebnis = Ergebnis()
    transaktion = Transaction(doc, t(u"pyMLG Legende aus DXF",
                                     u"pyMLG Legend from DXF",
                                     u"pyMLG Leyenda desde DXF"))
    transaktion.Start()
    try:
        ansicht = vorhandene_ansicht(doc, name) if ersetzen else None
        if ansicht is not None:
            ergebnis.geloescht = _leere_ansicht(doc, ansicht)
            ergebnis.ersetzt = True
            try:
                ansicht.Scale = int(massstab)
            except Exception:
                pass
        else:
            ansicht = _neue_ansicht(doc, name, massstab, vorlage)
        bauer = _Bauer(doc, ansicht, massstab, ergebnis, ueberschreiben)
        bauer.baue(plan)
        if transaktion.Commit() != TransactionStatus.Committed:
            raise LegendenFehler(t(u"Revit hat die Legende nicht übernommen.",
                                   u"Revit did not accept the legend.",
                                   u"Revit no aceptó la leyenda."))
        ergebnis.ansicht = ansicht
    except Exception:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        raise
    return ergebnis


def _neue_ansicht(doc, name, massstab, vorlage):
    if vorlage is not None:
        if not vorlage.CanViewBeDuplicated(ViewDuplicateOption.Duplicate):
            raise LegendenFehler(t(u"Die Legende \"%s\" lässt sich nicht duplizieren.",
                                   u"The legend \"%s\" cannot be duplicated.",
                                   u"La leyenda \"%s\" no se puede duplicar.")
                                 % _name(vorlage))
        ansicht = doc.GetElement(vorlage.Duplicate(ViewDuplicateOption.Duplicate))
    else:
        typ = next((v for v in FilteredElementCollector(doc).OfClass(ViewFamilyType)
                    if v.ViewFamily == ViewFamily.Drafting), None)
        if typ is None:
            raise LegendenFehler(t(u"Im Projekt gibt es keinen Typ für Zeichnungsansichten.",
                                   u"The project has no drafting view type.",
                                   u"El proyecto no tiene tipo de vista de diseño."))
        ansicht = ViewDrafting.Create(doc, typ.Id)
    try:
        ansicht.Scale = int(massstab)
    except Exception:
        pass
    vergeben = set(_name(v).lower() for v in FilteredElementCollector(doc).OfClass(View)
                   if v.Id != ansicht.Id)
    wunsch = lg.revit_name(name, 200)
    neuer_name, nummer = wunsch, 2
    while neuer_name.lower() in vergeben:
        neuer_name = u"%s (%d)" % (wunsch, nummer)
        nummer += 1
    ansicht.Name = neuer_name
    return ansicht


def _leere_ansicht(doc, ansicht):
    """Alle ansichtseigenen Elemente löschen (Linien, Texte, Füllbereiche,
    Legendenbauteile …). Rückgabe Anzahl."""
    ids = [e.Id for e in FilteredElementCollector(doc, ansicht.Id)
           .WhereElementIsNotElementType()
           if e.OwnerViewId == ansicht.Id and e.Category is not None]
    geloescht = 0
    for element_id in ids:
        # Abhängige Elemente können schon mit einem anderen verschwunden sein
        if doc.GetElement(element_id) is None:
            continue
        try:
            doc.Delete(element_id)
            geloescht += 1
        except Exception:
            pass
    return geloescht


NEU, GLEICH, GEAENDERT = u"neu", u"gleich", u"geaendert"


def _finde_oder_erzeuge(wunsch, finden, gleich, erzeugen, anpassen=None):
    """Element mit Wunschnamen wiederverwenden, wenn es passt. Passt es
    nicht: mit anpassen überschreiben, ohne unter freiem Namen (_2 …)
    anlegen. Rückgabe (element, NEU|GLEICH|GEAENDERT)."""
    name = wunsch
    for nummer in range(2, 1000):
        element = finden(name)
        if element is None:
            return erzeugen(name), NEU
        if gleich(element):
            return element, GLEICH
        if anpassen is not None:
            anpassen(element)
            return element, GEAENDERT
        name = u"%s_%d" % (wunsch, nummer)
    raise LegendenFehler(t(u"Kein freier Name für \"%s\" gefunden.",
                           u"No free name found for \"%s\".",
                           u"No se encontró un nombre libre para \"%s\".") % wunsch)


class _Bauer(object):

    def __init__(self, doc, ansicht, massstab, ergebnis, ueberschreiben=False):
        self.doc = doc
        self.ueberschreiben = ueberschreiben
        self.ansicht = ansicht
        self.massstab = float(massstab)
        self.ergebnis = ergebnis
        self.kurz = doc.Application.ShortCurveTolerance
        self.linien_kat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
        self._muster = {}
        self._stile = {}
        self._texttypen = {}
        self._fuellmuster = {}
        self._fuelltypen = {}
        self._vollton = None
        self._unsichtbar = None

    def _vorhanden(self, art, wunsch, finden, gleich, erzeugen, anpassen):
        """_finde_oder_erzeuge mit Zählung; anpassen nur beim Überschreiben."""
        element, zustand = _finde_oder_erzeuge(
            wunsch, finden, gleich, erzeugen,
            anpassen if self.ueberschreiben else None)
        if zustand == NEU:
            self.ergebnis.neu[art] += 1
        elif zustand == GEAENDERT:
            self.ergebnis.geaendert[art] += 1
        return element

    def baue(self, plan):
        # Füllungen zuerst, damit Linien und Texte darüber liegen
        for typ, schleifen in plan.flaechen:
            self._flaeche(plan.fuelltypen[typ], schleifen)
        for stil, punkte, geschlossen in plan.linien:
            self._linie(plan.stile[stil], punkte, geschlossen)
        for eintrag in plan.texte:
            self._text(plan.texttypen[eintrag.typ], eintrag)

    # -- Geometrie ---------------------------------------------------------------

    def _xyz(self, p):
        faktor = MM * self.massstab
        return XYZ(p[0] * faktor, p[1] * faktor, 0.0)

    def _kurven(self, punkte, geschlossen):
        """Polylinie (mm, mit bulge) -> Revit-Kurven; zu kurze fallen weg."""
        if geschlossen and len(punkte) == 2 and \
                abs(abs(punkte[0][2]) - 1.0) < 1e-9 and abs(abs(punkte[1][2]) - 1.0) < 1e-9 \
                and punkte[0][2] == punkte[1][2]:
            # Vollkreis aus zwei Halbbögen
            a, b = self._xyz(punkte[0]), self._xyz(punkte[1])
            mitte = a.Add(b).Divide(2.0)
            radius = a.DistanceTo(b) / 2.0
            if radius > self.kurz:
                return [Arc.Create(mitte, radius, 0.0, 2.0 * math.pi,
                                   XYZ.BasisX, XYZ.BasisY)]
        kurven = []
        for p1, p2, bulge in geo.abschnitte(punkte, geschlossen):
            a, b = self._xyz(p1), self._xyz(p2)
            if a.DistanceTo(b) < self.kurz:
                continue
            if abs(bulge) > 1e-9:
                kurven.append(Arc.Create(a, b, self._xyz(geo.bogenmitte(p1, p2, bulge))))
            else:
                kurven.append(Line.CreateBound(a, b))
        return kurven

    # -- Linien ------------------------------------------------------------------

    def _linienmuster(self, stil):
        if not stil.muster:
            return LinePatternElement.GetSolidPatternId()
        if stil.muster_name in self._muster:
            return self._muster[stil.muster_name]
        arten = {u"dash": LinePatternSegmentType.Dash,
                 u"space": LinePatternSegmentType.Space,
                 u"dot": LinePatternSegmentType.Dot}

        def muster(name):
            linienmuster = LinePattern(name)
            segmente = List[LinePatternSegment]()
            for art, mm in stil.muster:
                segmente.Add(LinePatternSegment(arten[art], mm * MM))
            linienmuster.SetSegments(segmente)
            return linienmuster

        def gleich(element):
            vorhanden = list(element.GetLinePattern().GetSegments())
            if len(vorhanden) != len(stil.muster):
                return False
            return all(s.Type == arten[art] and _nah(s.Length, mm * MM, 1e-5)
                       for s, (art, mm) in zip(vorhanden, stil.muster))

        element = self._vorhanden(
            u"muster", stil.muster_name,
            lambda n: LinePatternElement.GetLinePatternElementByName(self.doc, n),
            gleich,
            lambda n: LinePatternElement.Create(self.doc, muster(n)),
            lambda e: e.SetLinePattern(muster(_name(e))))
        self._muster[stil.muster_name] = element.Id
        return element.Id

    def _linienstil(self, stil):
        if stil.name in self._stile:
            return self._stile[stil.name]
        muster_id = self._linienmuster(stil)
        projektion = GraphicsStyleType.Projection

        def finden(name):
            for unter in self.linien_kat.SubCategories:
                if unter.Name.lower() == name.lower():
                    return unter
            return None

        def gleich(unter):
            return (_gleiche_farbe(unter.LineColor, stil.farbe)
                    and unter.GetLineWeight(projektion) == stil.stift
                    and unter.GetLinePatternId(projektion) == muster_id)

        def anpassen(unter):
            unter.LineColor = _revit_farbe(stil.farbe)
            unter.SetLineWeight(stil.stift, projektion)
            unter.SetLinePatternId(muster_id, projektion)

        def erzeugen(name):
            unter = self.doc.Settings.Categories.NewSubcategory(self.linien_kat, name)
            anpassen(unter)
            return unter

        unter = self._vorhanden(u"stile", stil.name, finden, gleich, erzeugen, anpassen)
        grafik = unter.GetGraphicsStyle(projektion)
        self._stile[stil.name] = grafik
        return grafik

    def _linie(self, stil, punkte, geschlossen):
        try:
            grafik = self._linienstil(stil)
        except Exception as fehler:
            self.ergebnis.fehlschlag(t(u"Linienstil", u"Line style", u"Estilo de línea"), fehler)
            return
        for kurve in self._kurven(punkte, geschlossen):
            try:
                detail = self.doc.Create.NewDetailCurve(self.ansicht, kurve)
                detail.LineStyle = grafik
                self.ergebnis.linien += 1
            except Exception as fehler:
                self.ergebnis.fehlschlag(t(u"Linie", u"Line", u"Línea"), fehler)

    # -- Texte -------------------------------------------------------------------

    def _texttyp(self, typ):
        if typ.name in self._texttypen:
            return self._texttypen[typ.name]

        def finden(name):
            for element in FilteredElementCollector(self.doc).OfClass(TextNoteType):
                if _name(element).lower() == name.lower():
                    return element
            return None

        def wert(element, parameter, art):
            p = element.get_Parameter(parameter)
            if p is None:
                return None
            return p.AsDouble() if art == u"d" else (
                p.AsInteger() if art == u"i" else p.AsString())

        def gleich(element):
            return (_nah(wert(element, BuiltInParameter.TEXT_SIZE, u"d") or 0.0,
                         typ.hoehe_mm * MM, 1e-6)
                    and (wert(element, BuiltInParameter.TEXT_FONT, u"s") or u"").lower()
                    == typ.schrift.lower()
                    and _nah(wert(element, BuiltInParameter.TEXT_WIDTH_SCALE, u"d") or 1.0,
                             typ.breitenfaktor, 1e-3)
                    and wert(element, BuiltInParameter.LINE_COLOR, u"i") == _farbwert(typ.farbe))

        def anpassen(element):
            for parameter, inhalt in (
                    (BuiltInParameter.TEXT_SIZE, typ.hoehe_mm * MM),
                    (BuiltInParameter.TEXT_FONT, typ.schrift),
                    (BuiltInParameter.TEXT_WIDTH_SCALE, float(typ.breitenfaktor)),
                    (BuiltInParameter.LINE_COLOR, _farbwert(typ.farbe)),
                    (BuiltInParameter.TEXT_BACKGROUND, 1),          # transparent
                    (BuiltInParameter.LEADER_OFFSET_SHEET, 0.0),
                    (BuiltInParameter.TEXT_BOX_VISIBILITY, 0),
                    (BuiltInParameter.TEXT_STYLE_BOLD, 0),
                    (BuiltInParameter.TEXT_STYLE_ITALIC, 0),
                    (BuiltInParameter.TEXT_STYLE_UNDERLINE, 0)):
                p = element.get_Parameter(parameter)
                if p is not None and not p.IsReadOnly:
                    try:
                        p.Set(inhalt)
                    except Exception:
                        pass

        def erzeugen(name):
            basis = self.doc.GetElement(
                self.doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType))
            if basis is None:
                basis = FilteredElementCollector(self.doc).OfClass(TextNoteType).FirstElement()
            neu = basis.Duplicate(name)
            anpassen(neu)
            return neu

        element = self._vorhanden(u"texttypen", typ.name, finden, gleich, erzeugen, anpassen)
        self._texttypen[typ.name] = element.Id
        return element.Id

    def _text(self, typ, eintrag):
        try:
            optionen = TextNoteOptions(self._texttyp(typ))
            optionen.HorizontalAlignment = {
                u"links": HorizontalTextAlignment.Left,
                u"mitte": HorizontalTextAlignment.Center,
                u"rechts": HorizontalTextAlignment.Right}[eintrag.h_ausr]
            optionen.VerticalAlignment = {
                u"oben": VerticalTextAlignment.Top,
                u"mitte": VerticalTextAlignment.Middle,
                u"unten": VerticalTextAlignment.Bottom}[eintrag.v_ausr]
            optionen.Rotation = eintrag.drehung
            optionen.KeepRotatedTextReadable = False
            TextNote.Create(self.doc, self.ansicht.Id, self._xyz((eintrag.x, eintrag.y)),
                            eintrag.inhalt.replace(u"\n", u"\r"), optionen)
            self.ergebnis.texte += 1
        except Exception as fehler:
            self.ergebnis.fehlschlag(t(u"Text", u"Text", u"Texto"), fehler)

    # -- Füllbereiche ------------------------------------------------------------

    def _vollton_id(self):
        if self._vollton is None:
            for element in FilteredElementCollector(self.doc).OfClass(FillPatternElement):
                muster = element.GetFillPattern()
                if muster.IsSolidFill:
                    self._vollton = element.Id
                    break
        if self._vollton is None:
            raise LegendenFehler(t(u"Im Projekt fehlt das Füllmuster <Vollton>.",
                                   u"The project has no <Solid fill> pattern.",
                                   u"Falta el patrón <Relleno sólido> en el proyecto."))
        return self._vollton

    def _fuellmuster_id(self, typ):
        if typ.muster_name is None:
            return self._vollton_id()
        if typ.muster_name in self._fuellmuster:
            return self._fuellmuster[typ.muster_name]

        def muster(name):
            fuellmuster = FillPattern(name, FillPatternTarget.Drafting,
                                      FillPatternHostOrientation.ToView)
            gitterliste = List[FillGrid]()
            for winkel, ox, oy, abstand, versatz, segmente in typ.gitter:
                gitter = FillGrid()
                gitter.Angle = winkel
                gitter.Origin = UV(ox * MM, oy * MM)
                gitter.Offset = abstand * MM
                gitter.Shift = versatz * MM
                if segmente:
                    laengen = List[Double]()
                    for laenge in segmente:
                        laengen.Add(laenge * MM)
                    gitter.SetSegments(laengen)
                gitterliste.Add(gitter)
            fuellmuster.SetFillGrids(gitterliste)
            return fuellmuster

        def gleich(element):
            gitter = list(element.GetFillPattern().GetFillGrids())
            if len(gitter) != len(typ.gitter):
                return False
            return all(_nah(g.Angle, w, 1e-4) and _nah(g.Offset, a * MM, 1e-5)
                       and _nah(g.Shift, v * MM, 1e-5)
                       for g, (w, _x, _y, a, v, _s) in zip(gitter, typ.gitter))

        element = self._vorhanden(
            u"fuellmuster", typ.muster_name,
            lambda n: FillPatternElement.GetFillPatternElementByName(
                self.doc, FillPatternTarget.Drafting, n),
            gleich,
            lambda n: FillPatternElement.Create(self.doc, muster(n)),
            lambda e: e.SetFillPattern(muster(_name(e))))
        self._fuellmuster[typ.muster_name] = element.Id
        return element.Id

    def _fuelltyp(self, typ):
        if typ.name in self._fuelltypen:
            return self._fuelltypen[typ.name]
        muster_id = self._fuellmuster_id(typ)

        def finden(name):
            for element in FilteredElementCollector(self.doc).OfClass(FilledRegionType):
                if _name(element).lower() == name.lower():
                    return element
            return None

        def gleich(element):
            return (element.ForegroundPatternId == muster_id
                    and _gleiche_farbe(element.ForegroundPatternColor, typ.farbe)
                    and element.BackgroundPatternId == ElementId.InvalidElementId)

        def anpassen(element):
            element.ForegroundPatternId = muster_id
            element.ForegroundPatternColor = _revit_farbe(typ.farbe)
            element.BackgroundPatternId = ElementId.InvalidElementId
            element.IsMasking = False

        def erzeugen(name):
            basis = FilteredElementCollector(self.doc).OfClass(FilledRegionType).FirstElement()
            if basis is None:
                raise LegendenFehler(t(u"Im Projekt gibt es keinen Füllbereichstyp.",
                                       u"The project has no filled region type.",
                                       u"El proyecto no tiene tipo de región rellenada."))
            neu = basis.Duplicate(name)
            anpassen(neu)
            return neu

        element = self._vorhanden(u"fuelltypen", typ.name, finden, gleich, erzeugen, anpassen)
        self._fuelltypen[typ.name] = element.Id
        return element.Id

    def _unsichtbare_linien(self):
        if self._unsichtbar is None:
            ziel = _UNSICHTBAR_ID
            try:
                ziel = int(BuiltInCategory.OST_InvisibleLines)
            except Exception:
                pass
            self._unsichtbar = ElementId.InvalidElementId
            for stil_id in FilledRegion.GetValidLineStyleIdsForFilledRegion(self.doc):
                stil = self.doc.GetElement(stil_id)
                kategorie = getattr(stil, "GraphicsStyleCategory", None)
                if kategorie is not None and id_wert(kategorie.Id) == ziel:
                    self._unsichtbar = stil_id
                    break
        return self._unsichtbar

    def _schleife(self, punkte):
        schleife = CurveLoop()
        for kurve in self._kurven(punkte, True):
            schleife.Append(kurve)
        return schleife

    def _fuellbereich(self, typ_id, schleifen):
        liste = List[CurveLoop]()
        for punkte in schleifen:
            liste.Add(self._schleife(punkte))
        bereich = FilledRegion.Create(self.doc, typ_id, self.ansicht.Id, liste)
        unsichtbar = self._unsichtbare_linien()
        if unsichtbar != ElementId.InvalidElementId:
            bereich.SetLineStyleId(unsichtbar)
        self.ergebnis.flaechen += 1

    def _flaeche(self, typ, schleifen):
        art = t(u"Füllbereich", u"Filled region", u"Región rellenada")
        try:
            typ_id = self._fuelltyp(typ)
        except Exception as fehler:
            self.ergebnis.fehlschlag(art, fehler)
            return
        try:
            self._fuellbereich(typ_id, schleifen)
            return
        except Exception as fehler:
            if len(schleifen) == 1:
                self.ergebnis.fehlschlag(art, fehler)
                return
        # Schleifen mit Inseln, die Revit ablehnt: jede einzeln versuchen
        for punkte in schleifen:
            try:
                self._fuellbereich(typ_id, [punkte])
            except Exception as fehler:
                self.ergebnis.fehlschlag(art, fehler)
