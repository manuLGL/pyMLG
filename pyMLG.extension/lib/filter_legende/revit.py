# -*- coding: utf-8 -*-
"""Revit-Teil von FilterLegend.

Lesen: Filter einer Ansicht mit ihren Grafik-Überschreibungen. Steuert
eine Ansichtsvorlage die Filter (V/G "Filter" in der Vorlage angehakt),
gelten die der Vorlage - genau das, was in der Ansicht zu sehen ist.

Erzeugen:
    1. vorhandene Legende ohne Inhalt duplizieren, gibt es keine:
       Zeichnungsansicht (wie DxfLegend) - oder mit "ersetzen" die
       gleichnamige Ansicht leeren
    2. je Filter einen Füllbereichstyp "Filterlegende - <Filter>" mit
       Vorder-/Hintergrundmuster und Farbe aus der Überschreibung.
       Gibt es den Namen schon mit gleichen Werten: wiederverwenden, mit
       anderen Werten: neuer Name _2, _3 … (andere Legenden behalten ihre
       Farben)
    3. je Filter ein rechteckiger Füllbereich, Text rechts daneben
Alles in einer Transaktion -> ein Rückgängig-Schritt.
"""

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    Color,
    CurveLoop,
    ElementId,
    ElementTypeGroup,
    FilledRegion,
    FilledRegionType,
    FillPatternElement,
    FillPatternTarget,
    FilteredElementCollector,
    HorizontalTextAlignment,
    Line,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    Transaction,
    TransactionStatus,
    VerticalTextAlignment,
    View,
    ViewType,
    XYZ,
)
from System import Byte
from System.Collections.Generic import List

from dxf_legende import logik as dxf_lg
from dxf_legende import revit as dxf_rv
from filter_legende import logik as lg
from mlg_sprache import t

MM = 1.0 / 304.8
TEXTABSTAND_MM = 3.0        # Farbfeld bis Text, Papier-mm
TYP_PRAEFIX = u"Filterlegende - "

LegendenFehler = dxf_rv.LegendenFehler
id_wert = dxf_rv.id_wert
name_von = dxf_rv._name

ANSICHTSARTEN = {
    ViewType.FloorPlan: (u"Grundriss", u"Floor plan", u"Planta"),
    ViewType.CeilingPlan: (u"Deckenplan", u"Ceiling plan", u"Plano de techo"),
    ViewType.EngineeringPlan: (u"Tragwerksplan", u"Structural plan",
                               u"Plano estructural"),
    ViewType.AreaPlan: (u"Flächenplan", u"Area plan", u"Plano de área"),
    ViewType.Section: (u"Schnitt", u"Section", u"Sección"),
    ViewType.Elevation: (u"Ansicht", u"Elevation", u"Alzado"),
    ViewType.Detail: (u"Detail", u"Detail", u"Detalle"),
    ViewType.ThreeD: (u"3D", u"3D", u"3D"),
    ViewType.DraftingView: (u"Zeichnungsansicht", u"Drafting view",
                            u"Vista de diseño"),
}
VORLAGE = (u"Ansichtsvorlage", u"View template", u"Plantilla de vista")


# -- Lesen ------------------------------------------------------------------------

def filterquelle(doc, ansicht):
    """Ansicht, deren Filter gelten: die Vorlage, wenn sie die Filter
    steuert, sonst die Ansicht selbst."""
    if ansicht.IsTemplate:
        return ansicht
    vorlage_id = ansicht.ViewTemplateId
    if vorlage_id is None or vorlage_id == ElementId.InvalidElementId:
        return ansicht
    vorlage = doc.GetElement(vorlage_id)
    if vorlage is None:
        return ansicht
    try:
        gesucht = id_wert(ElementId(BuiltInParameter.VIS_GRAPHICS_FILTERS))
        frei = set(id_wert(i) for i in vorlage.GetNonControlledTemplateParameterIds())
    except Exception:
        return vorlage
    return ansicht if gesucht in frei else vorlage


def _filter_ids(quelle):
    try:
        return list(quelle.GetOrderedFilters())     # Revit 2021+
    except Exception:
        pass
    try:
        return list(quelle.GetFilters())
    except Exception:
        return []


def quellen(doc):
    """[(ansicht, anzahl_filter, art_texte)] aller Ansichten und Vorlagen
    mit mindestens einem Filter - Ansichten nach Art und Name, Vorlagen
    zuletzt."""
    reihen = list(ANSICHTSARTEN)
    ergebnis = []
    for ansicht in FilteredElementCollector(doc).OfClass(View):
        if not ansicht.IsTemplate and ansicht.ViewType not in ANSICHTSARTEN:
            continue
        try:
            if not ansicht.AreGraphicsOverridesAllowed():
                continue
        except Exception:
            continue
        anzahl = len(_filter_ids(filterquelle(doc, ansicht)))
        if not anzahl:
            continue
        art = VORLAGE if ansicht.IsTemplate else ANSICHTSARTEN[ansicht.ViewType]
        ergebnis.append((ansicht, anzahl, art))

    def ordnung(eintrag):
        ansicht = eintrag[0]
        rang = len(reihen) + 1 if ansicht.IsTemplate else \
            reihen.index(ansicht.ViewType)
        return rang, name_von(ansicht).lower()
    return sorted(ergebnis, key=ordnung)


class _Muster(object):
    """Füllmuster des Projekts: Id-Zahl <-> ElementId, Name, Vollton."""

    def __init__(self, doc):
        self.ids = {}
        self.namen = {}
        self.vollton = set()
        self.vollton_zeichnung = None
        for element in FilteredElementCollector(doc).OfClass(FillPatternElement):
            wert = id_wert(element.Id)
            self.ids[wert] = element.Id
            self.namen[wert] = name_von(element)
            try:
                muster = element.GetFillPattern()
                if muster.IsSolidFill:
                    self.vollton.add(wert)
                    if muster.Target == FillPatternTarget.Drafting \
                            or self.vollton_zeichnung is None:
                        self.vollton_zeichnung = element.Id
            except Exception:
                pass

    def wert(self, element_id):
        """ElementId -> Zahl, VOLL oder None."""
        if element_id is None or element_id == ElementId.InvalidElementId:
            return None
        wert = id_wert(element_id)
        return lg.VOLL if wert in self.vollton else wert

    def element_id(self, wert):
        if wert is None:
            return ElementId.InvalidElementId
        if wert == lg.VOLL:
            if self.vollton_zeichnung is None:
                raise LegendenFehler(t(u"Im Projekt fehlt das Füllmuster <Vollton>.",
                                       u"The project has no <Solid fill> pattern.",
                                       u"Falta el patrón <Relleno sólido> en el proyecto."))
            return self.vollton_zeichnung
        return self.ids.get(wert, ElementId.InvalidElementId)

    def name(self, wert):
        if wert == lg.VOLL:
            return t(u"Vollton", u"Solid fill", u"Relleno sólido")
        return self.namen.get(wert, u"")


def muster(doc):
    """Füllmuster-Tabelle - je Lesevorgang neu, pyRevit kann Module
    zwischen zwei Läufen behalten."""
    return _Muster(doc)


def _farbe(farbe):
    try:
        if farbe is not None and farbe.IsValid:
            return (farbe.Red, farbe.Green, farbe.Blue)
    except Exception:
        pass
    return None


def _fuellung(tabelle, vg_sichtbar, vg_id, vg_farbe, hg_sichtbar, hg_id, hg_farbe):
    vg = (tabelle.wert(vg_id), _farbe(vg_farbe)) if vg_sichtbar else (None, None)
    hg = (tabelle.wert(hg_id), _farbe(hg_farbe)) if hg_sichtbar else (None, None)
    return lg.Fuellung(vg[0], vg[1], hg[0], hg[1])


def grafik(tabelle, ogs):
    """OverrideGraphicSettings -> logik.Grafik (tabelle = muster(doc))."""
    oberflaeche = _fuellung(
        tabelle,
        ogs.IsSurfaceForegroundPatternVisible, ogs.SurfaceForegroundPatternId,
        ogs.SurfaceForegroundPatternColor,
        ogs.IsSurfaceBackgroundPatternVisible, ogs.SurfaceBackgroundPatternId,
        ogs.SurfaceBackgroundPatternColor)
    schnitt = _fuellung(
        tabelle,
        ogs.IsCutForegroundPatternVisible, ogs.CutForegroundPatternId,
        ogs.CutForegroundPatternColor,
        ogs.IsCutBackgroundPatternVisible, ogs.CutBackgroundPatternId,
        ogs.CutBackgroundPatternColor)
    return lg.Grafik(oberflaeche, schnitt, _farbe(ogs.ProjectionLineColor),
                     _farbe(ogs.CutLineColor))


def lies_filter(doc, ansicht):
    """[logik.Eintrag] der Filter, die in der Ansicht gelten, in deren
    Reihenfolge."""
    quelle = filterquelle(doc, ansicht)
    tabelle = muster(doc)
    ergebnis = []
    for filter_id in _filter_ids(quelle):
        element = doc.GetElement(filter_id)
        if element is None:
            continue
        try:
            sichtbar = bool(quelle.GetFilterVisibility(filter_id))
        except Exception:
            sichtbar = True
        try:
            aktiv = bool(quelle.GetIsFilterEnabled(filter_id))   # Revit 2021+
        except Exception:
            aktiv = True
        eintrag = lg.Eintrag(id_wert(filter_id), filter_id, name_von(element),
                             grafik(tabelle, quelle.GetFilterOverrides(filter_id)),
                             sichtbar=sichtbar, aktiv=aktiv, ansicht=name_von(ansicht))
        eintrag.position = len(ergebnis)
        ergebnis.append(eintrag)
    return ergebnis


def texttypen(doc):
    """[(Name, Id)] aller Texttypen und die Id des Vorgabetyps."""
    typen = sorted(((name_von(e), e.Id) for e in
                    FilteredElementCollector(doc).OfClass(TextNoteType)),
                   key=lambda e: e[0].lower())
    vorgabe = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)
    return typen, vorgabe


def vorhandene_ansicht(doc, name):
    return dxf_rv.vorhandene_ansicht(doc, name)


def zielart(doc):
    """Was neu entsteht: u"legende" (Duplikat) oder u"zeichnung"."""
    return u"legende" if dxf_rv.legenden(doc) else u"zeichnung"


# -- Erzeugen ---------------------------------------------------------------------

class Ergebnis(object):

    def __init__(self):
        self.ansicht = None
        self.felder = 0
        self.texte = 0
        self.ersetzt = False
        self.geloescht = 0
        self.neue_typen = 0
        self.ohne_farbe = []        # Filternamen ohne Farbüberschreibung
        self.fehler = []            # Filternamen, deren Feld nicht entstand
        self.ersatzmuster = []      # Muster abgelehnt -> Vollton


def erzeuge(doc, posten, name, massstab, feld_breite, feld_hoehe, abstand,
            spalten=1, spaltenbreite=70.0, quelle=lg.QUELLE_AUTO, texttyp_id=None,
            rahmen=True, ersetzen=False):
    """Legende mit einem Farbfeld und Text je Posten [(Eintrag, Text)].
    Masse in Papier-mm. Rückgabe Ergebnis."""
    ergebnis = Ergebnis()
    transaktion = Transaction(doc, t(u"pyMLG Filterlegende", u"pyMLG Filter legend",
                                     u"pyMLG Leyenda de filtros"))
    transaktion.Start()
    try:
        ansicht = vorhandene_ansicht(doc, name) if ersetzen else None
        if ansicht is not None:
            ergebnis.geloescht = dxf_rv._leere_ansicht(doc, ansicht)
            ergebnis.ersetzt = True
            try:
                ansicht.Scale = int(massstab)
            except Exception:
                pass
        else:
            legenden = dxf_rv.legenden(doc)
            ansicht = dxf_rv._neue_ansicht(doc, name, massstab,
                                           legenden[0] if legenden else None)
        bauer = _Bauer(doc, ansicht, massstab, ergebnis, texttyp_id, rahmen)
        bauer.baue(posten, quelle, feld_breite, feld_hoehe, abstand, spalten,
                   spaltenbreite)
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


def _revit_farbe(rgb):
    return Color(Byte(int(rgb[0])), Byte(int(rgb[1])), Byte(int(rgb[2])))


def _gleiche_farbe(farbe, rgb):
    if rgb is None:
        return True         # ohne Muster ist die Farbe egal
    return _farbe(farbe) == tuple(rgb)


class _Bauer(object):

    def __init__(self, doc, ansicht, massstab, ergebnis, texttyp_id, rahmen):
        self.doc = doc
        self.ansicht = ansicht
        self.papier = MM * float(massstab)      # 1 mm Papier in Modelleinheiten
        self.ergebnis = ergebnis
        self.rahmen = rahmen
        self.muster = muster(doc)
        self._typen = {}
        if texttyp_id is None or texttyp_id == ElementId.InvalidElementId:
            texttyp_id = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)
        self.textoptionen = TextNoteOptions(texttyp_id)
        self.textoptionen.HorizontalAlignment = HorizontalTextAlignment.Left
        self.textoptionen.VerticalAlignment = VerticalTextAlignment.Middle
        self._unsichtbar = None

    def baue(self, posten, quelle, feld_breite, feld_hoehe, abstand, spalten,
             spaltenbreite):
        mit_farbe = []
        for eintrag, text in posten:
            fuellung = lg.darstellung(eintrag.grafik, quelle)
            if fuellung is None:
                self.ergebnis.ohne_farbe.append(eintrag.name)
            else:
                mit_farbe.append((eintrag, text, fuellung))
        ecken = lg.raster(len(mit_farbe), spalten, feld_hoehe, abstand, spaltenbreite)
        for (eintrag, text, fuellung), (x, y) in zip(mit_farbe, ecken):
            try:
                typ_id = self._fuelltyp(eintrag, fuellung)
                self._feld(typ_id, x, y, feld_breite, feld_hoehe)
            except LegendenFehler:
                raise
            except Exception:
                self.ergebnis.fehler.append(eintrag.name)
                continue
            if text:
                punkt = XYZ((x + feld_breite + TEXTABSTAND_MM) * self.papier,
                            (y - feld_hoehe / 2.0) * self.papier, 0.0)
                TextNote.Create(self.doc, self.ansicht.Id, punkt, text, self.textoptionen)
                self.ergebnis.texte += 1

    # -- Füllbereichstypen -----------------------------------------------------

    def _fuelltyp(self, eintrag, fuellung):
        if eintrag.schluessel in self._typen:
            return self._typen[eintrag.schluessel]
        ids = (self.muster.element_id(fuellung.vg_muster),
               self.muster.element_id(fuellung.hg_muster))

        def finden(name):
            for element in FilteredElementCollector(self.doc).OfClass(FilledRegionType):
                if name_von(element).lower() == name.lower():
                    return element
            return None

        def gleich(element):
            return (element.ForegroundPatternId == ids[0]
                    and element.BackgroundPatternId == ids[1]
                    and _gleiche_farbe(element.ForegroundPatternColor, fuellung.vg_farbe)
                    and _gleiche_farbe(element.BackgroundPatternColor, fuellung.hg_farbe)
                    and not element.IsMasking)

        def erzeugen(name):
            basis = FilteredElementCollector(self.doc).OfClass(FilledRegionType) \
                .FirstElement()
            if basis is None:
                raise LegendenFehler(t(u"Im Projekt gibt es keinen Füllbereichstyp.",
                                       u"The project has no filled region type.",
                                       u"El proyecto no tiene tipo de región rellenada."))
            neu = basis.Duplicate(name)
            self._setze(neu, eintrag, fuellung, ids)
            self.ergebnis.neue_typen += 1
            return neu

        wunsch = dxf_lg.revit_name(TYP_PRAEFIX + eintrag.name, 200)
        element, _art = dxf_rv._finde_oder_erzeuge(wunsch, finden, gleich, erzeugen)
        self._typen[eintrag.schluessel] = element.Id
        return element.Id

    def _setze(self, typ, eintrag, fuellung, ids):
        typ.IsMasking = False
        for vorne, muster_id, farbe in ((True, ids[0], fuellung.vg_farbe),
                                        (False, ids[1], fuellung.hg_farbe)):
            try:
                self._muster_setzen(typ, vorne, muster_id)
            except Exception:
                # z.B. Modellmuster, das der Typ nicht nimmt -> Vollton
                if muster_id == ElementId.InvalidElementId:
                    raise
                self._muster_setzen(typ, vorne, self.muster.element_id(lg.VOLL))
                if eintrag.name not in self.ergebnis.ersatzmuster:
                    self.ergebnis.ersatzmuster.append(eintrag.name)
            if farbe:
                if vorne:
                    typ.ForegroundPatternColor = _revit_farbe(farbe)
                else:
                    typ.BackgroundPatternColor = _revit_farbe(farbe)

    @staticmethod
    def _muster_setzen(typ, vorne, muster_id):
        if vorne:
            typ.ForegroundPatternId = muster_id
        else:
            typ.BackgroundPatternId = muster_id

    # -- Felder ------------------------------------------------------------------

    def _feld(self, typ_id, x, y, breite, hoehe):
        p = self.papier
        ecken = [XYZ(x * p, y * p, 0.0), XYZ((x + breite) * p, y * p, 0.0),
                 XYZ((x + breite) * p, (y - hoehe) * p, 0.0),
                 XYZ(x * p, (y - hoehe) * p, 0.0)]
        schleife = CurveLoop()
        for index in range(4):
            schleife.Append(Line.CreateBound(ecken[index], ecken[(index + 1) % 4]))
        schleifen = List[CurveLoop]()
        schleifen.Add(schleife)
        bereich = FilledRegion.Create(self.doc, typ_id, self.ansicht.Id, schleifen)
        if not self.rahmen:
            stil = self._unsichtbare_linien()
            if stil != ElementId.InvalidElementId:
                bereich.SetLineStyleId(stil)
        self.ergebnis.felder += 1

    def _unsichtbare_linien(self):
        if self._unsichtbar is None:
            self._unsichtbar = ElementId.InvalidElementId
            ziel = id_wert(ElementId(BuiltInCategory.OST_InvisibleLines))
            for stil_id in FilledRegion.GetValidLineStyleIdsForFilledRegion(self.doc):
                stil = self.doc.GetElement(stil_id)
                kategorie = getattr(stil, "GraphicsStyleCategory", None)
                if kategorie is not None and id_wert(kategorie.Id) == ziel:
                    self._unsichtbar = stil_id
                    break
        return self._unsichtbar
