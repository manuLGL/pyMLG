# -*- coding: utf-8 -*-
"""Revit-Teil von ComponentLegend.

Lesen: Alle in einer Ansicht sichtbaren Modellelemente mit Familientyp
(FilteredElementCollector mit Ansicht - berücksichtigt Sichtbarkeit und
Zuschnitt). Verschachtelte Familien und Kategorien, die Revit nicht als
Legendenbauteil zeigen kann (Räume, Raster, Ebenen …), fallen weg.

Erzeugen: Die API kann weder Legenden noch Legendenbauteile neu anlegen.
Deshalb:
    1. eine vorhandene Legende ohne Inhalt duplizieren (wie DxfLegend)
    2. ein vorhandenes Legendenbauteil als Muster in die neue Legende kopieren
    3. je Typ das Muster kopieren, Parameter "Bauteiltyp" (LEGEND_COMPONENT)
       und "Ansichtsrichtung" (LEGEND_COMPONENT_VIEW, Vorgabe Grundriss)
       setzen - kennt der Typ die gewählte Richtung nicht: Grundriss
    4. untereinander schieben: Oberkante an Unterkante des vorigen minus
       Abstand, linksbündig
    5. optional Text rechts daneben, auf Höhe der Bauteilmitte
Jeder Typ läuft in einer Untertransaktion - lehnt Revit einen Typ ab, fehlt
nur dieser. Alles in einer Transaktion -> ein Rückgängig-Schritt.
"""

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    CategoryType,
    CopyPasteOptions,
    ElementId,
    ElementTransformUtils,
    ElementTypeGroup,
    Curve,
    FilteredElementCollector,
    GeometryInstance,
    GraphicsStyleType,
    Line,
    Options,
    PolyLine,
    Solid,
    ReferenceArray,
    SpecTypeId,
    StorageType,
    HorizontalTextAlignment,
    SubTransaction,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    Transaction,
    TransactionStatus,
    Transform,
    VerticalTextAlignment,
    View,
    ViewType,
    XYZ,
)
from System.Collections.Generic import List

from bauteil_legende import logik as lg
from dxf_legende import revit as dxf_rv
from mlg_sprache import t

MM = 1.0 / 304.8

# Maßketten, in mm auf dem Papier
KETTE_ABSTAND_MM = 8.0     # Bauteilkante bis Maßlinie
HILFSLINIE_MM = 2.0        # Länge der unsichtbaren Hilfslinien
KETTE_PLATZ_MM = 12.0      # zusätzlicher Platz unter Bauteilen mit Breitenkette

LegendenFehler = dxf_rv.LegendenFehler
id_wert = dxf_rv.id_wert
name_von = dxf_rv._name

# Ansichten, aus denen gelesen werden kann
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
}

# Modellkategorien, die kein Legendenbauteil sein können. Namen statt
# Werte, damit eine in einer Revit-Version fehlende Kategorie nicht stört.
NICHT_MOEGLICH = (
    "OST_Rooms", "OST_Areas", "OST_MEPSpaces", "OST_HVAC_Zones",
    "OST_Grids", "OST_Levels", "OST_CLines", "OST_Lines", "OST_RvtLinks",
    "OST_Cameras", "OST_SectionBox", "OST_VolumeOfInterest",
    "OST_IOSModelGroups", "OST_Parts", "OST_Assemblies",
    "OST_ProjectBasePoint", "OST_SharedBasePoint", "OST_InternalOriginOverride",
    "OST_PipingSystem", "OST_DuctSystem", "OST_ElectricalCircuit",
    "OST_Rebar", "OST_AreaRein", "OST_PathRein", "OST_FabricAreas",
    "OST_FabricReinforcement", "OST_Topography", "OST_Coordination_Model",
    "OST_PointClouds", "OST_RasterImages", "OST_Materials",
)


# Kategorien, deren Paso (Breite × Höhe) angeschrieben werden kann
MIT_MASSEN = ("OST_Doors", "OST_Windows")


def _kategorie_ids(namen):
    ids = set()
    for name in namen:
        kategorie = getattr(BuiltInCategory, name, None)
        if kategorie is None:
            continue
        try:
            ids.add(id_wert(ElementId(kategorie)))
        except Exception:
            pass
    return ids


_GESPERRT = []


def _gesperrt():
    if not _GESPERRT:
        _GESPERRT.append(_kategorie_ids(NICHT_MOEGLICH))
    return _GESPERRT[0]


_MASSKATEGORIEN = []


def hat_masse(typ):
    """Tür oder Fenster - Typen, deren Paso angeschrieben werden kann."""
    if not _MASSKATEGORIEN:
        _MASSKATEGORIEN.append(_kategorie_ids(MIT_MASSEN))
    return typ.kat_id in _MASSKATEGORIEN[0]


def _ist_laenge(parameter):
    try:
        return (parameter.StorageType == StorageType.Double
                and parameter.Definition.GetDataType() == SpecTypeId.Length)
    except Exception:
        return False


def _quellen(doc, typ):
    """Typ und ein Exemplar - dort werden Parameter gesucht."""
    quellen = [doc.GetElement(typ.ref)]
    if typ.beispiel is not None:
        quellen.append(doc.GetElement(typ.beispiel))
    return [q for q in quellen if q is not None]


def laengenparameter(doc, typen):
    """Namen aller Längenparameter der Türen/Fenster (Typ und Exemplar),
    sortiert, und die Vorgaben (Breite, Höhe) - die Revit-Parameter
    Breite/Höhe in der Sprache des Projekts."""
    namen = set()
    vorgabe = [None, None]
    for typ in typen:
        if not hat_masse(typ):
            continue
        for quelle in _quellen(doc, typ):
            for parameter in quelle.Parameters:
                if _ist_laenge(parameter):
                    namen.add(parameter.Definition.Name)
            for index, eingebaut in enumerate((BuiltInParameter.FAMILY_WIDTH_PARAM,
                                               BuiltInParameter.FAMILY_HEIGHT_PARAM)):
                if vorgabe[index] is None:
                    parameter = quelle.get_Parameter(eingebaut)
                    if parameter is not None:
                        vorgabe[index] = parameter.Definition.Name
    return sorted(namen, key=lambda n: n.lower()), tuple(vorgabe)


def massen_mm(doc, typ, name):
    """Wert des Längenparameters name in mm - zuerst am Typ, dann am
    Exemplar. None, wenn es ihn nicht gibt oder er leer ist."""
    if not name:
        return None
    for quelle in _quellen(doc, typ):
        parameter = quelle.LookupParameter(name)
        if parameter is not None and parameter.HasValue and _ist_laenge(parameter):
            return parameter.AsDouble() * 304.8
    return None


def _standard_mm(doc, typ, eingebaut):
    """Revit-Parameter Breite/Höhe der Familie (Typ, dann Exemplar)."""
    for quelle in _quellen(doc, typ):
        parameter = quelle.get_Parameter(eingebaut)
        if parameter is not None and parameter.HasValue:
            return parameter.AsDouble() * 304.8
    return None


def masstext(doc, typ, masse):
    """u"825 × 2030" für Türen/Fenster, sonst u"". masse = (Breite, Höhe)
    als Parameternamen oder None."""
    return lg.masstext(*masswerte(doc, typ, masse))


def masswerte(doc, typ, masse):
    """(Breite, Höhe) in mm für Türen/Fenster, sonst (None, None). Hat eine
    Familie den gewählten Parameter nicht (z.B. "Breite Schiebeflügel" nur
    bei Schiebetüren), gilt ihre Standard-Breite bzw. -Höhe."""
    if not masse or not hat_masse(typ):
        return None, None
    werte = []
    for name, eingebaut in zip(masse, (BuiltInParameter.FAMILY_WIDTH_PARAM,
                                       BuiltInParameter.FAMILY_HEIGHT_PARAM)):
        wert = massen_mm(doc, typ, name)
        if wert is None:
            wert = _standard_mm(doc, typ, eingebaut)
        werte.append(wert)
    return tuple(werte)


def ansichten(doc):
    """Grafische Ansichten (ohne Vorlagen), nach Art und Name."""
    reihen = list(ANSICHTSARTEN)
    ergebnis = [v for v in FilteredElementCollector(doc).OfClass(View)
                if not v.IsTemplate and v.ViewType in ANSICHTSARTEN]
    return sorted(ergebnis, key=lambda v: (reihen.index(v.ViewType),
                                           name_von(v).lower()))


def ist_lesbar(ansicht):
    return (ansicht is not None and not ansicht.IsTemplate
            and ansicht.ViewType in ANSICHTSARTEN)


def _zuschnitt(ansicht):
    """(Umkehr-Transformation, Rahmen) des aktiven Zuschnitts oder None.
    3D-Ansichten bleiben aussen vor (Perspektive, Schnittbereich regelt
    der Collector selbst)."""
    if ansicht.ViewType == ViewType.ThreeD:
        return None
    try:
        if not ansicht.CropBoxActive:
            return None
        box = ansicht.CropBox
    except Exception:
        return None
    if box is None:
        return None
    return box.Transform.Inverse, (box.Min.X, box.Min.Y, box.Max.X, box.Max.Y)


def _im_zuschnitt(element, zuschnitt):
    box = element.get_BoundingBox(None)
    if box is None:
        return True
    umkehr, rahmen = zuschnitt
    punkte = []
    for x in (box.Min.X, box.Max.X):
        for y in (box.Min.Y, box.Max.Y):
            for z in (box.Min.Z, box.Max.Z):
                p = umkehr.OfPoint(XYZ(x, y, z))
                punkte.append((p.X, p.Y))
    return lg.ueberlappt(punkte, rahmen)


def lies_ansicht(doc, ansicht):
    """Funde für logik.sammle: alle sichtbaren Modellelemente mit Typ.

    Der Collector mit Ansicht lässt laut API-Doku Elemente knapp ausserhalb
    des Zuschnitts durch - die werden hier am Zuschnittrahmen aussortiert."""
    gesperrt = _gesperrt()
    zuschnitt = _zuschnitt(ansicht)
    funde = []
    namen = {}
    for element in (FilteredElementCollector(doc, ansicht.Id)
                    .WhereElementIsNotElementType()):
        kategorie = element.Category
        if kategorie is None:
            continue
        if kategorie.Parent is not None:
            kategorie = kategorie.Parent
        if kategorie.CategoryType != CategoryType.Model or element.ViewSpecific:
            continue
        kat_id = id_wert(kategorie.Id)
        if kat_id in gesperrt:
            continue
        # verschachtelte Familien gehören zur übergeordneten
        if getattr(element, "SuperComponent", None) is not None:
            continue
        if zuschnitt is not None and not _im_zuschnitt(element, zuschnitt):
            continue
        typ_id = element.GetTypeId()
        if typ_id is None or typ_id == ElementId.InvalidElementId:
            continue
        schluessel = id_wert(typ_id)
        if schluessel not in namen:
            typ = doc.GetElement(typ_id)
            if typ is None:
                namen[schluessel] = None
            else:
                namen[schluessel] = (getattr(typ, "FamilyName", u"") or u"",
                                     name_von(typ))
        if namen[schluessel] is None:
            continue
        familie, typname = namen[schluessel]
        funde.append((kat_id, kategorie.Name, schluessel, typ_id, familie, typname,
                      id_wert(element.Id), element.Id))
    return funde


def vorlage_bauteil(doc):
    """Ein vorhandenes Legendenbauteil in einer Legende - oder None."""
    for element in (FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_LegendComponents)
                    .WhereElementIsNotElementType()):
        ansicht = doc.GetElement(element.OwnerViewId)
        if ansicht is not None and ansicht.ViewType == ViewType.Legend \
                and not ansicht.IsTemplate:
            return element
    # Rückfall: jede Legende direkt durchsuchen
    gesucht = id_wert(ElementId(BuiltInCategory.OST_LegendComponents))
    for ansicht in dxf_rv.legenden(doc):
        for element in FilteredElementCollector(doc, ansicht.Id).WhereElementIsNotElementType():
            kategorie = element.Category
            if kategorie is not None and id_wert(kategorie.Id) == gesucht \
                    and element.OwnerViewId == ansicht.Id:
                return element
    return None


def texttypen(doc):
    """[(Name, Id)] aller Texttypen und die Id des Vorgabetyps."""
    typen = sorted(((name_von(e), e.Id) for e in
                    FilteredElementCollector(doc).OfClass(TextNoteType)),
                   key=lambda e: e[0].lower())
    vorgabe = doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)
    return typen, vorgabe


def vorhandene_legende(doc, name):
    ansicht = dxf_rv.vorhandene_ansicht(doc, name)
    if ansicht is not None and ansicht.ViewType == ViewType.Legend:
        return ansicht
    return None


class Ergebnis(object):

    def __init__(self):
        self.ansicht = None
        self.platziert = 0
        self.texte = 0
        self.ersetzt = False
        self.geloescht = 0
        self.ketten = 0              # erzeugte Maßketten
        self.ketten_fehler = 0
        self.abgelehnt = []          # Typen, die Revit nicht als Bauteil nimmt
        self.ohne_richtung = []      # Typen ohne die gewählte Ansichtsrichtung


def _ids(*element_ids):
    liste = List[ElementId]()
    for element_id in element_ids:
        liste.Add(element_id)
    return liste


def _kopiere_muster(doc, vorlage, ziel):
    if vorlage.OwnerViewId == ziel.Id:
        neu = ElementTransformUtils.CopyElement(doc, vorlage.Id, XYZ.Zero)
    else:
        quelle = doc.GetElement(vorlage.OwnerViewId)
        neu = ElementTransformUtils.CopyElements(quelle, _ids(vorlage.Id), ziel,
                                                 Transform.Identity, CopyPasteOptions())
    neu = list(neu)
    if not neu:
        raise LegendenFehler(t(u"Das Muster-Legendenbauteil liess sich nicht kopieren.",
                               u"The template legend component could not be copied.",
                               u"No se pudo copiar el componente de leyenda de muestra."))
    return doc.GetElement(neu[0])


def _leere(doc, ansicht, behalten):
    ids = [e.Id for e in FilteredElementCollector(doc, ansicht.Id)
           .WhereElementIsNotElementType()
           if e.OwnerViewId == ansicht.Id and e.Category is not None
           and e.Id != behalten]
    geloescht = 0
    for element_id in ids:
        if doc.GetElement(element_id) is None:
            continue
        try:
            doc.Delete(element_id)
            geloescht += 1
        except Exception:
            pass
    return geloescht


def erzeuge(doc, typen, name, massstab, abstand_mm, beschriftung=lg.BESCHRIFTUNG_KEINE,
            texttyp_id=None, ersetzen=False, richtung=lg.RICHTUNG_GRUNDRISS,
            masse=None, mass_art=lg.MASS_TEXT):
    """Legende mit einem Bauteil je Typ (in der gegebenen Reihenfolge)
    untereinander. masse = (Breite, Höhe) als Parameternamen: Paso von
    Türen/Fenstern als zweite Textzeile und/oder Maßkette (mass_art).
    Rückgabe Ergebnis."""
    vorlage = vorlage_bauteil(doc)
    if vorlage is None:
        raise LegendenFehler(fehlt_vorlage_text())
    ergebnis = Ergebnis()
    transaktion = Transaction(doc, t(u"pyMLG Legende aus Ansichten",
                                     u"pyMLG Legend from views",
                                     u"pyMLG Leyenda desde vistas"))
    transaktion.Start()
    try:
        ansicht = vorhandene_legende(doc, name) if ersetzen else None
        if ansicht is not None:
            # Muster zuerst kopieren - die Vorlage kann in dieser Legende liegen
            muster = _kopiere_muster(doc, vorlage, ansicht)
            ergebnis.geloescht = _leere(doc, ansicht, muster.Id)
            ergebnis.ersetzt = True
            try:
                ansicht.Scale = int(massstab)
            except Exception:
                pass
        else:
            ansicht = dxf_rv._neue_ansicht(doc, name, massstab,
                                           doc.GetElement(vorlage.OwnerViewId))
            muster = _kopiere_muster(doc, vorlage, ansicht)
        _Bauer(doc, ansicht, muster, massstab, abstand_mm, ergebnis, richtung).baue(
            typen, beschriftung, texttyp_id, masse, mass_art)
        doc.Delete(muster.Id)
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


def fehlt_vorlage_text():
    return t(u"Im Projekt ist noch kein Legendenbauteil platziert. Geladene "
             u"Familientypen zählen nicht - ein Legendenbauteil entsteht erst beim "
             u"Platzieren in einer Legende. Die Revit-API kann Legenden und "
             u"Legendenbauteile nicht neu anlegen, nur kopieren.\n\n"
             u"Einmal von Hand: Ansicht > Legenden > Legende anlegen, dann "
             u"Beschriften > Bauteil > Legendenbauteil und irgendein Bauteil "
             u"platzieren. Danach erledigt dieses Werkzeug den Rest.",
             u"No legend component is placed in the project yet. Loaded family "
             u"types do not count - a legend component only exists once it is "
             u"placed in a legend. The Revit API cannot "
             u"create legends or legend components, it can only copy them.\n\n"
             u"Once by hand: View > Legends > Legend, then Annotate > Component > "
             u"Legend Component and place any component. After that this tool "
             u"does the rest.",
             u"El proyecto aún no tiene ningún componente de leyenda colocado. Los "
             u"tipos de familia cargados no cuentan - un componente de leyenda solo "
             u"existe al colocarlo en una leyenda. La API de "
             u"Revit no puede crear leyendas ni componentes de leyenda, solo "
             u"copiarlos.\n\nUna vez a mano: Vista > Leyendas > Leyenda, después "
             u"Anotar > Componente > Componente de leyenda y colocar uno cualquiera. "
             u"Luego esta herramienta hace el resto.")


class _Bauer(object):

    def __init__(self, doc, ansicht, muster, massstab, abstand_mm, ergebnis,
                 richtung=lg.RICHTUNG_GRUNDRISS):
        self.doc = doc
        self.richtung = richtung
        self.ansicht = ansicht
        self.muster = muster
        self.papier = MM * float(massstab)          # 1 mm Papier in Modelleinheiten
        self.abstand = abstand_mm * self.papier
        self.ergebnis = ergebnis
        self._unsichtbar = None

    def baue(self, typen, beschriftung, texttyp_id, masse=None, mass_art=lg.MASS_TEXT):
        platziert = []          # (typ, breite, mitte_y)
        oben = 0.0
        kette = bool(masse) and lg.mit_kette(mass_art)
        for typ in typen:
            groesse = self._bauteil(typ, oben)
            if groesse is None:
                continue
            breite, hoehe = groesse
            platziert.append((typ, breite, oben - hoehe / 2.0))
            unten = oben - hoehe
            oben = unten - self.abstand
            if kette:
                werte = masswerte(self.doc, typ, masse)
                if self._kette(werte, breite, unten):
                    # Platz für die Breitenkette unter dem Bauteil
                    oben -= KETTE_PLATZ_MM * self.papier
        self.ergebnis.platziert = len(platziert)
        text_masse = masse if lg.mit_text(mass_art) else None
        if (beschriftung != lg.BESCHRIFTUNG_KEINE or text_masse) and platziert:
            self._texte(platziert, beschriftung, texttyp_id, text_masse)

    # -- Maßketten ----------------------------------------------------------------

    def _kette(self, werte, breite_bauteil, unten):
        """Maßketten für das Paso (mm) eines Bauteils mit Unterkante unten,
        links an x=0. Rückgabe, ob eine Breitenkette unter dem Bauteil liegt."""
        breite_mm, hoehe_mm = werte
        if self.richtung not in lg.RICHTUNGEN_MIT_HOEHE:
            hoehe_mm = None
        lage = lg.kettenlage(0.0, unten, breite_bauteil,
                             breite_mm * MM if breite_mm else None,
                             hoehe_mm * MM if hoehe_mm else None,
                             KETTE_ABSTAND_MM * self.papier)
        hilfe = HILFSLINIE_MM * self.papier
        if u"breite" in lage:
            x1, x2, y = lage[u"breite"]
            self._masslinie([Line.CreateBound(XYZ(x, unten, 0.0),
                                              XYZ(x, unten - hilfe, 0.0)) for x in (x1, x2)],
                            Line.CreateBound(XYZ(x1, y, 0.0), XYZ(x2, y, 0.0)))
        if u"hoehe" in lage:
            y1, y2, x = lage[u"hoehe"]
            self._masslinie([Line.CreateBound(XYZ(0.0, yy, 0.0),
                                              XYZ(-hilfe, yy, 0.0)) for yy in (y1, y2)],
                            Line.CreateBound(XYZ(x, y1, 0.0), XYZ(x, y2, 0.0)))
        return u"breite" in lage

    def _masslinie(self, hilfslinien, masslinie):
        """Bemaßung zwischen zwei unsichtbaren Hilfslinien - Legendenbauteile
        geben keine Kanten zum Bemaßen her."""
        doc = self.doc
        unter = SubTransaction(doc)
        unter.Start()
        try:
            bezuege = ReferenceArray()
            for kurve in hilfslinien:
                linie = doc.Create.NewDetailCurve(self.ansicht, kurve)
                stil = self._unsichtbare_linien()
                if stil is not None:
                    linie.LineStyle = stil
                bezuege.Append(linie.GeometryCurve.Reference)
            doc.Create.NewDimension(self.ansicht, masslinie, bezuege)
            unter.Commit()
            self.ergebnis.ketten += 1
        except Exception:
            if unter.HasStarted() and not unter.HasEnded():
                unter.RollBack()
            self.ergebnis.ketten_fehler += 1

    def _unsichtbare_linien(self):
        """Linienstil <Unsichtbare Linien> oder None."""
        if self._unsichtbar is None:
            self._unsichtbar = False
            linien = self.doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines)
            gesucht = id_wert(ElementId(BuiltInCategory.OST_InvisibleLines))
            for unter in linien.SubCategories:
                if id_wert(unter.Id) == gesucht:
                    self._unsichtbar = unter.GetGraphicsStyle(GraphicsStyleType.Projection)
                    break
        return self._unsichtbar or None

    def _bauteil(self, typ, oben):
        """Bauteil für typ mit Oberkante bei oben, linksbündig an x=0.
        Rückgabe (breite, hoehe) oder None, wenn Revit den Typ ablehnt."""
        doc = self.doc
        unter = SubTransaction(doc)
        unter.Start()
        try:
            neu = list(ElementTransformUtils.CopyElement(doc, self.muster.Id, XYZ.Zero))
            bauteil = doc.GetElement(neu[0])
            # Richtung vor und nach dem Typwechsel: Steht das Muster auf einer
            # Richtung, die der neue Typ nicht kennt, könnte Revit ihn ablehnen.
            self._setze_richtung(bauteil, lg.RICHTUNG_GRUNDRISS)
            parameter = bauteil.get_Parameter(BuiltInParameter.LEGEND_COMPONENT)
            parameter.Set(typ.ref)
            if id_wert(parameter.AsElementId()) != typ.schluessel:
                raise ValueError(u"Bauteiltyp nicht übernommen")
            gewaehlt = self._setze_richtung(bauteil, self.richtung)
            if not gewaehlt:
                # Richtung gibt es für diesen Typ nicht - Grundriss statt
                # dessen, was das Muster zufällig hatte
                self._setze_richtung(bauteil, lg.RICHTUNG_GRUNDRISS)
            doc.Regenerate()
            rahmen = self._sichtbar(bauteil)
            if rahmen is None:
                box = bauteil.get_BoundingBox(self.ansicht)
                if box is None:
                    raise ValueError(u"Bauteil ohne Ausdehnung")
                rahmen = (box.Min.X, box.Min.Y, box.Max.X, box.Max.Y)
            x0, y0, x1, y1 = rahmen
            ElementTransformUtils.MoveElement(doc, bauteil.Id, XYZ(-x0, oben - y1, 0.0))
            unter.Commit()
        except Exception:
            if unter.HasStarted() and not unter.HasEnded():
                unter.RollBack()
            self.ergebnis.abgelehnt.append(typ)
            return None
        if not gewaehlt:
            self.ergebnis.ohne_richtung.append(typ)
        return x1 - x0, y1 - y0

    def _sichtbar(self, bauteil):
        """(x0, y0, x1, y1) der sichtbaren Geometrie in der Legende - die
        Bounding-Box ist oft grösser (unsichtbare Teile der Familie), dann
        sässen Maßketten neben der Tür. None, wenn nichts lesbar ist."""
        optionen = Options()
        optionen.View = self.ansicht
        punkte = []

        def sammle(geometrie):
            if geometrie is None:
                return
            for objekt in geometrie:
                if isinstance(objekt, GeometryInstance):
                    sammle(objekt.GetInstanceGeometry())
                elif isinstance(objekt, Curve):
                    punkte.extend(objekt.Tessellate())
                elif isinstance(objekt, PolyLine):
                    punkte.extend(objekt.GetCoordinates())
                elif isinstance(objekt, Solid):
                    for kante in objekt.Edges:
                        punkte.extend(kante.Tessellate())

        try:
            sammle(bauteil.get_Geometry(optionen))
        except Exception:
            return None
        if not punkte:
            return None
        xs = [p.X for p in punkte]
        ys = [p.Y for p in punkte]
        return min(xs), min(ys), max(xs), max(ys)

    @staticmethod
    def _setze_richtung(bauteil, wert):
        """Ansichtsrichtung setzen. Rückgabe, ob das Bauteil sie jetzt hat."""
        richtung = bauteil.get_Parameter(BuiltInParameter.LEGEND_COMPONENT_VIEW)
        if richtung is None or richtung.IsReadOnly:
            return False
        try:
            richtung.Set(wert)
            return richtung.AsInteger() == wert
        except Exception:
            return False

    def _texte(self, platziert, art, texttyp_id, masse=None):
        if texttyp_id is None or texttyp_id == ElementId.InvalidElementId:
            texttyp_id = self.doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)
        x = max(breite for _typ, breite, _y in platziert) + max(self.abstand, 0.0)
        if self.abstand <= 0:
            x += 3.0 * MM * self.ansicht.Scale
        optionen = TextNoteOptions(texttyp_id)
        optionen.HorizontalAlignment = HorizontalTextAlignment.Left
        optionen.VerticalAlignment = VerticalTextAlignment.Middle
        for typ, _breite, mitte in platziert:
            text = lg.zeilen(lg.beschriftung(typ, art), masstext(self.doc, typ, masse))
            if not text:
                continue
            TextNote.Create(self.doc, self.ansicht.Id, XYZ(x, mitte, 0.0), text, optionen)
            self.ergebnis.texte += 1
