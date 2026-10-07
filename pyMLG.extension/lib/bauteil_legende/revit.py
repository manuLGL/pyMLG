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
       und "Ansichtsrichtung" (LEGEND_COMPONENT_VIEW = Grundriss) setzen
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
    FilteredElementCollector,
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


def lies_ansicht(doc, ansicht):
    """Funde für logik.sammle: alle sichtbaren Modellelemente mit Typ."""
    gesperrt = _gesperrt()
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
                      id_wert(element.Id)))
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
        self.abgelehnt = []          # Typen, die Revit nicht als Bauteil nimmt
        self.ohne_grundriss = []     # Typen ohne Ansichtsrichtung Grundriss


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
            texttyp_id=None, ersetzen=False):
    """Legende mit einem Bauteil je Typ (in der gegebenen Reihenfolge)
    untereinander. Rückgabe Ergebnis."""
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
        _Bauer(doc, ansicht, muster, massstab, abstand_mm, ergebnis).baue(
            typen, beschriftung, texttyp_id)
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

    def __init__(self, doc, ansicht, muster, massstab, abstand_mm, ergebnis):
        self.doc = doc
        self.ansicht = ansicht
        self.muster = muster
        self.abstand = abstand_mm * MM * float(massstab)
        self.ergebnis = ergebnis

    def baue(self, typen, beschriftung, texttyp_id):
        platziert = []          # (typ, breite, mitte_y)
        oben = 0.0
        for typ in typen:
            masse = self._bauteil(typ, oben)
            if masse is None:
                continue
            breite, hoehe = masse
            platziert.append((typ, breite, oben - hoehe / 2.0))
            oben -= hoehe + self.abstand
        self.ergebnis.platziert = len(platziert)
        if beschriftung != lg.BESCHRIFTUNG_KEINE and platziert:
            self._texte(platziert, beschriftung, texttyp_id)

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
            self._grundriss(bauteil)
            parameter = bauteil.get_Parameter(BuiltInParameter.LEGEND_COMPONENT)
            parameter.Set(typ.ref)
            if id_wert(parameter.AsElementId()) != typ.schluessel:
                raise ValueError(u"Bauteiltyp nicht übernommen")
            grundriss = self._grundriss(bauteil)
            doc.Regenerate()
            box = bauteil.get_BoundingBox(self.ansicht)
            if box is None:
                raise ValueError(u"Bauteil ohne Ausdehnung")
            ElementTransformUtils.MoveElement(
                doc, bauteil.Id, XYZ(-box.Min.X, oben - box.Max.Y, 0.0))
            unter.Commit()
        except Exception:
            if unter.HasStarted() and not unter.HasEnded():
                unter.RollBack()
            self.ergebnis.abgelehnt.append(typ)
            return None
        if not grundriss:
            self.ergebnis.ohne_grundriss.append(typ)
        return box.Max.X - box.Min.X, box.Max.Y - box.Min.Y

    @staticmethod
    def _grundriss(bauteil):
        """Ansichtsrichtung auf Grundriss stellen. Rückgabe, ob sie es ist."""
        richtung = bauteil.get_Parameter(BuiltInParameter.LEGEND_COMPONENT_VIEW)
        if richtung is None or richtung.IsReadOnly:
            return False
        try:
            richtung.Set(lg.RICHTUNG_GRUNDRISS)
            return richtung.AsInteger() == lg.RICHTUNG_GRUNDRISS
        except Exception:
            return False

    def _texte(self, platziert, art, texttyp_id):
        if texttyp_id is None or texttyp_id == ElementId.InvalidElementId:
            texttyp_id = self.doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType)
        x = max(breite for _typ, breite, _y in platziert) + max(self.abstand, 0.0)
        if self.abstand <= 0:
            x += 3.0 * MM * self.ansicht.Scale
        optionen = TextNoteOptions(texttyp_id)
        optionen.HorizontalAlignment = HorizontalTextAlignment.Left
        optionen.VerticalAlignment = VerticalTextAlignment.Middle
        for typ, _breite, mitte in platziert:
            text = lg.beschriftung(typ, art)
            if not text:
                continue
            TextNote.Create(self.doc, self.ansicht.Id, XYZ(x, mitte, 0.0), text, optionen)
            self.ergebnis.texte += 1
