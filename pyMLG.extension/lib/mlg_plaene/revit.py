# -*- coding: utf-8 -*-
"""Revit-Teil der Plan-Werkzeuge: Pläne, Plankopf, Ansichtsfenster.

Läuft unter IronPython (die MLGplans-Buttons nutzen pyrevit.forms).
ElementId.IntegerValue gibt es ab Revit 2026 nicht mehr, daher id_wert().
"""

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter as BIP,
    CategoryType,
    Curve,
    ElementType,
    ElementTypeGroup,
    FilteredElementCollector,
    GeometryInstance,
    Options,
    PolyLine,
    ScheduleSheetInstance,
    Viewport,
    ViewSheet,
    XYZ,
)
from Autodesk.Revit.UI.Selection import ISelectionFilter

from mlg_plaene import logik as lg
from mlg_sprache import t

CM_PRO_FUSS = 30.48

# Ohne Plankopf: Fläche eines A1-Blatts (84,1 x 59,4 cm)
A1 = (0.0, 0.0, 84.1 / CM_PRO_FUSS, 59.4 / CM_PRO_FUSS)


def id_wert(element_id):
    """Zahlenwert einer ElementId (Revit 2024+: .Value)."""
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def cm(wert_fuss):
    return wert_fuss * CM_PRO_FUSS


def fuss(wert_cm):
    return wert_cm / CM_PRO_FUSS


def zahl(text):
    """Zahl aus einem Eingabefeld, Komma oder Punkt. Wirft ValueError."""
    return float(text.strip().replace(u",", u"."))


# ---------------------------------------------------------------- Pläne

def plaene(doc):
    """Alle echten Pläne (keine Platzhalter), natürlich nach Nummer sortiert."""
    alle = [s for s in FilteredElementCollector(doc).OfClass(ViewSheet)
            if not s.IsPlaceholder and not s.IsTemplate]
    return sorted(alle, key=lambda s: lg.natuerlich(s.SheetNumber))


def plan_text(plan):
    return u"{} - {}".format(plan.SheetNumber, plan.Name)


def vergebene_nummern(doc):
    """Normierte Nummern aller Pläne, auch der Platzhalter."""
    return lg.nummern_menge(s.SheetNumber for s in
                            FilteredElementCollector(doc).OfClass(ViewSheet))


def gewaehlte_plaene(doc, uidoc):
    """Im Projektbrowser gewählte Pläne, sonst der aktive Plan, sonst []."""
    ids = uidoc.Selection.GetElementIds()
    gewaehlt = [doc.GetElement(i) for i in ids]
    gewaehlt = [s for s in gewaehlt if isinstance(s, ViewSheet) and not s.IsPlaceholder]
    if gewaehlt:
        return sorted(gewaehlt, key=lambda s: lg.natuerlich(s.SheetNumber))
    if isinstance(doc.ActiveView, ViewSheet):
        return [doc.ActiveView]
    return []


def plankoepfe(doc):
    """{"Familie: Typ": Typ} aller geladenen Plankopf-Typen."""
    koepfe = {}
    typen = (FilteredElementCollector(doc)
             .OfCategory(BuiltInCategory.OST_TitleBlocks)
             .WhereElementIsElementType())
    for typ in typen:
        name = typ.get_Parameter(BIP.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM).AsString()
        if name:
            koepfe[name] = typ
    return koepfe


def plankopf_auf(doc, plan):
    return (FilteredElementCollector(doc, plan.Id)
            .OfCategory(BuiltInCategory.OST_TitleBlocks)
            .WhereElementIsNotElementType()
            .FirstElement())


def _rechteck(rahmen):
    return (rahmen.Min.X, rahmen.Min.Y, rahmen.Max.X, rahmen.Max.Y)


# Obergrenze für die Linien des Plankopfs in der Vorschau (Logos mit
# vielen Kurven sollen das Fenster nicht bremsen)
MAX_LINIEN = 5000


def _linien_sammeln(geometrie, linien):
    for teil in geometrie:
        if len(linien) >= MAX_LINIEN:
            return
        if isinstance(teil, GeometryInstance):
            _linien_sammeln(teil.GetInstanceGeometry(), linien)
        elif isinstance(teil, Curve):
            linien.append([(p.X, p.Y) for p in teil.Tessellate()])
        elif isinstance(teil, PolyLine):
            linien.append([(p.X, p.Y) for p in teil.GetCoordinates()])


def plankopf_linien(doc, plan):
    """Sichtbare Linien des Plankopfs auf dem Plan: [[(x, y), ...], ...].

    Nur was auf diesem Plan sichtbar ist - per Parameter ausgeblendete
    Linien anderer Formate (A1/A2/A3 in einer Familie) fehlen.
    """
    kopf = plankopf_auf(doc, plan)
    if kopf is None:
        return []
    optionen = Options()
    optionen.View = plan
    geometrie = kopf.get_Geometry(optionen)
    linien = []
    if geometrie is not None:
        _linien_sammeln(geometrie, linien)
    return [l for l in linien if len(l) >= 2]


def plankopf_rechteck(doc, plan):
    """Umriss des Plankopfs auf dem Plan.

    Aus den sichtbaren Linien (passt zum gewählten Format), sonst aus dem
    Begrenzungsrahmen, ohne Plankopf ein A1-Blatt.
    """
    linien = plankopf_linien(doc, plan)
    if linien:
        punkte = [p for linie in linien for p in linie]
        rechteck = (min(p[0] for p in punkte), min(p[1] for p in punkte),
                    max(p[0] for p in punkte), max(p[1] for p in punkte))
        if lg.breite(rechteck) > lg.EPS and lg.hoehe(rechteck) > lg.EPS:
            return rechteck
    kopf = plankopf_auf(doc, plan)
    rahmen = kopf.get_BoundingBox(plan) if kopf is not None else None
    if rahmen is None:
        return A1
    return _rechteck(rahmen)


# ---------------------------------------------------------------- Ansichtsfenster

def ansichtsfenster_typen(doc):
    """{Name: Typ} der Ansichtsfenster-Typen (mit/ohne Titel usw.)."""
    standard = doc.GetElement(doc.GetDefaultElementTypeId(ElementTypeGroup.ViewportType))
    if standard is None:
        return {}
    typen = {}
    for typ in FilteredElementCollector(doc).OfClass(ElementType):
        if typ.FamilyName == standard.FamilyName:
            name = typ.get_Parameter(BIP.SYMBOL_NAME_PARAM)
            typen[name.AsString() if name is not None else typ.Name] = typ
    return typen


# Kategorien, die den Revit-Umriss aufblähen, ohne zur Zeichnung zu gehören
_NICHT_ZEICHNUNG = set()
for _name in ("OST_Grids", "OST_Levels", "OST_Elev", "OST_Sections",
              "OST_Callouts", "OST_ReferenceViewer", "OST_Viewers",
              "OST_Cameras", "OST_VolumeOfInterest", "OST_CLines",
              "OST_ReferencePoints", "OST_SectionBox"):
    if hasattr(BuiltInCategory, _name):
        _NICHT_ZEICHNUNG.add(int(getattr(BuiltInCategory, _name)))

# Zeichnungs-Umriss je Ansicht in Projektionskoordinaten (Papiermass, vom
# Platz auf dem Plan unabhängig) - einmal je Durchlauf gemessen
_ZEICHNUNG = {}


def messung_zuruecksetzen():
    """Zu Beginn eines Werkzeugs aufrufen - das Modell kann sich seit dem
    letzten Lauf geändert haben."""
    _ZEICHNUNG.clear()


def _projiziert(zu_projektion, rahmen):
    """Umriss eines Begrenzungsrahmens in Projektionskoordinaten."""
    xs, ys = [], []
    for x in (rahmen.Min.X, rahmen.Max.X):
        for y in (rahmen.Min.Y, rahmen.Max.Y):
            for z in (rahmen.Min.Z, rahmen.Max.Z):
                punkt = zu_projektion.OfPoint(rahmen.Transform.OfPoint(XYZ(x, y, z)))
                xs.append(punkt.X)
                ys.append(punkt.Y)
    return (min(xs), min(ys), max(xs), max(ys))


def _schnitt(a, b):
    """Gemeinsame Fläche zweier Rechtecke oder None."""
    r = (max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3]))
    return r if lg.breite(r) > lg.EPS and lg.hoehe(r) > lg.EPS else None


def _zuschnitt_projektion(ansicht, zu_projektion):
    if not ansicht.CropBoxActive:
        return None
    xs, ys = [], []
    for schleife in ansicht.GetCropRegionShapeManager().GetCropShape():
        for kurve in schleife:
            for punkt in kurve.Tessellate():
                p = zu_projektion.OfPoint(punkt)
                xs.append(p.X)
                ys.append(p.Y)
    return (min(xs), min(ys), max(xs), max(ys)) if xs else None


def zeichnung_projektion(doc, ansicht):
    """Umriss der Zeichnung in Projektionskoordinaten oder None.

    Gemessen werden die sichtbaren Elemente: Modellelemente und
    Detailbauteile begrenzt auf den Zuschneidebereich (Revit schneidet sie
    dort ab), Beschriftungen wie Maße und Texte voll. Achsen, Ebenen und
    Ansichtsmarken zählen nicht.
    """
    schluessel = id_wert(ansicht.Id)
    if schluessel in _ZEICHNUNG:
        return _ZEICHNUNG[schluessel]
    ergebnis = None
    try:
        trafos = ansicht.GetModelToProjectionTransforms()
        if trafos.Count:
            zu_projektion = trafos[0].GetModelToProjectionTransform()
            zuschnitt = _zuschnitt_projektion(ansicht, zu_projektion)
            teile = []
            elemente = (FilteredElementCollector(doc, ansicht.Id)
                        .WhereElementIsNotElementType())
            for elem in elemente:
                kategorie = elem.Category
                if kategorie is None or id_wert(kategorie.Id) in _NICHT_ZEICHNUNG:
                    continue
                rahmen = elem.get_BoundingBox(ansicht)
                if rahmen is None:
                    continue
                r = _projiziert(zu_projektion, rahmen)
                if zuschnitt is not None and kategorie.CategoryType != CategoryType.Annotation:
                    r = _schnitt(r, zuschnitt)
                    if r is None:
                        continue
                teile.append(r)
            ergebnis = lg.huelle(teile) if teile else zuschnitt
    except Exception:
        ergebnis = None
    _ZEICHNUNG[schluessel] = ergebnis
    return ergebnis


UMRISS, ZEICHNUNG = "umriss", "zeichnung"


def box_rechteck(fenster):
    """Revit-Umriss des Ansichtsfensters ohne Titel."""
    box = fenster.GetBoxOutline()
    return (box.MinimumPoint.X, box.MinimumPoint.Y,
            box.MaximumPoint.X, box.MaximumPoint.Y)


def _auf_plan(fenster, r):
    """Rechteck aus Projektions- in Plankoordinaten (samt Drehung)."""
    zu_plan = fenster.GetProjectionToSheetTransform()
    punkte = [zu_plan.OfPoint(XYZ(x, y, 0)) for x in (r[0], r[2]) for y in (r[1], r[3])]
    return (min(p.X for p in punkte), min(p.Y for p in punkte),
            max(p.X for p in punkte), max(p.Y for p in punkte))


def messen(fenster, doc=None):
    """(Umriss samt Titel, Messart).

    Ohne doc das Revit-Rechteck (UMRISS). Das umfasst alles in der Ansicht,
    auch weit entfernte Achsenenden oder Marken, und ist oft viel grösser als
    die Zeichnung. Mit doc zählt die Zeichnung selbst (ZEICHNUNG, siehe
    zeichnung_projektion), begrenzt auf das Revit-Rechteck.
    """
    box = box_rechteck(fenster)
    x0, y0, x1, y1 = box
    art = UMRISS
    if doc is not None:
        try:
            zeichnung = zeichnung_projektion(doc, doc.GetElement(fenster.ViewId))
            if zeichnung is not None:
                schnitt = _schnitt(_auf_plan(fenster, zeichnung), box)
                if schnitt is not None:
                    x0, y0, x1, y1 = schnitt
                    art = ZEICHNUNG
        except Exception:
            pass
    try:
        titel = fenster.GetLabelOutline()
        if titel is not None and not titel.IsEmpty:
            x0 = min(x0, titel.MinimumPoint.X)
            y0 = min(y0, titel.MinimumPoint.Y)
            x1 = max(x1, titel.MaximumPoint.X)
            y1 = max(y1, titel.MaximumPoint.Y)
    except Exception:
        pass
    return (x0, y0, x1, y1), art


def fenster_rechteck(fenster, doc=None):
    """Umriss eines Ansichtsfensters samt Titel, siehe messen()."""
    return messen(fenster, doc)[0]


def masse_text(r):
    return u"{:.0f} x {:.0f} cm".format(cm(lg.breite(r)), cm(lg.hoehe(r)))


def messung_text(doc, fenster):
    """Kurzbeschreibung für Hinweise: gemessene Grösse, Messart, Revit-Umriss."""
    umriss, art = messen(fenster, doc)
    arten = {
        ZEICHNUNG: t(u"Zeichnung", u"drawing", u"dibujo"),
        UMRISS: t(u"Revit-Umriss", u"Revit outline", u"contorno de Revit"),
    }
    return t(u"{} ({}, samt Titel; Revit-Umriss {})",
             u"{} ({}, incl. title; Revit outline {})",
             u"{} ({}, con título; contorno de Revit {})").format(
        masse_text(umriss), arten[art], masse_text(box_rechteck(fenster)))


def belegte_rechtecke(doc, plan, ausser=()):
    """Umrisse aller Ansichtsfenster und Bauteillisten auf dem Plan."""
    ausser = set(id_wert(i) for i in ausser)
    belegt = []
    for fid in plan.GetAllViewports():
        if id_wert(fid) not in ausser:
            belegt.append(fenster_rechteck(doc.GetElement(fid), doc))
    for liste in FilteredElementCollector(doc, plan.Id).OfClass(ScheduleSheetInstance):
        if liste.IsTitleblockRevisionSchedule or id_wert(liste.Id) in ausser:
            continue
        rahmen = liste.get_BoundingBox(plan)
        if rahmen is not None:
            belegt.append(_rechteck(rahmen))
    return belegt


def belegte_eintraege(doc, plan, ausser=()):
    """[(Name, Umriss)] aller Ansichtsfenster und Bauteillisten auf dem Plan -
    wie belegte_rechtecke(), mit Namen für die Vorschau."""
    ausser = set(id_wert(i) for i in ausser)
    eintraege = []
    for fid in plan.GetAllViewports():
        if id_wert(fid) not in ausser:
            fenster = doc.GetElement(fid)
            eintraege.append((doc.GetElement(fenster.ViewId).Name,
                              fenster_rechteck(fenster, doc)))
    for liste in FilteredElementCollector(doc, plan.Id).OfClass(ScheduleSheetInstance):
        if liste.IsTitleblockRevisionSchedule or id_wert(liste.Id) in ausser:
            continue
        rahmen = liste.get_BoundingBox(plan)
        if rahmen is not None:
            eintraege.append((doc.GetElement(liste.ScheduleId).Name, _rechteck(rahmen)))
    return eintraege


def vorschau_name(ansicht):
    """Name mit Maßstab für die Vorschau."""
    try:
        if ansicht.Scale > 1:
            return u"{}  1:{}".format(ansicht.Name, ansicht.Scale)
    except Exception:
        pass
    return ansicht.Name


def verschiebe_rechteck_auf(fenster, ist, soll):
    """Verschiebt das Ansichtsfenster so, dass Umriss ist auf soll liegt."""
    (ix, iy), (sx, sy) = lg.mitte(ist), lg.mitte(soll)
    fenster.SetBoxCenter(fenster.GetBoxCenter() + XYZ(sx - ix, sy - iy, 0))


def platzieren(doc, plan, ansicht, ziel, typ=None, belegt=None, abstand=0.0,
               erzwingen=False, ueberlappen=False):
    """Setzt eine Ansicht auf einen Plan.

    ziel ist entweder ein Punkt (x, y) für die Mitte des Ansichtsfensters
    oder ein Bereich (x0, y0, x1, y1). Mit belegt (Liste von Umrissen) wird
    im Bereich ein freier Platz gesucht, sonst wird im Bereich zentriert.

    Liefert (Ansichtsfenster, Umriss). War kein Platz, ist das
    Ansichtsfenster None und der Umriss der gemessene (für Hinweise) - mit
    erzwingen wird dann trotzdem zentriert, mit ueberlappen an die Stelle
    mit der kleinsten Überschneidung gesetzt.
    """
    if len(ziel) == 2:
        mittelpunkt = XYZ(ziel[0], ziel[1], 0)
    else:
        mx, my = lg.mitte(ziel)
        mittelpunkt = XYZ(mx, my, 0)
    fenster = Viewport.Create(doc, plan.Id, ansicht.Id, mittelpunkt)
    if typ is not None:
        fenster.ChangeTypeId(typ.Id)
    doc.Regenerate()
    if len(ziel) == 2:
        return fenster, fenster_rechteck(fenster, doc)

    ist = fenster_rechteck(fenster, doc)
    if belegt is None:
        soll = lg.zentriert(ist, ziel)
    elif ueberlappen:
        soll = lg.beste_position(ziel, belegt, lg.breite(ist), lg.hoehe(ist), abstand)
    else:
        soll = lg.freie_position(ziel, belegt, lg.breite(ist), lg.hoehe(ist), abstand)
        if soll is None:
            if not erzwingen:
                doc.Delete(fenster.Id)
                return None, ist
            soll = lg.zentriert(ist, ziel)
    verschiebe_rechteck_auf(fenster, ist, soll)
    return fenster, soll


def modell_auf_plan(fenster, ansicht, punkt=None):
    """Lage eines Modellpunkts (Standard: Projektursprung) auf dem Plan.

    None, wenn die Ansicht keine Modellansicht ist (Legende, Zeichenansicht)
    oder die API das nicht kann (vor Revit 2022).
    """
    punkt = punkt or XYZ(0, 0, 0)
    try:
        trafos = ansicht.GetModelToProjectionTransforms()
        if not trafos or trafos.Count == 0:
            return None
        projektion = trafos[0].GetModelToProjectionTransform().OfPoint(punkt)
        auf_plan = fenster.GetProjectionToSheetTransform().OfPoint(projektion)
        return XYZ(auf_plan.X, auf_plan.Y, 0)
    except Exception:
        return None


def fenster_eigenschaften(fenster):
    """Was beim Neu-Platzieren erhalten bleiben soll."""
    werte = {"typ": fenster.GetTypeId(), "mitte": fenster.GetBoxCenter(),
             "drehung": fenster.Rotation, "nummer": None,
             "titel_versatz": None, "titel_laenge": None}
    param = fenster.get_Parameter(BIP.VIEWPORT_DETAIL_NUMBER)
    if param is not None:
        werte["nummer"] = param.AsString()
    try:
        werte["titel_versatz"] = fenster.LabelOffset
        werte["titel_laenge"] = fenster.LabelLineLength
    except Exception:
        pass
    return werte


def eigenschaften_setzen(doc, fenster, werte):
    """Gegenstück zu fenster_eigenschaften(). Liefert Hinweistexte."""
    hinweise = []
    if id_wert(fenster.GetTypeId()) != id_wert(werte["typ"]):
        fenster.ChangeTypeId(werte["typ"])
    if fenster.Rotation != werte["drehung"]:
        try:
            fenster.Rotation = werte["drehung"]
        except Exception:
            pass
    doc.Regenerate()
    fenster.SetBoxCenter(werte["mitte"])
    if werte["titel_versatz"] is not None:
        try:
            fenster.LabelOffset = werte["titel_versatz"]
            fenster.LabelLineLength = werte["titel_laenge"]
        except Exception:
            pass
    param = fenster.get_Parameter(BIP.VIEWPORT_DETAIL_NUMBER)
    if werte["nummer"] and param is not None and param.AsString() != werte["nummer"]:
        # Eine doppelte Detailnummer lehnt Revit erst beim Abschliessen der
        # Transaktion mit einem Fehler ab, der alles zurücknimmt - deshalb
        # vorher prüfen und die von Revit vergebene Nummer behalten.
        if werte["nummer"].strip().lower() in detailnummern(doc, fenster):
            hinweise.append(t(u"Detailnummer {} ist auf dem Zielplan schon vergeben "
                              u"- neue Nummer {}",
                              u"detail number {} is already used on the target "
                              u"sheet - new number {}",
                              u"el número de detalle {} ya existe en el plano de "
                              u"destino: nuevo número {}")
                            .format(werte["nummer"], param.AsString()))
        else:
            param.Set(werte["nummer"])
    return hinweise


def detailnummern(doc, fenster):
    """Detailnummern der anderen Ansichtsfenster auf dem Plan von fenster."""
    plan = doc.GetElement(fenster.SheetId)
    nummern = set()
    for fid in plan.GetAllViewports():
        if id_wert(fid) == id_wert(fenster.Id):
            continue
        param = doc.GetElement(fid).get_Parameter(BIP.VIEWPORT_DETAIL_NUMBER)
        if param is not None and param.AsString():
            nummern.add(param.AsString().strip().lower())
    return nummern


class NurAnsichtsfenster(ISelectionFilter):
    """Auswahlfilter für PickObject(s): nur Ansichtsfenster."""

    def AllowElement(self, elem):
        return isinstance(elem, Viewport)

    def AllowReference(self, ref, punkt):
        return False


# ---------------------------------------------------------------- Werte für Muster

def ansichts_werte(ansicht):
    """Funktion Name -> Text für logik.muster_anwenden().

    Feste Namen (deutsch/englisch/spanisch): Ansicht, Ebene, Massstab;
    sonst ein Parameter der Ansicht mit diesem Namen.
    """
    def wert(name):
        schluessel = name.lower()
        if schluessel in (u"ansicht", u"view", u"vista"):
            return ansicht.Name
        if schluessel in (u"ebene", u"level", u"nivel"):
            ebene = getattr(ansicht, "GenLevel", None)
            return ebene.Name if ebene is not None else u""
        if schluessel in (u"massstab", u"maßstab", u"scale", u"escala"):
            try:
                return str(ansicht.Scale)
            except Exception:
                return u""
        param = ansicht.LookupParameter(name)
        if param is None:
            return None
        return param.AsString() or param.AsValueString() or u""
    return wert
