# -*- coding: utf-8 -*-
"""Modellansichten (Grundrisse, Schnitte, Ansichten, 3D) für TransferSingle.

Revit kann Modellansichten nicht zwischen Dokumenten kopieren. Sie werden im
Ziel neu angelegt und Schritt für Schritt angeglichen:

    1. Anlegen          Grundriss/Decke/Tragwerk: ViewPlan.Create auf der
                        passenden Ebene (gleicher Name, sonst gleiche Höhe)
                        Schnitt/Detail: CreateSection/CreateDetail mit
                        demselben Rahmen und derselben Tiefe
                        Ansicht: Ansichtsmarker, gedreht und verschoben
                        3D: Isometrie/Perspektive mit derselben Kamera
                        abhängige Ansicht: Hauptansicht übertragen, dann
                        Duplicate(AsDependent)
    2. Zuschnitt        Zuschnittform (Modellkoordinaten), bei 3D CropBox;
                        Beschriftungszuschnitt; 3D-Schnittbereich
    3. Ansichtsbereich  Ebenen über Name/Höhe, Versätze
    4. Parameter        alle beschreibbaren; Element-Ids (Phase, Phasenfilter,
                        Zuschnittbereich, Unterlage ...) über den Namen
    5. Sichtbarkeit     Kategorien (aus, Überschreibungen), Filter samt
                        Überschreibungen und Reihenfolge, Bearbeitungsbereiche
    6. Vorlage          gleichnamige im Ziel, sonst mitkopiert - zuletzt, damit
                        Werte, die sie nicht steuert, erhalten bleiben
    7. Ansichtselemente View -> View (Texte, Maße, Detaillinien,
                        Beschriftungen; was sich auf im Ziel fehlende Bauteile
                        bezieht, lässt Revit weg)

Die Transformation der Übertragung (keine / Verknüpfung / gemeinsame
Koordinaten) gilt für Lage, Ebenenhöhen, Zuschnitt, Kamera und Schnittbereich.

Nicht übertragen: in der Ansicht ausgeblendete Elemente und
Element-Überschreibungen (die Elemente sind im Ziel andere),
Verknüpfungs-Anzeigeeinstellungen und die Bindung von Ausschnitten an ihre
Bezugsansicht (ein Ausschnitt wird eine eigenständige Ansicht).
"""

from Autodesk.Revit.DB import (  # noqa: E402
    BoundingBoxXYZ,
    BuiltInCategory,
    BuiltInParameter,
    Category,
    Color,
    CurveLoop,
    ElementId,
    ElementTransformUtils,
    ElevationMarker,
    FillPatternElement,
    FilteredElementCollector,
    FilteredWorksetCollector,
    Level,
    Line,
    LinePatternElement,
    OverrideGraphicSettings,
    ParameterFilterElement,
    PlanViewPlane,
    StorageType,
    Transform,
    View,
    View3D,
    ViewDuplicateOption,
    ViewFamily,
    ViewFamilyType,
    ViewOrientation3D,
    ViewPlan,
    ViewSection,
    ViewType,
    WorksetKind,
    XYZ,
)

from mlg_sprache import t  # noqa: E402
from transfer_single import logik as lg  # noqa: E402
from transfer_single import revit as rv  # noqa: E402

# Werden eigens behandelt oder sind nicht übertragbar
_UEBERSPRINGEN = set([
    "VIEW_NAME", "VIEW_TEMPLATE", "VIEW_DEPENDENCY", "ELEM_TYPE_PARAM",
    "ELEM_FAMILY_AND_TYPE_PARAM", "ELEM_FAMILY_PARAM", "SYMBOL_ID_PARAM",
])

_MUSTER = (
    ("ProjectionLinePatternId", "SetProjectionLinePatternId"),
    ("CutLinePatternId", "SetCutLinePatternId"),
    ("SurfaceForegroundPatternId", "SetSurfaceForegroundPatternId"),
    ("SurfaceBackgroundPatternId", "SetSurfaceBackgroundPatternId"),
    ("CutForegroundPatternId", "SetCutForegroundPatternId"),
    ("CutBackgroundPatternId", "SetCutBackgroundPatternId"),
)

_KATEGORIEN_AUS = ("AreModelCategoriesHidden", "AreAnnotationCategoriesHidden",
                   "AreAnalyticalModelCategoriesHidden",
                   "AreImportCategoriesHidden", "ArePointCloudsHidden")

_GENAU = 1e-4       # Fuß


def _hinweis(hinweise, text):
    if text not in hinweise:
        hinweise.append(text)


# ---------------------------------------------------------------------------
# Einstieg
# ---------------------------------------------------------------------------

def kopiere(k, ansicht, hinweise):
    """Modellansicht im Ziel neu aufbauen (einmal je Lauf).
    k: revit.Kontext. Rückgabe: neue Ansicht."""
    schluessel = rv.id_wert(ansicht.Id)
    if schluessel in k.kopien:
        return k.kopien[schluessel]
    art = lg.modellansicht_art(ansicht.ViewType)
    if art is None:
        raise Exception(t(u"Ansichtsart %s wird nicht unterstützt",
                          u"view type %s is not supported",
                          u"el tipo de vista %s no es compatible")
                        % ansicht.ViewType)

    primaer_id = ansicht.GetPrimaryViewId()
    abhaengig = rv.id_wert(primaer_id) != -1
    if abhaengig:
        primaer = k.quelle.GetElement(primaer_id)
        if rv.id_wert(primaer_id) not in k.kopien:
            _hinweis(hinweise, t(
                u"Hauptansicht '%s' mit übertragen",
                u"primary view '%s' transferred as well",
                u"vista principal '%s' transferida también") % primaer.Name)
        neue_primaer = kopiere(k, primaer, hinweise)
        neue = k.ziel.GetElement(
            neue_primaer.Duplicate(ViewDuplicateOption.AsDependent))
    else:
        neue = _anlegen(k, ansicht, art, hinweise)
        if rv.id_wert(neue.ViewTemplateId) != -1:
            # Vorlage des Ansichtstyps - erst am Ende die richtige zuweisen
            neue.ViewTemplateId = ElementId.InvalidElementId

    _setze_namen(neue, ansicht.Name, hinweise)
    _zuschnitt(k, ansicht, neue, art, hinweise)
    if art in lg.PLAENE and not abhaengig:
        _ansichtsbereich(k, ansicht, neue, hinweise)
    _parameter(k, ansicht, neue, hinweise)
    if not abhaengig:       # abhängige Ansichten folgen ihrer Hauptansicht
        _sichtbarkeit(k, ansicht, neue)
        _filter(k, ansicht, neue, hinweise)
        _bearbeitungsbereiche(k, ansicht, neue)
        _vorlage(k, ansicht, neue, hinweise)
    if art == lg.DREI_D and ansicht.IsLocked:
        try:
            neue.SaveOrientationAndLock()
        except Exception as fehler:
            _hinweis(hinweise, t(u"Ansicht nicht gesperrt: %s",
                                 u"view not locked: %s",
                                 u"vista no bloqueada: %s")
                     % rv.fehlertext(fehler))
    if k.optionen.ansichtselemente and art != lg.DREI_D:
        _ansichtselemente(k, ansicht, neue, art, hinweise)
    k.kopien[schluessel] = neue
    return neue


# ---------------------------------------------------------------------------
# 1. Anlegen
# ---------------------------------------------------------------------------

def _anlegen(k, ansicht, art, hinweise):
    quell_typ = k.quelle.GetElement(ansicht.GetTypeId())
    typ_id = _ansichtstyp(k, quell_typ, hinweise)
    if art in lg.PLAENE:
        ebene_id = _ebene(k, ansicht.GenLevel, hinweise)
        if ebene_id is None:
            raise Exception(t(
                u"Ebene '%s' fehlt im Ziel (weder Name noch Höhe passen)",
                u"level '%s' missing in the target (neither name nor "
                u"elevation match)",
                u"falta el nivel '%s' en el destino (no coincide ni el "
                u"nombre ni la altura)") % ansicht.GenLevel.Name)
        return ViewPlan.Create(k.ziel, typ_id, ebene_id)
    if art in (lg.SCHNITT, lg.DETAIL):
        return _schnitt(k, ansicht, typ_id,
                        quell_typ.ViewFamily == ViewFamily.Detail)
    if art == lg.ANSICHT:
        return _ansicht(k, ansicht, typ_id)
    return _drei_d(k, ansicht, typ_id)


def _ansichtstyp(k, quell_typ, hinweise):
    """Gleichnamiger Ansichtstyp derselben Familie, sonst mitkopiert, sonst
    der erste dieser Familie."""
    name = rv._text(lambda: quell_typ.Name)
    typen = [v for v in rv._alle(k.ziel, ViewFamilyType)
             if v.ViewFamily == quell_typ.ViewFamily]
    for typ in typen:
        if rv._text(lambda: typ.Name) == name:
            return typ.Id
    neue_ids, fehler = rv._einheit(k.ziel, lambda: ElementTransformUtils
                                   .CopyElements(k.quelle,
                                                 rv.net_ids([quell_typ.Id]),
                                                 k.ziel, Transform.Identity,
                                                 rv._optionen()))
    if fehler is None and neue_ids is not None and neue_ids.Count:
        _hinweis(hinweise, t(u"Ansichtstyp '%s' mit übertragen",
                             u"view type '%s' transferred as well",
                             u"tipo de vista '%s' transferido también") % name)
        return list(neue_ids)[0]
    if typen:
        _hinweis(hinweise, t(
            u"Ansichtstyp '%s' fehlt im Ziel - '%s' verwendet",
            u"view type '%s' missing in the target - '%s' used",
            u"falta el tipo de vista '%s' en el destino - se usa '%s'")
            % (name, rv._text(lambda: typen[0].Name)))
        return typen[0].Id
    raise Exception(t(u"kein Ansichtstyp '%s' im Ziel",
                      u"no view type '%s' in the target",
                      u"no hay tipo de vista '%s' en el destino") % name)


def _ebenen(k):
    if k.ebenen is None:
        k.ebenen = [(rv._text(lambda: e.Name), e.ProjectElevation, e.Id)
                    for e in rv._alle(k.ziel, Level)]
    return k.ebenen


def _ebene(k, quell_ebene, hinweise):
    """Id der Ebene im Ziel oder None."""
    hoehe = quell_ebene.ProjectElevation + k.transform.Origin.Z
    ebene_id, wie = lg.passende_ebene(quell_ebene.Name, hoehe, _ebenen(k))
    if wie == "hoehe":
        _hinweis(hinweise, t(
            u"Ebene '%s' über die Höhe zugeordnet: '%s'",
            u"level '%s' matched by elevation: '%s'",
            u"nivel '%s' asignado por altura: '%s'")
            % (quell_ebene.Name, k.ziel.GetElement(ebene_id).Name))
    return ebene_id


def _rahmen(k, ansicht):
    """Ursprung, Rechts, Oben, Blickrichtung der Ansicht im Ziel."""
    tr = k.transform
    return (tr.OfPoint(ansicht.Origin), tr.OfVector(ansicht.RightDirection),
            tr.OfVector(ansicht.UpDirection), tr.OfVector(ansicht.ViewDirection))


def _zuschnitt_punkte(k, ansicht):
    try:
        schleifen = list(ansicht.GetCropRegionShapeManager().GetCropShape())
    except Exception:
        return []
    return [k.transform.OfPoint(kurve.GetEndPoint(0))
            for schleife in schleifen for kurve in schleife]


def _schnitt(k, ansicht, typ_id, detail):
    """Schnitt mit demselben Rahmen. Wo Revit die Schnittebene in den Rahmen
    legt, wird nachgemessen und notfalls einmal korrigiert neu angelegt."""
    ursprung, rechts, oben, blick = _rahmen(k, ansicht)
    punkte = _zuschnitt_punkte(k, ansicht)
    breite = max([abs((p - ursprung).DotProduct(rechts)) for p in punkte]
                 or [0.0]) or 10.0
    hoehe = max([abs((p - ursprung).DotProduct(oben)) for p in punkte]
                or [0.0]) or 10.0
    tiefe = 10.0
    try:
        tiefe = ansicht.get_Parameter(
            BuiltInParameter.VIEWER_BOUND_OFFSET_FAR).AsDouble() or 10.0
    except Exception:
        pass
    erzeuge = ViewSection.CreateDetail if detail else ViewSection.CreateSection

    def anlegen(versatz):
        rahmen = Transform.Identity
        rahmen.Origin = ursprung
        rahmen.BasisX = rechts
        rahmen.BasisY = oben
        rahmen.BasisZ = blick
        box = BoundingBoxXYZ()
        box.Transform = rahmen
        box.Min = XYZ(-breite, -hoehe, versatz - max(tiefe, 0.1))
        box.Max = XYZ(breite, hoehe, versatz)
        neue = erzeuge(k.ziel, typ_id, box)
        k.ziel.Regenerate()
        return neue

    neue = anlegen(0.0)
    abweichung = (neue.Origin - ursprung).DotProduct(blick)
    if abs(abweichung) > _GENAU:
        k.ziel.Delete(neue.Id)
        neue = anlegen(-abweichung)
    return neue


def _grundriss_fuer_marker(k, hoehe):
    plaene = [v for v in rv._alle(k.ziel, ViewPlan)
              if not v.IsTemplate and v.ViewType == ViewType.FloorPlan
              and v.GenLevel is not None]
    if not plaene:
        raise Exception(t(
            u"für Ansichten braucht das Ziel mindestens einen Grundriss",
            u"elevations need at least one floor plan in the target",
            u"los alzados necesitan al menos una planta en el destino"))
    return min(plaene, key=lambda v: abs(v.GenLevel.ProjectElevation - hoehe))


def _ansicht(k, ansicht, typ_id):
    """Ansichtsmarker setzen, Ansicht anlegen, Marker in Blickrichtung drehen
    und so verschieben, dass die Ansichtsebene stimmt."""
    ursprung, _r, _o, blick = _rahmen(k, ansicht)
    grundriss = _grundriss_fuer_marker(k, ursprung.Z)
    punkt = XYZ(ursprung.X, ursprung.Y, grundriss.GenLevel.ProjectElevation)
    marker = ElevationMarker.CreateElevationMarker(k.ziel, typ_id, punkt,
                                                   ansicht.Scale)
    neue = marker.CreateElevation(k.ziel, grundriss.Id, 0)
    k.ziel.Regenerate()
    ist = neue.ViewDirection
    winkel = lg.drehwinkel((ist.X, ist.Y), (blick.X, blick.Y))
    if abs(winkel) > 1e-9:
        ElementTransformUtils.RotateElement(
            k.ziel, marker.Id, Line.CreateBound(punkt, punkt + XYZ.BasisZ),
            winkel)
        k.ziel.Regenerate()
    abstand = (ursprung - neue.Origin).DotProduct(neue.ViewDirection)
    if abs(abstand) > _GENAU:
        ElementTransformUtils.MoveElement(
            k.ziel, marker.Id, neue.ViewDirection.Multiply(abstand))
        k.ziel.Regenerate()
    return neue


def _drei_d(k, ansicht, typ_id):
    if ansicht.IsPerspective:
        neue = View3D.CreatePerspective(k.ziel, typ_id)
    else:
        neue = View3D.CreateIsometric(k.ziel, typ_id)
    tr = k.transform
    lage = ansicht.GetOrientation()
    neue.SetOrientation(ViewOrientation3D(tr.OfPoint(lage.EyePosition),
                                          tr.OfVector(lage.UpDirection),
                                          tr.OfVector(lage.ForwardDirection)))
    if ansicht.IsSectionBoxActive:
        box = ansicht.GetSectionBox()
        box.Transform = tr.Multiply(box.Transform)
        neue.SetSectionBox(box)
    else:
        neue.IsSectionBoxActive = False
    k.ziel.Regenerate()
    return neue


def _setze_namen(neue, wunsch, hinweise):
    """Name der Quelle, bei Vergabe 'Name (2)', 'Name (3)' ..."""
    for nummer in range(1, 200):
        kandidat = lg.name_mit_nummer(wunsch, nummer)
        if neue.Name == kandidat:
            return
        try:
            neue.Name = kandidat
        except Exception:
            continue
        if nummer > 1:
            _hinweis(hinweise, t(u"Name im Ziel vergeben - neu: '%s'",
                                 u"name already used in the target - new: '%s'",
                                 u"nombre ya usado en el destino - nuevo: '%s'")
                     % kandidat)
        return


# ---------------------------------------------------------------------------
# 2. Zuschnitt, 3. Ansichtsbereich
# ---------------------------------------------------------------------------

def _zuschnitt(k, ansicht, neue, art, hinweise):
    if art == lg.DREI_D:
        quell_box = ansicht.CropBox
        if quell_box is not None:
            box = neue.CropBox
            box.Min = quell_box.Min
            box.Max = quell_box.Max
            neue.CropBox = box
    else:
        quell_mgr = ansicht.GetCropRegionShapeManager()
        ziel_mgr = neue.GetCropRegionShapeManager()
        if quell_mgr.Split:
            _hinweis(hinweise, t(
                u"geteilter Zuschnitt wird nicht übertragen",
                u"split crop region is not transferred",
                u"la región de recorte dividida no se transfiere"))
        elif ziel_mgr.CanHaveShape:
            schleifen = list(quell_mgr.GetCropShape())
            if schleifen:
                ziel_mgr.SetCropShape(_in_ansichtsebene(k, schleifen[0], neue,
                                                        ziel_mgr))
        _beschriftungszuschnitt(ansicht, neue, quell_mgr, ziel_mgr)
    neue.CropBoxActive = ansicht.CropBoxActive
    neue.CropBoxVisible = ansicht.CropBoxVisible


def _in_ansichtsebene(k, schleife, neue, ziel_mgr):
    """Zuschnittform der Quelle transformieren und in die Zuschnittebene der
    neuen Ansicht schieben (die Form muss parallel zur Ansichtsebene liegen,
    in der Höhe bzw. Tiefe darf sie nicht abweichen)."""
    kurven = [kurve.CreateTransformed(k.transform) for kurve in schleife]
    vorhanden = list(ziel_mgr.GetCropShape())
    if kurven and vorhanden:
        blick = neue.ViewDirection
        bezug = list(vorhanden[0])[0].GetEndPoint(0)
        versatz = (bezug - kurven[0].GetEndPoint(0)).DotProduct(blick)
        if abs(versatz) > 1e-9:
            schieben = Transform.CreateTranslation(blick.Multiply(versatz))
            kurven = [kurve.CreateTransformed(schieben) for kurve in kurven]
    neu = CurveLoop()
    for kurve in kurven:
        neu.Append(kurve)
    return neu


def _beschriftungszuschnitt(ansicht, neue, quell_mgr, ziel_mgr):
    try:
        aktiv = ansicht.get_Parameter(
            BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE).AsInteger()
        neue.get_Parameter(
            BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE).Set(aktiv)
    except Exception:
        return
    if not aktiv:
        return
    for seite in ("Left", "Right", "Top", "Bottom"):
        name = seite + "AnnotationCropOffset"
        try:
            setattr(ziel_mgr, name, getattr(quell_mgr, name))
        except Exception:
            pass


def _ansichtsbereich(k, ansicht, neue, hinweise):
    quell = ansicht.GetViewRange()
    bereich = neue.GetViewRange()
    for ebene in (PlanViewPlane.TopClipPlane, PlanViewPlane.CutPlane,
                  PlanViewPlane.BottomClipPlane, PlanViewPlane.ViewDepthPlane):
        ebene_id = quell.GetLevelId(ebene)
        if rv.id_wert(ebene_id) == rv.id_wert(ansicht.GenLevel.Id):
            ziel_id = neue.GenLevel.Id
        elif rv.id_wert(ebene_id) < 0:
            ziel_id = ebene_id      # "Ebene darüber", "Unbegrenzt" ...
        else:
            quell_ebene = k.quelle.GetElement(ebene_id)
            ziel_id = _ebene(k, quell_ebene, hinweise) \
                if quell_ebene is not None else None
            if ziel_id is None:
                _hinweis(hinweise, t(
                    u"Ansichtsbereich: Ebene '%s' fehlt im Ziel",
                    u"view range: level '%s' missing in the target",
                    u"rango de vista: falta el nivel '%s' en el destino")
                    % rv._text(lambda: quell_ebene.Name, u"?"))
                continue
        bereich.SetLevelId(ebene, ziel_id)
        bereich.SetOffset(ebene, quell.GetOffset(ebene))
    try:
        neue.SetViewRange(bereich)
    except Exception as fehler:
        _hinweis(hinweise, t(u"Ansichtsbereich nicht übernommen: %s",
                             u"view range not transferred: %s",
                             u"rango de vista no transferido: %s")
                 % rv.fehlertext(fehler))


# ---------------------------------------------------------------------------
# 4. Parameter
# ---------------------------------------------------------------------------

def gegenstueck(k, quell_id):
    """Id des gleichnamigen Elements (gleiche Klasse, gleiche Kategorie) im
    Ziel. Negative Ids (eingebaut, 'keine') gelten in jedem Dokument.
    None, wenn es im Ziel keines gibt."""
    if rv.id_wert(quell_id) < 0:
        return quell_id
    element = k.quelle.GetElement(quell_id)
    if element is None:
        return None
    if isinstance(element, Level):
        return _ebene(k, element, [])
    name = rv._text(lambda: element.Name)
    kategorie = rv.id_wert(element.Category.Id) \
        if element.Category is not None else None
    schluessel = (element.GetType().FullName, kategorie, name)
    if schluessel in k.namen:
        return k.namen[schluessel]
    try:
        kandidaten = FilteredElementCollector(k.ziel).OfClass(element.GetType())
    except Exception:       # Klassen wie Element lassen sich nicht filtern
        if element.Category is None:
            return None
        kandidaten = FilteredElementCollector(k.ziel).OfCategoryId(
            element.Category.Id)
    gefunden = None
    for kandidat in kandidaten:
        if rv._text(lambda: kandidat.Name) != name:
            continue
        if kategorie is not None and (kandidat.Category is None or rv.id_wert(
                kandidat.Category.Id) != kategorie):
            continue
        gefunden = kandidat.Id
        break
    k.namen[schluessel] = gefunden
    return gefunden


def _bip_name(parameter):
    try:
        bip = parameter.Definition.BuiltInParameter
    except Exception:
        return None
    name = u"%s" % bip
    return None if name == u"INVALID" else name


def _gegenparameter(neue, parameter):
    try:
        bip = parameter.Definition.BuiltInParameter
        if u"%s" % bip != u"INVALID":
            return neue.get_Parameter(bip)
        if parameter.IsShared:
            return neue.get_Parameter(parameter.GUID)
        return neue.LookupParameter(parameter.Definition.Name)
    except Exception:
        return None


def _setze_wert(k, quell_p, ziel_p):
    """True = übernommen oder schon gleich, False = Element fehlt im Ziel."""
    art = quell_p.StorageType
    if art == StorageType.String:
        wert = quell_p.AsString() or u""
        if (ziel_p.AsString() or u"") != wert:
            ziel_p.Set(wert)
    elif art == StorageType.Integer:
        wert = quell_p.AsInteger()
        if ziel_p.AsInteger() != wert:
            ziel_p.Set(wert)
    elif art == StorageType.Double:
        wert = quell_p.AsDouble()
        if abs(ziel_p.AsDouble() - wert) > 1e-9:
            ziel_p.Set(wert)
    elif art == StorageType.ElementId:
        wert = gegenstueck(k, quell_p.AsElementId())
        if wert is None:
            return False
        if rv.id_wert(ziel_p.AsElementId()) != rv.id_wert(wert):
            ziel_p.Set(wert)
    return True


def _parameter(k, ansicht, neue, hinweise):
    offen = []
    for quell_p in ansicht.Parameters:
        try:
            if quell_p.IsReadOnly or not quell_p.HasValue \
                    or _bip_name(quell_p) in _UEBERSPRINGEN:
                continue
        except Exception:
            continue
        ziel_p = _gegenparameter(neue, quell_p)
        if ziel_p is None or ziel_p.IsReadOnly \
                or ziel_p.StorageType != quell_p.StorageType:
            continue
        offen.append((quell_p, ziel_p))
    fehlend = []
    # Zweiter Durchgang für Werte, die von anderen abhängen
    for _durchgang in range(2):
        rest = []
        for quell_p, ziel_p in offen:
            try:
                if not _setze_wert(k, quell_p, ziel_p):
                    fehlend.append(u"%s (%s)" % (
                        quell_p.Definition.Name,
                        rv._text(lambda: quell_p.AsValueString(), u"?")))
            except Exception:
                rest.append((quell_p, ziel_p))
        offen = rest
        if not offen:
            break
    if fehlend:
        _hinweis(hinweise, t(
            u"im Ziel nicht vorhanden: %s",
            u"missing in the target: %s",
            u"no existe en el destino: %s") % u", ".join(fehlend))
    if offen:
        _hinweis(hinweise, t(
            u"Parameter nicht übernommen: %s",
            u"parameters not transferred: %s",
            u"parámetros no transferidos: %s")
            % u", ".join(rv._text(lambda: p.Definition.Name, u"?")
                         for p, _z in offen))


# ---------------------------------------------------------------------------
# 5. Sichtbarkeit/Grafiken
# ---------------------------------------------------------------------------

def _fingerabdruck(ogs):
    """Vergleichbare Werte einer OverrideGraphicSettings (alle Eigenschaften
    per Reflection, Farben als RGB)."""
    werte = []
    for eigenschaft in ogs.GetType().GetProperties():
        try:
            wert = eigenschaft.GetValue(ogs, None)
        except Exception:
            wert = None
        if isinstance(wert, Color):
            wert = (wert.Red, wert.Green, wert.Blue) if wert.IsValid else None
        elif isinstance(wert, ElementId):
            wert = rv.id_wert(wert)
        else:
            wert = u"%s" % wert
        werte.append((eigenschaft.Name, wert))
    return tuple(werte)


_STANDARD = []


def _ist_standard(ogs):
    if not _STANDARD:
        _STANDARD.append(_fingerabdruck(OverrideGraphicSettings()))
    return _fingerabdruck(ogs) == _STANDARD[0]


def _muster(k, muster_id):
    """Linien-/Füllmuster im Ziel (Name, Volltonfüllung), sonst mitkopiert."""
    schluessel = rv.id_wert(muster_id)
    if schluessel in k.muster:
        return k.muster[schluessel]
    element = k.quelle.GetElement(muster_id)
    gefunden = None
    try:
        if isinstance(element, FillPatternElement):
            muster = element.GetFillPattern()
            if muster.IsSolidFill:
                for kandidat in rv._alle(k.ziel, FillPatternElement):
                    if kandidat.GetFillPattern().IsSolidFill:
                        gefunden = kandidat.Id
                        break
            else:
                kandidat = FillPatternElement.GetFillPatternElementByName(
                    k.ziel, muster.Target, element.Name)
                gefunden = kandidat.Id if kandidat is not None else None
        elif isinstance(element, LinePatternElement):
            kandidat = LinePatternElement.GetLinePatternElementByName(
                k.ziel, element.Name)
            gefunden = kandidat.Id if kandidat is not None else None
    except Exception:
        gefunden = None
    if gefunden is None and element is not None:
        neue_ids, fehler = rv._einheit(k.ziel, lambda: ElementTransformUtils
                                       .CopyElements(k.quelle,
                                                     rv.net_ids([muster_id]),
                                                     k.ziel, Transform.Identity,
                                                     rv._optionen()))
        if fehler is None and neue_ids is not None and neue_ids.Count:
            gefunden = list(neue_ids)[0]
    k.muster[schluessel] = gefunden
    return gefunden


def _ogs(k, quell):
    """Überschreibungen für das Ziel - Muster-Ids umgeschlüsselt."""
    neu = OverrideGraphicSettings(quell)
    for lesen, setzen in _MUSTER:
        muster_id = getattr(quell, lesen)
        if rv.id_wert(muster_id) > 0:   # negative: Volllinie u.ä.
            getattr(neu, setzen)(_muster(k, muster_id)
                                 or ElementId.InvalidElementId)
    return neu


def _ziel_kategorie(k, kategorie):
    try:
        if rv.id_wert(kategorie.Id) < 0:
            return Category.GetCategory(k.ziel, kategorie.Id)
    except Exception:
        return None
    for kandidat in k.ziel.Settings.Categories:
        if kandidat.Name == kategorie.Name:
            return kandidat
    return None


def _kategorien(k):
    """[(Kategorie der Quelle, Id im Ziel)] samt Unterkategorien."""
    if k.kategorien is None:
        paare = []
        for kategorie in k.quelle.Settings.Categories:
            ziel_kat = _ziel_kategorie(k, kategorie)
            if ziel_kat is None:
                continue
            paare.append((kategorie, ziel_kat.Id))
            ziel_unter = dict((u.Name, u) for u in ziel_kat.SubCategories)
            for unter in kategorie.SubCategories:
                gegen = None
                if rv.id_wert(unter.Id) < 0:
                    gegen = _ziel_kategorie(k, unter)
                if gegen is None:
                    gegen = ziel_unter.get(unter.Name)
                if gegen is not None:
                    paare.append((unter, gegen.Id))
        k.kategorien = paare
    return k.kategorien


def _sichtbarkeit(k, ansicht, neue):
    for name in _KATEGORIEN_AUS:
        try:
            wert = getattr(ansicht, name)
            if getattr(neue, name) != wert:
                setattr(neue, name, wert)
        except Exception:
            pass
    for kategorie, ziel_id in _kategorien(k):
        try:
            if neue.CanCategoryBeHidden(ziel_id):
                aus = ansicht.GetCategoryHidden(kategorie.Id)
                if neue.GetCategoryHidden(ziel_id) != aus:
                    neue.SetCategoryHidden(ziel_id, aus)
        except Exception:
            pass
        try:
            if not neue.IsCategoryOverridable(ziel_id):
                continue
            # Die neue Ansicht hat (ohne Vorlage) keine Überschreibungen
            quell = ansicht.GetCategoryOverrides(kategorie.Id)
            if _ist_standard(quell):
                continue
            neue.SetCategoryOverrides(ziel_id, _ogs(k, quell))
        except Exception:
            pass


def _filter_im_ziel(k, filter_element, hinweise):
    ziel_id = gegenstueck(k, filter_element.Id)
    if ziel_id is not None:
        return ziel_id
    neue_ids, fehler = rv._einheit(k.ziel, lambda: ElementTransformUtils
                                   .CopyElements(k.quelle,
                                                 rv.net_ids([filter_element.Id]),
                                                 k.ziel, Transform.Identity,
                                                 rv._optionen()))
    name = rv._text(lambda: filter_element.Name, u"?")
    if fehler is not None or neue_ids is None or not neue_ids.Count:
        _hinweis(hinweise, t(u"Filter '%s' nicht übertragen: %s",
                             u"filter '%s' not transferred: %s",
                             u"filtro '%s' no transferido: %s")
                 % (name, fehler or u"?"))
        return None
    ziel_id = list(neue_ids)[0]
    k.namen[(filter_element.GetType().FullName, None, name)] = ziel_id
    text = t(u"Filter '%s' mit übertragen", u"filter '%s' transferred as well",
             u"filtro '%s' transferido también") % name
    if isinstance(filter_element, ParameterFilterElement):
        fehlend = rv._fehlende_filterparameter(k.quelle, filter_element, k.ziel)
        if fehlend:
            text += t(u" - Parameter fehlen im Ziel: %s",
                      u" - parameters missing in the target: %s",
                      u" - faltan parámetros en el destino: %s") \
                % u", ".join(fehlend)
    _hinweis(hinweise, text)
    return ziel_id


def _filter(k, ansicht, neue, hinweise):
    try:
        filter_ids = list(ansicht.GetOrderedFilters())
    except Exception:
        filter_ids = list(ansicht.GetFilters())
    for filter_id in filter_ids:
        filter_element = k.quelle.GetElement(filter_id)
        if filter_element is None:
            continue
        ziel_id = _filter_im_ziel(k, filter_element, hinweise)
        if ziel_id is None:
            continue
        try:
            if not neue.IsFilterApplied(ziel_id):
                neue.AddFilter(ziel_id)
            neue.SetFilterOverrides(
                ziel_id, _ogs(k, ansicht.GetFilterOverrides(filter_id)))
            neue.SetFilterVisibility(ziel_id,
                                     ansicht.GetFilterVisibility(filter_id))
            neue.SetIsFilterEnabled(ziel_id,
                                    ansicht.GetIsFilterEnabled(filter_id))
        except Exception as fehler:
            _hinweis(hinweise, t(u"Filter '%s': %s", u"filter '%s': %s",
                                 u"filtro '%s': %s")
                     % (rv._text(lambda: filter_element.Name, u"?"),
                        rv.fehlertext(fehler)))


def _bearbeitungsbereiche(k, ansicht, neue):
    if not (k.quelle.IsWorkshared and k.ziel.IsWorkshared):
        return
    im_ziel = dict((w.Name, w.Id) for w in FilteredWorksetCollector(k.ziel)
                   .OfKind(WorksetKind.UserWorkset))
    for bereich in FilteredWorksetCollector(k.quelle).OfKind(
            WorksetKind.UserWorkset):
        if bereich.Name not in im_ziel:
            continue
        try:
            neue.SetWorksetVisibility(im_ziel[bereich.Name],
                                      ansicht.GetWorksetVisibility(bereich.Id))
        except Exception:
            pass


# ---------------------------------------------------------------------------
# 6. Vorlage, 7. Ansichtselemente
# ---------------------------------------------------------------------------

def _vorlage(k, ansicht, neue, hinweise):
    vorlage_id = ansicht.ViewTemplateId
    if rv.id_wert(vorlage_id) == -1:
        return
    vorlage = k.quelle.GetElement(vorlage_id)
    name = rv._text(lambda: vorlage.Name, u"?")
    ziel_id = next((v.Id for v in rv._alle(k.ziel, View)
                    if v.IsTemplate and rv._text(lambda: v.Name) == name),
                   None)
    if ziel_id is None:
        neue_ids, fehler = rv._einheit(k.ziel, lambda: ElementTransformUtils
                                       .CopyElements(k.quelle,
                                                     rv.net_ids([vorlage_id]),
                                                     k.ziel, Transform.Identity,
                                                     rv._optionen()))
        if fehler is not None or neue_ids is None or not neue_ids.Count:
            _hinweis(hinweise, t(
                u"Ansichtsvorlage '%s' nicht übertragen: %s",
                u"view template '%s' not transferred: %s",
                u"plantilla de vista '%s' no transferida: %s")
                % (name, fehler or u"?"))
            return
        ziel_id = list(neue_ids)[0]
        _hinweis(hinweise, t(u"Ansichtsvorlage '%s' mit übertragen",
                             u"view template '%s' transferred as well",
                             u"plantilla de vista '%s' transferida también")
                 % name)
    try:
        neue.ViewTemplateId = ziel_id
    except Exception as fehler:
        _hinweis(hinweise, t(u"Ansichtsvorlage '%s' nicht zuweisbar: %s",
                             u"view template '%s' cannot be assigned: %s",
                             u"no se puede asignar la plantilla '%s': %s")
                 % (name, rv.fehlertext(fehler)))


def _in_der_ebene(transform):
    """Anteil der Transformation in der Grundrissebene (ohne Höhenversatz -
    den Ebenenwechsel erledigt Revit beim Kopieren View -> View)."""
    neu = Transform(transform)
    neu.Origin = XYZ(transform.Origin.X, transform.Origin.Y, 0.0)
    return neu


def _ansichtselemente(k, ansicht, neue, art, hinweise):
    # Schnitt-/Ansichtssymbole gehören zu anderen Ansichten - nicht kopieren
    sichtfenster = rv.id_wert(ElementId(BuiltInCategory.OST_Viewers))
    ids = []
    for element_id in rv._eigene_elemente(k.quelle, ansicht):
        element = k.quelle.GetElement(element_id)
        if not isinstance(element, ElevationMarker) \
                and rv.id_wert(element.Category.Id) != sichtfenster:
            ids.append(element_id)
    if not ids:
        return
    zusatz = _in_der_ebene(k.transform) if art in lg.PLAENE \
        else Transform.Identity

    def kopieren(auswahl):
        return rv._einheit(k.ziel, lambda: ElementTransformUtils.CopyElements(
            ansicht, rv.net_ids(auswahl), neue, zusatz, rv._optionen()))

    _e, fehler = kopieren(ids)
    if fehler is None:
        return
    # Einzeln, damit ein Element, das Revit nicht einfügen kann (z.B. eine
    # Beschriftung eines Bauteils, das im Ziel fehlt), nicht alle verhindert
    fehlend = sum(1 for i in ids if kopieren([i])[1] is not None)
    if fehlend:
        _hinweis(hinweise, t(
            u"%d von %d Ansichtselementen nicht übertragen (z.B. "
            u"Beschriftungen oder Maße zu Bauteilen, die im Ziel fehlen)",
            u"%d of %d view elements not transferred (e.g. tags or "
            u"dimensions of elements missing in the target)",
            u"%d de %d elementos de vista no transferidos (p. ej. etiquetas "
            u"o cotas de elementos que faltan en el destino)")
            % (fehlend, len(ids)))
