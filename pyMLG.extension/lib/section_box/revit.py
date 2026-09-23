# -*- coding: utf-8 -*-
# Schnittbox der 3D-Ansicht lesen und setzen, dazu die Umrechnung zwischen
# Projektkoordinaten und gemeinsamen Koordinaten (Vermessungspunkt).
#
# Laengen wandern in Metern durch logik.py, Revit rechnet intern in Fuss.
# Die Achsen der Lage sind Richtungen und werden nicht umgerechnet.

from Autodesk.Revit.DB import (BoundingBoxXYZ, FilteredElementCollector,
                               Transaction, Transform, View3D, ViewFamily,
                               ViewFamilyType, XYZ)

from mlg_sprache import t

from section_box import logik as lg


def _m(wert):
    return wert * lg.M_PRO_FUSS


def _fuss(wert):
    return wert / lg.M_PRO_FUSS


def _punkt_m(xyz):
    return (_m(xyz.X), _m(xyz.Y), _m(xyz.Z))


def _richtung_m(xyz):
    return (xyz.X, xyz.Y, xyz.Z)


def xyz_aus_m(punkt):
    return XYZ(_fuss(punkt[0]), _fuss(punkt[1]), _fuss(punkt[2]))


def _xyz_richtung(punkt):
    return XYZ(punkt[0], punkt[1], punkt[2])


def lage_aus_transform(transform):
    """(Ursprung, X, Y, Z) in Metern bzw. als Richtungen."""
    return (_punkt_m(transform.Origin), _richtung_m(transform.BasisX),
            _richtung_m(transform.BasisY), _richtung_m(transform.BasisZ))


def transform_aus_lage(lage):
    transform = Transform.Identity
    transform.Origin = xyz_aus_m(lage[0])
    transform.BasisX = _xyz_richtung(lage[1])
    transform.BasisY = _xyz_richtung(lage[2])
    transform.BasisZ = _xyz_richtung(lage[3])
    return transform


def sauber(transform):
    """Achsen exakt rechtwinklig und einheitslang machen.

    Aus dem Text gelesene oder aus Punkten abgetastete Achsen sind nur
    ungefaehr orthonormal. BoundingBoxXYZ.Transform prueft das streng und
    wirft sonst eine ArgumentException.
    """
    achse_x = transform.BasisX.Normalize()
    achse_z = transform.BasisZ.Normalize()
    achse_y = achse_z.CrossProduct(achse_x)
    if (achse_x.IsZeroLength() or achse_z.IsZeroLength()
            or achse_y.IsZeroLength()):
        return transform
    achse_y = achse_y.Normalize()
    geputzt = Transform.Identity
    geputzt.Origin = transform.Origin
    geputzt.BasisX = achse_x
    geputzt.BasisY = achse_y
    geputzt.BasisZ = achse_x.CrossProduct(achse_y)
    return geputzt


def nach_gemeinsam(doc):
    """Transform Projektkoordinaten -> gemeinsame Koordinaten.

    Abgetastet an Nullpunkt und Einheitsvektoren: GetProjectPosition liefert
    zu einem internen Punkt seine gemeinsamen Koordinaten, damit steht die
    Abbildung fest - unabhaengig davon, in welche Richtung
    GetTotalTransform in der jeweiligen Revit-Version zeigt.
    """
    ort = doc.ActiveProjectLocation
    try:
        def gemeinsam(punkt):
            lage = ort.GetProjectPosition(punkt)
            return XYZ(lage.EastWest, lage.NorthSouth, lage.Elevation)

        ursprung = gemeinsam(XYZ.Zero)
        transform = Transform.Identity
        transform.Origin = ursprung
        transform.BasisX = gemeinsam(XYZ.BasisX) - ursprung
        transform.BasisY = gemeinsam(XYZ.BasisY) - ursprung
        transform.BasisZ = gemeinsam(XYZ.BasisZ) - ursprung
        return sauber(transform)
    except Exception:
        return sauber(ort.GetTotalTransform().Inverse)


def lies_box(doc, ansicht, quelle=u""):
    """Die Schnittbox der 3D-Ansicht als logik.Box."""
    kasten = ansicht.GetSectionBox()
    intern = kasten.Transform
    return lg.Box(lage_aus_transform(nach_gemeinsam(doc).Multiply(intern)),
                  lage_aus_transform(intern),
                  _punkt_m(kasten.Min), _punkt_m(kasten.Max), quelle=quelle)


def baue_kasten(doc, box, gemeinsam=True):
    """BoundingBoxXYZ in den Projektkoordinaten dieses Dokuments."""
    if gemeinsam:
        lage = nach_gemeinsam(doc).Inverse.Multiply(
            transform_aus_lage(box.gemeinsam))
    else:
        lage = transform_aus_lage(box.intern)
    kasten = BoundingBoxXYZ()
    kasten.Transform = sauber(lage)
    kasten.Min = xyz_aus_m(box.min)
    kasten.Max = xyz_aus_m(box.max)
    return kasten


def mittelpunkt(kasten):
    """Mitte des Kastens in Projektkoordinaten (XYZ, Fuss)."""
    mitte = XYZ((kasten.Min.X + kasten.Max.X) / 2.0,
                (kasten.Min.Y + kasten.Max.Y) / 2.0,
                (kasten.Min.Z + kasten.Max.Z) / 2.0)
    return kasten.Transform.OfPoint(mitte)


def abstand_m(kasten_a, kasten_b):
    """Abstand der Mittelpunkte zweier Kaesten in Metern."""
    return _m(mittelpunkt(kasten_a).DistanceTo(mittelpunkt(kasten_b)))


def setze_box(doc, ansicht, kasten, titel):
    """Schnittbox setzen und einschalten - in einer eigenen Transaktion."""
    transaktion = Transaction(doc, titel)
    transaktion.Start()
    try:
        ansicht.SetSectionBox(kasten)
        if not ansicht.IsSectionBoxActive:
            ansicht.IsSectionBoxActive = True
        transaktion.Commit()
    except Exception:
        if transaktion.HasStarted():
            transaktion.RollBack()
        raise


def zeige(uidoc, ansicht, kasten):
    """Die Ansicht auf den Kasten zoomen - schlaegt es fehl, bleibt die
    Ansicht wie sie ist."""
    try:
        ecke_a = kasten.Transform.OfPoint(kasten.Min)
        ecke_b = kasten.Transform.OfPoint(kasten.Max)
        for ui_ansicht in uidoc.GetOpenUIViews():
            if ui_ansicht.ViewId == ansicht.Id:
                ui_ansicht.ZoomAndCenterRectangle(ecke_a, ecke_b)
                return
    except Exception:
        pass


def _standardname(doc):
    """So nennt Revit die Standard-3D-Ansicht: {3D - Anwender}."""
    return u"{3D - %s}" % doc.Application.Username


def _vorhandene(doc):
    return [ansicht for ansicht
            in FilteredElementCollector(doc).OfClass(View3D)
            if not ansicht.IsTemplate]


def _erzeuge(doc):
    typ = None
    for kandidat in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if kandidat.ViewFamily == ViewFamily.ThreeDimensional:
            typ = kandidat
            break
    if typ is None:
        return None

    transaktion = Transaction(doc, t(u"Standard-3D-Ansicht",
                                     u"Default 3D view",
                                     u"Vista 3D predeterminada"))
    transaktion.Start()
    try:
        ansicht = View3D.CreateIsometric(doc, typ.Id)
        try:
            ansicht.Name = _standardname(doc)
        except Exception:
            pass
        transaktion.Commit()
        return ansicht
    except Exception:
        if transaktion.HasStarted():
            transaktion.RollBack()
        return None


def standard_3d(uidoc, doc):
    """Die Standard-3D-Ansicht des Anwenders - sonst eine andere vorhandene,
    sonst eine neue. Sie wird zur aktiven Ansicht. None, wenn nichts geht."""
    vorhanden = _vorhandene(doc)
    nach_namen = dict((ansicht.Name, ansicht) for ansicht in vorhanden)

    ansicht = nach_namen.get(_standardname(doc)) or nach_namen.get(u"{3D}")
    if ansicht is None:
        gerade = [kandidat for kandidat in vorhanden
                  if not kandidat.IsPerspective]
        ansicht = (gerade or vorhanden or [None])[0]
    if ansicht is None:
        ansicht = _erzeuge(doc)
    if ansicht is None:
        return None

    try:
        uidoc.ActiveView = ansicht
    except Exception:
        return None
    return ansicht
