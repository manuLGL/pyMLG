# -*- coding: utf-8 -*-
# Eine Schnittbox um die gewaehlten Elemente - eigene wie verknuepfte.
#
# Ein mit Tab angetipptes Element einer Verknuepfung steht in der Auswahl nur
# als Referenz (siehe lib/linked_ids). Seine Ausdehnung gilt im verknuepften
# Dokument und muss ueber die Transformationen der durchlaufenen
# Verknuepfungen ins Wirtsmodell gerechnet werden - bei verschachtelten
# Verknuepfungen eine nach der anderen.

from Autodesk.Revit.DB import BoundingBoxXYZ, Transform, XYZ
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ObjectType

from linked_ids import revit as li
from mlg_sprache import t
from section_box import logik as lg
from section_box import revit as sb

# Rand um die Elemente und kleinste Kantenlaenge, beides in Metern
RAND = 0.5
MINDESTMASS = 1.0


def _ecken(kasten, transform=None):
    """Die acht Ecken einer BoundingBoxXYZ, bei Bedarf umgerechnet."""
    ecken = []
    for x in (kasten.Min.X, kasten.Max.X):
        for y in (kasten.Min.Y, kasten.Max.Y):
            for z in (kasten.Min.Z, kasten.Max.Z):
                punkt = kasten.Transform.OfPoint(XYZ(x, y, z))
                if transform is not None:
                    punkt = transform.OfPoint(punkt)
                ecken.append((punkt.X * lg.M_PRO_FUSS,
                              punkt.Y * lg.M_PRO_FUSS,
                              punkt.Z * lg.M_PRO_FUSS))
    return ecken


def _element_ecken(element, transform=None):
    if element is None:
        return []
    try:
        kasten = element.get_BoundingBox(None)
    except Exception:
        kasten = None
    if kasten is None:
        return []
    return _ecken(kasten, transform)


def _ketten_transform(kette):
    """Die Transformationen der durchlaufenen Verknuepfungen verketten."""
    gesamt = None
    for instanz in kette:
        if instanz is None:
            return None
        try:
            einzeln = instanz.GetTotalTransform()
        except Exception:
            return None
        gesamt = einzeln if gesamt is None else gesamt.Multiply(einzeln)
    return gesamt


def _aus_referenz(doc, referenz):
    dokument, element, _wert, kette = li.aufloesen(doc, referenz)
    if element is None:
        return []
    if not kette:
        return _element_ecken(element)
    transform = _ketten_transform(kette)
    if transform is None:
        return []
    return _element_ecken(element, transform)


def _aus_ids(uidoc, doc):
    ecken = []
    # Eine ganz gewaehlte Verknuepfung zaehlt mit ihrer vollen Ausdehnung -
    # die liegt schon in den Koordinaten des Wirtsmodells.
    for element_id in uidoc.Selection.GetElementIds():
        ecken.extend(_element_ecken(doc.GetElement(element_id)))
    return ecken


def ecken_der_auswahl(uidoc):
    """Alle Eckpunkte der laufenden Auswahl in Metern, Wirtskoordinaten."""
    doc = uidoc.Document
    try:
        referenzen = list(uidoc.Selection.GetReferences())
    except Exception:
        referenzen = []          # Revit vor 2023

    if referenzen:
        ecken = []
        for referenz in referenzen:
            ecken.extend(_aus_referenz(doc, referenz))
        return ecken
    return _aus_ids(uidoc, doc)


def picke(uidoc):
    """Elemente in Verknuepfungen antippen lassen.

    Rueckgabe: Eckpunkte, oder None bei Abbruch (Esc).
    """
    hinweis = t(u"Elemente in einer Verknüpfung wählen - mit Tab das "
                u"einzelne Element antippen, Fertigstellen beendet die Wahl",
                u"Select elements inside a link - press Tab to reach the "
                u"single element, Finish ends the selection",
                u"Seleccione elementos de un vínculo: con Tab llega al "
                u"elemento concreto, Finalizar termina la selección")
    try:
        referenzen = uidoc.Selection.PickObjects(ObjectType.LinkedElement,
                                                 hinweis)
    except OperationCanceledException:
        return None
    doc = uidoc.Document
    ecken = []
    for referenz in referenzen:
        ecken.extend(_aus_referenz(doc, referenz))
    return ecken


def kasten_um(ecken, rand=RAND, mindestmass=MINDESTMASS):
    """Achsparallele BoundingBoxXYZ um die Punkte (Meter) - oder None."""
    grenzen = lg.huelle(ecken, rand, mindestmass)
    if grenzen is None:
        return None
    klein, gross = grenzen
    kasten = BoundingBoxXYZ()
    kasten.Transform = Transform.Identity
    kasten.Min = sb.xyz_aus_m(klein)
    kasten.Max = sb.xyz_aus_m(gross)
    return kasten
