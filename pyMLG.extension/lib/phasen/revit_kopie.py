# -*- coding: utf-8 -*-
"""Revit-Teil der Phasen-Werkzeuge: kopieren, Phasen und Ebenen übertragen.

Revit setzt bei kopierten Elementen "Phase erstellt" auf die Phase der aktiven
Ansicht. Diese Funktionen kopieren Elemente und stellen die Phasen des
Originals wieder her. Die Zuordnung Kopie -> Original läuft über die Lage
(siehe phasen.zuordnung), weil CopyElements() keine Zuordnung liefert.

Läuft unter IronPython und CPython. ElementId.IntegerValue gibt es ab
Revit 2026 nicht mehr, daher id_wert().
"""

from Autodesk.Revit.DB import (
    BuiltInParameter as BIP,
    CopyPasteOptions,
    ElementId,
    ElementTransformUtils,
    FamilyInstance,
    FilteredElementCollector,
    Level,
    LocationCurve,
    LocationPoint,
    StorageType,
    Transform,
    XYZ,
)
from System.Collections.Generic import List

from phasen.zuordnung import Eintrag, zuordnen
from mlg_sprache import t

# Ebenen gelten als gleich hoch, wenn sie weniger als ca. 0,3 mm abweichen.
HOEHEN_TOLERANZ = 1e-3

# Ebenen-Parameter, die die Unterkante bestimmen
BASIS_EBENEN = (
    BIP.WALL_BASE_CONSTRAINT,
    BIP.FAMILY_BASE_LEVEL_PARAM,
    BIP.FAMILY_LEVEL_PARAM,
    BIP.SCHEDULE_LEVEL_PARAM,
    BIP.INSTANCE_REFERENCE_LEVEL_PARAM,
    BIP.LEVEL_PARAM,
    BIP.ROOF_BASE_LEVEL_PARAM,
    BIP.ROOF_CONSTRAINT_LEVEL_PARAM,
    BIP.STAIRS_BASE_LEVEL_PARAM,
    BIP.GROUP_LEVEL,
    BIP.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM,
)

# Ebenen-Parameter, die die Oberkante bestimmen
OBERE_EBENEN = (
    BIP.WALL_HEIGHT_TYPE,
    BIP.FAMILY_TOP_LEVEL_PARAM,
    BIP.STAIRS_TOP_LEVEL_PARAM,
)

# Versätze zur Ebene. Revit behält beim Ebenenwechsel je nach Kategorie den
# Versatz oder die absolute Höhe bei - nach dem Wechsel werden sie deshalb
# ausdrücklich auf die Werte des Originals gesetzt.
VERSAETZE = (
    BIP.WALL_BASE_OFFSET,
    BIP.WALL_TOP_OFFSET,
    BIP.WALL_USER_HEIGHT_PARAM,
    BIP.FAMILY_BASE_LEVEL_OFFSET_PARAM,
    BIP.FAMILY_TOP_LEVEL_OFFSET_PARAM,
    BIP.FLOOR_HEIGHTABOVELEVEL_PARAM,
    BIP.CEILING_HEIGHTABOVELEVEL_PARAM,
    BIP.INSTANCE_ELEVATION_PARAM,
    BIP.INSTANCE_FREE_HOST_OFFSET_PARAM,
    BIP.ROOF_LEVEL_OFFSET_PARAM,
    BIP.STAIRS_BASE_OFFSET,
    BIP.STAIRS_TOP_OFFSET,
    BIP.GROUP_OFFSET_FROM_LEVEL,
)


def id_wert(element_id):
    """Zahlenwert einer ElementId (Revit 2024+: .Value)."""
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def _gueltig(element_id):
    return element_id is not None and id_wert(element_id) != -1


# ---------------------------------------------------------------- Auswahl

def modell_elemente(doc, element_ids):
    """Trennt Modellelemente von ansichtsspezifischen (Beschriftungen usw.).

    Liefert (Liste der Modell-Ids, Anzahl übersprungener Elemente).
    """
    modell, uebersprungen = [], 0
    for eid in element_ids:
        elem = doc.GetElement(eid)
        if elem is None:
            continue
        if _gueltig(elem.OwnerViewId):
            uebersprungen += 1
        else:
            modell.append(elem.Id)
    return modell, uebersprungen


def auswahl_setzen(uidoc, element_ids):
    uidoc.Selection.SetElementIds(List[ElementId](element_ids))


# ---------------------------------------------------------------- Zuordnung

def _xyz(p):
    return (p.X, p.Y, p.Z)


def eintrag(elem):
    """Zuordnungs-Eintrag: Kategorie, Typ und Lage eines Elements."""
    punkte = []
    try:
        ort = elem.Location
        if isinstance(ort, LocationPoint):
            punkte = [_xyz(ort.Point)]
        elif isinstance(ort, LocationCurve):
            kurve = ort.Curve
            punkte = [_xyz(kurve.GetEndPoint(0)), _xyz(kurve.GetEndPoint(1))]
    except Exception:
        punkte = []
    if not punkte:
        rahmen = elem.get_BoundingBox(None)
        if rahmen is not None:
            punkte = [_xyz(rahmen.Min), _xyz(rahmen.Max)]
    kategorie = id_wert(elem.Category.Id) if elem.Category is not None else None
    return Eintrag(id_wert(elem.Id), (kategorie, id_wert(elem.GetTypeId())), punkte)


def _hat_phasen(elem):
    return elem.get_Parameter(BIP.PHASE_CREATED) is not None


def _mit_abhaengigen(doc, elemente):
    """{id: Element} der Elemente samt abhängiger Elemente mit Phasen (z.B. Türen)."""
    ergebnis = {}
    for elem in elemente:
        if elem is None:
            continue
        ergebnis[id_wert(elem.Id)] = elem
        try:
            abhaengige = elem.GetDependentElements(None)
        except Exception:
            abhaengige = []
        for eid in abhaengige:
            wert = id_wert(eid)
            if wert in ergebnis:
                continue
            kind = doc.GetElement(eid)
            if kind is not None and _hat_phasen(kind):
                ergebnis[wert] = kind
    return ergebnis


def kopieren(doc, element_ids, verschiebung):
    """Kopiert Modellelemente um verschiebung (XYZ), mit Neu-Hosting.

    Muss in einer offenen Transaktion laufen. Liefert
    (Ids der direkten Kopien, Liste von (Kopie, Original), Anzahl Kopien mit
    Phasen ohne gefundenes Original). Die Paare umfassen auch abhängige
    Elemente wie Türen und Fenster.
    """
    originale = _mit_abhaengigen(doc, [doc.GetElement(i) for i in element_ids])
    eintraege = [eintrag(e) for e in originale.values()]

    neue_ids = ElementTransformUtils.CopyElements(
        doc, List[ElementId](element_ids), doc,
        Transform.CreateTranslation(verschiebung), CopyPasteOptions())
    doc.Regenerate()

    kopien = _mit_abhaengigen(doc, [doc.GetElement(i) for i in neue_ids])
    zuordnung = zuordnen(eintraege, [eintrag(e) for e in kopien.values()],
                         _xyz(verschiebung))
    paare = [(kopien[k], originale[o]) for k, o in zuordnung.items()]
    ohne_original = sum(1 for k, e in kopien.items()
                        if k not in zuordnung and _hat_phasen(e))
    return list(neue_ids), paare, ohne_original


# ---------------------------------------------------------------- Phasen

def phasen_uebertragen(kopie, original):
    """Setzt "Phase erstellt/abgebrochen" der Kopie auf die des Originals.

    Liefert False, wenn das Element keine änderbaren Phasen hat.
    """
    ziel_erstellt = kopie.get_Parameter(BIP.PHASE_CREATED)
    quelle_erstellt = original.get_Parameter(BIP.PHASE_CREATED)
    if ziel_erstellt is None or quelle_erstellt is None or ziel_erstellt.IsReadOnly:
        return False
    ziel_abgebr = kopie.get_Parameter(BIP.PHASE_DEMOLISHED)
    quelle_abgebr = original.get_Parameter(BIP.PHASE_DEMOLISHED)
    abgebr_aenderbar = (ziel_abgebr is not None and quelle_abgebr is not None
                        and not ziel_abgebr.IsReadOnly)

    erstellt = quelle_erstellt.AsElementId()
    abgebrochen = quelle_abgebr.AsElementId() if abgebr_aenderbar else None

    # Erst "abgebrochen" leeren: sonst lehnt Revit eine "erstellt"-Phase ab,
    # die nach der vorhandenen "abgebrochen"-Phase liegt.
    if abgebr_aenderbar and _gueltig(ziel_abgebr.AsElementId()):
        ziel_abgebr.Set(ElementId.InvalidElementId)
    if id_wert(ziel_erstellt.AsElementId()) != id_wert(erstellt):
        ziel_erstellt.Set(erstellt)
    if abgebr_aenderbar and _gueltig(abgebrochen):
        ziel_abgebr.Set(abgebrochen)
    return True


# ---------------------------------------------------------------- Ebenen

def ebenen(doc):
    return list(FilteredElementCollector(doc).OfClass(Level))


def ebene_auf_hoehe(alle_ebenen, hoehe):
    for ebene in alle_ebenen:
        if abs(ebene.Elevation - hoehe) < HOEHEN_TOLERANZ:
            return ebene
    return None


def ist_eingefuegt(elem):
    """Tür, Fenster o.ä. in einem Wirt, der keine Ebene ist - folgt dem Wirt."""
    return (isinstance(elem, FamilyInstance) and elem.Host is not None
            and not isinstance(elem.Host, Level))


def basis_ebene(doc, elem):
    """Ebene, auf der das Element steht (eingefügte Elemente: die des Wirts)."""
    if ist_eingefuegt(elem):
        elem = elem.Host
    for bip in BASIS_EBENEN:
        param = elem.get_Parameter(bip)
        if param is not None and param.StorageType == StorageType.ElementId:
            ebene = doc.GetElement(param.AsElementId())
            if isinstance(ebene, Level):
                return ebene
    ebene = doc.GetElement(elem.LevelId)
    return ebene if isinstance(ebene, Level) else None


def _unterkante(elem):
    punkte = eintrag(elem).punkte
    return min(p[2] for p in punkte) if punkte else None


def ebenen_anpassen(doc, kopie, original, delta, alle_ebenen):
    """Hängt die Kopie an die Ebenen, die um delta über denen des Originals liegen.

    Die Kopie ist bereits um delta verschoben. Liefert eine Liste von
    Hinweistexten (leer = alles in Ordnung).
    """
    if abs(delta) < HOEHEN_TOLERANZ or ist_eingefuegt(kopie):
        return []
    hinweise = []
    geaendert = False
    # Nach oben zuerst die Oberkante versetzen, nach unten zuerst die Basis -
    # so liegt die Oberkante nie unter der Basis.
    reihenfolge = OBERE_EBENEN + BASIS_EBENEN if delta > 0 else BASIS_EBENEN + OBERE_EBENEN
    for bip in reihenfolge:
        quelle = original.get_Parameter(bip)
        ziel = kopie.get_Parameter(bip)
        if (quelle is None or ziel is None or ziel.IsReadOnly
                or ziel.StorageType != StorageType.ElementId):
            continue
        alte = doc.GetElement(quelle.AsElementId())
        if not isinstance(alte, Level):
            continue
        neue = ebene_auf_hoehe(alle_ebenen, alte.Elevation + delta)
        try:
            if neue is not None:
                if id_wert(ziel.AsElementId()) != id_wert(neue.Id):
                    ziel.Set(neue.Id)
                geaendert = True
            elif bip == BIP.WALL_HEIGHT_TYPE:
                # Keine passende obere Ebene: Wand "nicht verbunden" mit
                # gleicher Höhe (WALL_USER_HEIGHT_PARAM folgt bei den Versätzen)
                ziel.Set(ElementId.InvalidElementId)
                geaendert = True
            else:
                hinweise.append(t(u"keine Ebene {:.3f} m über '{}'", u"no level {:.3f} m above '{}'", u"ningún nivel {:.3f} m por encima de '{}'").format(
                    delta * 0.3048, alte.Name))
        except Exception as fehler:
            hinweise.append(t(u"Ebene nicht änderbar ({})", u"level cannot be changed ({})", u"no se puede cambiar el nivel ({})").format(fehler))

    if geaendert:
        for bip in VERSAETZE:
            quelle = original.get_Parameter(bip)
            ziel = kopie.get_Parameter(bip)
            if (quelle is None or ziel is None or ziel.IsReadOnly
                    or ziel.StorageType != StorageType.Double):
                continue
            if abs(ziel.AsDouble() - quelle.AsDouble()) > 1e-9:
                try:
                    ziel.Set(quelle.AsDouble())
                except Exception:
                    pass

    return hinweise


def hoehen_korrigieren(doc, paare, delta):
    """Absicherung nach ebenen_anpassen() für Kategorien, deren Versatz
    VERSAETZE nicht erfasst: Kopie so verschieben, dass sie genau delta über
    dem Original liegt. Liefert die Anzahl korrigierter Elemente."""
    doc.Regenerate()
    korrigiert = 0
    for kopie, original in paare:
        if ist_eingefuegt(kopie):
            continue
        soll, ist = _unterkante(original), _unterkante(kopie)
        if soll is None or ist is None:
            continue
        abweichung = soll + delta - ist
        if abs(abweichung) > HOEHEN_TOLERANZ:
            ElementTransformUtils.MoveElement(doc, kopie.Id, XYZ(0, 0, abweichung))
            korrigiert += 1
    return korrigiert
