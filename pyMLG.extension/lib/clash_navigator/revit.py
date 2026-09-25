# -*- coding: utf-8 -*-
"""Revit-Seite des Clash Navigators: Elemente finden, Schnittbox setzen.

Läuft nur im API-Kontext - das Panel ruft alles über sein ExternalEvent auf.

Zuordnung eines Elements aus dem Bericht:
  1. Die Datei aus Navisworks ('TGA.nwc') wird mit Titel und Dateinamen des
     aktiven Modells und aller geladenen Verknüpfungen verglichen.
  2. In den passenden Modellen wird die Element-ID gesucht. Ist die Datei
     unbekannt, in allen - das aktive Modell zuerst.
  3. Steht dieselbe Verknüpfung mehrfach im Projekt, gewinnt die Instanz,
     deren Element dem Kollisionspunkt am nächsten liegt.

Der Kollisionspunkt stammt aus Navisworks. Meist sind das die gemeinsamen
Koordinaten (Standard beim NWC-Export aus Revit), manchmal die internen.
"auto" nimmt je Kollision die Deutung, die näher an den Elementen liegt.
"""

import os

import clr

from Autodesk.Revit.DB import (BoundingBoxXYZ, ElementId, ElementType,
                               FilteredElementCollector, ModelPathUtils,
                               Reference, RevitLinkInstance, Transaction,
                               Transform, View3D, ViewDetailLevel, ViewFamily,
                               ViewFamilyType, XYZ)

from mlg_sprache import t
from section_box import revit as sb

from clash_navigator import logik as lg

clr.AddReference("System")
from System import Int64  # noqa: E402
from System.Collections.Generic import List  # noqa: E402

GEMEINSAM = u"gemeinsam"
INTERN = u"intern"
AUTO = u"auto"


class Fehler(Exception):
    """Fachlicher Fehler mit Klartext für den Benutzer."""


class Modell(object):
    """Das aktive Modell oder eine geladene Verknüpfungsinstanz."""

    def __init__(self, dokument, instanz=None):
        self.dokument = dokument
        self.instanz = instanz
        self.transform = instanz.GetTotalTransform() if instanz else None
        self.namen = dokumentnamen(dokument)

    @property
    def ist_wirt(self):
        return self.instanz is None


class Fund(object):
    def __init__(self, objekt, element, modell, kasten):
        self.objekt = objekt
        self.element = element
        self.modell = modell
        self.kasten = kasten          # (min, max) in Metern, Wirtssystem


def dokumentnamen(dokument):
    """Titel, Dateiname und Name des Zentralmodells."""
    namen = []
    try:
        namen.append(dokument.Title)
    except Exception:
        pass
    try:
        if dokument.PathName:
            namen.append(os.path.basename(dokument.PathName.replace(u"\\",
                                                                    u"/")))
    except Exception:
        pass
    try:
        if dokument.IsWorkshared:
            zentral = ModelPathUtils.ConvertModelPathToUserVisiblePath(
                dokument.GetWorksharingCentralModelPath())
            if zentral:
                namen.append(os.path.basename(zentral.replace(u"\\", u"/")))
    except Exception:
        pass
    return [name for name in namen if name]


def modelle(doc):
    liste = [Modell(doc)]
    for instanz in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        try:
            dokument = instanz.GetLinkDocument()
        except Exception:
            dokument = None
        if dokument is not None:
            liste.append(Modell(dokument, instanz))
    return liste


def _element_id(wert):
    try:
        return ElementId(Int64(wert))
    except Exception:
        return ElementId(int(wert))


def id_wert_text(element_id):
    """ElementId als Text (Revit 2024+: .Value, davor .IntegerValue)."""
    try:
        return u"%d" % element_id.Value
    except AttributeError:
        return u"%d" % element_id.IntegerValue


def _m(xyz):
    return (xyz.X * lg.M_PRO_FUSS, xyz.Y * lg.M_PRO_FUSS,
            xyz.Z * lg.M_PRO_FUSS)


def _kasten(element, transform):
    """Hüllkasten des Elements in Metern, ins Wirtsmodell umgerechnet."""
    try:
        kasten = element.get_BoundingBox(None)
    except Exception:
        kasten = None
    if kasten is None:
        return None
    ecken = []
    for x in (kasten.Min.X, kasten.Max.X):
        for y in (kasten.Min.Y, kasten.Max.Y):
            for z in (kasten.Min.Z, kasten.Max.Z):
                punkt = kasten.Transform.OfPoint(XYZ(x, y, z))
                if transform is not None:
                    punkt = transform.OfPoint(punkt)
                ecken.append(_m(punkt))
    return lg.vereinigung([(ecke, ecke) for ecke in ecken])


def passende_modelle(objekt, liste):
    """Die Modelle, deren Namen zur Datei aus Navisworks passen."""
    if not objekt.datei:
        return []
    return [modell for modell in liste
            if lg.modell_passt(objekt.datei, modell.namen)]


def finde(objekt, liste, punkt=None):
    """Das Element zum Objekt aus dem Bericht - Fund oder None.

    Passt die Datei zu keinem Modell, wird überall gesucht. Element-IDs
    wiederholen sich aber zwischen Modellen - dann entscheidet die Nähe zum
    Kollisionspunkt, ohne Punkt das aktive Modell (steht vorn).
    """
    if objekt.element_id is None:
        return None
    kandidaten = []
    for modell in passende_modelle(objekt, liste) or liste:
        element = modell.dokument.GetElement(_element_id(objekt.element_id))
        if element is None or isinstance(element, ElementType):
            continue
        kasten = _kasten(element, modell.transform)
        if kasten is None:
            continue
        kandidaten.append(Fund(objekt, element, modell, kasten))
    if not kandidaten:
        return None
    if punkt is not None and len(kandidaten) > 1:
        kandidaten.sort(key=lambda fund: lg.abstand_zu_kasten(punkt,
                                                               fund.kasten))
    return kandidaten[0]


def _nach_intern(doc):
    """Transform gemeinsame -> interne Koordinaten (Fuss)."""
    return sb.nach_gemeinsam(doc).Inverse


def punkt_intern(doc, punkt, modus, kaesten):
    """Kollisionspunkt (Meter, Navisworks) -> Wirtssystem (Meter).

    Rückgabe: (punkt, verwendete Deutung) oder (None, None).
    """
    if punkt is None:
        return None, None
    intern = tuple(punkt)
    xyz = sb.xyz_aus_m(punkt)
    gemeinsam = _m(_nach_intern(doc).OfPoint(xyz))
    if modus == INTERN:
        return intern, INTERN
    if modus == GEMEINSAM or not kaesten:
        return gemeinsam, GEMEINSAM
    huelle = lg.vereinigung(kaesten)
    if lg.abstand_zu_kasten(intern, huelle) < \
            lg.abstand_zu_kasten(gemeinsam, huelle):
        return intern, INTERN
    return gemeinsam, GEMEINSAM


# ---------------------------------------------------------------------------
# Ansicht
# ---------------------------------------------------------------------------

def ansichtsname(doc):
    # Diese Zeichen verbietet Revit in Ansichtsnamen
    benutzer = u"".join(zeichen for zeichen in doc.Application.Username
                        if zeichen not in u"\\:{}[]|;<>?`~")
    return u"pyMLG Clash - %s" % (benutzer.strip() or u"Benutzer")


def _erzeuge_ansicht(doc, name):
    typ = None
    for kandidat in FilteredElementCollector(doc).OfClass(ViewFamilyType):
        if kandidat.ViewFamily == ViewFamily.ThreeDimensional:
            typ = kandidat
            break
    if typ is None:
        raise Fehler(t(u"Im Projekt gibt es keinen 3D-Ansichtstyp.",
                       u"The project has no 3D view type.",
                       u"El proyecto no tiene ningún tipo de vista 3D."))
    transaktion = Transaction(doc, t(u"Clash-Ansicht anlegen",
                                     u"Create clash view",
                                     u"Crear vista de conflictos"))
    transaktion.Start()
    try:
        ansicht = View3D.CreateIsometric(doc, typ.Id)
        ansicht.Name = name
        try:
            ansicht.ViewTemplateId = ElementId.InvalidElementId
        except Exception:
            pass
        try:
            ansicht.DetailLevel = ViewDetailLevel.Fine
        except Exception:
            pass
        transaktion.Commit()
        return ansicht
    except Exception:
        if transaktion.HasStarted():
            transaktion.RollBack()
        raise


def clash_ansicht(uidoc, aktive_verwenden):
    """Die 3D-Ansicht für die Schnittbox - aktiv geschaltet.

    Standard: eine eigene Ansicht je Benutzer ('pyMLG Clash - name'), damit
    im Teammodell keine fremde Ansicht verändert wird.
    """
    doc = uidoc.Document
    aktiv = uidoc.ActiveView
    if aktive_verwenden and isinstance(aktiv, View3D) and \
            not aktiv.IsTemplate:
        return aktiv
    name = ansichtsname(doc)
    ansicht = None
    for kandidat in FilteredElementCollector(doc).OfClass(View3D):
        if not kandidat.IsTemplate and kandidat.Name == name:
            ansicht = kandidat
            break
    if ansicht is None:
        ansicht = _erzeuge_ansicht(doc, name)
    if aktiv is None or aktiv.Id != ansicht.Id:
        uidoc.ActiveView = ansicht
    return ansicht


def _waehle(uidoc, funde):
    referenzen = List[Reference]()
    eigene = List[ElementId]()
    for fund in funde:
        referenz = Reference(fund.element)
        if fund.modell.ist_wirt:
            eigene.Add(fund.element.Id)
        else:
            referenz = referenz.CreateLinkReference(fund.modell.instanz)
        referenzen.Add(referenz)
    try:
        uidoc.Selection.SetReferences(referenzen)      # Revit 2023+
    except Exception:
        uidoc.Selection.SetElementIds(eigene)


# ---------------------------------------------------------------------------
# Zur Kollision springen
# ---------------------------------------------------------------------------

class Ergebnis(object):
    def __init__(self):
        self.gefunden = 0
        self.gesucht = 0
        self.deutung = None
        self.meldungen = []


def gehe_zu(uiapp, clash, optionen):
    """Schnittbox um die Kollision legen und hinzoomen.

    optionen: dict mit rand_cm, schnitt, aktive_ansicht, auswaehlen,
              koordinaten (auto/gemeinsam/intern)
    """
    uidoc = uiapp.ActiveUIDocument
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        raise Fehler(t(u"Kein Projekt geöffnet.", u"No project open.",
                       u"No hay ningún proyecto abierto."))

    ergebnis = Ergebnis()
    liste = modelle(doc)
    rohpunkt = clash.punkt

    # Punkt grob vorab, damit bei mehrfach verknüpften Modellen die richtige
    # Instanz gewählt wird
    vorab, _ = punkt_intern(doc, rohpunkt, optionen.get(u"koordinaten"), [])
    funde = []
    for objekt in clash.objekte:
        if objekt.element_id is None:
            continue
        ergebnis.gesucht += 1
        if objekt.datei and not passende_modelle(objekt, liste):
            ergebnis.meldungen.append(t(
                u"Datei %s passt zu keinem geladenen Modell - Element %d "
                u"wird in allen gesucht.",
                u"File %s matches no loaded model - element %d is searched "
                u"in all of them.",
                u"El archivo %s no corresponde a ningún modelo cargado; el "
                u"elemento %d se busca en todos.")
                % (objekt.datei, objekt.element_id))
        fund = finde(objekt, liste, vorab)
        if fund is None:
            ergebnis.meldungen.append(t(
                u"Element %d (%s) nicht gefunden.",
                u"Element %d (%s) not found.",
                u"Elemento %d (%s) no encontrado.")
                % (objekt.element_id, objekt.datei or u"?"))
        else:
            funde.append(fund)
    ergebnis.gefunden = len(funde)

    kaesten = [fund.kasten for fund in funde]
    punkt, ergebnis.deutung = punkt_intern(
        doc, rohpunkt, optionen.get(u"koordinaten") or AUTO, kaesten)
    kasten = lg.kasten_fuer(kaesten, punkt,
                            rand=float(optionen.get(u"rand_cm", 15)) / 100.0,
                            schnitt=bool(optionen.get(u"schnitt", True)))
    if kasten is None:
        raise Fehler(t(u"Zu dieser Kollision gibt es weder gefundene "
                       u"Elemente noch einen Kollisionspunkt.",
                       u"This clash has neither elements that were found "
                       u"nor a clash point.",
                       u"Este conflicto no tiene elementos encontrados ni "
                       u"punto de conflicto."))

    ansicht = clash_ansicht(uidoc, bool(optionen.get(u"aktive_ansicht")))
    box = BoundingBoxXYZ()
    box.Transform = Transform.Identity
    box.Min = sb.xyz_aus_m(kasten[0])
    box.Max = sb.xyz_aus_m(kasten[1])
    try:
        sb.setze_box(doc, ansicht, box, t(u"Clash Navigator: Schnittbox",
                                          u"Clash Navigator: section box",
                                          u"Clash Navigator: caja de "
                                          u"sección"))
    except Exception as fehler:
        raise Fehler(t(u"Die Schnittbox ließ sich nicht setzen: %s\n"
                       u"Steuert eine Ansichtsvorlage die Ansicht?",
                       u"The section box could not be set: %s\n"
                       u"Is a view template controlling the view?",
                       u"No se pudo aplicar la caja de sección: %s\n"
                       u"¿Una plantilla de vista controla la vista?")
                     % fehler)
    sb.zeige(uidoc, ansicht, box)
    if optionen.get(u"auswaehlen", True) and funde:
        try:
            _waehle(uidoc, funde)
        except Exception:
            pass
    return ergebnis
