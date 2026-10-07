# -*- coding: utf-8 -*-
"""Revit-Zugriffe von WallProfileCopy.

Ablauf je Zielwand (Revit 2022+):

    Wand noch rechteckig: Transaktion CreateProfileSketch + Regenerate
    (StartWithNewSketch lehnt Wände ab), scheitert der Rest: wieder weg
    SketchEditScope.Start(SketchId)
    Transaktion: alte Skizzenlinien löschen, Linien der Quellwand
                 transformiert auf die Skizzenebene der Zielwand setzen
    SketchEditScope.Commit(...)

Alle Wände laufen in einer Transaktionsgruppe -> ein Rückgängig-Schritt.
Lässt Revit den SketchEditScope in der Gruppe nicht zu, wird die Gruppe
verworfen und jede Wand einzeln übertragen.
"""

import io
import os
import traceback
import uuid

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInParameter,
    CurveElement,
    ElementId,
    FailureProcessingResult,
    FailureSeverity,
    IFailuresPreprocessor,
    Line,
    SketchEditScope,
    Transaction,
    TransactionGroup,
    TransactionStatus,
    Transform,
    Wall,
    XYZ,
)
from Autodesk.Revit.Exceptions import (  # noqa: E402
    InvalidOperationException,
    OperationCanceledException,
)
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType  # noqa: E402

from mlg_sprache import t  # noqa: E402
from wand_profil import logik as lg  # noqa: E402

# Abstand zur Skizzenebene, ab dem das Profil auf die Ebene geschoben wird
EBENEN_TOLERANZ = 1e-6

_klassen = {}

PROTOKOLL = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "pyMLG",
    "WallProfileCopy_Protokoll.txt")
_zeilen = []


def _notiere(text):
    _zeilen.append(text)


def schreibe_protokoll():
    """Ablauf des letzten Laufs (Fehler mit Traceback) in die Datei."""
    try:
        if not os.path.isdir(os.path.dirname(PROTOKOLL)):
            os.makedirs(os.path.dirname(PROTOKOLL))
        with io.open(PROTOKOLL, "w", encoding="utf-8") as datei:
            datei.write(u"\n".join(_zeilen) + u"\n")
    except Exception:
        pass
_pruefungen = {}


class Abbruch(Exception):
    """Der Nutzer hat die Auswahl abgebrochen."""


class StartFehler(Exception):
    """Revit lässt den Skizzenmodus nicht starten."""


def id_wert(element_id):
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def hat_profil(wand):
    return (isinstance(wand, Wall)
            and wand.SketchId != ElementId.InvalidElementId)


def _achse(wand):
    """(start, ende) der geraden Wandachse oder None."""
    ort = getattr(wand, "Location", None)
    kurve = getattr(ort, "Curve", None)
    if not isinstance(kurve, Line):
        return None
    a, b = kurve.GetEndPoint(0), kurve.GetEndPoint(1)
    return (a.X, a.Y, a.Z), (b.X, b.Y, b.Z)


def _basis(doc, wand):
    """Absolute Höhe der Wandunterkante (Ebene + Basisversatz)."""
    ebene = doc.GetElement(wand.LevelId)
    hoehe = ebene.Elevation if ebene is not None else 0.0
    versatz = wand.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)
    if versatz is not None:
        hoehe += versatz.AsDouble()
    return hoehe


def beschreibe(wand):
    try:
        return u"%s [%d]" % (wand.Name, id_wert(wand.Id))
    except Exception:
        return u"[%d]" % id_wert(wand.Id)


# ---------------------------------------------------------------- Auswahl

def _filter(name, erlaubt):
    """ISelectionFilter als .NET-Klasse. Der Namensraum ist je Sitzung
    eindeutig, weil die CPython-Engine im Revit-Prozess weiterlebt."""
    if name not in _klassen:
        namensraum = "pyMLG.WandProfil_%s_%s" % (name, uuid.uuid4().hex)

        class Filter(ISelectionFilter):
            __namespace__ = namensraum

            def AllowElement(self, element):
                try:
                    return bool(_pruefungen[name](element))
                except Exception:
                    return False

            def AllowReference(self, referenz, punkt):
                return False

        _klassen[name] = Filter
    _pruefungen[name] = erlaubt
    return _klassen[name]()


def waehle_quelle(uidoc):
    try:
        referenz = uidoc.Selection.PickObject(
            ObjectType.Element, _filter("quelle", hat_profil),
            t(u"Wand mit bearbeitetem Profil wählen",
              u"Select the wall with the edited profile",
              u"Seleccione el muro con el perfil editado"))
    except OperationCanceledException:
        raise Abbruch()
    return uidoc.Document.GetElement(referenz.ElementId)


def waehle_ziele(uidoc, quelle):
    def erlaubt(element):
        return (isinstance(element, Wall) and element.Id != quelle.Id
                and element.CanHaveProfileSketch())
    try:
        referenzen = uidoc.Selection.PickObjects(
            ObjectType.Element, _filter("ziele", erlaubt),
            t(u"Zielwände wählen, dann Fertig stellen",
              u"Select the target walls, then Finish",
              u"Seleccione los muros de destino y pulse Finalizar"))
    except OperationCanceledException:
        raise Abbruch()
    doc = uidoc.Document
    return [doc.GetElement(r.ElementId) for r in referenzen]


def aufteilen(doc, auswahl):
    """Vorauswahl -> (quelle oder None, ziele).

    Quelle ist die einzige gewählte Wand mit bearbeitetem Profil. Haben
    mehrere oder keine ein Profil, bleibt die Quelle offen."""
    waende = [doc.GetElement(i) for i in auswahl]
    waende = [w for w in waende if isinstance(w, Wall)]
    mit_profil = [w for w in waende if hat_profil(w)]
    if len(mit_profil) == 1:
        quelle = mit_profil[0]
        return quelle, [w for w in waende if w.Id != quelle.Id]
    return None, waende


def pruefe_ziele(doc, quelle, ziele):
    """-> (passende Wände, [(Wand, Grund)])."""
    achse_q = _achse(quelle)
    gut, schlecht = [], []
    for wand in ziele:
        if wand.Id == quelle.Id:
            continue
        achse = _achse(wand)
        if not wand.CanHaveProfileSketch() or achse is None:
            schlecht.append((wand, t(u"kein Profil möglich (gebogen, "
                                     u"geneigt oder Vorhangfassade)",
                                     u"cannot have a profile (curved, "
                                     u"tapered or curtain wall)",
                                     u"no admite perfil (curvo, inclinado "
                                     u"o muro cortina)")))
        elif not lg.gleich_lang(achse_q, achse):
            schlecht.append((wand, t(u"andere Länge", u"different length",
                                     u"longitud distinta")))
        else:
            gut.append(wand)
    return gut, schlecht


# ---------------------------------------------------------------- Übertragen

def _vorverarbeiter():
    """Warnungen still bestätigen - sonst ein Dialog je Wand."""
    if "fehler" not in _klassen:
        namensraum = "pyMLG.WandProfil_%s" % uuid.uuid4().hex

        class Vorverarbeiter(IFailuresPreprocessor):
            __namespace__ = namensraum

            def PreprocessFailures(self, accessor):
                for meldung in list(accessor.GetFailureMessages()):
                    if meldung.GetSeverity() == FailureSeverity.Warning:
                        accessor.DeleteWarning(meldung)
                return FailureProcessingResult.Continue

        _klassen["fehler"] = Vorverarbeiter
    return _klassen["fehler"]()


def _profil_kurven(doc, wand):
    skizze = doc.GetElement(wand.SketchId)
    return [kurve for schleife in skizze.Profile for kurve in schleife]


def _trafo(doc, quelle, ziel):
    winkel, (dx, dy, dz) = lg.lage(_achse(quelle), _basis(doc, quelle),
                                   _achse(ziel), _basis(doc, ziel))
    return Transform.CreateTranslation(XYZ(dx, dy, dz)).Multiply(
        Transform.CreateRotation(XYZ.BasisZ, winkel))


def _auf_ebene(kurven, ebene):
    """Kurven senkrecht auf die Skizzenebene schieben (z. B. wenn die
    Wände eine andere Basislinie haben)."""
    flaeche = ebene.GetPlane()
    abstand = (kurven[0].GetEndPoint(0) - flaeche.Origin).DotProduct(
        flaeche.Normal)
    if abs(abstand) <= EBENEN_TOLERANZ:
        return kurven
    schub = Transform.CreateTranslation(flaeche.Normal.Multiply(-abstand))
    return [k.CreateTransformed(schub) for k in kurven]


def _in_transaktion(doc, name, aktion):
    transaktion = Transaction(doc, name)
    transaktion.Start()
    try:
        optionen = transaktion.GetFailureHandlingOptions()
        optionen.SetFailuresPreprocessor(_vorverarbeiter())
        transaktion.SetFailureHandlingOptions(optionen)
        aktion()
    except Exception:
        transaktion.RollBack()
        raise
    transaktion.Commit()


def _profil_entfernen(doc, wand, name):
    """Die eben angelegte Rechteckskizze wieder weg - die Wand bleibt wie
    vorher."""
    try:
        if wand.SketchId != ElementId.InvalidElementId:
            _in_transaktion(doc, name, wand.RemoveProfileSketch)
    except Exception:
        _notiere(u"%s: Skizze nicht entfernt\n%s" % (
            beschreibe(wand), traceback.format_exc()))


def _uebertrage_eine(doc, quellkurven, quelle, ziel, name):
    trafo = _trafo(doc, quelle, ziel)
    kurven = [k.CreateTransformed(trafo) for k in quellkurven]

    # Rechteckige Wand: erst eine Profilskizze anlegen. StartWithNewSketch
    # lehnt Wände ab ("already has a sketch defined").
    neu = ziel.SketchId == ElementId.InvalidElementId
    if neu:
        _in_transaktion(doc, name, lambda: (ziel.CreateProfileSketch(),
                                            doc.Regenerate()))
    try:
        bereich = SketchEditScope(doc, name)
        try:
            bereich.Start(ziel.SketchId)
        except InvalidOperationException as fehler_obj:
            raise StartFehler(_kurz(fehler_obj))
    except Exception:
        if neu:
            _profil_entfernen(doc, ziel, name)
        raise
    try:
        transaktion = Transaction(doc, name)
        transaktion.Start()
        try:
            optionen = transaktion.GetFailureHandlingOptions()
            optionen.SetFailuresPreprocessor(_vorverarbeiter())
            transaktion.SetFailureHandlingOptions(optionen)
            skizze = doc.GetElement(ziel.SketchId)
            if skizze is None:
                raise RuntimeError(t(u"Revit hat keine Skizze angelegt",
                                     u"Revit did not create a sketch",
                                     u"Revit no ha creado un boceto"))
            ebene = skizze.SketchPlane
            alte = [i for i in skizze.GetAllElements()
                    if isinstance(doc.GetElement(i), CurveElement)]
            for element_id in alte:
                doc.Delete(element_id)
            for kurve in _auf_ebene(kurven, ebene):
                doc.Create.NewModelCurve(kurve, ebene)
        except Exception:
            transaktion.RollBack()
            raise
        if transaktion.Commit() != TransactionStatus.Committed:
            raise RuntimeError(t(u"Skizze ungültig", u"Invalid sketch",
                                 u"Boceto no válido"))
        bereich.Commit(_vorverarbeiter())
    except Exception:
        if bereich.IsActive:
            bereich.Cancel()
        if neu:
            _profil_entfernen(doc, ziel, name)
        raise
    if ziel.SketchId == ElementId.InvalidElementId:
        raise RuntimeError(t(u"Revit hat das Profil verworfen",
                             u"Revit discarded the profile",
                             u"Revit ha descartado el perfil"))


def _alle(doc, quellkurven, quelle, ziele, name, erste_streng=False):
    """erste_streng: scheitert bei der ersten Wand schon der Start des
    Skizzenmodus, StartFehler weiterreichen statt zu sammeln."""
    fertig, fehler = [], []
    for nummer, ziel in enumerate(ziele):
        try:
            _uebertrage_eine(doc, quellkurven, quelle, ziel, name)
            fertig.append(ziel)
        except StartFehler as fehler_obj:
            _notiere(u"%s: Start\n%s" % (beschreibe(ziel),
                                         traceback.format_exc()))
            if erste_streng and nummer == 0:
                raise
            fehler.append((ziel, _kurz(fehler_obj)))
        except Exception as fehler_obj:
            _notiere(u"%s\n%s" % (beschreibe(ziel), traceback.format_exc()))
            fehler.append((ziel, _kurz(fehler_obj)))
        else:
            _notiere(u"%s: ok" % beschreibe(ziel))
    return fertig, fehler


def _kurz(fehler_obj):
    zeilen = (u"%s" % fehler_obj).strip().splitlines()
    return zeilen[0] if zeilen else type(fehler_obj).__name__


def uebertrage(doc, quelle, ziele):
    """Profil der Quelle auf alle Ziele. -> (fertig, [(Wand, Grund)])."""
    name = t(u"Wandprofil übertragen", u"Transfer wall profile",
             u"Transferir perfil de muro")
    del _zeilen[:]
    quellkurven = _profil_kurven(doc, quelle)
    _notiere(u"Quelle %s, %d Kurven, Ziele: %s" % (
        beschreibe(quelle), len(quellkurven),
        u", ".join(beschreibe(z) for z in ziele)))
    try:
        return _uebertrage_alle(doc, quellkurven, quelle, ziele, name)
    finally:
        schreibe_protokoll()


def _uebertrage_alle(doc, quellkurven, quelle, ziele, name):
    gruppe = TransactionGroup(doc, name)
    gruppe.Start()
    try:
        ergebnis = _alle(doc, quellkurven, quelle, ziele, name,
                         erste_streng=True)
    except StartFehler:
        # Vielleicht duldet Revit den Skizzenmodus nicht in einer Gruppe:
        # einzeln ohne Gruppe (dann ein Rückgängig-Schritt je Wand)
        gruppe.RollBack()
        _notiere(u"-> ohne Transaktionsgruppe neu")
        return _alle(doc, quellkurven, quelle, ziele, name)
    except Exception:
        gruppe.RollBack()
        raise
    gruppe.Assimilate()
    return ergebnis
