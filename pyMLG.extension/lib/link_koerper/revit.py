# -*- coding: utf-8 -*-
"""Revit-Zugriffe von LinkKoerper.

Ablauf je Element der Verknüpfung (siehe logik.py für den Abgleich):

    1. Volumenkörper des Elements lesen und mit der Transformation der
       Verknüpfung(en) ins Hauptmodell rechnen
    2. Familiendokument aus der Vorlage "Allgemeines Modell" anlegen, die
       Körper - um ihre Mitte verschoben - als FreeFormElement einsetzen
    3. Familie unter ihrem Namen speichern, laden, Datei wieder löschen
    4. Eine Instanz in der Mitte des Körpers platzieren

Document.LoadFamily darf nur ohne offene Transaktion laufen - eine
TransactionGroup stört nicht. Der ganze Lauf ist deshalb ein Rückgängig-
Schritt, obwohl er aus vielen Transaktionen besteht.
"""

import io
import os
import time
import uuid

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    Category,
    CategoryType,
    ElementId,
    ElementTransformUtils,
    Family,
    FamilyInstance,
    FailureProcessingResult,
    FailureSeverity,
    FilteredElementCollector,
    FreeFormElement,
    GeometryInstance,
    IFailuresPreprocessor,
    Level,
    Options,
    RevitLinkInstance,
    SaveAsOptions,
    Solid,
    SolidUtils,
    Transaction,
    TransactionGroup,
    Transform,
    ViewDetailLevel,
    XYZ,
)
from Autodesk.Revit.DB.Structure import StructuralType  # noqa: E402
from Autodesk.Revit.Exceptions import OperationCanceledException  # noqa: E402
from Autodesk.Revit.UI.Selection import ObjectType  # noqa: E402

from link_koerper import logik as lg  # noqa: E402
from linked_ids import revit as li  # noqa: E402
from mlg_sprache import t  # noqa: E402

M_PRO_FUSS = 0.3048
M3_PRO_FUSS3 = M_PRO_FUSS ** 3

# Kategorien ohne sinnvollen Volumenkörper bzw. mit doppelter Geometrie
# (Gruppen enthalten ihre Elemente noch einmal)
_OHNE = (
    "OST_Rooms", "OST_MEPSpaces", "OST_Areas", "OST_Lines", "OST_Cameras",
    "OST_RvtLinks", "OST_IOSModelGroups", "OST_ProjectBasePoint",
    "OST_SharedBasePoint", "OST_SketchLines", "OST_HVAC_Zones",
    "OST_PointClouds", "OST_Levels", "OST_Grids",
)

# Vorlagen "Allgemeines Modell" nach Sprache der Revit-Installation
_VORLAGEN = (
    u"metric generic model.rft", u"m_generic model.rft",
    u"generic model.rft", u"allgemeines modell.rft",
    u"m_allgemeines modell.rft", u"allgemeines modell metrisch.rft",
    u"modelo genérico métrico.rft", u"m_modelo genérico.rft",
    u"modelo genérico.rft",
)
_VORLAGE_WORT = (u"generic model", u"allgemeines modell", u"modelo genérico",
                 u"modelo generico")
# Varianten mit Host oder Sonderverhalten
_VORLAGE_NICHT = (u"face", u"line", u"wall", u"ceiling", u"floor", u"roof",
                  u"adaptive", u"pattern", u"hosted", u"two level",
                  u"fläche", u"flaeche", u"linie", u"wand", u"decke",
                  u"boden", u"dach", u"adaptiv", u"muster", u"basiert",
                  u"cara", u"línea", u"linea", u"muro", u"techo", u"suelo",
                  u"cubierta", u"adaptativo", u"patrón", u"basado")

PROTOKOLL = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "pyMLG",
    "LinkKoerper_Protokoll.txt")

id_wert = li.id_wert


def _ohne_ids():
    ids = set()
    for name in _OHNE:
        kategorie = getattr(BuiltInCategory, name, None)
        if kategorie is not None:
            ids.add(id_wert(ElementId(kategorie)))
    return ids


# ---------------------------------------------------------------- Quellen

class Quelle(object):
    """Ein Element der Verknüpfung mit seinen Körpern im Hauptmodell."""

    def __init__(self, link_ids, link_titel, element, transform):
        self.link_ids = link_ids
        self.link_titel = link_titel
        self.element = element
        self.element_id = id_wert(element.Id)
        self.transform = transform
        kategorie = element.Category
        self.kategorie = kategorie.Name if kategorie is not None else u""
        self.solids = []           # Hauptmodell-Koordinaten
        self.mitte = None          # XYZ in Fuss
        self.mitte_m = None        # (x, y, z) in Metern
        self.form = None
        self.fehler = None

    @property
    def schluessel(self):
        return lg.schluessel(self.link_ids, self.element_id)

    @property
    def familienname(self):
        return lg.familienname(self.link_ids, self.element_id,
                               self.link_titel, self.kategorie)

    @property
    def beschreibung(self):
        return u"%s %d (%s)" % (self.kategorie, self.element_id,
                                self.link_titel)


def verknuepfungen(doc):
    """Geladene Verknüpfungen: [(Anzeigetext, RevitLinkInstance), ...]."""
    ergebnis = []
    for instanz in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        try:
            link_doc = instanz.GetLinkDocument()
        except Exception:
            link_doc = None
        if link_doc is None:
            continue
        ergebnis.append((u"%s  (%s)" % (link_doc.Title, instanz.Name),
                         instanz))
    ergebnis.sort(key=lambda paar: paar[0].lower())
    return ergebnis


def elemente_nach_kategorie(link_doc):
    """{Kategorie-ID: (Name, [Elemente])} der Modellelemente der Verknüpfung.

    Nur Elemente mit Ausdehnung und ohne Ansichtsbezug - ob sie einen
    Volumenkörper haben, zeigt sich erst beim Erstellen.
    """
    ohne = _ohne_ids()
    gruppen = {}
    for element in FilteredElementCollector(link_doc) \
            .WhereElementIsNotElementType():
        try:
            kategorie = element.Category
            if kategorie is None or \
                    kategorie.CategoryType != CategoryType.Model:
                continue
            schluessel = id_wert(kategorie.Id)
            if schluessel in ohne or isinstance(element, RevitLinkInstance):
                continue
            if id_wert(element.OwnerViewId) > 0:
                continue
            if element.get_BoundingBox(None) is None:
                continue
        except Exception:
            continue
        if schluessel not in gruppen:
            gruppen[schluessel] = (kategorie.Name, [])
        gruppen[schluessel][1].append(element)
    return gruppen


def aus_kategorien(instanz, gruppen, kategorie_ids):
    link_doc = instanz.GetLinkDocument()
    transform = instanz.GetTotalTransform()
    quellen = []
    for schluessel in kategorie_ids:
        for element in gruppen.get(schluessel, (u"", []))[1]:
            quellen.append(Quelle([id_wert(instanz.Id)], link_doc.Title,
                                  element, transform))
    return quellen


def _ketten_transform(kette):
    gesamt = None
    for instanz in kette:
        if instanz is None:
            return None
        einzeln = instanz.GetTotalTransform()
        gesamt = einzeln if gesamt is None else gesamt.Multiply(einzeln)
    return gesamt


def aus_referenzen(doc, referenzen):
    """Referenzen auf Elemente in Verknüpfungen -> Quellen (ohne doppelte).

    Elemente des Hauptmodells und nicht geladene Verknüpfungen fallen weg.
    """
    quellen = []
    gesehen = set()
    for referenz in referenzen or []:
        dokument, element, _wert, kette = li.aufloesen(doc, referenz)
        if element is None or not kette or \
                isinstance(element, RevitLinkInstance):
            continue
        transform = _ketten_transform(kette)
        if transform is None:
            continue
        quelle = Quelle([id_wert(instanz.Id) for instanz in kette],
                        dokument.Title, element, transform)
        if quelle.schluessel in gesehen:
            continue
        gesehen.add(quelle.schluessel)
        quellen.append(quelle)
    return quellen


def referenzen_der_auswahl(uidoc):
    """Angetippte Elemente in Verknüpfungen (Revit 2023+), sonst []."""
    try:
        referenzen = list(uidoc.Selection.GetReferences())
    except Exception:
        return []
    return [referenz for referenz in referenzen
            if id_wert(getattr(referenz, "LinkedElementId", None)) > 0]


def waehle(uidoc):
    """Elemente in Verknüpfungen picken. None bei Abbruch (Esc)."""
    hinweis = t(u"Elemente in der Verknüpfung wählen - mit Tab das einzelne "
                u"Element antippen, Fertigstellen beendet die Wahl",
                u"Select elements inside the link - press Tab to reach the "
                u"single element, Finish ends the selection",
                u"Seleccione elementos del vínculo: con Tab llega al "
                u"elemento concreto, Finalizar termina la selección")
    try:
        return list(uidoc.Selection.PickObjects(ObjectType.LinkedElement,
                                                hinweis))
    except OperationCanceledException:
        return None


# ---------------------------------------------------------------- Geometrie

def _optionen():
    optionen = Options()
    optionen.DetailLevel = ViewDetailLevel.Fine
    optionen.ComputeReferences = False
    optionen.IncludeNonVisibleObjects = False
    return optionen


def _sammle_solids(geometrie, sammlung):
    if geometrie is None:
        return
    for objekt in geometrie:
        if isinstance(objekt, Solid):
            try:
                if objekt.Volume > 1e-9 and objekt.Faces.Size > 0:
                    sammlung.append(objekt)
            except Exception:
                pass
        elif isinstance(objekt, GeometryInstance):
            # GetInstanceGeometry liegt schon in den Koordinaten des
            # Dokuments, in dem das Element steht
            _sammle_solids(objekt.GetInstanceGeometry(), sammlung)


def lies_geometrie(quelle, optionen=None):
    """Körper, Mitte und Formkennung der Quelle bestimmen.

    Rückgabe: True, wenn die Quelle einen Volumenkörper hat.
    """
    roh = []
    _sammle_solids(quelle.element.get_Geometry(optionen or _optionen()), roh)
    punkte = []
    volumen = 0.0
    for solid in roh:
        try:
            solid = SolidUtils.CreateTransformed(solid, quelle.transform)
        except Exception:
            continue
        quelle.solids.append(solid)
        volumen += solid.Volume
        for kante in solid.Edges:
            for punkt in kante.Tessellate():
                punkte.append((punkt.X * M_PRO_FUSS, punkt.Y * M_PRO_FUSS,
                               punkt.Z * M_PRO_FUSS))
    if not quelle.solids or not punkte:
        quelle.solids = []
        return False
    quelle.mitte_m = lg.mitte(punkte)
    quelle.mitte = XYZ(*(wert / M_PRO_FUSS for wert in quelle.mitte_m))
    relativ = [tuple(p[i] - quelle.mitte_m[i] for i in range(3))
               for p in punkte]
    quelle.form = lg.formkennung(relativ, volumen * M3_PRO_FUSS3)
    return True


# ---------------------------------------------------------------- Bestand

class Vorhanden(object):
    """Schon erstellter Körper im Hauptmodell."""

    def __init__(self, familie):
        self.familie = familie
        self.instanz = None
        self.form = None
        self.mitte_m = None


def vorhandene(doc):
    """{Schlüssel: Vorhanden} der Körper früherer Läufe."""
    nach_schluessel = {}
    nach_familie = {}
    for familie in FilteredElementCollector(doc).OfClass(Family):
        schluessel = lg.schluessel_aus_name(familie.Name)
        if schluessel is None:
            continue
        eintrag = Vorhanden(familie)
        nach_schluessel[schluessel] = eintrag
        nach_familie[id_wert(familie.Id)] = eintrag
    if not nach_familie:
        return nach_schluessel
    for instanz in FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_GenericModel) \
            .OfClass(FamilyInstance):
        try:
            symbol = instanz.Symbol
            eintrag = nach_familie.get(id_wert(symbol.Family.Id))
        except Exception:
            continue
        if eintrag is None or eintrag.instanz is not None:
            continue
        eintrag.instanz = instanz
        eintrag.form = lg.form_aus_typname(symbol.Name)
        punkt = getattr(instanz.Location, "Point", None)
        if punkt is not None:
            eintrag.mitte_m = (punkt.X * M_PRO_FUSS, punkt.Y * M_PRO_FUSS,
                               punkt.Z * M_PRO_FUSS)
    return nach_schluessel


def zaehle_verwaiste(bestand, instanz):
    """Körper zu dieser Verknüpfung, deren Element dort fehlt."""
    link_doc = instanz.GetLinkDocument()
    kette = u"%d" % id_wert(instanz.Id)
    anzahl = 0
    for (link_teil, element_id) in bestand:
        if link_teil != kette:
            continue
        try:
            if link_doc.GetElement(ElementId(element_id)) is None:
                anzahl += 1
        except Exception:
            pass
    return anzahl


# ---------------------------------------------------------------- Vorlage

def finde_vorlage(app, gespeichert=None):
    """Pfad der Familienvorlage "Allgemeines Modell", oder None."""
    if gespeichert and os.path.isfile(gespeichert):
        return gespeichert
    wurzel = u""
    try:
        wurzel = app.FamilyTemplatePath or u""
    except Exception:
        pass
    if not wurzel or not os.path.isdir(wurzel):
        return None
    kandidaten = []
    for ordner, _unter, dateien in os.walk(wurzel):
        for datei in dateien:
            if datei.lower().endswith(u".rft"):
                kandidaten.append(os.path.join(ordner, datei))
    for name in _VORLAGEN:
        for pfad in kandidaten:
            if os.path.basename(pfad).lower() == name:
                return pfad
    for pfad in kandidaten:
        name = os.path.basename(pfad).lower()
        if any(wort in name for wort in _VORLAGE_WORT) and \
                not any(wort in name for wort in _VORLAGE_NICHT):
            return pfad
    return None


# ---------------------------------------------------------------- Warnungen

_klassen = {}


def _vorverarbeiter():
    """Warnungen still bestätigen - sonst ein Dialog je Transaktion.

    Der Namensraum ist je Sitzung eindeutig, weil die CPython-Engine im
    Revit-Prozess weiterlebt und ein zweiter Typ gleichen Namens scheitert.
    """
    if "fehler" not in _klassen:
        namensraum = "pyMLG.LinkKoerper_%s" % uuid.uuid4().hex

        class Vorverarbeiter(IFailuresPreprocessor):
            __namespace__ = namensraum

            def PreprocessFailures(self, accessor):
                for meldung in list(accessor.GetFailureMessages()):
                    if meldung.GetSeverity() == FailureSeverity.Warning:
                        accessor.DeleteWarning(meldung)
                return FailureProcessingResult.Continue

        _klassen["fehler"] = Vorverarbeiter
    return _klassen["fehler"]()


def _transaktion(dokument, name):
    transaktion = Transaction(dokument, name)
    transaktion.Start()
    try:
        optionen = transaktion.GetFailureHandlingOptions()
        optionen.SetFailuresPreprocessor(_vorverarbeiter())
        transaktion.SetFailureHandlingOptions(optionen)
    except Exception:
        pass
    return transaktion


# ---------------------------------------------------------------- Erstellen

def _arbeitsordner():
    ordner = os.path.join(os.path.dirname(PROTOKOLL), "LinkKoerper")
    if not os.path.isdir(ordner):
        os.makedirs(ordner)
    return ordner


def _loesche_dateien(ordner, name):
    """Die gespeicherte Familie samt Sicherungen (name.0001.rfa)."""
    for datei in os.listdir(ordner):
        if datei == name + u".rfa" or (datei.startswith(name + u".")
                                       and datei.endswith(u".rfa")):
            try:
                os.remove(os.path.join(ordner, datei))
            except Exception:
                pass


def _baue_familie(app, doc, vorlage, quelle, ordner):
    """Familie mit den Körpern der Quelle bauen und in doc laden."""
    name = quelle.familienname
    fam_doc = app.NewFamilyDocument(vorlage)
    if fam_doc is None:
        raise RuntimeError(t(u"Die Familienvorlage lässt sich nicht öffnen: ",
                             u"The family template cannot be opened: ",
                             u"No se puede abrir la plantilla de familia: ")
                           + vorlage)
    try:
        transaktion = _transaktion(fam_doc, u"pyMLG Link")
        try:
            try:
                fam_doc.OwnerFamily.FamilyCategory = Category.GetCategory(
                    fam_doc, BuiltInCategory.OST_GenericModel)
            except Exception:
                pass
            verschiebung = Transform.CreateTranslation(quelle.mitte.Negate())
            eingesetzt = 0
            letzter_fehler = None
            for solid in quelle.solids:
                try:
                    FreeFormElement.Create(
                        fam_doc,
                        SolidUtils.CreateTransformed(solid, verschiebung))
                    eingesetzt += 1
                except Exception as fehler:
                    letzter_fehler = fehler
            if not eingesetzt:
                raise RuntimeError(t(
                    u"Revit nimmt die Geometrie nicht als Körper an",
                    u"Revit does not accept the geometry as a solid",
                    u"Revit no acepta la geometría como sólido")
                    + (u": %s" % letzter_fehler if letzter_fehler else u""))
            verwaltung = fam_doc.FamilyManager
            if verwaltung.CurrentType is None:
                verwaltung.NewType(lg.typname(quelle.form))
            else:
                verwaltung.RenameCurrentType(lg.typname(quelle.form))
            transaktion.Commit()
        except Exception:
            if transaktion.HasStarted() and not transaktion.HasEnded():
                transaktion.RollBack()
            raise
        optionen = SaveAsOptions()
        optionen.OverwriteExistingFile = True
        fam_doc.SaveAs(os.path.join(ordner, name + u".rfa"), optionen)
        return fam_doc.LoadFamily(doc)
    finally:
        fam_doc.Close(False)
        _loesche_dateien(ordner, name)


def _ebenen(doc):
    ebenen = list(FilteredElementCollector(doc).OfClass(Level))
    ebenen.sort(key=lambda ebene: ebene.Elevation)
    return ebenen


def _ebene_unter(ebenen, hoehe):
    """Höchste Ebene unter hoehe, sonst die unterste."""
    passend = [ebene for ebene in ebenen if ebene.Elevation <= hoehe + 1e-6]
    if passend:
        return passend[-1]
    return ebenen[0] if ebenen else None


def _platziere(doc, familie, mitte, ebenen):
    symbol = None
    for symbol_id in familie.GetFamilySymbolIds():
        symbol = doc.GetElement(symbol_id)
        break
    if symbol is None:
        raise RuntimeError(t(u"Die geladene Familie hat keinen Typ",
                             u"The loaded family has no type",
                             u"La familia cargada no tiene tipo"))
    if not symbol.IsActive:
        symbol.Activate()
        doc.Regenerate()
    ebene = _ebene_unter(ebenen, mitte.Z)
    if ebene is not None:
        instanz = doc.Create.NewFamilyInstance(
            mitte, symbol, ebene, StructuralType.NonStructural)
    else:
        instanz = doc.Create.NewFamilyInstance(
            mitte, symbol, StructuralType.NonStructural)
    doc.Regenerate()
    # Je nach Vorlage rechnet Revit die Höhe relativ zur Ebene - nachführen
    punkt = getattr(instanz.Location, "Point", None)
    if punkt is not None:
        rest = mitte.Subtract(punkt)
        if rest.GetLength() > 1e-6:
            ElementTransformUtils.MoveElement(doc, instanz.Id, rest)
    return instanz


def _verschiebe(doc, instanz, ziel):
    punkt = instanz.Location.Point
    gepinnt = instanz.Pinned
    if gepinnt:
        instanz.Pinned = False
    ElementTransformUtils.MoveElement(doc, instanz.Id, ziel.Subtract(punkt))
    if gepinnt:
        instanz.Pinned = True


def plane(quellen, bestand, ergebnis):
    """[(Quelle, Aktion, Vorhanden oder None)] - liest dafür die Geometrie."""
    optionen = _optionen()
    plan = []
    for quelle in quellen:
        try:
            hat_koerper = lies_geometrie(quelle, optionen)
        except Exception as fehler:
            ergebnis.fehler.append(u"%s: %s" % (quelle.beschreibung, fehler))
            continue
        if not hat_koerper:
            ergebnis.ohne_geometrie += 1
            continue
        alt = bestand.get(quelle.schluessel)
        aktion = lg.entscheide(alt.form if alt else None,
                               alt.mitte_m if alt else None,
                               quelle.form, quelle.mitte_m)
        plan.append((quelle, aktion, alt))
    return plan


def ausfuehren(doc, app, plan, vorlage, ergebnis, fortschritt=None):
    """Den Plan umsetzen - ein Rückgängig-Schritt für alles."""
    gruppe = TransactionGroup(doc, t(u"Link als Allgemeines Modell",
                                     u"Link to generic model",
                                     u"Vínculo a modelo genérico"))
    gruppe.Start()
    try:
        _ausfuehren(doc, app, plan, vorlage, ergebnis, fortschritt)
    except Exception:
        gruppe.RollBack()
        raise
    gruppe.Assimilate()


def _ausfuehren(doc, app, plan, vorlage, ergebnis, fortschritt):
    for quelle, aktion, _alt in plan:
        if aktion == lg.GLEICH:
            ergebnis.zaehle(aktion)

    verschieben = [(q, alt) for q, aktion, alt in plan
                   if aktion == lg.VERSCHIEBEN]
    if verschieben:
        transaktion = _transaktion(doc, u"pyMLG Link: verschieben")
        for quelle, alt in verschieben:
            try:
                _verschiebe(doc, alt.instanz, quelle.mitte)
                ergebnis.zaehle(lg.VERSCHIEBEN)
            except Exception as fehler:
                ergebnis.fehler.append(u"%s: %s" % (quelle.beschreibung,
                                                    fehler))
        transaktion.Commit()

    ersetzen = [alt for _q, aktion, alt in plan if aktion == lg.ERSETZEN]
    if ersetzen:
        # Die Familie samt Instanz löschen, damit der Name frei wird
        transaktion = _transaktion(doc, u"pyMLG Link: alte Körper löschen")
        for alt in ersetzen:
            doc.Delete(alt.familie.Id)
        transaktion.Commit()

    bauen = [(q, aktion) for q, aktion, _alt in plan
             if aktion in (lg.NEU, lg.ERSETZEN)]
    if not bauen:
        return
    ordner = _arbeitsordner()
    ebenen = _ebenen(doc)
    for nummer, (quelle, aktion) in enumerate(bauen):
        if fortschritt is not None:
            fortschritt.aktualisiere(nummer, len(bauen))
        try:
            familie = _baue_familie(app, doc, vorlage, quelle, ordner)
            transaktion = _transaktion(doc, u"pyMLG Link: platzieren")
            try:
                instanz = _platziere(doc, familie, quelle.mitte, ebenen)
                transaktion.Commit()
            except Exception:
                transaktion.RollBack()
                raise
            ergebnis.neue_ids.append(instanz.Id)
            ergebnis.zaehle(aktion)
        except Exception as fehler:
            ergebnis.fehler.append(u"%s: %s" % (quelle.beschreibung,
                                                _fehlertext(fehler)))


def _fehlertext(fehler):
    text = getattr(fehler, "Message", None) or u"%s" % fehler
    zeilen = [zeile.strip() for zeile in text.splitlines() if zeile.strip()]
    return zeilen[0] if zeilen else fehler.__class__.__name__


def schreibe_protokoll(ergebnis, dauer):
    try:
        with io.open(PROTOKOLL, "w", encoding="utf-8") as datei:
            datei.write(u"pyMLG LinkKoerper  %s  (%.0f s)\n\n" % (
                time.strftime("%Y-%m-%d %H:%M"), dauer))
            datei.write(ergebnis.text() + u"\n")
            if ergebnis.fehler:
                datei.write(u"\n" + u"\n".join(ergebnis.fehler) + u"\n")
        return PROTOKOLL
    except Exception:
        return None
