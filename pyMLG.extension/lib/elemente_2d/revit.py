# -*- coding: utf-8 -*-
"""Revit-Zugriffe von Elements2D.

2D-Element = ansichtsspezifisches Element (ViewSpecific, mit Besitzer-
Ansicht) einer Modell- oder Beschriftungskategorie. Nicht dazu gehören:

    - Ansichten und alles, was eine Ansicht darstellt oder auf einem Plan
      platziert (Ansichtsfenster, Planköpfe, Bauteillisten auf Plänen,
      Schnitt-/Ansichts-/Detailmarken, Zuschneidebereiche) - Löschen würde
      Ansichten oder Pläne zerstören
    - Mitglieder von Gruppen - die Gruppe selbst steht im Baum
    - interne Elemente (Skizzen, Arbeitsebenenraster)

Ein Durchlauf über alle Elemente des Projekts; Ansichts- und Typangaben
werden zwischengespeichert.
"""

import uuid

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    CategoryType,
    CurveElement,
    ElementId,
    ElementType,
    ElevationMarker,
    FailureProcessingResult,
    FailureSeverity,
    FilteredElementCollector,
    Group,
    IFailuresPreprocessor,
    ImportInstance,
    ScheduleSheetInstance,
    Sketch,
    SketchPlane,
    SubTransaction,
    TextNote,
    Transaction,
    TransactionStatus,
    View,
    ViewSheet,
    Viewport,
)

from elemente_2d import logik as lg  # noqa: E402
from filter_more.baum import natuerlich  # noqa: E402
from filter_more.revit import Namen, id_wert, net_liste  # noqa: E402
from mlg_sprache import t  # noqa: E402

AUSGESCHLOSSEN = (
    "OST_Viewports", "OST_TitleBlocks", "OST_ScheduleGraphics", "OST_Views",
    "OST_Viewers", "OST_Elev", "OST_ElevationMarks", "OST_Sections",
    "OST_SectionHeads", "OST_Callouts", "OST_CalloutHeads",
    "OST_ReferenceViewer", "OST_ReferenceViewerSymbol", "OST_CropBoundary",
    "OST_AnnotationCrop", "OST_SketchLines", "OST_IOSSketchGrid",
    "OST_Cameras", "OST_SectionBox",
)

AUSGESCHLOSSENE_KLASSEN = (View, Viewport, ScheduleSheetInstance,
                           ElevationMarker)

IMPORTE = t(u"Importierte CAD-Dateien", u"Imported CAD files",
            u"Archivos CAD importados")
LINIENSTIL = t(u"Linienstil", u"Line style", u"Estilo de línea")

ANSICHTSARTEN = {
    "FloorPlan": (u"Grundrisse", u"Floor plans", u"Plantas"),
    "CeilingPlan": (u"Deckenpläne", u"Ceiling plans", u"Planos de techo"),
    "EngineeringPlan": (u"Tragwerkspläne", u"Structural plans",
                        u"Planos estructurales"),
    "AreaPlan": (u"Flächenpläne", u"Area plans", u"Planos de área"),
    "Section": (u"Schnitte", u"Sections", u"Secciones"),
    "Elevation": (u"Ansichten", u"Elevations", u"Alzados"),
    "Detail": (u"Detailansichten", u"Detail views", u"Vistas de detalle"),
    "DraftingView": (u"Zeichenansichten", u"Drafting views",
                     u"Vistas de diseño"),
    "Legend": (u"Legenden", u"Legends", u"Leyendas"),
    "DrawingSheet": (u"Pläne", u"Sheets", u"Planos"),
    "ThreeD": (u"3D-Ansichten", u"3D views", u"Vistas 3D"),
    "Rendering": (u"Renderings", u"Renderings", u"Renderizados"),
}


def _text(funktion, ersatz):
    try:
        wert = funktion()
        return u"%s" % wert if wert else ersatz
    except Exception:
        return ersatz


def _ids(sammlung):
    try:
        return set(id_wert(i) for i in sammlung)
    except Exception:
        return set()


def _ausgeschlossene_kategorien():
    ergebnis = set()
    for name in AUSGESCHLOSSEN:
        kategorie = getattr(BuiltInCategory, name, None)
        if kategorie is None:
            continue
        try:
            ergebnis.add(id_wert(ElementId(kategorie)))
        except Exception:
            pass
    return ergebnis


def platzierte_ansichten(doc):
    """{Ansicht-Id: {Plan-Id}} der Ansichten, die auf einem Plan liegen. Bei
    abhängigen Ansichten zählt auch die Hauptansicht - ihr gehören die
    Beschriftungen."""
    ergebnis = {}
    for viewport in FilteredElementCollector(doc).OfClass(Viewport):
        try:
            plan = id_wert(viewport.SheetId)
            ansichten = [id_wert(viewport.ViewId)]
            ansicht = doc.GetElement(viewport.ViewId)
            if ansicht is not None:
                ansichten.append(id_wert(ansicht.GetPrimaryViewId()))
        except Exception:
            continue
        for wert in ansichten:
            if wert != -1:
                ergebnis.setdefault(wert, set()).add(plan)
    return ergebnis


def gewaehlte_plaene(uidoc):
    """Plan-Ids der aktuellen Revit-Auswahl (z.B. im Projektbrowser
    markierte Pläne); None, wenn kein Plan ausgewählt ist."""
    doc = uidoc.Document
    ergebnis = set()
    for element_id in uidoc.Selection.GetElementIds():
        if isinstance(doc.GetElement(element_id), ViewSheet):
            ergebnis.add(id_wert(element_id))
    return ergebnis or None


def plaene(doc):
    """[('Nummer - Name', Plan-Id)] aller Pläne, natürlich sortiert."""
    ergebnis = []
    for plan in FilteredElementCollector(doc).OfClass(ViewSheet):
        try:
            if plan.IsTemplate:
                continue
        except Exception:
            pass
        ergebnis.append((u"%s - %s" % (_text(lambda: plan.SheetNumber, u"?"),
                                       _text(lambda: plan.Name, u"?")),
                         id_wert(plan.Id)))
    return sorted(ergebnis, key=lambda e: natuerlich(e[0]))


class _Ansichten(object):
    """Zwischenspeicher Ansicht-Id -> (Name, Art, ist Plan) oder None."""

    def __init__(self, doc):
        self.doc = doc
        self._daten = {}

    def info(self, wert, element_id):
        if wert not in self._daten:
            self._daten[wert] = self._lies(element_id)
        return self._daten[wert]

    def _lies(self, element_id):
        ansicht = self.doc.GetElement(element_id)
        if not isinstance(ansicht, View):
            return None
        try:
            if ansicht.IsTemplate:
                return None
        except Exception:
            pass
        ist_plan = isinstance(ansicht, ViewSheet)
        if ist_plan:
            name = u"%s - %s" % (_text(lambda: ansicht.SheetNumber, u"?"),
                                 _text(lambda: ansicht.Name, u"?"))
        else:
            name = _text(lambda: ansicht.Name, u"?")
        art_schluessel = _text(lambda: ansicht.ViewType, u"?")
        texte = ANSICHTSARTEN.get(art_schluessel)
        art = t(*texte) if texte else art_schluessel
        return name, art, ist_plan


def _ist_2d(element, ausgeschlossen):
    if isinstance(element, Group):
        # Detailgruppen und angehängte Detailgruppen: Kategorie und
        # ViewSpecific sind bei Gruppen nicht verlässlich - maßgeblich ist
        # die Besitzer-Ansicht. Verschachtelte Gruppen stecken in ihrer
        # äußeren Gruppe.
        try:
            return (id_wert(element.OwnerViewId) != -1
                    and id_wert(element.GroupId) == -1)
        except Exception:
            return False
    try:
        if not element.ViewSpecific:
            return False
    except Exception:
        return False
    if isinstance(element, AUSGESCHLOSSENE_KLASSEN):
        return False
    if id_wert(element.GroupId) != -1:
        return False
    kategorie = element.Category
    if kategorie is None or id_wert(kategorie.Id) in ausgeschlossen:
        return False
    if isinstance(element, ImportInstance):
        return True
    try:
        return kategorie.CategoryType in (CategoryType.Model,
                                          CategoryType.Annotation)
    except Exception:
        return False


def _eintrag(element, namen, ansichten, platziert):
    besitzer = element.OwnerViewId
    wert = id_wert(besitzer)
    if wert == -1:
        return None
    info = ansichten.info(wert, besitzer)
    if info is None:
        return None
    ansicht, art, ist_plan = info

    _kid, kategorie = namen.kategorie(element)
    typ_id, familie, typ = namen.typ(element)
    name = None
    if isinstance(element, ImportInstance):
        familie, typ, kategorie = kategorie, kategorie, IMPORTE
    elif typ_id == -1 and isinstance(element, CurveElement):
        familie = LINIENSTIL
        typ = _text(lambda: element.LineStyle.Name, typ)
    if isinstance(element, TextNote):
        name = lg.kurz(_text(lambda: element.Text, u""))
    elif isinstance(element, Group):
        name = _text(lambda: element.Name, typ)
    auf_plaenen = {wert} if ist_plan else platziert.get(wert, ())
    return lg.Eintrag(id_wert(element.Id), kategorie, familie, typ, name,
                      ansicht, art, auf_plan=ist_plan, plaene=auf_plaenen)


def sammle(doc):
    """-> ([Eintrag], {Id-Wert: Element}) aller 2D-Elemente des Projekts."""
    ausgeschlossen = _ausgeschlossene_kategorien()
    platziert = platzierte_ansichten(doc)
    namen = Namen(doc)
    ansichten = _Ansichten(doc)
    eintraege, elemente = [], {}
    for element in FilteredElementCollector(doc) \
            .WhereElementIsNotElementType():
        if not _ist_2d(element, ausgeschlossen):
            continue
        try:
            eintrag = _eintrag(element, namen, ansichten, platziert)
        except Exception:
            continue
        if eintrag is not None:
            eintraege.append(eintrag)
            elemente[eintrag.element_id] = element
    return eintraege, elemente


# ---------------------------------------------------------------------------
# Auswählen, Zeigen, Löschen
# ---------------------------------------------------------------------------

def waehle_aus(uidoc, ids):
    uidoc.Selection.SetElementIds(net_liste(ElementId,
                                            [ElementId(i) for i in ids]))


def zeige(uidoc, element):
    """Öffnet die Ansicht des Elements und zoomt darauf."""
    uidoc.ShowElements(element)
    uidoc.Selection.SetElementIds(net_liste(ElementId, [element.Id]))


_vorverarbeiter_klasse = None

# Fehlertexte, wegen derer Revit das Löschen zurückgerollt hat
REVIT_FEHLER = []


def _vorverarbeiter():
    """Warnungen löschen; bei einem Fehler zurückrollen statt ihn
    aufzulösen - eine Auflösung kann weitere Elemente löschen, die der
    Probelauf nicht gesehen hat. Namensraum je Sitzung eindeutig (siehe
    join_multiple/revit.py)."""
    global _vorverarbeiter_klasse
    if _vorverarbeiter_klasse is None:
        namensraum = "pyMLG.Elements2D_%s" % uuid.uuid4().hex

        class Vorverarbeiter(IFailuresPreprocessor):
            __namespace__ = namensraum

            def PreprocessFailures(self, accessor):
                fehler = False
                for meldung in list(accessor.GetFailureMessages()):
                    if meldung.GetSeverity() == FailureSeverity.Warning:
                        accessor.DeleteWarning(meldung)
                    else:
                        fehler = True
                        try:
                            REVIT_FEHLER.append(
                                meldung.GetDescriptionText())
                        except Exception:
                            pass
                if fehler:
                    return FailureProcessingResult.ProceedWithRollBack
                return FailureProcessingResult.Continue

        _vorverarbeiter_klasse = Vorverarbeiter
    return _vorverarbeiter_klasse()


def _fehlertext(fehler):
    text = getattr(fehler, "Message", None) or u"%s" % fehler
    zeilen = [z.strip() for z in text.splitlines() if z.strip()]
    return zeilen[0] if zeilen else fehler.__class__.__name__


def _starte_transaktion(doc, name):
    transaktion = Transaction(doc, name)
    transaktion.Start()
    try:
        optionen = transaktion.GetFailureHandlingOptions()
        optionen.SetFailuresPreprocessor(_vorverarbeiter())
        optionen.SetClearAfterRollback(True)
        transaktion.SetFailureHandlingOptions(optionen)
    except Exception:
        pass
    return transaktion


def _loesche_in_transaktion(doc, ids):
    """Zuerst alle auf einmal; scheitert das (z.B. ein Element gehört im
    Workset-Modell einem anderen Nutzer), einzeln - die übrigen werden
    trotzdem gelöscht.
    Rückgabe: (gelöschte Id-Werte inkl. abhängiger Elemente,
               [(Id-Wert, Fehlertext)])"""
    fehler = []
    alle = SubTransaction(doc)
    alle.Start()
    try:
        geloescht = _ids(doc.Delete(net_liste(
            ElementId, [ElementId(i) for i in ids])))
        alle.Commit()
        return geloescht, fehler
    except Exception:
        alle.RollBack()
    geloescht = set()
    for wert in ids:
        element_id = ElementId(wert)
        if wert in geloescht or doc.GetElement(element_id) is None:
            continue
        try:
            geloescht |= _ids(doc.Delete(element_id))
        except Exception as f:
            fehler.append((wert, _fehlertext(f)))
    return geloescht, fehler


def _ist_geschuetzt(element):
    """Modell-Element, das nie mitgelöscht werden darf: nicht ansichts-
    spezifisch (3D-Elemente, Modellgruppen, Raster, Ebenen,
    Referenzebenen ...). Typen, Ansichten, Skizzen und interne Elemente
    zählen nicht."""
    if element is None or isinstance(element, (ElementType, View, Sketch,
                                               SketchPlane)):
        return False
    kategorie = element.Category
    if kategorie is None or id_wert(kategorie.Id) in _skizzen_kategorien():
        return False
    try:
        if kategorie.CategoryType not in (CategoryType.Model,
                                          CategoryType.Annotation,
                                          CategoryType.AnalyticalModel):
            return False
        if isinstance(element, Group):
            return id_wert(element.OwnerViewId) == -1
        return not element.ViewSpecific
    except Exception:
        return True     # im Zweifel schützen


_skizzen = []


def _skizzen_kategorien():
    if not _skizzen:
        werte = set()
        for name in ("OST_SketchLines", "OST_IOSSketchGrid"):
            kategorie = getattr(BuiltInCategory, name, None)
            if kategorie is not None:
                werte.add(id_wert(ElementId(kategorie)))
        _skizzen.append(werte)
    return _skizzen[0]


def beschreibe(doc, wert):
    element = doc.GetElement(ElementId(wert))
    if element is None:
        return u"[%d]" % wert
    kategorie = _text(lambda: element.Category.Name, u"?")
    return u"%s: %s  [%d]" % (kategorie, _text(lambda: element.Name, u"?"),
                              wert)


def pruefe_loeschen(doc, ids):
    """Probelauf: löscht in einer Transaktion, die sofort zurückgenommen
    wird, und prüft, was Revit dabei mitgelöscht hätte.

    Rückgabe: ([(Id-Wert, Text)] der geschützten Modell-Elemente, die
               verschwinden würden,
               Id-Werte der markierten Elemente, an denen sie hängen)"""
    transaktion = _starte_transaktion(doc, u"pyMLG Probelauf")
    try:
        geloescht, _fehler = _loesche_in_transaktion(doc, ids)
    finally:
        transaktion.RollBack()
    gefaehrdet = set(w for w in geloescht - set(ids)
                     if _ist_geschuetzt(doc.GetElement(ElementId(w))))
    if not gefaehrdet:
        return [], set()
    ursachen = set()
    for wert in ids:
        element = doc.GetElement(ElementId(wert))
        try:
            if _ids(element.GetDependentElements(None)) & gefaehrdet:
                ursachen.add(wert)
        except Exception:
            continue
    return ([(w, beschreibe(doc, w)) for w in sorted(gefaehrdet)],
            ursachen)


def loesche(doc, ids):
    """Löscht die Elemente in EINER Transaktion (ein Strg+Z).
    Vorher pruefe_loeschen() aufrufen.
    Rückgabe: (gelöschte Id-Werte inkl. abhängiger Elemente,
               [(Id-Wert, Fehlertext)])
    """
    del REVIT_FEHLER[:]
    transaktion = _starte_transaktion(doc, t(u"2D-Elemente löschen",
                                             u"Delete 2D elements",
                                             u"Eliminar elementos 2D"))
    try:
        geloescht, fehler = _loesche_in_transaktion(doc, ids)
        status = transaktion.Commit()
    except Exception:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        raise
    if status != TransactionStatus.Committed:
        return set(), [(None, t(
            u"Revit hat das Löschen zurückgenommen - nichts wurde gelöscht.",
            u"Revit rolled the deletion back - nothing was deleted.",
            u"Revit ha deshecho la eliminación; no se ha eliminado nada."))
        ] + [(None, text) for text in REVIT_FEHLER]
    return geloescht, fehler
