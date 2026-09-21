# -*- coding: utf-8 -*-
"""Revit-Zugriffe von TransferSingle.

Kopiert wird mit ElementTransformUtils.CopyElements:
    Typen, Vorlagen, Filter, Materialien, Muster ...   Dokument -> Dokument
    Modellelemente                                      mit Transformation
    Zeichenansichten, Legenden, Bauteillisten           Ansicht als Ganzes;
                                                        Ansichtselemente bei
                                                        Bedarf View -> View
    Pläne                                               neuer Plan im Ziel,
                                                        Planelemente (Plankopf,
                                                        Texte) View -> View,
                                                        übertragbare Ansichten
                                                        neu platziert

Modellansichten (Grundrisse, Schnitte, 3D) lassen sich per API nicht zwischen
Dokumenten kopieren - sie werden im Protokoll genannt.

Doppelte Typnamen: IDuplicateTypeNamesHandler (Zieltypen verwenden, abbrechen
oder nachfragen). Revit-Warnungen beim Übertragen löscht auf Wunsch ein
IFailuresPreprocessor. Beide .NET-Klassen entstehen je Sitzung mit
eindeutigem Namensraum (siehe join_multiple/revit.py).

Jede Übertragungseinheit läuft in einer Untertransaktion: Scheitert sie,
wird nur sie zurückgenommen und im Protokoll vermerkt.
"""

import uuid

import clr

clr.AddReference("System")
from System.Collections.Generic import List  # noqa: E402

from Autodesk.Revit.DB import (  # noqa: E402
    CategoryType,
    CopyPasteOptions,
    DuplicateTypeAction,
    ElementId,
    ElementTransformUtils,
    FailureProcessingResult,
    FailureSeverity,
    FillPatternElement,
    FilteredElementCollector,
    IDuplicateTypeNamesHandler,
    IFailuresPreprocessor,
    LinePatternElement,
    Material,
    ParameterFilterElement,
    RevitLinkInstance,
    ScheduleSheetInstance,
    SelectionFilterElement,
    SubTransaction,
    Transaction,
    Transform,
    View,
    ViewSchedule,
    ViewSheet,
    ViewType,
    Viewport,
)
from Autodesk.Revit.UI import (  # noqa: E402
    TaskDialog,
    TaskDialogCommandLinkId,
    TaskDialogCommonButtons,
    TaskDialogResult,
)

from mlg_sprache import t  # noqa: E402
from transfer_single import logik as lg  # noqa: E402

# Transformation für Modellelemente
KEINE = "keine"
VERKNUEPFUNG = "verknuepfung"
GEMEINSAM = "gemeinsam"

# Doppelte Typnamen
DUP_ZIEL = "ziel"
DUP_ABBRECHEN = "abbrechen"
DUP_FRAGEN = "fragen"

T_TYPEN = t(u"Typen", u"Types", u"Tipos")
T_VORLAGEN = t(u"Ansichtsvorlagen", u"View templates", u"Plantillas de vista")
T_PLAENE = t(u"Pläne", u"Sheets", u"Planos")
T_ZEICHEN = t(u"Zeichenansichten", u"Drafting views", u"Vistas de diseño")
T_LEGENDEN = t(u"Legenden", u"Legends", u"Leyendas")
T_LISTEN = t(u"Bauteillisten", u"Schedules", u"Tablas de planificación")
T_FILTER = t(u"Filter", u"Filters", u"Filtros")
T_MATERIALIEN = t(u"Materialien", u"Materials", u"Materiales")
T_LINIENMUSTER = t(u"Linienmuster", u"Line patterns", u"Patrones de línea")
T_FUELLMUSTER = t(u"Füllmuster", u"Fill patterns", u"Patrones de relleno")
T_BROWSER = t(u"Browser-Organisation", u"Browser organization",
              u"Organización del navegador")
T_MODELL = t(u"Modell", u"Model", u"Modelo")


def id_wert(element_id):
    """Zahlenwert einer ElementId (Revit 2024+: .Value)."""
    if element_id is None:
        return -1
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def net_ids(ids):
    """ElementIds oder Zahlen -> List[ElementId] (pythonnet nimmt keine
    Python-Liste im Konstruktor an)."""
    liste = List[ElementId]()
    for wert in ids:
        liste.Add(wert if isinstance(wert, ElementId) else ElementId(wert))
    return liste


def _text(funktion, ersatz=u""):
    try:
        wert = funktion()
        return u"%s" % wert if wert else ersatz
    except Exception:
        return ersatz


def fehlertext(ausnahme):
    text = getattr(ausnahme, "Message", None) or u"%s" % ausnahme
    return (text.strip().splitlines() or [u"?"])[0]


# ---------------------------------------------------------------------------
# Quellen und Ziele
# ---------------------------------------------------------------------------

class Quelle(object):
    def __init__(self, doc, titel, link=None, host=None):
        self.doc = doc
        self.titel = titel
        self.link = link        # RevitLinkInstance oder None
        self.host = host        # Dokument, in dem die Verknüpfung liegt

    @property
    def ist_verknuepfung(self):
        return self.link is not None


def _projekte(app):
    return [d for d in app.Documents
            if not d.IsFamilyDocument and not d.IsLinked]


def quellen(app, mit_verknuepfungen):
    ergebnis = [Quelle(d, d.Title) for d in _projekte(app)]
    if mit_verknuepfungen:
        for host in _projekte(app):
            for instanz in FilteredElementCollector(host).OfClass(
                    RevitLinkInstance):
                link_doc = instanz.GetLinkDocument()
                if link_doc is None:
                    continue
                ergebnis.append(Quelle(
                    link_doc, u"%s  [%s %s: %s]" % (
                        link_doc.Title,
                        t(u"Verknüpfung in", u"link in", u"vínculo en"),
                        host.Title, _text(lambda: instanz.Name)),
                    instanz, host))
    return ergebnis


def ziele(app, quelle):
    """Offene Projekte außer der Quelle selbst."""
    return [d for d in _projekte(app) if not d.Equals(quelle.doc)]


# ---------------------------------------------------------------------------
# Elemente sammeln
# ---------------------------------------------------------------------------

def _typname(element):
    familie = _text(lambda: element.FamilyName)
    name = _text(lambda: element.Name, u"?")
    return u"%s: %s" % (familie, name) if familie else name


def _alle(doc, klasse):
    try:
        return list(FilteredElementCollector(doc).OfClass(klasse))
    except Exception:
        return []


def datensaetze(doc, mit_modell=False):
    """[(Gruppe, Bezeichnung, Id)] aller übertragbaren Elemente."""
    ergebnis = []

    def neu(gruppe, element, name):
        ergebnis.append((gruppe, name, id_wert(element.Id)))

    for typ in FilteredElementCollector(doc).WhereElementIsElementType():
        kategorie = _text(lambda: typ.Category.Name) or \
            _text(lambda: typ.GetType().Name, u"?")
        neu(u"%s – %s" % (kategorie, T_TYPEN), typ, _typname(typ))

    for ansicht in _alle(doc, View):
        try:
            if ansicht.IsTemplate:
                neu(T_VORLAGEN, ansicht, u"%s  (%s)" % (
                    ansicht.Name, ansicht.ViewType))
            elif isinstance(ansicht, ViewSheet):
                neu(T_PLAENE, ansicht, u"%s - %s" % (ansicht.SheetNumber,
                                                     ansicht.Name))
            elif ansicht.ViewType == ViewType.DraftingView:
                neu(T_ZEICHEN, ansicht, ansicht.Name)
            elif ansicht.ViewType == ViewType.Legend:
                neu(T_LEGENDEN, ansicht, ansicht.Name)
            elif isinstance(ansicht, ViewSchedule) \
                    and not ansicht.IsTitleblockRevisionSchedule \
                    and not getattr(ansicht, "IsInternalKeynoteSchedule",
                                    False):
                neu(T_LISTEN, ansicht, ansicht.Name)
        except Exception:
            continue

    for klasse, gruppe in ((ParameterFilterElement, T_FILTER),
                           (SelectionFilterElement, T_FILTER),
                           (Material, T_MATERIALIEN),
                           (LinePatternElement, T_LINIENMUSTER),
                           (FillPatternElement, T_FUELLMUSTER)):
        for element in _alle(doc, klasse):
            neu(gruppe, element, _text(lambda: element.Name, u"?"))
    try:
        from Autodesk.Revit.DB import BrowserOrganization
        for element in _alle(doc, BrowserOrganization):
            neu(T_BROWSER, element, _text(lambda: element.Name, u"?"))
    except ImportError:
        pass

    if mit_modell:
        typnamen = {}
        for element in FilteredElementCollector(doc) \
                .WhereElementIsNotElementType():
            try:
                kategorie = element.Category
                if kategorie is None or element.ViewSpecific \
                        or kategorie.CategoryType != CategoryType.Model \
                        or isinstance(element, (View, RevitLinkInstance)) \
                        or id_wert(element.GroupId) != -1:
                    continue
                typ_wert = id_wert(element.GetTypeId())
                if typ_wert not in typnamen:
                    typnamen[typ_wert] = _text(
                        lambda: doc.GetElement(element.GetTypeId()).Name,
                        _text(lambda: element.Name, u"?"))
                neu(u"%s: %s" % (T_MODELL, kategorie.Name), element,
                    u"%s  [%d]" % (typnamen[typ_wert], id_wert(element.Id)))
            except Exception:
                continue
    return ergebnis


# ---------------------------------------------------------------------------
# .NET-Hilfsklassen
# ---------------------------------------------------------------------------

_zustand = {"modus": DUP_ZIEL, "immer": False, "abgebrochen": False}
_klassen = {}
# Meldungen, die der Vorverarbeiter beim Speichern behandelt hat
REVIT_MELDUNGEN = []


def _frage_duplikate():
    """0 = Zieltypen verwenden, 1 = ... und nicht mehr fragen, 2 = abbrechen"""
    dialog = TaskDialog(u"TransferSingle")
    dialog.MainInstruction = t(
        u"Im Ziel gibt es bereits Typen mit gleichem Namen.",
        u"Types with the same name already exist in the target.",
        u"En el destino ya existen tipos con el mismo nombre.")
    dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, t(
        u"Zieltypen verwenden", u"Use destination types",
        u"Usar los tipos del destino"))
    dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, t(
        u"Zieltypen verwenden und nicht mehr fragen",
        u"Use destination types and do not ask again",
        u"Usar los tipos del destino y no volver a preguntar"))
    dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, t(
        u"Diesen Schritt abbrechen", u"Cancel this step",
        u"Cancelar este paso"))
    # "None" ist in Python ein Schlüsselwort
    dialog.CommonButtons = getattr(TaskDialogCommonButtons, "None")
    ergebnis = dialog.Show()
    if ergebnis == TaskDialogResult.CommandLink1:
        return 0
    if ergebnis == TaskDialogResult.CommandLink2:
        return 1
    return 2


def _duplikat_handler():
    if "duplikate" not in _klassen:
        namensraum = "pyMLG.TransferSingle_%s" % uuid.uuid4().hex

        class Duplikate(IDuplicateTypeNamesHandler):
            __namespace__ = namensraum

            def OnDuplicateTypeNamesFound(self, args):
                modus = _zustand["modus"]
                if modus == DUP_ABBRECHEN:
                    _zustand["abgebrochen"] = True
                    return DuplicateTypeAction.Abort
                if modus == DUP_FRAGEN and not _zustand["immer"]:
                    antwort = _frage_duplikate()
                    if antwort == 2:
                        _zustand["abgebrochen"] = True
                        return DuplicateTypeAction.Abort
                    if antwort == 1:
                        _zustand["immer"] = True
                return DuplicateTypeAction.UseDestinationTypes

        _klassen["duplikate"] = Duplikate
    return _klassen["duplikate"]()


def _vorverarbeiter():
    if "fehler" not in _klassen:
        namensraum = "pyMLG.TransferSingleFehler_%s" % uuid.uuid4().hex

        class Vorverarbeiter(IFailuresPreprocessor):
            __namespace__ = namensraum

            def PreprocessFailures(self, accessor):
                aufgeloest = False
                for meldung in list(accessor.GetFailureMessages()):
                    text = _text(lambda: meldung.GetDescriptionText(), u"?")
                    if meldung.GetSeverity() == FailureSeverity.Warning:
                        REVIT_MELDUNGEN.append(t(
                            u"Warnung bestätigt: ", u"Warning accepted: ",
                            u"Aviso aceptado: ") + text)
                        accessor.DeleteWarning(meldung)
                    elif meldung.HasResolutions():
                        REVIT_MELDUNGEN.append(t(
                            u"Fehler von Revit aufgelöst: ",
                            u"Error resolved by Revit: ",
                            u"Error resuelto por Revit: ") + text)
                        accessor.ResolveFailure(meldung)
                        aufgeloest = True
                    else:
                        REVIT_MELDUNGEN.append(t(u"Fehler: ", u"Error: ",
                                                 u"Error: ") + text)
                if aufgeloest:
                    return FailureProcessingResult.ProceedWithCommit
                return FailureProcessingResult.Continue

        _klassen["fehler"] = Vorverarbeiter
    return _klassen["fehler"]()


# ---------------------------------------------------------------------------
# Übertragen
# ---------------------------------------------------------------------------

class Optionen(object):
    def __init__(self, transformation=KEINE, duplikate=DUP_ZIEL,
                 dialoge_bestaetigen=True, plaene_mit_ansichten=True,
                 ansichtselemente=True):
        self.transformation = transformation
        self.duplikate = duplikate
        self.dialoge_bestaetigen = dialoge_bestaetigen
        self.plaene_mit_ansichten = plaene_mit_ansichten
        self.ansichtselemente = ansichtselemente


class Protokoll(object):
    def __init__(self):
        self.zeilen = []
        self.ok = 0
        self.fehler = 0
        # Erst nach dem Speichern prüfbar: [(Text, [neue ElementIds], Quelle)]
        self.ausstehend = []

    def vormerken(self, text, neue_ids, quelle=None):
        self.ausstehend.append((text, list(neue_ids or []), quelle))

    def erfolg(self, text):
        self.ok += 1
        self.zeilen.append(u"  OK      " + text)

    def fehlschlag(self, text):
        self.fehler += 1
        self.zeilen.append(u"  " + t(u"FEHLER", u"ERROR", u"ERROR")
                           + u"  " + text)

    def hinweis(self, text):
        self.zeilen.append(u"  " + text)

    def kopf(self, text):
        self.zeilen.append(u"")
        self.zeilen.append(text)


def bezeichnung(element):
    if isinstance(element, ViewSheet):
        return u"%s - %s" % (element.SheetNumber, element.Name)
    kategorie = _text(lambda: element.Category.Name)
    name = _typname(element) if _text(lambda: element.FamilyName) \
        else _text(lambda: element.Name, u"?")
    return u"%s [%d]%s" % (name, id_wert(element.Id),
                          u" (%s)" % kategorie if kategorie else u"")


def transformation(quelle, ziel, art, protokoll):
    if art == VERKNUEPFUNG:
        if quelle.link is not None and ziel.Equals(quelle.host):
            return quelle.link.GetTotalTransform()
        protokoll.hinweis(t(
            u"Transformation 'Verknüpfung' gilt nur für eine Verknüpfung als "
            u"Quelle und ihr Hostprojekt als Ziel - ohne Transformation "
            u"übertragen.",
            u"'Link' transform only applies to a link as source and its host "
            u"as target - transferred without transform.",
            u"La transformación 'Vínculo' solo vale para un vínculo como "
            u"origen y su proyecto anfitrión como destino: se transfiere sin "
            u"transformación."))
    elif art == GEMEINSAM:
        # GetTotalTransform: gemeinsame Koordinaten -> interne Koordinaten
        quell_t = quelle.doc.ActiveProjectLocation.GetTotalTransform()
        ziel_t = ziel.ActiveProjectLocation.GetTotalTransform()
        return ziel_t.Multiply(quell_t.Inverse)
    return Transform.Identity


def _optionen():
    optionen = CopyPasteOptions()
    try:
        optionen.SetDuplicateTypeNamesHandler(_duplikat_handler())
    except Exception:
        pass
    return optionen


def _einheit(ziel, funktion):
    """Funktion in einer Untertransaktion. Rückgabe (Ergebnis, Fehlertext)."""
    st = SubTransaction(ziel)
    st.Start()
    try:
        ergebnis = funktion()
        st.Commit()
        return ergebnis, None
    except Exception as ausnahme:
        if st.HasStarted() and not st.HasEnded():
            st.RollBack()
        if _zustand["abgebrochen"]:
            _zustand["abgebrochen"] = False
            return None, t(u"doppelte Typnamen - abgebrochen",
                           u"duplicate type names - cancelled",
                           u"nombres de tipo duplicados - cancelado")
        return None, fehlertext(ausnahme)


def _gleichnamig_im_ziel(ziel, element):
    """Gibt es im Ziel schon ein Element derselben Klasse mit diesem Namen?
    Nur für benannte Nicht-Typen (Filter, Materialien, Muster, Vorlagen) -
    Typen regelt der Duplikat-Dialog."""
    klassen = (ParameterFilterElement, SelectionFilterElement, Material,
               LinePatternElement, FillPatternElement)
    name = _text(lambda: element.Name)
    if not name:
        return False
    if isinstance(element, View) and element.IsTemplate:
        return any(v.IsTemplate and v.Name == name for v in _alle(ziel, View))
    for klasse in klassen:
        if isinstance(element, klasse):
            return any(_text(lambda: e.Name) == name
                       for e in _alle(ziel, klasse))
    return False


def _parametername(doc, param_id):
    element = doc.GetElement(param_id)
    if element is None:
        return None
    try:
        return element.GetDefinition().Name
    except Exception:
        return _text(lambda: element.Name) or None


def _filterparameter(doc, filter_element):
    """Namen der Projekt-/gemeinsamen Parameter in den Filterregeln
    (eingebaute Parameter gibt es in jedem Projekt)."""
    try:
        ids = filter_element.GetElementFilterParameters()
    except Exception:
        return set()
    return set(n for n in (_parametername(doc, i) for i in ids
                           if id_wert(i) > 0) if n)


def _fehlende_filterparameter(quelle_doc, element, ziel):
    """Parameter der Filterregeln, die es im Ziel nicht gibt."""
    if not isinstance(element, ParameterFilterElement):
        return []
    try:
        from Autodesk.Revit.DB import ParameterElement
        im_ziel = set(_text(lambda: p.GetDefinition().Name)
                      for p in _alle(ziel, ParameterElement))
    except Exception:
        return []
    return sorted(n for n in _filterparameter(quelle_doc, element)
                  if n not in im_ziel)


def _kopiere_gruppe(quelle_doc, elemente, ziel, transform, protokoll):
    """Element für Element, damit das Protokoll jedes einzeln prüfen kann.
    Ob die Kopie wirklich bleibt, zeigt sich erst nach dem Speichern
    (pruefe_nach_speichern)."""
    for element in elemente:
        name = bezeichnung(element)
        if _gleichnamig_im_ziel(ziel, element):
            protokoll.hinweis(t(
                u"Hinweis  %s: im Ziel gibt es bereits ein Element mit diesem "
                u"Namen",
                u"Note     %s: an element with this name already exists in "
                u"the target",
                u"Nota     %s: en el destino ya existe un elemento con este "
                u"nombre") % name)
        fehlend = _fehlende_filterparameter(quelle_doc, element, ziel)
        if fehlend:
            protokoll.hinweis(t(
                u"Hinweis  %s: Parameter fehlen im Ziel: %s - Regeln damit "
                u"können nicht übernommen werden",
                u"Note     %s: parameters missing in the target: %s - rules "
                u"using them cannot be transferred",
                u"Nota     %s: faltan parámetros en el destino: %s - las "
                u"reglas que los usan no se pueden transferir")
                % (name, u", ".join(fehlend)))
        neue, fehler = _einheit(ziel, lambda: ElementTransformUtils
                                .CopyElements(quelle_doc,
                                              net_ids([element.Id]), ziel,
                                              transform, _optionen()))
        if fehler is None:
            protokoll.vormerken(name, neue, element)
        else:
            protokoll.fehlschlag(u"%s: %s" % (name, fehler))


def pruefe_nach_speichern(quelle_doc, ziel, status, protokoll):
    """Vorgemerkte Kopien prüfen: noch vorhanden? Filterregeln vollständig?"""
    committed = u"%s" % status == u"Committed"
    for text, neue_ids, quelle in protokoll.ausstehend:
        if not committed:
            protokoll.fehlschlag(t(
                u"%s: Revit hat die Übertragung beim Speichern zurückgenommen",
                u"%s: Revit rolled back the transfer when saving",
                u"%s: Revit deshizo la transferencia al guardar") % text)
            continue
        vorhanden = [ziel.GetElement(i) for i in neue_ids]
        if not neue_ids or any(e is None or not e.IsValidObject
                               for e in vorhanden):
            protokoll.fehlschlag(t(
                u"%s: nach dem Speichern im Ziel nicht mehr vorhanden",
                u"%s: no longer present in the target after saving",
                u"%s: ya no existe en el destino después de guardar") % text)
            continue
        protokoll.erfolg(text)
        if isinstance(quelle, ParameterFilterElement):
            fehlend = sorted(_filterparameter(quelle_doc, quelle)
                             - _filterparameter(ziel, vorhanden[0]))
            if fehlend:
                protokoll.hinweis(t(
                    u"          Regeln fehlen im Ziel (Parameter: %s)",
                    u"          rules missing in the target (parameters: %s)",
                    u"          faltan reglas en el destino (parámetros: %s)")
                    % u", ".join(fehlend))
    protokoll.ausstehend = []


def _eigene_elemente(doc, ansicht):
    """Ids der Elemente, die der Ansicht gehören (ohne Viewports und
    platzierte Bauteillisten)."""
    ergebnis = []
    for element in FilteredElementCollector(doc, ansicht.Id) \
            .WhereElementIsNotElementType():
        try:
            if id_wert(element.OwnerViewId) != id_wert(ansicht.Id) \
                    or element.Category is None \
                    or isinstance(element, (Viewport, ScheduleSheetInstance)):
                continue
            ergebnis.append(element.Id)
        except Exception:
            continue
    return ergebnis


def _kopiere_ansicht(quelle_doc, ansicht, ziel, optionen, kopien):
    """Ansicht übertragen (einmal je Lauf). Rückgabe: neue Ansicht."""
    schluessel = id_wert(ansicht.Id)
    if schluessel in kopien:
        return kopien[schluessel]
    neue_ids = ElementTransformUtils.CopyElements(
        quelle_doc, net_ids([ansicht.Id]), ziel, Transform.Identity,
        _optionen())
    neue = ziel.GetElement(list(neue_ids)[0])
    if optionen.ansichtselemente and not isinstance(ansicht, ViewSchedule):
        quell_ids = _eigene_elemente(quelle_doc, ansicht)
        # Kopiert Revit die Ansichtselemente bereits mit, nicht verdoppeln
        if quell_ids and not _eigene_elemente(ziel, neue):
            ElementTransformUtils.CopyElements(
                ansicht, net_ids(quell_ids), neue, Transform.Identity,
                _optionen())
    kopien[schluessel] = neue
    return neue


UEBERTRAGBARE_ANSICHTEN = (ViewType.DraftingView, ViewType.Legend)


def _kopiere_plan(quelle_doc, plan, ziel, optionen, kopien, nummern,
                  protokoll):
    def ausfuehren():
        neu = ViewSheet.Create(ziel, ElementId.InvalidElementId)
        neu.SheetNumber = lg.freie_nummer(plan.SheetNumber, nummern)
        neu.Name = plan.Name
        eigene = _eigene_elemente(quelle_doc, plan)
        if eigene:      # Plankopf, Texte, Linien ...
            ElementTransformUtils.CopyElements(
                plan, net_ids(eigene), neu, Transform.Identity, _optionen())
        hinweise = []
        if optionen.plaene_mit_ansichten:
            for vp_id in plan.GetAllViewports():
                viewport = quelle_doc.GetElement(vp_id)
                ansicht = quelle_doc.GetElement(viewport.ViewId)
                if ansicht.ViewType not in UEBERTRAGBARE_ANSICHTEN:
                    hinweise.append(t(
                        u"Modellansicht '%s' nicht übertragbar",
                        u"model view '%s' cannot be transferred",
                        u"la vista de modelo '%s' no se puede transferir")
                        % ansicht.Name)
                    continue
                neue = _kopiere_ansicht(quelle_doc, ansicht, ziel, optionen,
                                        kopien)
                if Viewport.CanAddViewToSheet(ziel, neu.Id, neue.Id):
                    Viewport.Create(ziel, neu.Id, neue.Id,
                                    viewport.GetBoxCenter())
                else:
                    hinweise.append(t(
                        u"'%s' ist bereits auf einem anderen Plan",
                        u"'%s' is already on another sheet",
                        u"'%s' ya está en otro plano") % neue.Name)
            for instanz in FilteredElementCollector(quelle_doc, plan.Id) \
                    .OfClass(ScheduleSheetInstance):
                if instanz.IsTitleblockRevisionSchedule:
                    continue
                liste = quelle_doc.GetElement(instanz.ScheduleId)
                neue = _kopiere_ansicht(quelle_doc, liste, ziel, optionen,
                                        kopien)
                ScheduleSheetInstance.Create(ziel, neu.Id, neue.Id,
                                             instanz.Point)
        return neu, hinweise

    ergebnis, fehler = _einheit(ziel, ausfuehren)
    if fehler is not None:
        protokoll.fehlschlag(u"%s: %s" % (bezeichnung(plan), fehler))
        return
    neu, hinweise = ergebnis
    for hinweis in hinweise:
        protokoll.hinweis(u"%s: %s" % (bezeichnung(plan), hinweis))
    protokoll.vormerken(u"%s -> %s" % (bezeichnung(plan), neu.SheetNumber),
                        [neu.Id], plan)


def _ist_modellelement(element):
    """Platziertes Modellelement (bekommt die Transformation). Typen,
    Materialien, Filter usw. haben keine Lage."""
    try:
        return (element.Category is not None and not element.ViewSpecific
                and element.Category.CategoryType == CategoryType.Model
                and element.Location is not None
                and id_wert(element.GetTypeId()) != -1)
    except Exception:
        return False


def uebertrage(quelle, ziel_docs, id_werte, optionen):
    """Markierte Elemente in alle Ziele übertragen. Rückgabe: Protokoll."""
    protokoll = Protokoll()
    doc = quelle.doc
    elemente = [e for e in (doc.GetElement(ElementId(w)) for w in id_werte)
                if e is not None]
    plaene, ansichten, modell, rest = [], [], [], []
    for element in elemente:
        if isinstance(element, ViewSheet):
            plaene.append(element)
        elif isinstance(element, View) and not element.IsTemplate:
            ansichten.append(element)
        elif _ist_modellelement(element):
            modell.append(element)
        else:
            rest.append(element)

    for ziel in ziel_docs:
        _zustand.update(modus=optionen.duplikate, immer=False,
                        abgebrochen=False)
        del REVIT_MELDUNGEN[:]
        protokoll.kopf(u"%s  ->  %s" % (quelle.titel, ziel.Title))
        transaktion = Transaction(ziel, t(u"pyMLG Elemente übertragen",
                                          u"pyMLG Transfer elements",
                                          u"pyMLG Transferir elementos"))
        transaktion.Start()
        if optionen.dialoge_bestaetigen:
            try:
                fehleroptionen = transaktion.GetFailureHandlingOptions()
                fehleroptionen.SetFailuresPreprocessor(_vorverarbeiter())
                fehleroptionen.SetClearAfterRollback(True)
                transaktion.SetFailureHandlingOptions(fehleroptionen)
            except Exception:
                pass
        try:
            _kopiere_gruppe(doc, rest, ziel, Transform.Identity, protokoll)
            if modell:
                _kopiere_gruppe(doc, modell, ziel, transformation(
                    quelle, ziel, optionen.transformation, protokoll),
                    protokoll)
            kopien = {}
            for ansicht in ansichten:
                neu, fehler = _einheit(ziel, lambda: _kopiere_ansicht(
                    doc, ansicht, ziel, optionen, kopien))
                if fehler is None:
                    protokoll.vormerken(bezeichnung(ansicht), [neu.Id],
                                        ansicht)
                else:
                    protokoll.fehlschlag(u"%s: %s" % (bezeichnung(ansicht),
                                                      fehler))
            if plaene:
                nummern = set(p.SheetNumber.casefold() for p in
                              FilteredElementCollector(ziel).OfClass(ViewSheet))
                for plan in plaene:
                    _kopiere_plan(doc, plan, ziel, optionen, kopien, nummern,
                                  protokoll)
            status = transaktion.Commit()
        except Exception as ausnahme:
            if transaktion.HasStarted() and not transaktion.HasEnded():
                transaktion.RollBack()
            protokoll.ausstehend = []
            protokoll.fehlschlag(t(u"Übertragung abgebrochen: %s",
                                   u"Transfer cancelled: %s",
                                   u"Transferencia cancelada: %s")
                                 % fehlertext(ausnahme))
            continue
        pruefe_nach_speichern(doc, ziel, status, protokoll)
        for meldung in REVIT_MELDUNGEN:
            protokoll.hinweis(u"Revit: " + meldung)
    return protokoll


# ---------------------------------------------------------------------------
# Markierte in der Quelle verwalten
# ---------------------------------------------------------------------------

def loesche(doc, id_werte):
    protokoll = Protokoll()
    transaktion = Transaction(doc, t(u"pyMLG Elemente löschen",
                                     u"pyMLG Delete elements",
                                     u"pyMLG Eliminar elementos"))
    transaktion.Start()
    try:
        for wert in id_werte:
            element = doc.GetElement(ElementId(wert))
            if element is None:
                continue        # schon mit einem anderen Element gelöscht
            name = bezeichnung(element)
            _e, fehler = _einheit(doc, lambda: doc.Delete(element.Id))
            if fehler is None:
                protokoll.erfolg(name)
            else:
                protokoll.fehlschlag(u"%s: %s" % (name, fehler))
        transaktion.Commit()
    except Exception:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        raise
    return protokoll


def benenne_um(doc, id_werte, aktion, a=u"", b=u""):
    protokoll = Protokoll()
    lg.neuer_name(u"", aktion, a, b)    # ungültige Eingabe -> ValueError vorab
    transaktion = Transaction(doc, t(u"pyMLG Elemente umbenennen",
                                     u"pyMLG Rename elements",
                                     u"pyMLG Renombrar elementos"))
    transaktion.Start()
    try:
        for wert in id_werte:
            element = doc.GetElement(ElementId(wert))
            if element is None:
                continue
            alt = _text(lambda: element.Name)
            neu = lg.neuer_name(alt, aktion, a, b)
            if neu == alt:
                continue

            def setzen():
                element.Name = neu
            _e, fehler = _einheit(doc, setzen)
            if fehler is None:
                protokoll.erfolg(u"%s -> %s" % (alt, neu))
            else:
                protokoll.fehlschlag(u"%s: %s" % (alt, fehler))
        transaktion.Commit()
    except Exception:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        raise
    return protokoll
