# -*- coding: utf-8 -*-
"""Revit-Zugriffe von FilterMore.

Erweitern der Auswahl: "gleiche Kategorie/Familie/Typ/Workset" und "gehostete
Elemente" brauchen einen Durchlauf über alle Kandidaten (Modell oder
Ansicht) - alle gewählten Kriterien werden in EINEM Durchlauf geprüft.
Host, verschachtelte, verbundene, übergeordnete und abhängige Elemente
ergeben sich direkt aus den markierten Elementen.
"""

import clr

clr.AddReference("System")
from System.Collections.Generic import List  # noqa: E402

from Autodesk.Revit.DB import (  # noqa: E402
    CategoryType,
    ElementId,
    ElementType,
    FilteredElementCollector,
    JoinGeometryUtils,
    ParameterFilterElement,
    View,
)

from filter_more import baum as bm  # noqa: E402

AUSWAHL = 0
ANSICHT = 1
MODELL = 2

# Schlüssel der Erweiterungen (Kontrollkästchen im Fenster)
GLEICHE_KATEGORIE = "kategorie"
GLEICHE_FAMILIE = "familie"
GLEICHER_TYP = "typ"
GLEICHES_WORKSET = "workset"
HOST = "host"
GEHOSTETE = "gehostete"
VERSCHACHTELTE = "verschachtelte"
VERBUNDENE = "verbundene"
UEBERGEORDNETE = "uebergeordnete"
ABHAENGIGE = "abhaengige"

MIT_DURCHLAUF = (GLEICHE_KATEGORIE, GLEICHE_FAMILIE, GLEICHER_TYP,
                 GLEICHES_WORKSET, GEHOSTETE)

OHNE_KATEGORIE = u"(ohne Kategorie)"
OHNE_FAMILIE = u"(ohne Familie)"
OHNE_TYP = u"(ohne Typ)"


def id_wert(element_id):
    """Zahlenwert einer ElementId (Revit 2024+: .Value)."""
    if element_id is None:
        return -1
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def net_liste(typ, werte):
    """Python-Iterable -> List[typ] (pythonnet nimmt keine Python-Liste im
    Konstruktor an)."""
    liste = List[typ]()
    for wert in werte:
        liste.Add(wert)
    return liste


def _text(funktion, ersatz):
    try:
        wert = funktion()
        return u"%s" % wert if wert else ersatz
    except Exception:
        return ersatz


# ---------------------------------------------------------------------------
# Sammeln
# ---------------------------------------------------------------------------

def gueltig(element, nur_3d=False, streng=True):
    """Gehört das Element in den Baum?

    streng  (Ansicht/Modell/Erweiterung) nur Modell- und Beschriftungs-
            kategorien - sonst landen interne Elemente wie Skizzen oder
            Sonneneinstellungen in der Liste. Die Revit-Auswahl wird
            nicht streng geprüft: was der Nutzer gewählt hat, bleibt.
    nur_3d  nur modellierte Elemente (Modellkategorie, nicht
            ansichtsspezifisch)
    """
    if element is None or isinstance(element, (ElementType, View)):
        return False
    kategorie = element.Category
    if kategorie is None:
        return False
    try:
        art = kategorie.CategoryType
        if nur_3d:
            return art == CategoryType.Model and not element.ViewSpecific
        if streng:
            return art in (CategoryType.Model, CategoryType.Annotation)
    except Exception:
        return False
    return True


def sammle(doc, uidoc, modus, nur_3d=False):
    if modus == AUSWAHL:
        return [e for e in (doc.GetElement(i)
                            for i in uidoc.Selection.GetElementIds())
                if gueltig(e, nur_3d, streng=False)]
    return [e for e in _kandidaten(doc, uidoc, modus == ANSICHT)
            if gueltig(e, nur_3d)]


def _kandidaten(doc, uidoc, nur_ansicht):
    sammler = (FilteredElementCollector(doc, uidoc.ActiveView.Id)
               if nur_ansicht else FilteredElementCollector(doc))
    return sammler.WhereElementIsNotElementType()


class Namen(object):
    """Zwischenspeicher für Kategorie- und Typangaben.

    Tausende Elemente teilen sich wenige Typen - ohne Speicher würde für
    jedes Element der Typ erneut aus dem Dokument geholt."""

    def __init__(self, doc):
        self.doc = doc
        self._typen = {}         # Typ-Id -> (Familie, Typname)
        self._kategorien = {}    # Kategorie-Id -> Name

    def kategorie(self, element):
        """(Kategorie-Id oder None, Name)."""
        kategorie = element.Category
        if kategorie is None:
            return None, OHNE_KATEGORIE
        wert = id_wert(kategorie.Id)
        if wert not in self._kategorien:
            self._kategorien[wert] = _text(lambda: kategorie.Name,
                                           OHNE_KATEGORIE)
        return wert, self._kategorien[wert]

    def typ(self, element):
        """(Typ-Id, Familie, Typname); Typ-Id -1 ohne Typ."""
        try:
            typ_id = element.GetTypeId()
            wert = id_wert(typ_id)
        except Exception:
            return -1, OHNE_FAMILIE, OHNE_TYP
        if wert == -1:
            return -1, OHNE_FAMILIE, OHNE_TYP
        if wert not in self._typen:
            typ = self.doc.GetElement(typ_id)
            if typ is None:
                self._typen[wert] = (OHNE_FAMILIE, OHNE_TYP)
            else:
                self._typen[wert] = (
                    _text(lambda: typ.FamilyName, OHNE_FAMILIE),
                    _text(lambda: typ.Name, OHNE_TYP))
        familie, typname = self._typen[wert]
        return wert, familie, typname

    def familienschluessel(self, element):
        kategorie_id, _name = self.kategorie(element)
        typ_id, familie, _typname = self.typ(element)
        if kategorie_id is None or typ_id == -1:
            return None
        return (kategorie_id, familie)


def datensatz(namen, element):
    """(Kategorie, Familie, Typ, Bezeichnung, Id) für baum.baue()."""
    _kid, kategorie = namen.kategorie(element)
    _tid, familie, typname = namen.typ(element)
    name = _text(lambda: element.Name, typname)
    element_id = id_wert(element.Id)
    return (kategorie, familie, typname, u"%s  [%d]" % (name, element_id),
            element_id)


def baue_baum(doc, elemente):
    namen = Namen(doc)
    return bm.baue([datensatz(namen, e) for e in elemente])


# ---------------------------------------------------------------------------
# Auswahl erweitern
# ---------------------------------------------------------------------------

def _ids(sammlung):
    try:
        return [id_wert(i) for i in sammlung]
    except Exception:
        return []


def erweitere(doc, uidoc, quellen, was, nur_ansicht=False, nur_3d=False,
              ohne_gruppe=False, ohne_baugruppe=False):
    """Elemente, die zu den markierten Elementen (quellen) passen.

    was          Menge der Erweiterungs-Schlüssel
    nur_ansicht  nur Elemente, die in der aktuellen Ansicht sichtbar sind
    Rückgabe:    [Element] ohne die Quellen selbst
    """
    quell_ids = set(id_wert(e.Id) for e in quellen)
    treffer = set()
    namen = Namen(doc)
    sichtbar = None          # Ids der Ansicht, sobald einmal durchlaufen

    # Direkte Beziehungen
    for e in quellen:
        if HOST in was:
            host = getattr(e, "Host", None)
            if host is not None:
                treffer.add(id_wert(host.Id))
        if VERSCHACHTELTE in was and hasattr(e, "GetSubComponentIds"):
            treffer.update(_ids(e.GetSubComponentIds()))
        if UEBERGEORDNETE in was:
            ueber = getattr(e, "SuperComponent", None)
            if ueber is not None:
                treffer.add(id_wert(ueber.Id))
        if VERBUNDENE in was:
            try:
                treffer.update(_ids(
                    JoinGeometryUtils.GetJoinedElements(doc, e)))
            except Exception:
                pass
        if ABHAENGIGE in was:
            try:
                treffer.update(_ids(e.GetDependentElements(None)))
            except Exception:
                pass

    # Ein Durchlauf für alle "gleichen" und die gehosteten Elemente
    if any(k in was for k in MIT_DURCHLAUF):
        kategorien = set(id_wert(e.Category.Id) for e in quellen
                         if e.Category is not None)
        familien = set(namen.familienschluessel(e) for e in quellen)
        familien.discard(None)
        typen = set(id_wert(e.GetTypeId()) for e in quellen)
        typen.discard(-1)
        worksets = set()
        if GLEICHES_WORKSET in was and doc.IsWorkshared:
            worksets = set(e.WorksetId.IntegerValue for e in quellen)
        if nur_ansicht:
            sichtbar = set()
        for c in _kandidaten(doc, uidoc, nur_ansicht):
            c_id = id_wert(c.Id)
            if sichtbar is not None:
                sichtbar.add(c_id)
            if not gueltig(c, nur_3d):
                continue
            if c_id in treffer or c_id in quell_ids:
                continue
            try:
                if (GLEICHE_KATEGORIE in was
                        and id_wert(c.Category.Id) in kategorien) \
                        or (GLEICHER_TYP in was
                            and id_wert(c.GetTypeId()) in typen) \
                        or (worksets
                            and c.WorksetId.IntegerValue in worksets) \
                        or (GLEICHE_FAMILIE in was
                            and namen.familienschluessel(c) in familien):
                    treffer.add(c_id)
                    continue
                if GEHOSTETE in was:
                    host = getattr(c, "Host", None)
                    if host is not None and id_wert(host.Id) in quell_ids:
                        treffer.add(c_id)
            except Exception:
                continue

    treffer -= quell_ids
    treffer.discard(-1)
    if nur_ansicht and sichtbar is None:
        sichtbar = set(id_wert(c.Id)
                       for c in _kandidaten(doc, uidoc, True))

    ergebnis = []
    for wert in treffer:
        if sichtbar is not None and wert not in sichtbar:
            continue
        element = doc.GetElement(ElementId(wert))
        if not gueltig(element, nur_3d):
            continue
        if ohne_gruppe and id_wert(element.GroupId) != -1:
            continue
        if ohne_baugruppe and id_wert(element.AssemblyInstanceId) != -1:
            continue
        ergebnis.append(element)
    return ergebnis


# ---------------------------------------------------------------------------
# Anzeigefilter
# ---------------------------------------------------------------------------

def anzeigefilter(doc):
    """[(Name, ParameterFilterElement)] sortiert."""
    filter_ = [(f.Name, f) for f in
               FilteredElementCollector(doc).OfClass(ParameterFilterElement)]
    return sorted(filter_, key=lambda e: bm.natuerlich(e[0]))


def filter_treffer(filter_element, elemente):
    """Ids der Elemente, die der Anzeigefilter erfasst: Kategorie des
    Filters UND Regeln erfüllt - wie in einer Ansicht."""
    kategorien = set(id_wert(k) for k in filter_element.GetCategories())
    regeln = filter_element.GetElementFilter()
    ergebnis = set()
    for e in elemente:
        try:
            if e.Category is None or id_wert(e.Category.Id) not in kategorien:
                continue
            if regeln is None or regeln.PassesFilter(e):
                ergebnis.add(id_wert(e.Id))
        except Exception:
            continue
    return ergebnis


def waehle_aus(uidoc, elemente):
    uidoc.Selection.SetElementIds(net_liste(ElementId,
                                            [e.Id for e in elemente]))
