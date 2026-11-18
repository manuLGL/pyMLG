# -*- coding: utf-8 -*-
"""Revit-Zugriffe von EmptyTags.

Beschriftungen sind IndependentTag (Bauteilbeschriftungen, auch an
Elementen aus Verknüpfungen) und SpatialElementTag (Raum-, Flächen- und
MEP-Raumbeschriftungen). Beide haben einen Typ (FamilySymbol), dessen
Family die Beschriftungsfamilie ist.

Der Auswahlfilter für PickObjects ist eine .NET-Klasse; pythonnet erlaubt je
Sitzung nur eine Klasse pro Namensraum. Deshalb bekommt sie einmal einen
eindeutigen Namensraum und wird dann wiederverwendet.
"""

import uuid

from Autodesk.Revit.DB import ElementId, FilteredElementCollector, \
    IndependentTag, SpatialElementTag  # noqa: E402
from Autodesk.Revit.Exceptions import OperationCanceledException  # noqa: E402
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType  # noqa: E402
from System.Collections.Generic import List  # noqa: E402

from leere_beschriftung import logik as lg  # noqa: E402
from mlg_sprache import t  # noqa: E402

_KLASSEN = []


def id_wert(element_id):
    """Zahlenwert einer ElementId (Revit 2024+: .Value)."""
    if element_id is None:
        return -1
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def ist_beschriftung(element):
    return isinstance(element, (IndependentTag, SpatialElementTag))


def tag_text(element):
    """Angezeigter Text einer Beschriftung, None wenn nicht lesbar."""
    try:
        text = element.TagText
    except Exception:
        return None
    return u"%s" % text if text is not None else None


def schluessel(doc, element, ganze_familie):
    """Gruppe, innerhalb der verglichen wird: Typ oder ganze Familie."""
    typ_id = element.GetTypeId()
    if ganze_familie:
        typ = doc.GetElement(typ_id)
        familie = getattr(typ, "Family", None) if typ is not None else None
        if familie is not None:
            return (u"F", id_wert(familie.Id))
    return (u"T", id_wert(typ_id))


def name_von(doc, element, ganze_familie):
    """"Familie: Typ" (bzw. nur Familie) für Meldungen."""
    typ = doc.GetElement(element.GetTypeId())
    if typ is None:
        return u""
    familie = u""
    try:
        familie = typ.FamilyName or u""
    except Exception:
        pass
    if ganze_familie:
        return familie
    try:
        typname = typ.Name or u""
    except Exception:
        typname = u""
    return u"%s: %s" % (familie, typname) if familie else typname


def beschriftungen_der_ansicht(doc, ansicht):
    """Alle Beschriftungen, die in der Ansicht liegen."""
    ergebnis = []
    for klasse in (IndependentTag, SpatialElementTag):
        ergebnis.extend(FilteredElementCollector(doc, ansicht.Id)
                        .OfClass(klasse).ToElements())
    return ergebnis


def beschriftungen_der_auswahl(uidoc):
    doc = uidoc.Document
    elemente = [doc.GetElement(id_) for id_ in uidoc.Selection.GetElementIds()]
    return [element for element in elemente
            if element is not None and ist_beschriftung(element)]


def _filterklasse():
    if not _KLASSEN:
        class NurBeschriftungen(ISelectionFilter):
            __namespace__ = "pyMLG.EmptyTags_%s" % uuid.uuid4().hex

            def AllowElement(self, element):
                try:
                    return ist_beschriftung(element)
                except Exception:
                    return False

            def AllowReference(self, referenz, punkt_):
                return False

        _KLASSEN.append(NurBeschriftungen)
    return _KLASSEN[0]


def waehle_muster(uidoc):
    """Muster-Beschriftungen picken. None bei Abbruch (Esc)."""
    hinweis = t(u"Beschriftung(en) mit \"?\" als Muster wählen, dann "
                u"Fertigstellen",
                u"Select tag(s) showing \"?\" as a sample, then Finish",
                u"Seleccione etiqueta(s) con \"?\" como muestra y pulse "
                u"Finalizar")
    try:
        referenzen = uidoc.Selection.PickObjects(
            ObjectType.Element, _filterklasse()(), hinweis)
    except OperationCanceledException:
        return None
    doc = uidoc.Document
    return [doc.GetElement(referenz.ElementId) for referenz in referenzen]


def passende(doc, ansicht, beispiele, ganze_familie):
    """ElementIds aller Beschriftungen der Ansicht, die wie ein Beispiel
    aussehen."""
    vorlagen = lg.muster([(schluessel(doc, element, ganze_familie),
                           tag_text(element)) for element in beispiele])
    nach_wert = {}
    kandidaten = []
    for element in beschriftungen_der_ansicht(doc, ansicht):
        wert = id_wert(element.Id)
        nach_wert[wert] = element.Id
        kandidaten.append((wert, schluessel(doc, element, ganze_familie),
                           tag_text(element)))
    return [nach_wert[wert] for wert in lg.treffer(kandidaten, vorlagen)]


def setze_auswahl(uidoc, element_ids):
    ids = List[ElementId]()
    for element_id in element_ids:
        ids.Add(element_id)
    uidoc.Selection.SetElementIds(ids)
