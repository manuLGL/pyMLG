# -*- coding: utf-8 -*-
"""Revit-Zugriffe von LinkedIds.

Ein mit Tab angetipptes Element einer Verknüpfung taucht in
Selection.GetElementIds() nur als RevitLinkInstance auf - die eigene ID des
Elements steht ausschliesslich in der Referenz:

    Reference.ElementId        die RevitLinkInstance im Wirtsmodell
    Reference.LinkedElementId  das Element im verknüpften Dokument

Selection.GetReferences() gibt es ab Revit 2023. Fehlt es, bleibt das Wählen
über PickObjects(ObjectType.LinkedElement) - der übliche Weg, wenn beim Start
nichts ausgewählt ist.

Verschachtelte Verknüpfungen: löst LinkedElementId wieder eine
RevitLinkInstance auf, führt Reference.CreateReferenceInLink() eine Ebene
tiefer. MAX_TIEFE begrenzt den Abstieg, damit ein Ring aus Verknüpfungen das
Werkzeug nicht aufhängt.
"""

from Autodesk.Revit.DB import RevitLinkInstance  # noqa: E402
from Autodesk.Revit.Exceptions import OperationCanceledException  # noqa: E402
from Autodesk.Revit.UI.Selection import ObjectType  # noqa: E402

from linked_ids import logik as lg  # noqa: E402
from mlg_sprache import t  # noqa: E402

MAX_TIEFE = 8


def id_wert(element_id):
    """Zahlenwert einer ElementId (Revit 2024+: .Value)."""
    if element_id is None:
        return -1
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def _name(element):
    """Element.Name - bei manchen Elementen wirft der Zugriff."""
    if element is None:
        return u""
    try:
        return element.Name or u""
    except Exception:
        return u""


def _kategorie(element):
    try:
        kategorie = element.Category
    except Exception:
        return u""
    return kategorie.Name if kategorie is not None else u""


def _typ_text(dokument, element):
    """"Familie: Typ" des Elements, soweit vorhanden."""
    if dokument is None or element is None:
        return u""
    try:
        typ_id = element.GetTypeId()
    except Exception:
        return u""
    if id_wert(typ_id) <= 0:
        return u""
    typ = dokument.GetElement(typ_id)
    if typ is None:
        return u""
    familie = u""
    try:
        familie = typ.FamilyName or u""
    except Exception:
        pass
    name = _name(typ)
    if familie and name:
        return u"%s: %s" % (familie, name)
    return familie or name


def _unique_id(element):
    try:
        return element.UniqueId or u""
    except Exception:
        return u""


def _verknuepfungsname(instanz):
    """Dateiname einer Verknüpfung - auch wenn sie nicht geladen ist.

    Element.Document ist das Dokument, in dem die Instanz liegt - bei
    verschachtelten Verknüpfungen nicht das geöffnete Modell.
    """
    if instanz is None:
        return u""
    try:
        typ = instanz.Document.GetElement(instanz.GetTypeId())
    except Exception:
        typ = None
    return _name(typ) or _name(instanz)


def _aufloesen(doc, referenz):
    """Folgt einer Referenz bis zum Element - auch durch verschachtelte
    Verknüpfungen.

    Rückgabe: (dokument, element, id_wert, kette). dokument und element sind
    None, wenn die Verknüpfung nicht geladen ist; die ID ist dann trotzdem
    bekannt. kette enthält die durchlaufenen RevitLinkInstance-Objekte.
    """
    aktuell = doc
    ref = referenz
    kette = []
    for _ in range(MAX_TIEFE):
        verknuepft = getattr(ref, "LinkedElementId", None)
        if id_wert(verknuepft) <= 0:
            element = aktuell.GetElement(ref.ElementId) if aktuell else None
            return aktuell, element, id_wert(ref.ElementId), kette

        instanz = aktuell.GetElement(ref.ElementId) if aktuell else None
        kette.append(instanz)
        verknuepftes_doc = None
        if instanz is not None:
            try:
                verknuepftes_doc = instanz.GetLinkDocument()
            except Exception:
                verknuepftes_doc = None
        if verknuepftes_doc is None:
            return None, None, id_wert(verknuepft), kette

        element = verknuepftes_doc.GetElement(verknuepft)
        if not isinstance(element, RevitLinkInstance):
            return verknuepftes_doc, element, id_wert(verknuepft), kette

        # Das Element ist selbst eine Verknüpfung: eine Ebene tiefer weiter
        try:
            ref = ref.CreateReferenceInLink()
        except Exception:
            return verknuepftes_doc, element, id_wert(verknuepft), kette
        aktuell = verknuepftes_doc
    return aktuell, None, -1, kette


def eintrag_aus_referenz(doc, referenz):
    """Referenz -> Eintrag, oder None, wenn nichts Brauchbares dahinter steht."""
    dokument, element, wert, kette = _aufloesen(doc, referenz)
    if wert <= 0:
        return None
    namen = [_name(instanz) for instanz in kette if instanz is not None]
    verknuepfung = u" > ".join(name for name in namen if name)

    if dokument is None:
        # Verknüpfung nicht geladen - nur die ID steht fest
        letzte = kette[-1] if kette else None
        name = _verknuepfungsname(letzte) or t(
            u"Nicht geladene Verknüpfung", u"Link not loaded",
            u"Vínculo no cargado")
        return lg.Eintrag(name, wert, verknuepfung=verknuepfung,
                          ist_verknuepfung=True, geladen=False)

    return lg.Eintrag(dokument.Title, wert,
                      kategorie=_kategorie(element),
                      typ=_typ_text(dokument, element),
                      name=_name(element),
                      unique_id=_unique_id(element),
                      verknuepfung=verknuepfung,
                      ist_verknuepfung=bool(kette))


def _eintrag_aus_element(doc, element):
    if element is None:
        return None
    return lg.Eintrag(doc.Title, id_wert(element.Id),
                      kategorie=_kategorie(element),
                      typ=_typ_text(doc, element),
                      name=_name(element),
                      unique_id=_unique_id(element))


def sammle_aus_auswahl(uidoc):
    """Einträge zur laufenden Auswahl - Verknüpfungen wie eigene Elemente."""
    doc = uidoc.Document
    try:
        referenzen = list(uidoc.Selection.GetReferences())
    except Exception:
        referenzen = []          # Revit vor 2023

    eintraege = []
    if referenzen:
        lg.ergaenze(eintraege, [eintrag for eintrag in
                                (eintrag_aus_referenz(doc, referenz)
                                 for referenz in referenzen)
                                if eintrag is not None])
        return eintraege

    # Ohne Referenzen steht ein angetipptes Element der Verknüpfung nur als
    # RevitLinkInstance in der Auswahl - die hilft nicht weiter und bleibt
    # draussen; starte() fragt dann nach den Elementen in der Verknüpfung.
    elemente = [doc.GetElement(id_) for id_ in uidoc.Selection.GetElementIds()]
    lg.ergaenze(eintraege, [_eintrag_aus_element(doc, element)
                            for element in elemente
                            if element is not None
                            and not isinstance(element, RevitLinkInstance)])
    return eintraege


def waehle(uidoc):
    """Elemente in Verknüpfungen picken.

    Rückgabe: Liste der Einträge, oder None bei Abbruch (Esc).
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
    eintraege = []
    lg.ergaenze(eintraege, [eintrag for eintrag in
                            (eintrag_aus_referenz(doc, referenz)
                             for referenz in referenzen)
                            if eintrag is not None])
    return eintraege
