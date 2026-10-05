# -*- coding: utf-8 -*-
"""Revit-Seite von LinkedViews: Verknüpfungen, ihre Ansichten und die
Anzeigeeinstellungen für RVT-Verknüpfungen (ab Revit 2024 in der API:
View.Get/Set/RemoveLinkOverrides, RevitLinkGraphicsSettings).

Eine Einstellung gilt für einen RevitLinkType (alle Exemplare) oder für ein
einzelnes RevitLinkInstance. Hat ein Exemplar keine eigene Einstellung,
liefert GetLinkOverrides dafür None - dann gilt die des Typs.
"""

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    LinkVisibility,
    RevitLinkGraphicsSettings,
    RevitLinkInstance,
    RevitLinkType,
    View,
)

from filter_manager.aufloesung import id_wert
from link_ansichten import logik as lg
from mlg_sprache import t

# Zustand einer Verknüpfung in einer Ansicht
HOST = "host"          # Nach Basisbauteilansicht
LINK = "link"          # Nach verknüpfter Ansicht
EIGEN = "eigen"        # Benutzerdefiniert


class LinkAnsicht(object):
    """Eine Ansicht im verknüpften Modell."""

    def __init__(self, verknuepfung, ansicht, typen):
        self.verknuepfung = verknuepfung
        self.id = ansicht.Id
        self.wert = id_wert(ansicht.Id)
        self.name = ansicht.Name
        self.typ = u"%s" % ansicht.ViewType
        self.art = typen.get(self.typ, self.typ)
        self.gruppe = lg.gruppe(self.typ)
        self.ebene = u""
        self.vorlage = u""
        try:
            ebene = ansicht.GenLevel
            if ebene is not None:
                self.ebene = ebene.Name
        except Exception:
            pass
        try:
            vorlage = ansicht.Document.GetElement(ansicht.ViewTemplateId)
            if vorlage is not None:
                self.vorlage = vorlage.Name
        except Exception:
            pass
        self.suchtext = u" ".join([self.art, self.name, self.ebene,
                                   self.vorlage, verknuepfung.name]).lower()

    @property
    def text(self):
        return u"%s: %s" % (self.art, self.name)


class Verknuepfung(object):
    """Ein Link-Typ oder ein einzelnes Exemplar - so wie die Zeilen im
    Reiter "Revit-Verknüpfungen" von Sichtbarkeit/Grafiken."""

    def __init__(self, element, typ, link_doc, exemplar=False):
        self.element = element
        self.id = element.Id
        self.wert = id_wert(element.Id)
        self.typ_id = typ.Id
        self.typ_name = typ.Name
        self.exemplar = exemplar
        self.name = element.Name if exemplar else typ.Name
        self.link_doc = link_doc
        self.ansichten = []

    @property
    def geladen(self):
        return self.link_doc is not None


def _link_doc(exemplare):
    for exemplar in exemplare:
        try:
            dokument = exemplar.GetLinkDocument()
        except Exception:
            dokument = None
        if dokument is not None:
            return dokument
    return None


def verknuepfungen(doc):
    """Alle RVT-Verknüpfungen erster Ebene: je Typ ein Eintrag, bei mehreren
    Exemplaren zusätzlich je Exemplar einer (gleich hinter dem Typ)."""
    exemplare = {}
    for exemplar in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
        exemplare.setdefault(id_wert(exemplar.GetTypeId()), []).append(
            exemplar)

    typen = lg.ansichtstypen()
    ergebnis = []
    link_typen = list(FilteredElementCollector(doc).OfClass(RevitLinkType))
    link_typen.sort(key=lambda typ: typ.Name.lower())
    for typ in link_typen:
        try:
            if typ.IsNestedLink:
                continue
        except Exception:
            pass
        eigene = exemplare.get(id_wert(typ.Id), [])
        if not eigene:
            continue
        link_doc = _link_doc(eigene)
        eintrag = Verknuepfung(typ, typ, link_doc)
        eintrag.ansichten = _ansichten(eintrag, link_doc, typen)
        ergebnis.append(eintrag)
        if len(eigene) > 1:
            for exemplar in sorted(eigene, key=lambda e: e.Name.lower()):
                einzeln = Verknuepfung(exemplar, typ, link_doc, exemplar=True)
                einzeln.ansichten = _ansichten(einzeln, link_doc, typen)
                ergebnis.append(einzeln)
    return ergebnis


def _ansichten(verknuepfung, link_doc, typen):
    if link_doc is None:
        return []
    ergebnis = []
    for ansicht in FilteredElementCollector(link_doc).OfClass(View):
        try:
            if ansicht.IsTemplate or lg.gruppe(ansicht.ViewType) is None:
                continue
            ergebnis.append(LinkAnsicht(verknuepfung, ansicht, typen))
        except Exception:
            continue
    ergebnis.sort(key=lambda a: (lg.GRUPPEN_REIHENFOLGE.index(a.gruppe),
                                 a.art.lower(), a.name.lower()))
    return ergebnis


# ---------------------------------------------------------------------------
# Ansichten im aktuellen Modell
# ---------------------------------------------------------------------------

def unterstuetzt(ansicht):
    """Kann diese Ansicht (oder Vorlage) Link-Einstellungen haben?"""
    try:
        if lg.gruppe(ansicht.ViewType) is None:
            return False
        return ansicht.IsTemplate or ansicht.AreGraphicsOverridesAllowed()
    except Exception:
        return False


def host_ansichten(doc):
    """Ansichten und Vorlagen, in denen sich eine verknüpfte Ansicht setzen
    lässt."""
    return [a for a in FilteredElementCollector(doc).OfClass(View)
            if unterstuetzt(a)]


def steuernde_vorlage(doc, ansicht):
    """Vorlage, die die RVT-Verknüpfungen dieser Ansicht steuert, sonst None
    (Haken bei "V/G-Überschreibungen RVT-Verknüpfungen" in der Vorlage)."""
    try:
        if ansicht.IsTemplate:
            return None
        vorlage_id = ansicht.ViewTemplateId
        if vorlage_id is None or id_wert(vorlage_id) == -1:
            return None
        vorlage = doc.GetElement(vorlage_id)
        if vorlage is None:
            return None
        param = id_wert(ElementId(BuiltInParameter.VIS_GRAPHICS_RVT_LINKS))
        frei = set(id_wert(i) for i in
                   vorlage.GetNonControlledTemplateParameterIds())
        return None if param in frei else vorlage
    except Exception:
        return None


def zustand(ansicht, verknuepfung):
    """(Art, Id-Wert der verknüpften Ansicht oder None, geerbt)

    geerbt = True, wenn ein Exemplar keine eigene Einstellung hat und die des
    Typs gilt.
    """
    geerbt = False
    einstellung = None
    try:
        einstellung = ansicht.GetLinkOverrides(verknuepfung.id)
        if einstellung is None and verknuepfung.exemplar:
            geerbt = True
            einstellung = ansicht.GetLinkOverrides(verknuepfung.typ_id)
    except Exception:
        einstellung = None
    if einstellung is None:
        return HOST, None, geerbt
    art = einstellung.LinkVisibilityType
    if art == LinkVisibility.ByLinkView:
        return LINK, id_wert(einstellung.LinkedViewId), geerbt
    if art == LinkVisibility.Custom:
        return EIGEN, None, geerbt
    return HOST, None, geerbt


def _einstellung(link_ansicht_id, vorsichtig):
    einstellung = RevitLinkGraphicsSettings()
    einstellung.LinkVisibilityType = LinkVisibility.ByLinkView
    einstellung.LinkedViewId = link_ansicht_id
    if vorsichtig:
        # Ansichten ohne Farbfüllung/Ansichtsbereich lehnen "nach
        # verknüpfter Ansicht" für diese Teile ab
        einstellung.ColorFill = LinkVisibility.ByHostView
        einstellung.ViewRange = LinkVisibility.ByHostView
    return einstellung


def setze(ansicht, verknuepfung, link_ansicht):
    """Nach verknüpfter Ansicht: link_ansicht. Wirft bei Ablehnung."""
    try:
        ansicht.SetLinkOverrides(verknuepfung.id,
                                 _einstellung(link_ansicht.id, False))
        return
    except Exception as erster:
        fehler = erster
    try:
        einstellung = _einstellung(link_ansicht.id, False)
        if not RevitLinkGraphicsSettings.IsViewRangeSupported(ansicht):
            einstellung.ViewRange = LinkVisibility.ByHostView
        ansicht.SetLinkOverrides(verknuepfung.id, einstellung)
        return
    except Exception:
        pass
    try:
        ansicht.SetLinkOverrides(verknuepfung.id,
                                 _einstellung(link_ansicht.id, True))
    except Exception:
        raise fehler


def zuruecksetzen(ansicht, verknuepfung):
    """Typ: zurück auf "Nach Basisbauteilansicht". Exemplar: eigene
    Einstellung weg, es gilt wieder die des Typs."""
    ansicht.RemoveLinkOverrides(verknuepfung.id)


def exemplare_mit_eigener(ansicht, verknuepfungen, typ_eintrag):
    """Exemplare des Typs, die in der Ansicht eine eigene Einstellung haben -
    für sie ändert eine Einstellung am Typ nichts."""
    namen = []
    for v in verknuepfungen:
        if not v.exemplar or id_wert(v.typ_id) != id_wert(typ_eintrag.id):
            continue
        try:
            if ansicht.GetLinkOverrides(v.id) is not None:
                namen.append(v.name)
        except Exception:
            pass
    return namen


def verwendung(doc, ansichten, verknuepfungen):
    """{(Verknüpfungs-Wert, Ansichts-Wert): [Ansichtsnamen]} - wo welche
    verknüpfte Ansicht schon eingestellt ist.

    Ansichten, deren Vorlage die RVT-Verknüpfungen steuert, zählen mit dem
    Wert ihrer Vorlage.
    """
    typen = lg.ansichtstypen()
    vorlage_text = t(u"Vorlage", u"Template", u"Plantilla")
    ergebnis = {}
    zwischen = {}
    for ansicht in ansichten:
        quelle = steuernde_vorlage(doc, ansicht) or ansicht
        if ansicht.IsTemplate:
            name = u"%s: %s" % (vorlage_text, ansicht.Name)
        else:
            name = u"%s: %s" % (typen.get(u"%s" % ansicht.ViewType,
                                          u"%s" % ansicht.ViewType),
                                ansicht.Name)
        for v in verknuepfungen:
            schluessel = (id_wert(quelle.Id), v.wert)
            if schluessel not in zwischen:
                zwischen[schluessel] = zustand(quelle, v)
            art, wert, _geerbt = zwischen[schluessel]
            if art == LINK and wert is not None:
                ergebnis.setdefault((v.wert, wert), []).append(name)
    return ergebnis
