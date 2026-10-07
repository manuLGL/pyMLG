# -*- coding: utf-8 -*-
"""Plangrösse ändern: Plankopf-Typ und Format-Parameter auf vielen Plänen
setzen, den Planinhalt mitnehmen und die Pläne wie ein Vorbild ausrichten.

Format-Parameter sind die beschreibbaren Exemplar-Parameter des Plankopfs
vom Typ Ja/Nein, Ganzzahl oder Zahl/Länge - so werden Familien mit
mehreren Formaten (A0/A1/A2 per Ja/Nein, Breite/Höhe als Länge)
abgedeckt. Textparameter (Gezeichnet, Datum ...) gehören zum einzelnen
Plan und bleiben unberührt.

Genutzt von "Plangrösse" (Reiter MLGplans). Die Änderungen müssen in
einer offenen Transaktion laufen. Läuft unter IronPython.
"""

from System.Collections.Generic import List

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter as BIP,
    ElementId,
    ElementTransformUtils,
    FilteredElementCollector,
    ScheduleSheetInstance,
    SpecTypeId,
    StorageType,
    XYZ,
)

from mlg_plaene import ausrichten as ar
from mlg_plaene import logik as lg
from mlg_plaene import revit as rv
from mlg_sprache import t

JA_NEIN, GANZZAHL, ZAHL = "ja_nein", "ganzzahl", "zahl"


# ---------------------------------------------------------------- Plankopf

def plankoepfe_auf(doc, plan):
    """Alle Plankopf-Exemplare auf dem Plan (meist einer)."""
    return list(FilteredElementCollector(doc, plan.Id)
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .WhereElementIsNotElementType())


def typ_name(doc, kopf):
    """"Familie: Typ" eines Plankopf-Exemplars - derselbe Parameter wie in
    rv.plankoepfe, damit der Name dort gefunden wird. (Element.Name geht
    unter IronPython bei Typen nicht: AttributeError.)"""
    typ = doc.GetElement(kopf.GetTypeId())
    param = typ.get_Parameter(BIP.SYMBOL_FAMILY_AND_TYPE_NAMES_PARAM)
    return param.AsString() if param is not None else typ.FamilyName


def beispiel_kopf(doc, typ):
    """Ein Plankopf-Exemplar dieses Typs, sonst derselben Familie, sonst None
    - um die Format-Parameter einer Familie zu lesen."""
    familie = None
    for kopf in (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_TitleBlocks)
                 .WhereElementIsNotElementType()):
        if rv.id_wert(kopf.GetTypeId()) == rv.id_wert(typ.Id):
            return kopf
        if familie is None and doc.GetElement(kopf.GetTypeId()).FamilyName == typ.FamilyName:
            familie = kopf
    return familie


def _art(param):
    if param.StorageType == StorageType.Double:
        return ZAHL
    try:
        if param.Definition.GetDataType() == SpecTypeId.Boolean.YesNo:
            return JA_NEIN
    except Exception:
        pass
    return GANZZAHL


def wert_text(param, art):
    if art == ZAHL:
        return param.AsValueString() or u"{}".format(param.AsDouble())
    return u"{}".format(param.AsInteger())


def format_parameter(kopf):
    """[(Name, Art, Wert als Text)] der Format-Parameter, nach Name.

    Eingebaute Parameter (negative Id) zählen nicht, ebenso schreibgeschützte
    (per Formel oder Typ bestimmt).
    """
    gesehen, ergebnis = set(), []
    for param in kopf.Parameters:
        if param.IsReadOnly or param.StorageType not in (StorageType.Integer,
                                                         StorageType.Double):
            continue
        if rv.id_wert(param.Id) < 0:
            continue
        name = param.Definition.Name
        if name in gesehen:
            continue
        gesehen.add(name)
        art = _art(param)
        ergebnis.append((name, art, wert_text(param, art)))
    return sorted(ergebnis, key=lambda e: lg.natuerlich(e[0]))


def format_schluessel(doc, kopf):
    """Gleicher Schlüssel = gleiches Format (Typ und Format-Parameter)."""
    return (rv.id_wert(kopf.GetTypeId()),
            tuple((n, w) for n, _, w in format_parameter(kopf)))


def parameter_setzen(param, art, text):
    """Setzt einen Format-Parameter aus dem Eingabetext.

    Liefert True bei Änderung, False wenn der Wert schon stimmt. Wirft
    ValueError bei unlesbarer Eingabe.
    """
    if wert_text(param, art) == text:
        return False
    if art == JA_NEIN:
        param.Set(1 if text.strip() in (u"1", u"ja", u"yes", u"sí", u"si") else 0)
    elif art == GANZZAHL:
        param.Set(int(text.strip()))
    elif not param.SetValueString(text):
        raise ValueError(text)
    return True


# ---------------------------------------------------------------- Inhalt

def inhalt_ids(doc, plan):
    """Alles, was auf dem Plan selbst liegt, ohne Plankopf: Ansichtsfenster,
    Bauteillisten, Texte, Linien, Bilder, Revisionswolken ..."""
    plan_id = rv.id_wert(plan.Id)
    kopf_kategorie = int(BuiltInCategory.OST_TitleBlocks)
    ids = []
    for elem in FilteredElementCollector(doc, plan.Id).WhereElementIsNotElementType():
        if elem.Category is None or rv.id_wert(elem.OwnerViewId) != plan_id:
            continue
        if rv.id_wert(elem.Category.Id) == kopf_kategorie:
            continue
        if isinstance(elem, ScheduleSheetInstance) and elem.IsTitleblockRevisionSchedule:
            continue
        if elem.GroupId != ElementId.InvalidElementId:
            continue                      # Gruppenmitglied - die Gruppe wandert
        ids.append(elem.Id)
    return ids


def _id_liste(ids):
    liste = List[ElementId]()
    for i in ids:
        liste.Add(i)
    return liste


def inhalt_verschieben(doc, plan, dx, dy):
    """Verschiebt den Planinhalt um (dx, dy). Angeheftete Elemente werden
    dafür kurz gelöst. Liefert (Anzahl, Hinweise)."""
    ids = inhalt_ids(doc, plan)
    if not ids or (abs(dx) < lg.EPS and abs(dy) < lg.EPS):
        return 0, []
    angeheftet = [doc.GetElement(i) for i in ids]
    angeheftet = [e for e in angeheftet if e.Pinned]
    for elem in angeheftet:
        elem.Pinned = False
    weg, hinweise, anzahl = XYZ(dx, dy, 0), [], 0
    try:
        ElementTransformUtils.MoveElements(doc, _id_liste(ids), weg)
        anzahl = len(ids)
    except Exception:
        # einzeln, damit ein sperriges Element nicht alles aufhält
        for i in ids:
            try:
                ElementTransformUtils.MoveElement(doc, i, weg)
                anzahl += 1
            except Exception:
                elem = doc.GetElement(i)
                hinweise.append(t(u"'{}' ließ sich nicht verschieben",
                                  u"'{}' could not be moved",
                                  u"'{}' no se pudo mover").format(
                    elem.Category.Name))
    for elem in angeheftet:
        elem.Pinned = True
    return anzahl, hinweise


# ---------------------------------------------------------------- Ablauf

class Aenderung(object):
    """Was auf jeden Plan angewendet wird.

    typ:        Plankopf-Typ oder None (unverändert)
    werte:      [(Name, Art, Text)] zu setzende Format-Parameter
    anker:      lg.OBEN_LINKS ... oder None (Inhalt nicht mitnehmen)
    """

    def __init__(self, typ=None, werte=(), anker=None):
        self.typ = typ
        self.werte = list(werte)
        self.anker = anker

    def leer(self):
        return self.typ is None and not self.werte


def format_aendern(doc, plan, aenderung):
    """Setzt Plankopf-Typ und Format-Parameter auf einem Plan und nimmt den
    Inhalt mit. Liefert (geändert?, Hinweise)."""
    koepfe = plankoepfe_auf(doc, plan)
    if not koepfe:
        return False, [t(u"kein Plankopf - übersprungen",
                         u"no title block - skipped",
                         u"sin cajetín: omitido")]
    hinweise = []
    if len(koepfe) > 1:
        hinweise.append(t(u"{} Planköpfe - nur der erste wird geändert",
                          u"{} title blocks - only the first one is changed",
                          u"{} cajetines: sólo se cambia el primero").format(len(koepfe)))
    kopf = koepfe[0]
    alt = rv.plankopf_rechteck(doc, plan)

    geaendert = False
    if aenderung.typ is not None and \
            rv.id_wert(kopf.GetTypeId()) != rv.id_wert(aenderung.typ.Id):
        kopf.ChangeTypeId(aenderung.typ.Id)
        geaendert = True
    for name, art, text in aenderung.werte:
        param = kopf.LookupParameter(name)
        if param is None or param.IsReadOnly:
            hinweise.append(t(u"Parameter '{}' fehlt oder ist schreibgeschützt",
                              u"parameter '{}' is missing or read-only",
                              u"el parámetro '{}' falta o es de sólo lectura").format(name))
            continue
        try:
            geaendert = parameter_setzen(param, art, text) or geaendert
        except Exception:
            hinweise.append(t(u"'{}' = '{}' ließ sich nicht setzen",
                              u"'{}' = '{}' could not be set",
                              u"'{}' = '{}' no se pudo asignar").format(name, text))
    if not geaendert:
        return False, hinweise

    doc.Regenerate()
    if aenderung.anker:
        neu = rv.plankopf_rechteck(doc, plan)
        dx, dy = lg.anker_versatz(alt, neu, aenderung.anker)
        _, nicht = inhalt_verschieben(doc, plan, dx, dy)
        hinweise.extend(nicht)
    return True, hinweise


def listen_wie_vorbild(doc, vorbild, ziel):
    """Bauteillisten, die auch auf dem Vorbild liegen, an dieselbe Stelle.
    Liefert die Anzahl."""
    def listen(plan):
        return [l for l in FilteredElementCollector(doc, plan.Id)
                .OfClass(ScheduleSheetInstance)
                if not l.IsTitleblockRevisionSchedule]

    vorlage = dict((rv.id_wert(l.ScheduleId), l) for l in listen(vorbild))
    anzahl = 0
    for liste in listen(ziel):
        muster = vorlage.get(rv.id_wert(liste.ScheduleId))
        if muster is not None and not muster.Point.IsAlmostEqualTo(liste.Point):
            liste.Point = muster.Point
            anzahl += 1
    return anzahl


def wie_vorbild(doc, vorbild, ziel, art):
    """Ansichtsfenster und Bauteillisten von ziel wie auf vorbild.
    Liefert (Anzahl, Hinweise)."""
    vorbilder = [doc.GetElement(i) for i in vorbild.GetAllViewports()]
    ziele = [doc.GetElement(i) for i in ziel.GetAllViewports()]
    anzahl, hinweise = 0, []
    if vorbilder and ziele:
        anzahl, hinweise = ar.wie_vorbild(doc, vorbilder, ziele, art, melden="keine")
    return anzahl + listen_wie_vorbild(doc, vorbild, ziel), hinweise
