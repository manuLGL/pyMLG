# -*- coding: utf-8 -*-
"""Revit-Zugriffe von StackValues.

Quellen sind die gewählten Stützen, Wände und Unterzüge. Ziele sind alle
übrigen Elemente derselben Kategorien im aktiven Modell (keine
Verknüpfungen). Die Lage ist der Grundriss (logik.py), die Ebene die
Basis- bzw. Referenzebene des Elements.
"""

import uuid

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    BuiltInParameter,
    ElementId,
    FailureProcessingResult,
    FailureSeverity,
    FilteredElementCollector,
    IFailuresPreprocessor,
    StorageType,
    Transaction,
)
from Autodesk.Revit.Exceptions import OperationCanceledException  # noqa: E402
from Autodesk.Revit.UI.Selection import ISelectionFilter, ObjectType  # noqa: E402
from System import Guid  # noqa: E402

from mlg_sprache import t  # noqa: E402
from werte_stapel import logik as lg  # noqa: E402

FUSS = 0.3048

# (Kürzel, Kategorie, Lage als Punkt?)
KATEGORIEN = (
    ("tragwerk", BuiltInCategory.OST_StructuralColumns, True),
    ("architektur", BuiltInCategory.OST_Columns, True),
    ("wand", BuiltInCategory.OST_Walls, False),
    ("unterzug", BuiltInCategory.OST_StructuralFraming, False),
)


def kategorie_name(kuerzel, anzahl=2):
    namen = {
        "tragwerk": (t(u"Tragwerksstütze", u"structural column",
                       u"pilar estructural"),
                     t(u"Tragwerksstützen", u"structural columns",
                       u"pilares estructurales")),
        "architektur": (t(u"Architekturstütze", u"architectural column",
                          u"pilar arquitectónico"),
                        t(u"Architekturstützen", u"architectural columns",
                          u"pilares arquitectónicos")),
        "wand": (t(u"Wand", u"wall", u"muro"),
                 t(u"Wände", u"walls", u"muros")),
        "unterzug": (t(u"Unterzug", u"beam", u"viga"),
                     t(u"Unterzüge", u"beams", u"vigas")),
    }[kuerzel]
    return namen[0] if anzahl == 1 else namen[1]


# Diese eingebauten Parameter stehen in der Liste oben
BEVORZUGT = ("ALL_MODEL_MARK", "ALL_MODEL_INSTANCE_COMMENTS")

# Basis- bzw. Referenzebene - der erste vorhandene zählt
EBENEN_PARAMETER = ("WALL_BASE_CONSTRAINT", "FAMILY_BASE_LEVEL_PARAM",
                    "INSTANCE_REFERENCE_LEVEL_PARAM", "SCHEDULE_LEVEL_PARAM")

OHNE_EBENE = -1

_WERTTYPEN = (StorageType.String, StorageType.Integer, StorageType.Double)

_klassen = {}
_pruefungen = {}


def id_wert(element_id):
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def beschreibe(element):
    try:
        return u"%s [%d]" % (element.Name, id_wert(element.Id))
    except Exception:
        return u"[%d]" % id_wert(element.Id)


# ---------------------------------------------------------------- Elemente

_KATEGORIE_IDS = None


def kategorie(element):
    """Kürzel aus KATEGORIEN oder None."""
    global _KATEGORIE_IDS
    if _KATEGORIE_IDS is None:
        _KATEGORIE_IDS = dict((id_wert(ElementId(bic)), kuerzel)
                              for kuerzel, bic, _ in KATEGORIEN)
    kat = getattr(element, "Category", None)
    if kat is None:
        return None
    return _KATEGORIE_IDS.get(id_wert(kat.Id))


def _xy(p):
    return (p.X, p.Y)


def lage(element, kuerzel):
    """Grundriss-Lage (logik.py) oder None."""
    ort = element.Location
    if ort is None:
        return None
    kurve = getattr(ort, "Curve", None)
    als_punkt = dict((k, p) for k, _, p in KATEGORIEN)[kuerzel]
    if als_punkt:
        p = getattr(ort, "Point", None)
        if p is None and kurve is not None:
            # geneigte Stütze: unterer Endpunkt
            a, b = kurve.GetEndPoint(0), kurve.GetEndPoint(1)
            p = a if a.Z <= b.Z else b
        return lg.punkt_lage(_xy(p)) if p is not None else None
    if kurve is None:
        return None
    try:
        mitte = kurve.Evaluate(0.5, True)
    except Exception:
        return None
    return lg.linien_lage(_xy(kurve.GetEndPoint(0)),
                          _xy(kurve.GetEndPoint(1)), _xy(mitte))


def ebene_id(element):
    """Id-Wert der Basis-/Referenzebene oder OHNE_EBENE."""
    for name in EBENEN_PARAMETER:
        bip = getattr(BuiltInParameter, name, None)
        if bip is None:
            continue
        try:
            p = element.get_Parameter(bip)
        except Exception:
            p = None
        if p is not None and p.StorageType == StorageType.ElementId:
            wert = id_wert(p.AsElementId())
            if wert > 0:
                return wert
    try:
        wert = id_wert(element.LevelId)
        return wert if wert > 0 else OHNE_EBENE
    except Exception:
        return OHNE_EBENE


# ---------------------------------------------------------------- Parameter

def schluessel_von(p):
    """'bip:ALL_MODEL_MARK', 'guid:<GUID>' (gemeinsam genutzt) oder
    'name:<Name>' (Projektparameter)."""
    definition = p.Definition
    eingebaut = getattr(definition, "BuiltInParameter", None)
    if eingebaut is not None and eingebaut != BuiltInParameter.INVALID:
        return u"bip:%s" % eingebaut.ToString()
    try:
        if p.IsShared:
            return u"guid:%s" % p.GUID.ToString()
    except Exception:
        pass
    return u"name:%s" % definition.Name


def parameter(element, schluessel):
    art, _, name = schluessel.partition(u":")
    try:
        if art == u"bip":
            bip = getattr(BuiltInParameter, name, None)
            return element.get_Parameter(bip) if bip is not None else None
        if art == u"guid":
            return element.get_Parameter(Guid(name))
        return element.LookupParameter(name)
    except Exception:
        return None


def _wert(p):
    """Wert für den Vergleich: Text, Zahl oder None (leer)."""
    art = p.StorageType
    if art == StorageType.String:
        return p.AsString() or u""
    if not p.HasValue:
        return None
    if art == StorageType.Integer:
        return p.AsInteger()
    return p.AsDouble()


def _anzeige(p):
    """Wert so, wie Revit ihn anzeigt."""
    if p.StorageType == StorageType.String:
        return p.AsString() or u""
    try:
        text = p.AsValueString()
        if text:
            return text
    except Exception:
        pass
    wert = _wert(p)
    return u"" if wert is None else u"%s" % wert


class Option(object):
    def __init__(self, schluessel, name, kategorien):
        self.schluessel = schluessel
        self.name = name
        self.kategorien = kategorien    # set der Kürzel, die ihn haben
        self.anzeige = name


def parameter_optionen(elemente):
    """Beschreibbare Exemplarparameter (Text, Zahl, Ja/Nein) der Quellen."""
    gefunden = {}
    for element in elemente:
        kuerzel = kategorie(element)
        for p in element.Parameters:
            if p.StorageType not in _WERTTYPEN or p.IsReadOnly:
                continue
            schluessel = schluessel_von(p)
            option = gefunden.get(schluessel)
            if option is None:
                option = gefunden[schluessel] = Option(
                    schluessel, p.Definition.Name, set())
            option.kategorien.add(kuerzel)

    alle = set(kategorie(e) for e in elemente)
    namen = {}
    for option in gefunden.values():
        namen[option.name] = namen.get(option.name, 0) + 1
    for option in gefunden.values():
        text = option.name
        if namen[option.name] > 1:
            art = option.schluessel.partition(u":")[0]
            text += {u"bip": u"", u"guid": t(u" (gemeinsam genutzt)",
                                             u" (shared)",
                                             u" (compartido)")}.get(
                art, t(u" (Projekt)", u" (project)", u" (proyecto)"))
        if option.kategorien != alle:
            text += u"  – " + t(u"nur", u"only", u"solo") + u" " + \
                u", ".join(kategorie_name(k) for k, _, _ in KATEGORIEN
                           if k in option.kategorien)
        option.anzeige = text

    def rang(option):
        for i, bip in enumerate(BEVORZUGT):
            if option.schluessel == u"bip:" + bip:
                return (i, u"")
        return (len(BEVORZUGT), option.name.lower())
    return sorted(gefunden.values(), key=rang)


# ---------------------------------------------------------------- Auswahl

def _filter(name, erlaubt):
    """ISelectionFilter als .NET-Klasse. Der Namensraum ist je Sitzung
    eindeutig, weil die CPython-Engine im Revit-Prozess weiterlebt."""
    if name not in _klassen:
        namensraum = "pyMLG.WerteStapel_%s_%s" % (name, uuid.uuid4().hex)

        class Filter(ISelectionFilter):
            __namespace__ = namensraum

            def AllowElement(self, element):
                try:
                    return bool(_pruefungen[name](element))
                except Exception:
                    return False

            def AllowReference(self, referenz, punkt_):
                return False

        _klassen[name] = Filter
    _pruefungen[name] = erlaubt
    return _klassen[name]()


def auswahl(uidoc):
    """(passende gewählte Elemente, Anzahl übergangener)."""
    doc = uidoc.Document
    elemente = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    elemente = [e for e in elemente if e is not None]
    passend = [e for e in elemente if kategorie(e) is not None]
    return passend, len(elemente) - len(passend)


def waehle_quellen(uidoc):
    """Quellen per Rechteck/Klick wählen - [] bei ESC."""
    try:
        referenzen = uidoc.Selection.PickObjects(
            ObjectType.Element,
            _filter("quelle", lambda e: kategorie(e) is not None),
            t(u"Stützen, Wände und Unterzüge wählen, deren Werte übertragen "
              u"werden - dann 'Fertig stellen'",
              u"Select the columns, walls and beams whose values are copied "
              u"- then 'Finish'",
              u"Seleccione los pilares, muros y vigas cuyos valores se "
              u"copian; luego 'Finalizar'"))
    except OperationCanceledException:
        return []
    doc = uidoc.Document
    return [doc.GetElement(r.ElementId) for r in referenzen]


# ---------------------------------------------------------------- Bestand

class Bestand(object):
    """Quellen und mögliche Ziele, in Revit und als logik.Element."""

    def __init__(self, doc, quell_elemente):
        self.doc = doc
        self.quell_revit = []
        self.quellen = []
        self.ohne_lage = []          # gewählte Elemente ohne Grundriss-Lage
        for e in quell_elemente:
            kuerzel = kategorie(e)
            ort = lage(e, kuerzel)
            if ort is None:
                self.ohne_lage.append(e)
                continue
            self.quell_revit.append(e)
            self.quellen.append(lg.Element(id_wert(e.Id), kuerzel, ort,
                                           ebene_id(e)))

        quell_ids = set(q.kennung for q in self.quellen)
        self.ziel_revit = []
        self.ziele = []
        for kuerzel, bic, _ in KATEGORIEN:
            if not any(q.kategorie == kuerzel for q in self.quellen):
                continue
            sammler = FilteredElementCollector(doc).OfCategory(bic) \
                .WhereElementIsNotElementType()
            for e in sammler:
                kennung = id_wert(e.Id)
                if kennung in quell_ids:
                    continue
                ort = lage(e, kuerzel)
                if ort is None:
                    continue
                self.ziel_revit.append(e)
                self.ziele.append(lg.Element(kennung, kuerzel, ort,
                                             ebene_id(e)))
        # (Schlüssel, repr(Wert)) -> Text wie in Revit, für den Bericht
        self.anzeigen = {}

    def anzahl_je_kategorie(self):
        ergebnis = {}
        for q in self.quellen:
            ergebnis[q.kategorie] = ergebnis.get(q.kategorie, 0) + 1
        return ergebnis

    def lies_quellen(self, schluessel):
        for e, q in zip(self.quell_revit, self.quellen):
            q.werte = {}
            for s in schluessel:
                p = parameter(e, s)
                if p is None or p.StorageType not in _WERTTYPEN:
                    q.werte[s] = lg.FEHLT
                    continue
                q.werte[s] = _wert(p)
                self.anzeigen[(s, repr(q.werte[s]))] = _anzeige(p)

    def lies_ziele(self, indizes, schluessel):
        for zi in indizes:
            e, z = self.ziel_revit[zi], self.ziele[zi]
            z.werte = {}
            for s in schluessel:
                p = parameter(e, s)
                if p is None or p.StorageType not in _WERTTYPEN:
                    z.werte[s] = lg.FEHLT
                elif p.IsReadOnly:
                    z.werte[s] = lg.SCHREIBGESCHUETZT
                else:
                    z.werte[s] = _wert(p)
                    self.anzeigen[(s, repr(z.werte[s]))] = _anzeige(p)

    def anzeige(self, schluessel, wert):
        if wert is None:
            return u""
        return self.anzeigen.get((schluessel, repr(wert)), u"%s" % wert)


# ---------------------------------------------------------------- Schreiben

def _vorverarbeiter():
    """Warnungen (z.B. doppeltes Kennzeichen bei übereinanderstehenden
    Stützen) still bestätigen."""
    if "fehler" not in _klassen:
        namensraum = "pyMLG.WerteStapel_%s" % uuid.uuid4().hex

        class Vorverarbeiter(IFailuresPreprocessor):
            __namespace__ = namensraum

            def PreprocessFailures(self, accessor):
                for meldung in list(accessor.GetFailureMessages()):
                    if meldung.GetSeverity() == FailureSeverity.Warning:
                        accessor.DeleteWarning(meldung)
                return FailureProcessingResult.Continue

        _klassen["fehler"] = Vorverarbeiter
    return _klassen["fehler"]()


def _setze(p, wert):
    art = p.StorageType
    if art == StorageType.String:
        return p.Set(u"%s" % wert)
    if art == StorageType.Integer:
        return p.Set(int(wert))
    return p.Set(float(wert))


def schreibe(bestand, auftraege):
    """Aufträge (logik.Plan.auftraege) in einer Transaktion schreiben.
    Rückgabe: [(Ziel-Index, Schlüssel, Grund)] der gescheiterten."""
    fehler = []
    transaktion = Transaction(bestand.doc, t(u"Werte übertragen",
                                             u"Copy values",
                                             u"Copiar valores"))
    transaktion.Start()
    try:
        optionen = transaktion.GetFailureHandlingOptions()
        optionen.SetFailuresPreprocessor(_vorverarbeiter())
        transaktion.SetFailureHandlingOptions(optionen)
        for zi, s, wert in auftraege:
            p = parameter(bestand.ziel_revit[zi], s)
            try:
                if p is None or not _setze(p, wert):
                    fehler.append((zi, s, t(u"nicht gesetzt", u"not set",
                                            u"no asignado")))
            except Exception as ausnahme:
                text = getattr(ausnahme, "Message", None) or u"%s" % ausnahme
                fehler.append((zi, s, text.strip().splitlines()[0]
                               if text.strip() else u"?"))
    except Exception:
        transaktion.RollBack()
        raise
    transaktion.Commit()
    return fehler
