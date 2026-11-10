# -*- coding: utf-8 -*-
"""Revit-Zugriffe von ApartmentAreas.

Jeder Eingangspfeil (wohnung_top.logik.MARKER_FAMILIE) steckt in einer
Modellgruppe (Wohnung oder Gewerbe). Gelesen werden die Räume der
obersten Gruppe, zu der der Pfeil gehört (auch aus verschachtelten
Gruppen), summiert nach Raumkategorie (logik.REGELN). Geschrieben wird nur
in den Pfeil - Gruppen und Räume bleiben unverändert.

Der Pfeil liegt in einer Gruppe: seine Parameter müssen "Werte je nach
Gruppeninstanz unterschiedlich" erlauben, sonst lehnt Revit das Schreiben
ab. Das Werkzeug prüft vorher, fragt nach und rollt bei Fehlern zurück -
Gruppen werden nie aufgelöst.
"""

from Autodesk.Revit.DB import (  # noqa: E402
    ElementId,
    FamilyInstance,
    FilteredElementCollector,
    Group,
    StorageType,
    Transaction,
    TransactionGroup,
    TransactionStatus,
)
from Autodesk.Revit.DB.Architecture import Room  # noqa: E402

from mlg_sprache import t  # noqa: E402
from wohnung_flaechen import logik as lg  # noqa: E402
from wohnung_top import logik as top  # noqa: E402
from wohnung_top import revit as wt  # noqa: E402

KATEGORIE_PARAMETER = u"Raumkategorie"

id_wert = wt.id_wert
parameter = wt.parameter
pruefe_parameter = wt.pruefe_parameter
GruppenFehler = wt.GruppenFehler


def kategorie(raum):
    """Raumkategorie als Text - Schlüssel (ElementId) oder Text."""
    p = parameter(raum, KATEGORIE_PARAMETER)
    if p is None:
        return u""
    try:
        if p.StorageType == StorageType.ElementId:
            zeile = raum.Document.GetElement(p.AsElementId())
            return zeile.Name if zeile is not None else u""
        if p.StorageType == StorageType.String:
            return p.AsString() or u""
        return p.AsValueString() or u""
    except Exception:
        return u""


def _raeume(doc, gruppe, ergebnis):
    for i in gruppe.GetMemberIds():
        element = doc.GetElement(i)
        if isinstance(element, Room):
            ergebnis.append(element)
        elif isinstance(element, Group):
            _raeume(doc, element, ergebnis)
    return ergebnis


def oberste_gruppe(doc, element):
    gruppe = None
    gruppen_id = element.GroupId
    while gruppen_id is not None and gruppen_id != ElementId.InvalidElementId:
        gruppe = doc.GetElement(gruppen_id)
        if gruppe is None:
            break
        gruppen_id = gruppe.GroupId
    return gruppe


class Pfeil(object):
    """Ein Eingangspfeil mit den Summen seiner Gruppe."""

    def __init__(self, element, gruppe, werte, ungezaehlt, anzahl):
        self.element = element
        self.gruppe = gruppe
        self.werte = werte              # {Parameter: Fläche in Fuss²}
        self.ungezaehlt = ungezaehlt    # {Raumkategorie: Anzahl}
        self.raeume = anzahl

    def beschreibung(self):
        name = u""
        if self.gruppe is not None and self.gruppe.GroupType is not None:
            name = self.gruppe.GroupType.Name
        return u"%s [%d]" % (name or u"-", id_wert(self.element.Id))


def pfeile(doc, ids=None):
    """(Pfeile mit Gruppe, Pfeile ohne Gruppe). ids: nur diese Elemente
    (Auswahl), sonst alle im Projekt."""
    if ids:
        elemente = [doc.GetElement(i) for i in ids]
    else:
        elemente = FilteredElementCollector(doc).OfClass(FamilyInstance) \
            .WhereElementIsNotElementType()
    mit, ohne = [], []
    cache = {}
    for element in elemente:
        if not isinstance(element, FamilyInstance):
            continue
        if not top.ist_marker(wt.familienname(element)):
            continue
        gruppe = oberste_gruppe(doc, element)
        if gruppe is None:
            ohne.append(element)
            continue
        schluessel = id_wert(gruppe.Id)
        if schluessel not in cache:
            raeume = _raeume(doc, gruppe, [])
            werte, ungezaehlt = lg.summen(
                [(kategorie(r), r.Area) for r in raeume])
            cache[schluessel] = (werte, ungezaehlt, len(raeume))
        werte, ungezaehlt, anzahl = cache[schluessel]
        mit.append(Pfeil(element, gruppe, dict(werte), ungezaehlt, anzahl))
    return mit, ohne


class Lauf(object):
    """Werte in die Pfeile schreiben - eine Transaktionsgruppe, am Ende
    übernehmen (ein Rückgängig-Schritt) oder verwerfen."""

    def __init__(self, doc):
        self.doc = doc
        self.pfeile = 0            # Pfeile, die Werte bekommen haben
        self.fehler = []           # [(Pfeil, Parameter, Grund)]
        self.gruppe = TransactionGroup(
            doc, t(u"Wohnungsflächen in Eingangspfeile",
                   u"Apartment areas into entrance arrows",
                   u"Superficies en flechas de entrada"))
        self.gruppe.Start()

    def _transaktion(self, name, arbeit):
        gesammelt = []
        transaktion = Transaction(self.doc, name)
        transaktion.Start()
        optionen = transaktion.GetFailureHandlingOptions()
        optionen.SetFailuresPreprocessor(wt._vorverarbeiter(gesammelt))
        transaktion.SetFailureHandlingOptions(optionen)
        try:
            arbeit()
        except Exception:
            transaktion.RollBack()
            raise
        if transaktion.Commit() != TransactionStatus.Committed:
            raise GruppenFehler(u"\n".join(gesammelt) or t(
                u"Revit hat die Änderung abgelehnt.",
                u"Revit rejected the change.",
                u"Revit ha rechazado el cambio."))

    def erlaube_varianz(self, definitionen):
        def arbeit():
            for definition in definitionen:
                definition.SetAllowVaryBetweenGroups(self.doc, True)
        self._transaktion(t(u"Werte je Gruppeninstanz erlauben",
                            u"Allow values per group instance",
                            u"Permitir valores por instancia de grupo"),
                          arbeit)

    def schreibe(self, liste, namen):
        def arbeit():
            for pfeil in liste:
                geschrieben = False
                for name in namen:
                    p = parameter(pfeil.element, name)
                    if p is None:
                        continue
                    grund = _setze_flaeche(p, pfeil.werte[name])
                    if grund:
                        self.fehler.append((pfeil, name, grund))
                    else:
                        geschrieben = True
                if geschrieben:
                    self.pfeile += 1
        self._transaktion(t(u"Wohnungsflächen schreiben",
                            u"Write apartment areas",
                            u"Escribir superficies"), arbeit)

    def uebernehmen(self):
        self.gruppe.Assimilate()

    def verwerfen(self):
        if self.gruppe.HasStarted():
            self.gruppe.RollBack()


def _setze_flaeche(p, flaeche_qf):
    """Fläche (Fuss², intern) setzen - Fehlertext oder None. Ein
    Textparameter bekommt '12,34 m²'."""
    if p.IsReadOnly:
        return t(u"schreibgeschützt", u"read-only", u"de solo lectura")
    try:
        if p.StorageType == StorageType.Double:
            p.Set(float(flaeche_qf))
        elif p.StorageType == StorageType.String:
            p.Set(lg.qm_text(flaeche_qf))
        else:
            return t(u"falscher Parametertyp", u"wrong parameter type",
                     u"tipo de parámetro incorrecto")
    except Exception as fehler:
        return u"%s" % fehler
    return None
