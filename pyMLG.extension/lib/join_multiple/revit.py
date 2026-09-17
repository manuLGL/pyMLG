# -*- coding: utf-8 -*-
"""Revit-Zugriffe von JoinMultiple: Elemente sammeln, verbinden, lösen.

Kandidaten werden über Bounding-Boxen gesucht (logik.kandidaten_paare) statt
über einen FilteredElementCollector je Element - das ist bei vielen Elementen
um Größenordnungen schneller. Ob sich zwei Elemente wirklich schneiden, prüft
danach ElementIntersectsElementFilter.

Warnungen wie "Elemente sind verbunden, schneiden sich aber nicht" würden
beim Commit je Paar einen Dialog erzeugen. Ein IFailuresPreprocessor löscht
sie; Fehler mit Lösungsvorschlag (meist "Verbindung aufheben") werden
automatisch aufgelöst.
"""

import io
import os
import time
import uuid

import clr

clr.AddReference("System")
from System.Collections.Generic import List  # noqa: E402

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    CategoryType,
    ElementId,
    ElementIntersectsElementFilter,
    ElementMulticategoryFilter,
    FailureProcessingResult,
    FailureSeverity,
    FilteredElementCollector,
    IFailuresPreprocessor,
    JoinGeometryUtils,
    StorageType,
    Transaction,
)

from join_multiple import logik as lg  # noqa: E402
from mlg_sprache import t  # noqa: E402

AUSWAHL = 0
ANSICHT = 1
MODELL = 2


def id_wert(element_id):
    """Zahlenwert einer ElementId (Revit 2024+: .Value)."""
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def net_liste(typ, werte):
    """Python-Iterable -> List[typ]. pythonnet (CPython) nimmt eine
    Python-Liste nicht direkt im Konstruktor an."""
    liste = List[typ]()
    for wert in werte:
        liste.Add(wert)
    return liste


def standard_kategorien():
    """{Kategorie-Id-Wert: (BuiltInCategory, Name)} der Vorgabeliste.
    Kategorien, die es in dieser Revit-Version nicht gibt, fehlen."""
    ergebnis = {}
    for name, _prio in lg.STANDARD_PRIORITAETEN:
        bic = getattr(BuiltInCategory, name, None)
        if bic is not None:
            ergebnis[id_wert(ElementId(bic))] = (bic, name)
    return ergebnis


def _passt(element, erlaubt, alle_kategorien):
    kategorie = element.Category
    if kategorie is None or element.ViewSpecific:
        return False
    if alle_kategorien:
        return kategorie.CategoryType == CategoryType.Model
    return id_wert(kategorie.Id) in erlaubt


def sammle(doc, uidoc, modus, alle_kategorien=False):
    """Elemente des Bereichs, die sich grundsätzlich verbinden lassen."""
    erlaubt = standard_kategorien()
    if modus == AUSWAHL:
        kandidaten = [doc.GetElement(i)
                      for i in uidoc.Selection.GetElementIds()]
    else:
        if modus == ANSICHT:
            sammler = FilteredElementCollector(doc, uidoc.ActiveView.Id)
        else:
            sammler = FilteredElementCollector(doc)
        sammler = sammler.WhereElementIsNotElementType()
        if not alle_kategorien:
            bics = net_liste(BuiltInCategory,
                             [b for b, _n in erlaubt.values()])
            sammler = sammler.WherePasses(ElementMulticategoryFilter(bics))
        kandidaten = list(sammler)
    return [e for e in kandidaten
            if e is not None and _passt(e, erlaubt, alle_kategorien)]


def kategorie_schluessel(element):
    return id_wert(element.Category.Id)


def parameter_zahl(doc, element, name):
    """Zahlenwert des Parameters am Exemplar, ersatzweise am Typ."""
    if not name:
        return None
    quellen = [element]
    try:
        typ = doc.GetElement(element.GetTypeId())
        if typ is not None:
            quellen.append(typ)
    except Exception:
        pass
    for quelle in quellen:
        parameter = quelle.LookupParameter(name)
        if parameter is None or not parameter.HasValue:
            continue
        art = parameter.StorageType
        if art == StorageType.Integer:
            return float(parameter.AsInteger())
        if art == StorageType.Double:
            return parameter.AsDouble()
        if art == StorageType.String:
            wert = lg.zahl_aus_wert(parameter.AsString())
            if wert is not None:
                return wert
    return None


# ---------------------------------------------------------------------------
# Warnungen unterdrücken
# ---------------------------------------------------------------------------

_vorverarbeiter_klasse = None

# Meldungen, die der Vorverarbeiter behandelt hat (für das Protokoll)
REVIT_MELDUNGEN = []

PROTOKOLL = os.path.join(
    os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "pyMLG",
    "JoinMultiple_Protokoll.txt")


def _vorverarbeiter():
    """IFailuresPreprocessor als .NET-Klasse (pythonnet). Der Namensraum ist
    je Sitzung eindeutig, weil die CPython-Engine im Revit-Prozess
    weiterlebt und ein zweiter Typ gleichen Namens scheitern würde.
    Gelingt das nicht, läuft alles ohne - Revit zeigt dann die Warnungen."""
    global _vorverarbeiter_klasse
    if _vorverarbeiter_klasse is None:
        namensraum = "pyMLG.JoinMultiple_%s" % uuid.uuid4().hex

        class Vorverarbeiter(IFailuresPreprocessor):
            __namespace__ = namensraum

            def PreprocessFailures(self, accessor):
                aufgeloest = False
                for meldung in list(accessor.GetFailureMessages()):
                    try:
                        text = meldung.GetDescriptionText()
                    except Exception:
                        text = u"?"
                    if meldung.GetSeverity() == FailureSeverity.Warning:
                        REVIT_MELDUNGEN.append(u"Warnung gelöscht: " + text)
                        accessor.DeleteWarning(meldung)
                    elif meldung.HasResolutions():
                        REVIT_MELDUNGEN.append(u"Fehler aufgelöst: " + text)
                        accessor.ResolveFailure(meldung)
                        aufgeloest = True
                    else:
                        REVIT_MELDUNGEN.append(u"Fehler: " + text)
                if aufgeloest:
                    return FailureProcessingResult.ProceedWithCommit
                return FailureProcessingResult.Continue

        _vorverarbeiter_klasse = Vorverarbeiter
    return _vorverarbeiter_klasse()


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


# ---------------------------------------------------------------------------
# Verbinden / Lösen
# ---------------------------------------------------------------------------

class Ergebnis(object):
    def __init__(self):
        self.anzahl = dict((a, 0) for a in (
            lg.NEU, lg.BEREITS, lg.UEBERSPRUNGEN, lg.UMGEDREHT, lg.RICHTIG,
            lg.GLEICH, lg.NICHT_GEPRUEFT, lg.ABGELEHNT))
        self.geloest = 0
        self.fehler = []           # Texte
        self.abgelehnt = []        # Texte
        self.bearbeitet = set()    # Id-Werte
        self.falsch_nach_speichern = []   # Texte
        self.protokoll = []        # Zeilen für JoinMultiple_Protokoll.txt

    @staticmethod
    def _paar(a, b):
        return u"%s (%s) + %s (%s)" % (a.Category.Name, id_wert(a.Id),
                                       b.Category.Name, id_wert(b.Id))

    def fehlschlag(self, a, b, ausnahme):
        text = getattr(ausnahme, "Message", None) or u"%s" % ausnahme
        text = (text.strip().splitlines() or [u""])[0]
        self.fehler.append(u"%s: %s" % (self._paar(a, b), text))


def _schneidet(doc, a, b):
    """Text für das Protokoll: wer schneidet wen."""
    try:
        if not JoinGeometryUtils.AreElementsJoined(doc, a, b):
            return u"nicht verbunden"
        return (u"%s schneidet" % id_wert(a.Id)
                if JoinGeometryUtils.IsCuttingElementInJoin(doc, a, b)
                else u"%s schneidet" % id_wert(b.Id))
    except Exception as ausnahme:
        return u"Fehler: %s" % ausnahme


def schreibe_protokoll(zeilen):
    try:
        ordner = os.path.dirname(PROTOKOLL)
        if not os.path.isdir(ordner):
            os.makedirs(ordner)
        with io.open(PROTOKOLL, "w", encoding="utf-8") as datei:
            datei.write(u"\n".join(zeilen) + u"\n")
        return PROTOKOLL
    except Exception:
        return None


def _schneiden_sich(a, b):
    for x, y in ((a, b), (b, a)):
        try:
            if ElementIntersectsElementFilter(x).PassesFilter(y):
                return True
        except Exception:
            continue
    return False


def _box(element):
    try:
        box = element.get_BoundingBox(None)
    except Exception:
        return None
    if box is None:
        return None
    return ((box.Min.X, box.Min.Y, box.Min.Z),
            (box.Max.X, box.Max.Y, box.Max.Z))


def verbinde(doc, elemente, prioritaet, nur_bei_schnitt=True,
             gleiche_kategorie=True, bestehende_anpassen=False):
    """Kandidatenpaare verbinden und die Schnittreihenfolge nach Priorität
    setzen.

    prioritaet  Funktion Element -> Zahl

    Drei Schritte in einer Transaktion:
      1. verbinden (JoinGeometry) - noch ohne Reihenfolge
      2. nach doc.Regenerate() die Reihenfolge prüfen und bei Bedarf
         tauschen. Direkt nach JoinGeometry liefert IsCuttingElementInJoin
         nicht verlässlich den neuen Zustand - ein Tausch auf veralteter
         Grundlage dreht die Reihenfolge genau falsch herum.
      3. nach erneutem Regenerate nachkontrollieren: Paare, bei denen Revit
         die Reihenfolge nicht übernimmt, werden gemeldet.
    """
    ergebnis = Ergebnis()
    prios = [prioritaet(e) for e in elemente]
    kategorien = [kategorie_schluessel(e) for e in elemente]
    paare = lg.kandidaten_paare([_box(e) for e in elemente])
    zu_pruefen = []            # (i, j) für Durchgang 2
    del REVIT_MELDUNGEN[:]
    log = ergebnis.protokoll
    log.append(u"JoinMultiple-Protokoll %s" % time.strftime("%Y-%m-%d %H:%M"))
    log.append(u"Elemente: %d, Kandidatenpaare: %d, nur bei Schnitt: %s, "
               u"gleiche Kategorie: %s, bestehende anpassen: %s" % (
                   len(elemente), len(paare), nur_bei_schnitt,
                   gleiche_kategorie, bestehende_anpassen))
    log.append(u"")

    transaktion = _starte_transaktion(doc, t(u"pyMLG Elemente verbinden", u"pyMLG Join elements", u"pyMLG Unir elementos"))
    try:
        # 1. Verbinden
        for i, j in paare:
            if not lg.paar_erlaubt(kategorien[i], kategorien[j],
                                   gleiche_kategorie):
                continue
            a, b = elemente[i], elemente[j]
            try:
                verbunden = JoinGeometryUtils.AreElementsJoined(doc, a, b)
                schneiden = (_schneiden_sich(a, b)
                             if not verbunden and nur_bei_schnitt else None)
                aktion = lg.verbinden_aktion(verbunden, schneiden,
                                             nur_bei_schnitt)
                if aktion == lg.NEU:
                    JoinGeometryUtils.JoinGeometry(doc, a, b)
                    ergebnis.bearbeitet.update((id_wert(a.Id),
                                                id_wert(b.Id)))
            except Exception as ausnahme:
                ergebnis.fehlschlag(a, b, ausnahme)
                continue
            ergebnis.anzahl[aktion] += 1
            log.append(u"[1] %s | Prio %s / %s | %s | danach: %s" % (
                ergebnis._paar(a, b), prios[i], prios[j], aktion,
                _schneidet(doc, a, b)))
            if lg.reihenfolge_pruefen(aktion, bestehende_anpassen):
                zu_pruefen.append((i, j))
            elif aktion == lg.BEREITS:
                ergebnis.anzahl[lg.NICHT_GEPRUEFT] += 1

        # 2. Reihenfolge nach Priorität
        doc.Regenerate()
        getauscht = []
        for i, j in zu_pruefen:
            reihenfolge = lg.oben_unten(prios[i], prios[j])
            if reihenfolge is None:
                ergebnis.anzahl[lg.GLEICH] += 1
                continue
            paar = (elemente[i], elemente[j])
            oben, unten = paar[reihenfolge[0]], paar[reihenfolge[1]]
            try:
                vorher = _schneidet(doc, oben, unten)
                verbunden = JoinGeometryUtils.AreElementsJoined(doc, oben,
                                                                unten)
                if verbunden and not JoinGeometryUtils.IsCuttingElementInJoin(
                        doc, oben, unten):
                    JoinGeometryUtils.SwitchJoinOrder(doc, oben, unten)
                    getauscht.append((oben, unten))
                    schritt = u"getauscht"
                else:
                    ergebnis.anzahl[lg.RICHTIG] += 1
                    schritt = u"nicht getauscht"
                log.append(u"[2] soll: %s schneidet %s | vorher: %s | %s | "
                           u"danach: %s" % (id_wert(oben.Id), id_wert(unten.Id),
                                            vorher, schritt,
                                            _schneidet(doc, oben, unten)))
            except Exception as ausnahme:
                ergebnis.fehlschlag(oben, unten, ausnahme)

        # 3. Nachkontrolle
        if getauscht:
            doc.Regenerate()
        for oben, unten in getauscht:
            try:
                ok = JoinGeometryUtils.IsCuttingElementInJoin(doc, oben,
                                                              unten)
            except Exception:
                ok = False
            if ok:
                ergebnis.anzahl[lg.UMGEDREHT] += 1
                ergebnis.bearbeitet.update((id_wert(oben.Id),
                                            id_wert(unten.Id)))
            else:
                ergebnis.anzahl[lg.ABGELEHNT] += 1
                ergebnis.abgelehnt.append(ergebnis._paar(oben, unten))
        status = transaktion.Commit()
    except Exception:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        raise

    # 4. Kontrolle nach dem Speichern: Hier sieht man, ob die
    # Fehlerbehandlung beim Commit Verbindungen wieder verändert hat
    log.append(u"")
    log.append(u"Commit: %s" % status)
    for meldung in REVIT_MELDUNGEN:
        log.append(u"Revit: " + meldung)
    for i, j in zu_pruefen:
        reihenfolge = lg.oben_unten(prios[i], prios[j])
        if reihenfolge is None:
            continue
        paar = (elemente[i], elemente[j])
        oben, unten = paar[reihenfolge[0]], paar[reihenfolge[1]]
        ist = _schneidet(doc, oben, unten)
        log.append(u"[4] soll: %s schneidet %s | nach Commit: %s" % (
            id_wert(oben.Id), id_wert(unten.Id), ist))
        if ist != u"%s schneidet" % id_wert(oben.Id):
            ergebnis.falsch_nach_speichern.append(
                u"%s: %s" % (ergebnis._paar(oben, unten), ist))
    schreibe_protokoll(log)
    return ergebnis


def loese(doc, elemente, gleiche_kategorie=True):
    """Verbindungen zwischen Elementen des Bereichs aufheben. Verbindungen
    zu Elementen außerhalb (andere Kategorien, nicht ausgewählt) bleiben."""
    ergebnis = Ergebnis()
    index = dict((id_wert(e.Id), i) for i, e in enumerate(elemente))
    kategorien = [kategorie_schluessel(e) for e in elemente]

    transaktion = _starte_transaktion(doc, t(u"pyMLG Verbindungen lösen", u"pyMLG Unjoin elements", u"pyMLG Desunir elementos"))
    try:
        for i, a in enumerate(elemente):
            try:
                partner = list(JoinGeometryUtils.GetJoinedElements(doc, a))
            except Exception:
                continue
            for partner_id in partner:
                j = index.get(id_wert(partner_id))
                if j is None or j <= i or not lg.paar_erlaubt(
                        kategorien[i], kategorien[j], gleiche_kategorie):
                    continue
                b = elemente[j]
                try:
                    JoinGeometryUtils.UnjoinGeometry(doc, a, b)
                except Exception as ausnahme:
                    ergebnis.fehlschlag(a, b, ausnahme)
                    continue
                ergebnis.geloest += 1
                ergebnis.bearbeitet.update((id_wert(a.Id), id_wert(b.Id)))
        transaktion.Commit()
    except Exception:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        raise
    return ergebnis


def waehle_aus(uidoc, elemente, id_werte):
    """Die Elemente mit den genannten Id-Werten in Revit auswählen."""
    uidoc.Selection.SetElementIds(net_liste(
        ElementId, [e.Id for e in elemente if id_wert(e.Id) in id_werte]))
