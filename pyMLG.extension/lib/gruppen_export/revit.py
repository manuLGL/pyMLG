# -*- coding: utf-8 -*-
"""Revit-Zugriffe von GroupExport und GroupToLink.

GroupExport: Die Revit-API kennt kein "Gruppe speichern". Nachgebaut wie
in der Oberfläche: je Gruppe ein neues, leeres Projekt (metrisch, ohne
Vorlage - eine Vorlage brächte gleichnamige Typen mit, und Revit würde deren
Eigenschaften statt der Quelltypen verwenden), eine Instanz hineinkopieren,
in Grundstellung bringen (Ursprung auf 0/0/0, ohne Drehung), die Gruppe
auflösen (verschachtelte Gruppen bleiben) und als .rvt speichern. Das aktive
Projekt wird dabei nicht verändert.

GroupToLink: Jede Instanz einer Gruppe durch eine Verknüpfung der
gleichnamigen .rvt ersetzen - Ursprung auf Ursprung, dann gedreht und
verschoben wie die Gruppe, Bearbeitungsbereich der Gruppe übernommen.
Alles in einer Transaktion - ein Rückgängig-Schritt.
"""

import uuid

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInCategory,
    BuiltInParameter,
    CopyPasteOptions,
    DuplicateTypeAction,
    ElementId,
    ElementTransformUtils,
    FailureProcessingResult,
    FailureSeverity,
    FilteredElementCollector,
    Group,
    IDuplicateTypeNamesHandler,
    IFailuresPreprocessor,
    Line,
    ModelPathUtils,
    RevitLinkInstance,
    RevitLinkOptions,
    RevitLinkType,
    SaveAsOptions,
    SubTransaction,
    Transaction,
    Transform,
    UnitSystem,
    XYZ,
)
from System.Collections.Generic import List  # noqa: E402

from mlg_sprache import t  # noqa: E402

_klassen = {}


def id_wert(element_id):
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def _ist_modellgruppe(element):
    return (isinstance(element, Group) and element.Category is not None
            and id_wert(element.Category.Id)
            == int(BuiltInCategory.OST_IOSModelGroups))


def _je_typ(gruppen):
    """{Gruppenname: eine Instanz} - je Typ wird nur einmal exportiert."""
    ergebnis = {}
    for gruppe in gruppen:
        name = gruppe.GroupType.Name
        if name not in ergebnis:
            ergebnis[name] = gruppe
    return ergebnis


def gewaehlte_gruppen(doc, ids):
    """Modellgruppen der aktuellen Auswahl, je Typ eine Instanz."""
    gruppen = [doc.GetElement(i) for i in ids]
    return _je_typ(g for g in gruppen if _ist_modellgruppe(g))


def alle_gruppen(doc):
    """Alle platzierten Modellgruppen, je Typ eine Instanz."""
    gruppen = (FilteredElementCollector(doc)
               .OfCategory(BuiltInCategory.OST_IOSModelGroups)
               .WhereElementIsNotElementType())
    return _je_typ(g for g in gruppen if _ist_modellgruppe(g))


# ------------------------------------------------------------ Hilfsklassen

def _vorverarbeiter(fehler):
    """Warnungen im neuen Projekt still bestätigen, Fehler sammeln und
    zurückrollen."""
    if "fehler" not in _klassen:
        namensraum = "pyMLG.GruppenExport_%s" % uuid.uuid4().hex

        class Vorverarbeiter(IFailuresPreprocessor):
            __namespace__ = namensraum

            def PreprocessFailures(self, accessor):
                abbruch = False
                for meldung in list(accessor.GetFailureMessages()):
                    if meldung.GetSeverity() == FailureSeverity.Warning:
                        accessor.DeleteWarning(meldung)
                    else:
                        abbruch = True
                        try:
                            _klassen["sammlung"].append(
                                meldung.GetDescriptionText())
                        except Exception:
                            pass
                if abbruch:
                    return FailureProcessingResult.ProceedWithRollBack
                return FailureProcessingResult.Continue

        _klassen["fehler"] = Vorverarbeiter
    _klassen["sammlung"] = fehler
    return _klassen["fehler"]()


def _duplikat_handler():
    """Gleichnamige Typen im leeren Projekt: dessen Typen verwenden statt
    nachzufragen."""
    if "duplikate" not in _klassen:
        namensraum = "pyMLG.GruppenExportDuplikate_%s" % uuid.uuid4().hex

        class Duplikate(IDuplicateTypeNamesHandler):
            __namespace__ = namensraum

            def OnDuplicateTypeNamesFound(self, args):
                return DuplicateTypeAction.UseDestinationTypes

        _klassen["duplikate"] = Duplikate
    return _klassen["duplikate"]()


def _optionen():
    optionen = CopyPasteOptions()
    try:
        optionen.SetDuplicateTypeNamesHandler(_duplikat_handler())
    except Exception:
        pass
    return optionen


# ------------------------------------------------------------------ Export

class ExportFehler(Exception):
    """Eine Gruppe konnte nicht exportiert werden."""


def ursprung(gruppe):
    """(Ursprungspunkt, Drehung um Z in Radiant) einer Gruppeninstanz."""
    ort = gruppe.Location
    punkt = getattr(ort, "Point", None) or XYZ.Zero
    try:
        drehung = ort.Rotation
    except Exception:
        drehung = 0.0
    return punkt, drehung


def _verschiebung(gruppe):
    """Gruppe in Grundstellung: Ursprung auf 0/0/0, ohne Drehung - so lässt
    sich die Datei an jeder Instanz als Verknüpfung einsetzen (GroupToLink)."""
    punkt, drehung = ursprung(gruppe)
    return Transform.CreateRotationAtPoint(
        XYZ.BasisZ, -drehung, XYZ.Zero).Multiply(
        Transform.CreateTranslation(punkt.Negate()))


def exportiere(app, doc, gruppe, pfad):
    """Gruppe als eigenes Modell unter pfad speichern (überschreibt)."""
    neu = app.NewProjectDocument(UnitSystem.Metric)
    try:
        fehler = []
        tr = Transaction(neu, t(u"Gruppe übernehmen", u"Take over group",
                                u"Tomar grupo"))
        optionen = tr.GetFailureHandlingOptions()
        optionen.SetFailuresPreprocessor(_vorverarbeiter(fehler))
        optionen.SetClearAfterRollback(True)
        tr.SetFailureHandlingOptions(optionen)
        tr.Start()
        try:
            ids = ElementTransformUtils.CopyElements(
                doc, List[ElementId]([gruppe.Id]), neu,
                _verschiebung(gruppe), _optionen())
            for neue_id in ids:
                kopie = neu.GetElement(neue_id)
                if isinstance(kopie, Group):
                    kopie.UngroupMembers()
        except Exception:
            tr.RollBack()
            raise
        tr.Commit()
        if fehler:
            raise ExportFehler(u"; ".join(fehler))
        if not ids or ids.Count == 0:
            raise ExportFehler(t(u"nichts kopiert", u"nothing copied",
                                 u"no se ha copiado nada"))
        speichern = SaveAsOptions()
        speichern.OverwriteExistingFile = True
        neu.SaveAs(pfad, speichern)
    finally:
        neu.Close(False)


# ------------------------------------------------------- Gruppe -> Link

def instanzen(doc):
    """({Gruppenname: [Instanzen]}, {Gruppenname: Anzahl verschachtelt}) -
    nur oberste Instanzen lassen sich ersetzen, verschachtelte gehören zu
    einer anderen Gruppe."""
    oben, verschachtelt = {}, {}
    gruppen = (FilteredElementCollector(doc)
               .OfCategory(BuiltInCategory.OST_IOSModelGroups)
               .WhereElementIsNotElementType())
    for gruppe in gruppen:
        if not _ist_modellgruppe(gruppe):
            continue
        name = gruppe.GroupType.Name
        if gruppe.GroupId != ElementId.InvalidElementId:
            verschachtelt[name] = verschachtelt.get(name, 0) + 1
        else:
            oben.setdefault(name, []).append(gruppe)
    return oben, verschachtelt


def _link_typ(doc, pfad):
    """Vorhandener Linktyp mit diesem Dateinamen oder ein neu geladener."""
    datei = pfad.replace(u"\\", u"/").split(u"/")[-1].lower()
    for typ in FilteredElementCollector(doc).OfClass(RevitLinkType):
        if typ.Name.lower() == datei:
            return typ.Id
    ergebnis = RevitLinkType.Create(
        doc, ModelPathUtils.ConvertUserVisiblePathToModelPath(pfad),
        RevitLinkOptions(False))
    if ergebnis.ElementId == ElementId.InvalidElementId:
        raise ExportFehler(t(u"Datei konnte nicht verknüpft werden (%s)",
                             u"file could not be linked (%s)",
                             u"no se pudo vincular el archivo (%s)")
                           % ergebnis.LoadResult)
    return ergebnis.ElementId


def _bearbeitungsbereich(gruppe, link):
    if not gruppe.Document.IsWorkshared:
        return
    try:
        ziel = link.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
        if ziel is not None and not ziel.IsReadOnly:
            ziel.Set(gruppe.WorksetId.IntegerValue)
    except Exception:
        pass


def _setze_link(doc, typ_id, gruppe):
    """Verknüpfung Ursprung auf Ursprung, dann wie die Gruppe drehen und
    verschieben."""
    punkt, drehung = ursprung(gruppe)
    link = RevitLinkInstance.Create(doc, typ_id)
    if abs(drehung) > 1e-9:
        ElementTransformUtils.RotateElement(
            doc, link.Id, Line.CreateBound(XYZ.Zero, XYZ.BasisZ), drehung)
    if not punkt.IsZeroLength():
        ElementTransformUtils.MoveElement(doc, link.Id, punkt)
    _bearbeitungsbereich(gruppe, link)
    return link


def ersetze(doc, auftraege):
    """auftraege: [(Gruppenname, Pfad, [Instanzen])].

    Rückgabe ({Gruppenname: ersetzte Instanzen}, [Fehlertext]). Eine
    fehlerhafte Gruppe wird ausgelassen (Untertransaktion), die anderen
    laufen weiter; alles bildet einen Rückgängig-Schritt."""
    fertig, fehler, meldungen = {}, [], []
    tr = Transaction(doc, t(u"Gruppen durch Verknüpfungen ersetzen",
                            u"Replace groups with links",
                            u"Reemplazar grupos por vínculos"))
    optionen = tr.GetFailureHandlingOptions()
    optionen.SetFailuresPreprocessor(_vorverarbeiter(meldungen))
    optionen.SetClearAfterRollback(True)
    tr.SetFailureHandlingOptions(optionen)
    tr.Start()
    try:
        for name, pfad, gruppen in auftraege:
            st = SubTransaction(doc)
            st.Start()
            try:
                typ_id = _link_typ(doc, pfad)
                for gruppe in gruppen:
                    _setze_link(doc, typ_id, gruppe)
                doc.Delete(List[ElementId]([g.Id for g in gruppen]))
                st.Commit()
                fertig[name] = len(gruppen)
            except Exception as ausnahme:
                if st.HasStarted() and not st.HasEnded():
                    st.RollBack()
                fehler.append(u"%s: %s" % (name, ausnahme))
    except Exception:
        tr.RollBack()
        raise
    tr.Commit()
    if meldungen:
        # Revit hat beim Abschluss abgelehnt - alles zurückgerollt
        return {}, fehler + meldungen
    return fertig, fehler
