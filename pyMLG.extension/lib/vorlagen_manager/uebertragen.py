# -*- coding: utf-8 -*-
"""Einzelne Einstellungen von einer Ansichtsvorlage auf mehrere übertragen.

Das ist der Kern des Werkzeugs: nicht "die ganze Vorlage anwenden", sondern
genau die Zeilen abhaken, die wandern sollen - ein Parameter, eine einzelne
Kategorie, ein einzelner Filter, ein Bearbeitungsbereich.

Jede übertragbare Sache ist ein Posten. Ein Posten weiß,

    * wie der Wert in der Quelle aussieht (Text für die Liste),
    * wie er in einer Zielvorlage aussieht (für "nur Unterschiede"),
    * wie er geschrieben wird.

Geschrieben wird direkt - Wert lesen, Wert setzen. Nicht über
ApplyViewTemplateParameters(): dessen Doku lässt offen, welche Parameter
wirklich übernommen werden, und es überträgt immer ganze Parameter, nie
eine einzelne Kategorie. Nur für die Einstellungen, die die API weder lesen
noch schreiben lässt (Modelldarstellung, Schatten, Skizzenlinien,
Beleuchtung, Fotobelichtung), bleibt es der einzige Weg - diese Posten
sammelt UndurchsichtigerPosten ein und wendet sie gemeinsam an.
"""

from Autodesk.Revit.DB import (
    ElementId,
    FilteredElementCollector,
    FilteredWorksetCollector,
    ParameterFilterElement,
    StorageType,
    ViewPlan,
    WorksetId,
    WorksetKind,
)

from filter_manager.aufloesung import id_liste, id_wert
from mlg_sprache import t
from vorlagen_manager import bloecke as bl
from vorlagen_manager import logik as lg
from vorlagen_manager import revit as rv

GLEICH = t(u"gleich", u"same", u"igual")
VERSCHIEDEN = t(u"verschieden", u"varies", u"varía")
UNBEKANNT = u"?"


class Posten(object):
    """Eine einzeln übertragbare Einstellung."""

    def __init__(self, gruppe, name, quelltext):
        self.gruppe = gruppe
        self.name = name
        self.quelltext = quelltext
        self.undurchsichtig = False

    @property
    def anzeigename(self):
        return u"%s › %s" % (self.gruppe, self.name) if self.gruppe \
            else self.name

    def suchtext(self):
        return (u"%s %s" % (self.anzeigename, self.quelltext)).lower()

    # -- von den Unterklassen zu füllen ---------------------------------

    def quellwert(self):
        """Vergleichbarer Wert in der Quelle."""
        raise NotImplementedError

    def zielwert(self, vorlage):
        """Vergleichbarer Wert in einer Zielvorlage."""
        raise NotImplementedError

    def zieltext(self, vorlage):
        """Lesbarer Wert in einer Zielvorlage."""
        return u"%s" % self.zielwert(vorlage)

    def schreibe(self, vorlage):
        raise NotImplementedError

    # -- Vergleich ------------------------------------------------------

    def vergleich(self, ziele):
        """(Text, unterschiedlich) über alle Zielvorlagen."""
        try:
            quelle = self.quellwert()
            werte = [self.zielwert(z) for z in ziele]
        except Exception:
            return UNBEKANNT, True
        einheitlich = lg.vereine(werte)
        if einheitlich is lg.VERSCHIEDEN:
            return VERSCHIEDEN, True
        if einheitlich == quelle:
            return GLEICH, False
        try:
            return self.zieltext(ziele[0]), True
        except Exception:
            return UNBEKANNT, True


# ---------------------------------------------------------------------------
# Parameter
# ---------------------------------------------------------------------------

class ParameterPosten(Posten):
    """Ein gewöhnlicher Parameterwert der Vorlage."""

    def __init__(self, modell, quelle, eintrag):
        Posten.__init__(self, t(u"Parameter", u"Parameter", u"Parámetro"),
                        eintrag.name, eintrag.text or u"–")
        self.modell = modell
        self.quelle = quelle
        self.pid = eintrag.pid

    def _wert(self, vorlage):
        parameter = self.modell.parameter(vorlage, self.pid)
        if parameter is None:
            return None
        art = parameter.StorageType
        if art == StorageType.String:
            return parameter.AsString() or u""
        if art == StorageType.Double:
            return round(parameter.AsDouble(), 9)
        if art == StorageType.ElementId:
            return id_wert(parameter.AsElementId())
        if art == StorageType.Integer:
            return parameter.AsInteger()
        return None

    def quellwert(self):
        return self._wert(self.quelle)

    def zielwert(self, vorlage):
        return self._wert(vorlage)

    def zieltext(self, vorlage):
        parameter = self.modell.parameter(vorlage, self.pid)
        if parameter is None:
            return t(u"nicht vorhanden", u"not present", u"no existe")
        try:
            if parameter.StorageType == StorageType.String:
                return parameter.AsString() or u"–"
            if parameter.StorageType == StorageType.ElementId:
                return rv.elementname(self.modell.doc, parameter.AsElementId())
            return parameter.AsValueString() or u"%s" % self._wert(vorlage)
        except Exception:
            return UNBEKANNT

    def schreibe(self, vorlage):
        quelle = self.modell.parameter(self.quelle, self.pid)
        ziel = self.modell.parameter(vorlage, self.pid)
        if quelle is None or ziel is None or ziel.IsReadOnly:
            return
        art = quelle.StorageType
        if art == StorageType.String:
            ziel.Set(quelle.AsString() or u"")
        elif art == StorageType.Double:
            ziel.Set(quelle.AsDouble())
        elif art == StorageType.Integer:
            ziel.Set(quelle.AsInteger())
        elif art == StorageType.ElementId:
            ziel.Set(quelle.AsElementId())


# ---------------------------------------------------------------------------
# Kategorien
# ---------------------------------------------------------------------------

class KategoriePosten(Posten):
    """Sichtbarkeit und Grafiküberschreibungen einer einzelnen Kategorie."""

    def __init__(self, doc, quelle, gruppe, name, kategorie_id):
        self.doc = doc
        self.quelle = quelle
        self.kid = kategorie_id
        Posten.__init__(self, gruppe, name, self._text(quelle))

    def _zustand(self, vorlage):
        kategorie_id = ElementId(self.kid)
        versteckt = vorlage.GetCategoryHidden(kategorie_id)
        ogs = vorlage.GetCategoryOverrides(kategorie_id)
        return versteckt, ogs

    def _text(self, vorlage):
        try:
            versteckt, ogs = self._zustand(vorlage)
        except Exception:
            return UNBEKANNT
        sicht = (t(u"ausgeblendet", u"hidden", u"oculta") if versteckt
                 else t(u"sichtbar", u"visible", u"visible"))
        kurz = bl.ogs_kurz(self.doc, ogs)
        return sicht if kurz == u"–" else u"%s, %s" % (sicht, kurz)

    def quellwert(self):
        return self._text(self.quelle)

    def zielwert(self, vorlage):
        return self._text(vorlage)

    def zieltext(self, vorlage):
        return self._text(vorlage)

    def schreibe(self, vorlage):
        kategorie_id = ElementId(self.kid)
        versteckt, ogs = self._zustand(self.quelle)
        try:
            vorlage.SetCategoryHidden(kategorie_id, versteckt)
        except Exception:
            pass
        vorlage.SetCategoryOverrides(kategorie_id, ogs)


# ---------------------------------------------------------------------------
# Filter
# ---------------------------------------------------------------------------

class FilterPosten(Posten):
    """Ein Filter: angewendet, sichtbar, Überschreibungen."""

    def __init__(self, doc, quelle, name, filter_id):
        self.doc = doc
        self.quelle = quelle
        self.fid = filter_id
        Posten.__init__(self, t(u"Filter", u"Filter", u"Filtro"), name,
                        self._text(quelle))

    def _text(self, vorlage):
        filter_id = ElementId(self.fid)
        try:
            if not vorlage.IsFilterApplied(filter_id):
                return t(u"nicht angewendet", u"not applied", u"no aplicado")
            sicht = (t(u"sichtbar", u"visible", u"visible")
                     if vorlage.GetFilterVisibility(filter_id)
                     else t(u"unsichtbar", u"hidden", u"oculto"))
            kurz = bl.ogs_kurz(self.doc, vorlage.GetFilterOverrides(filter_id))
            return sicht if kurz == u"–" else u"%s, %s" % (sicht, kurz)
        except Exception:
            return UNBEKANNT

    def quellwert(self):
        return self._text(self.quelle)

    def zielwert(self, vorlage):
        return self._text(vorlage)

    def zieltext(self, vorlage):
        return self._text(vorlage)

    def schreibe(self, vorlage):
        filter_id = ElementId(self.fid)
        if not self.quelle.IsFilterApplied(filter_id):
            if vorlage.IsFilterApplied(filter_id):
                vorlage.RemoveFilter(filter_id)
            return
        if not vorlage.IsFilterApplied(filter_id):
            vorlage.AddFilter(filter_id)
        vorlage.SetFilterVisibility(
            filter_id, self.quelle.GetFilterVisibility(filter_id))
        vorlage.SetFilterOverrides(
            filter_id, self.quelle.GetFilterOverrides(filter_id))


# ---------------------------------------------------------------------------
# Bearbeitungsbereiche
# ---------------------------------------------------------------------------

class WorksetPosten(Posten):

    def __init__(self, quelle, name, workset_nummer):
        self.quelle = quelle
        self.nummer = workset_nummer
        Posten.__init__(self, t(u"Bearbeitungsbereich", u"Workset",
                                u"Subproyecto"), name, self._text(quelle))

    def _sichtbarkeit(self, vorlage):
        return vorlage.GetWorksetVisibility(WorksetId(self.nummer))

    def _text(self, vorlage):
        try:
            wert = self._sichtbarkeit(vorlage)
        except Exception:
            return UNBEKANNT
        for kandidat, beschriftung in bl.WORKSET_WAHL:
            if kandidat == wert:
                return beschriftung
        return u"%s" % wert

    def quellwert(self):
        return self._text(self.quelle)

    def zielwert(self, vorlage):
        return self._text(vorlage)

    def zieltext(self, vorlage):
        return self._text(vorlage)

    def schreibe(self, vorlage):
        vorlage.SetWorksetVisibility(WorksetId(self.nummer),
                                     self._sichtbarkeit(self.quelle))


# ---------------------------------------------------------------------------
# Ansichtsbereich
# ---------------------------------------------------------------------------

class AnsichtsbereichPosten(Posten):
    """Eine Ebene des Ansichtsbereichs (Ebene und Versatz)."""

    def __init__(self, quelle, ebene, beschriftung):
        self.quelle = quelle
        self.ebene = ebene
        Posten.__init__(self, t(u"Ansichtsbereich", u"View range",
                                u"Rango de vista"), beschriftung,
                        self._text(quelle))

    def _werte(self, vorlage):
        bereich = vorlage.GetViewRange()
        return (id_wert(bereich.GetLevelId(self.ebene)),
                round(bereich.GetOffset(self.ebene), 9))

    def _text(self, vorlage):
        try:
            _ebene_id, versatz = self._werte(vorlage)
        except Exception:
            return UNBEKANNT
        return u"%.1f mm" % (versatz * bl.MM_JE_FUSS)

    def quellwert(self):
        try:
            return self._werte(self.quelle)
        except Exception:
            return None

    def zielwert(self, vorlage):
        if not isinstance(vorlage, ViewPlan):
            return None
        try:
            return self._werte(vorlage)
        except Exception:
            return None

    def zieltext(self, vorlage):
        if not isinstance(vorlage, ViewPlan):
            return t(u"kein Grundriss", u"not a plan view", u"no es un plano")
        return self._text(vorlage)

    def schreibe(self, vorlage):
        if not isinstance(vorlage, ViewPlan):
            return
        ebene_id, versatz = self._werte(self.quelle)
        bereich = vorlage.GetViewRange()
        bereich.SetLevelId(self.ebene, ElementId(ebene_id))
        bereich.SetOffset(self.ebene, versatz)
        vorlage.SetViewRange(bereich)


# ---------------------------------------------------------------------------
# Was die API nicht hergibt
# ---------------------------------------------------------------------------

class UndurchsichtigerPosten(Posten):
    """Block, den nur ApplyViewTemplateParameters() übertragen kann.

    Lesen lässt er sich nicht - deshalb bleibt der Vergleich leer und der
    Posten wird gemeinsam mit den anderen seiner Art angewendet.
    """

    def __init__(self, name, pid):
        Posten.__init__(self, t(u"Nur übertragbar", u"Transfer only",
                                u"Solo transferible"), name,
                        t(u"Bearbeiten…", u"Edit…", u"Editar…"))
        self.pid = pid
        self.undurchsichtig = True

    def quellwert(self):
        return None

    def zielwert(self, vorlage):
        return None

    def zieltext(self, vorlage):
        return u"–"

    def vergleich(self, ziele):
        return t(u"nicht lesbar", u"not readable", u"no legible"), True

    def schreibe(self, vorlage):
        pass          # wird gesammelt über uebertrage() angewendet


# ---------------------------------------------------------------------------
# Posten einer Quellvorlage sammeln
# ---------------------------------------------------------------------------

def sammle(modell, quelle, ziele):
    """Alle übertragbaren Posten einer Quellvorlage."""
    doc = modell.doc
    posten = []
    eintraege = modell.eintraege(quelle, mit_bloecken=False)
    gemeinsam = set(eintraege)
    for vorlage in ziele:
        gemeinsam &= set(modell.eintraege(vorlage, mit_bloecken=False))

    for pid in sorted(gemeinsam, key=lambda p: eintraege[p].sortierung()):
        eintrag = eintraege[pid]
        if eintrag.art == rv.ART_BLOCK:
            if bl.editor_fuer(eintrag.bip_name) is None:
                posten.append(UndurchsichtigerPosten(eintrag.name, pid))
            continue
        posten.append(ParameterPosten(modell, quelle, eintrag))

    for bip_name, typ in rv.KATEGORIE_BLOECKE.items():
        gruppe = bl.UEBERSCHRIFTEN.get(bip_name, bip_name)
        for name, kategorie, _einzug in bl.kategorien(doc, quelle, typ):
            posten.append(KategoriePosten(doc, quelle, gruppe, name,
                                          id_wert(kategorie.Id)))

    elemente = list(FilteredElementCollector(doc).OfClass(
        ParameterFilterElement))
    elemente.sort(key=lambda f: lg.natuerlich(f.Name))
    for element in elemente:
        posten.append(FilterPosten(doc, quelle, u"%s" % element.Name,
                                   id_wert(element.Id)))

    if doc.IsWorkshared:
        worksets = list(FilteredWorksetCollector(doc).OfKind(
            WorksetKind.UserWorkset))
        worksets.sort(key=lambda w: lg.natuerlich(w.Name))
        for workset in worksets:
            posten.append(WorksetPosten(quelle, u"%s" % workset.Name,
                                        workset.Id.IntegerValue))

    if isinstance(quelle, ViewPlan):
        for ebene, beschriftung in bl.EBENEN:
            posten.append(AnsichtsbereichPosten(quelle, ebene, beschriftung))

    return posten


# ---------------------------------------------------------------------------
# Auswahlfenster
# ---------------------------------------------------------------------------

def waehle_posten(besitzer, modell, quelle, ziele):
    """Liste aller Posten zum Abhaken. Rückgabe: [Posten] oder None.

    Der Vergleich mit den Zielen liest jede Einstellung in jeder
    Zielvorlage - das kostet bei vielen Kategorien spürbar Zeit und
    passiert deshalb erst auf Knopfdruck.
    """
    alle = sammle(modell, quelle, ziele)
    fenster = bl.Blockfenster(
        besitzer,
        t(u"Einstellungen übertragen", u"Transfer settings",
          u"Transferir ajustes"),
        t(u"Von \"%s\" auf %d Vorlage(n). Abhaken, was übertragen werden "
          u"soll - alles andere bleibt in den Zielvorlagen unangetastet. "
          u"\"Unterschiede suchen\" vergleicht die Ziele und hakt an, was "
          u"abweicht.",
          u"From \"%s\" to %d template(s). Tick what should be transferred - "
          u"everything else stays untouched in the targets. \"Find "
          u"differences\" compares the targets and ticks what differs.",
          u"De \"%s\" a %d plantilla(s). Marque lo que se debe transferir; "
          u"el resto queda intacto en los destinos. \"Buscar diferencias\" "
          u"compara los destinos y marca lo que difiere.")
        % (quelle.Name, len(ziele)))
    fenster.zaehlwort = t(u"angehakt", u"ticked", u"marcados")

    verglichen = [False]

    for eintrag in alle:
        gitter = bl._gitter("auto", "1.5", "1.2", "1")
        haken = bl._kaestchen(0)
        gitter.Children.Add(haken)
        gitter.Children.Add(bl._text(eintrag.anzeigename, 1))
        gitter.Children.Add(bl._text(eintrag.quelltext, 2))
        vergleich = bl._text(u"", 3, grau=True)
        gitter.Children.Add(vergleich)
        zeile = fenster.zeile(gitter, eintrag.suchtext(), eintrag)
        zeile.haken = haken
        zeile.vergleich = vergleich
        zeile.abweichend = None        # noch nicht verglichen

        def bei_haken(sender, args, zeile=zeile):
            zeile.geaendert = bool(sender.IsChecked)
            fenster.zaehle()

        haken.Click += fenster._h(bei_haken)

    nur_abweichende = fenster.schalter(
        t(u"nur Abweichungen", u"differences only", u"solo diferencias"), None)

    def zusatzfilter():
        if not nur_abweichende.IsChecked:
            return None
        return lambda z: z.abweichend is not False

    def bei_schalter(sender, args):
        if nur_abweichende.IsChecked and not verglichen[0]:
            vergleiche()
        fenster.setze_zusatzfilter(zusatzfilter())

    nur_abweichende.Click += fenster._h(bei_schalter)

    def vergleiche():
        fenster.c("zaehler").Text = t(u"Vergleiche…", u"Comparing…",
                                      u"Comparando…")
        for zeile in fenster.zeilen:
            text, abweichend = zeile.schluessel.vergleich(ziele)
            zeile.vergleich.Text = text
            zeile.vergleich.Foreground = bl.BLAU if abweichend else bl.GRAU
            zeile.abweichend = abweichend
        verglichen[0] = True
        fenster.zaehle()

    def bei_unterschiede(sender, args):
        vergleiche()
        for zeile in fenster.sichtbare():
            if zeile.abweichend:
                zeile.haken.IsChecked = True
                zeile.geaendert = True
        fenster.setze_zusatzfilter(zusatzfilter())

    def markiere(an):
        for zeile in fenster.sichtbare():
            zeile.haken.IsChecked = an
            zeile.geaendert = an
        fenster.zaehle()

    fenster.knopf(t(u"Alle anhaken", u"Tick all", u"Marcar todo"),
                  lambda s, a: markiere(True))
    fenster.knopf(t(u"Keine", u"None", u"Ninguno"),
                  lambda s, a: markiere(False))
    fenster.knopf(t(u"Unterschiede suchen", u"Find differences",
                    u"Buscar diferencias"), bei_unterschiede,
                  t(u"Vergleicht jede Zeile mit allen Zielvorlagen und hakt "
                    u"die Abweichungen an.",
                    u"Compares every row with all target templates and ticks "
                    u"the differences.",
                    u"Compara cada fila con todas las plantillas de destino y "
                    u"marca las diferencias."))

    if not fenster.zeige():
        return None
    gewaehlt = [z.schluessel for z in fenster.zeilen if z.geaendert]
    return gewaehlt or None


# ---------------------------------------------------------------------------
# Übertragen
# ---------------------------------------------------------------------------

def uebertrage(modell, quelle, ziele, posten):
    """Posten von der Quelle auf alle Ziele schreiben.

    Rückgabe: (Anzahl geschriebener Posten, [(Vorlage, Posten, Grund)])
    """
    fehler = []
    geschrieben = 0
    einzeln = [p for p in posten if not p.undurchsichtig]
    undurchsichtig = [p for p in posten if p.undurchsichtig]

    for vorlage in ziele:
        for eintrag in einzeln:
            try:
                eintrag.schreibe(vorlage)
                geschrieben += 1
            except Exception as ausnahme:
                fehler.append((vorlage.Name, eintrag.anzeigename,
                               rv.fehlertext(ausnahme)))

    if undurchsichtig:
        anzahl, weitere = _uebertrage_undurchsichtig(
            quelle, ziele, [p.pid for p in undurchsichtig])
        geschrieben += anzahl * len(undurchsichtig)
        fehler.extend(weitere)
    modell.vergiss()
    return geschrieben, fehler


def _uebertrage_undurchsichtig(quelle, ziele, pids):
    """Nicht lesbare Blöcke über ApplyViewTemplateParameters() übertragen.

    Die "Einschließen"-Haken der Quelle werden kurz auf die gewählten
    Parameter beschränkt und danach wieder hergestellt - auch wenn beim
    Übertragen etwas schiefgeht.
    """
    alle = list(quelle.GetTemplateParameterIds())
    behalten = set(pids)
    vorher = list(quelle.GetNonControlledTemplateParameterIds())
    aussen = [i for i in alle if id_wert(i) not in behalten]
    erfolgreich, fehler = 0, []
    quelle.SetNonControlledTemplateParameterIds(id_liste(aussen))
    try:
        for ziel in ziele:
            try:
                ziel.ApplyViewTemplateParameters(quelle)
                erfolgreich += 1
            except Exception as ausnahme:
                fehler.append((ziel.Name,
                               t(u"nicht lesbare Blöcke",
                                 u"settings that cannot be read",
                                 u"ajustes no legibles"),
                               rv.fehlertext(ausnahme)))
    finally:
        quelle.SetNonControlledTemplateParameterIds(id_liste(vorher))
    return erfolgreich, fehler
