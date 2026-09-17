# -*- coding: utf-8 -*-
"""Löst Ids aus Filterregeln gegen das Revit-Dokument in Klartext auf.

Parameter-Ids in Filterregeln (FilterRule.GetRuleParameter()):

    Wert < 0   eingebauter Parameter. Der Zahlenwert IST der
               BuiltInParameter-Enumwert (z.B. -1002001).
               Name: LabelUtils.GetLabelForBuiltInParameter(ForgeTypeId)
    Wert > 0   Element im Dokument:
               SharedParameterElement  (gemeinsam genutzter Parameter)
               ParameterElement        (Projektparameter)
               GlobalParameter         (nur bei Verknüpfungsregeln)
               Name: ParameterElement.GetDefinition().Name

Sonderfall "Bearbeitungsbereich" (ELEM_PARTITION_PARAM): Revit speichert ihn
als Ganzzahl = WorksetId. Angezeigt und ausgewählt wird der Workset-Name wie
im nativen Dialog.

Zahlen: Die Projekteinheiten runden (z.B. 1,125 m -> "1,12 m"). Lässt sich
der gerundete Text nicht wieder exakt einlesen, wird mit feinerer Genauigkeit
formatiert - sonst würde der gerundete Wert beim Speichern in den Filter
zurückgeschrieben.

Für eingebaute Parameter wird einmalig eine Tabelle Zahlenwert -> Name aus
ParameterUtils.GetAllBuiltInParameters() aufgebaut. Der Schlüssel entsteht
über ElementId(BuiltInParameter).Value - so ist keine Umwandlung
int -> Enum nötig, die unter pythonnet fehleranfällig ist
(siehe schedule_sync.revit_helpers.parameter_map).

Die Spezifikation (Einheit) eines eingebauten Parameters kennt nur ein
Parameter am Element. Sie wird deshalb erst bei Bedarf (Zahlenwerte) über
Element.GetParameter(ForgeTypeId) an wenigen Elementen der
Filterkategorien ermittelt - große Modelle werden nicht durchsucht.
"""

from Autodesk.Revit.DB import (
    BuiltInParameter,
    Category,
    ElementId,
    ElementMulticategoryFilter,
    FilteredElementCollector,
    FilteredWorksetCollector,
    FormatOptions,
    FormatValueOptions,
    LabelUtils,
    ParameterUtils,
    StorageType,
    UnitFormatUtils,
    UnitUtils,
    WorksetKind,
)
from System.Collections.Generic import List

from mlg_sprache import t

QUELLE_EINGEBAUT = t(u"eingebaut", u"built-in", u"integrado")
QUELLE_GEMEINSAM = t(u"gemeinsam genutzt", u"shared", u"compartido")
QUELLE_PROJEKT = t(u"Projektparameter", u"project parameter", u"parámetro de proyecto")
QUELLE_GLOBAL = t(u"globaler Parameter", u"global parameter", u"parámetro global")
QUELLE_UNBEKANNT = t(u"unbekannt", u"unknown", u"desconocido")

# Höchstens so viele Elemente je Kategorienmenge nach der Einheit fragen
MAX_ELEMENTE_SPEC = 200


def id_wert(element_id):
    """Zahlenwert einer ElementId (Revit 2024+: .Value)."""
    if element_id is None:
        return None
    if isinstance(element_id, int):
        return element_id
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def id_liste(ids):
    """Python-Iterable -> List[ElementId] für die Revit-API."""
    liste = List[ElementId]()
    for eid in ids:
        liste.Add(eid if not isinstance(eid, int) else ElementId(eid))
    return liste


BEARBEITUNGSBEREICH = id_wert(ElementId(BuiltInParameter.ELEM_PARTITION_PARAM))

# Toleranz (Fuß), ab der ein formatierter Wert als verlustfrei gilt
TOLERANZ_FORMAT = 1e-9


def speicherart_aus_spec(spec):
    """StorageType aus einer Spezifikation (ForgeTypeId) ableiten."""
    if spec is None:
        return None
    try:
        typ_id = spec.TypeId or u""
    except Exception:
        return None
    if not typ_id:
        return None
    if u"spec.string" in typ_id:
        return StorageType.String
    if u"spec.bool" in typ_id or u"spec.int" in typ_id:
        return StorageType.Integer
    try:
        if UnitUtils.IsMeasurableSpec(spec):
            return StorageType.Double
    except Exception:
        pass
    try:
        if Category.IsBuiltInCategory(spec):
            return StorageType.ElementId     # Verweis, z.B. Material
    except Exception:
        pass
    return None


def ist_ja_nein(spec):
    try:
        return spec is not None and u"spec.bool" in (spec.TypeId or u"")
    except Exception:
        return False


class ParameterInfo(object):
    """Klartextangaben zu einer Parameter-Id."""

    def __init__(self, wert, name, quelle, speicherart=None, spec=None,
                 forge_id=None):
        self.wert = wert                # Zahlenwert der Id
        self.name = name
        self.quelle = quelle
        self.speicherart = speicherart  # StorageType oder None
        self.spec = spec                # ForgeTypeId oder None
        self.forge_id = forge_id        # nur eingebaute Parameter
        self.bearbeitungsbereich = wert == BEARBEITUNGSBEREICH

    def __repr__(self):
        return u"<Parameter %s '%s' (%s)>" % (self.wert, self.name, self.quelle)


class RevitAufloeser(object):
    """Auflöser für regelbaum.zerlege_filter() - mit Zwischenspeicher."""

    def __init__(self, doc):
        self.doc = doc
        self._eingebaut = None          # int -> (Name, ForgeTypeId)
        self._infos = {}                # int -> ParameterInfo
        self._spec_gesucht = set()      # (Parameter, Kategorien) schon gesucht
        self._kontext = []              # Kategorie-Ids des aktuellen Filters
        self._kategorien = {}           # int -> Name
        self._worksets = None           # int -> Name

    # -- Kontext ------------------------------------------------------------

    def setze_kontext(self, kategorie_ids):
        """Kategorien des gerade bearbeiteten Filters (für die Einheit)."""
        self._kontext = list(kategorie_ids or [])

    def id_wert(self, element_id):
        return id_wert(element_id)

    # -- Parameter ----------------------------------------------------------

    def _tabelle_eingebaut(self):
        if self._eingebaut is None:
            tabelle = {}
            for forge_id in ParameterUtils.GetAllBuiltInParameters():
                try:
                    bip = ParameterUtils.GetBuiltInParameter(forge_id)
                    wert = id_wert(ElementId(bip))
                    name = LabelUtils.GetLabelForBuiltInParameter(forge_id)
                except Exception:
                    continue
                tabelle[wert] = (name, forge_id)
            self._eingebaut = tabelle
        return self._eingebaut

    def info(self, param_id):
        """ParameterInfo zu einer ElementId oder einem Zahlenwert.

        Bei eingebauten Parametern ist .spec zunächst None - dafür
        info_mit_spec() verwenden (sucht an Elementen, nur wenn nötig).
        """
        wert = id_wert(param_id)
        info = self._infos.get(wert)
        if info is None:
            info = self._ermittle(wert)
            self._infos[wert] = info
        return info

    def info_mit_spec(self, param_id):
        """Wie info(), ermittelt aber bei Zahlen-/Ganzzahlparametern auch die
        Spezifikation (Einheit bzw. Ja/Nein)."""
        info = self.info(param_id)
        if (info.spec is None and info.quelle == QUELLE_EINGEBAUT
                and info.speicherart in (StorageType.Double,
                                         StorageType.Integer)):
            self._suche_spec(info)
        return info

    def _ermittle(self, wert):
        if wert is None or wert == -1:
            return ParameterInfo(wert, None, QUELLE_UNBEKANNT)

        if wert < 0:
            treffer = self._tabelle_eingebaut().get(wert)
            if treffer is None:
                return ParameterInfo(
                    wert, t(u"<BuiltInParameter %s nicht aufgelöst>", u"<BuiltInParameter %s not resolved>", u"<BuiltInParameter %s no resuelto>") % wert,
                    QUELLE_UNBEKANNT)
            name, forge_id = treffer
            try:
                speicherart = self.doc.GetTypeOfStorage(forge_id)
            except Exception:
                speicherart = None
            return ParameterInfo(wert, name, QUELLE_EINGEBAUT, speicherart,
                                 None, forge_id)

        element = self.doc.GetElement(ElementId(wert))
        if element is None:
            return ParameterInfo(
                wert, t(u"<Parameter-Id %s nicht gefunden>", u"<parameter id %s not found>", u"<id de parámetro %s no encontrado>") % wert,
                QUELLE_UNBEKANNT)
        typ = element.GetType().Name
        quelle = {"GlobalParameter": QUELLE_GLOBAL,
                  "SharedParameterElement": QUELLE_GEMEINSAM,
                  "ParameterElement": QUELLE_PROJEKT}.get(typ, QUELLE_PROJEKT)
        name, spec = None, None
        try:
            definition = element.GetDefinition()
            name = definition.Name
            spec = definition.GetDataType()
        except Exception:
            pass
        if not name:
            try:
                name = element.Name
            except Exception:
                name = u"<Element %s>" % wert
        return ParameterInfo(wert, name, quelle, speicherart_aus_spec(spec),
                             spec)

    def _suche_spec(self, info):
        """Einheit eines eingebauten Parameters an einem Element ablesen."""
        if not self._kontext or info.forge_id is None:
            return
        schluessel = (info.wert, tuple(sorted(id_wert(k)
                                              for k in self._kontext)))
        if schluessel in self._spec_gesucht:
            return
        self._spec_gesucht.add(schluessel)
        kategorienfilter = ElementMulticategoryFilter(id_liste(self._kontext))
        for nur_typen in (False, True):
            sammler = FilteredElementCollector(self.doc).WherePasses(
                kategorienfilter)
            sammler = (sammler.WhereElementIsElementType() if nur_typen
                       else sammler.WhereElementIsNotElementType())
            for anzahl, element in enumerate(sammler):
                if anzahl >= MAX_ELEMENTE_SPEC:
                    break
                try:
                    param = element.GetParameter(info.forge_id)
                    if param is not None:
                        info.spec = param.Definition.GetDataType()
                        return
                except Exception:
                    continue

    def parametername(self, param_id):
        return self.info(param_id).name


    # -- Kategorien und Elemente -------------------------------------------

    def kategoriename(self, kategorie_id):
        wert = id_wert(kategorie_id)
        if wert not in self._kategorien:
            try:
                kategorie = Category.GetCategory(self.doc, ElementId(wert))
                name = kategorie.Name if kategorie is not None else None
            except Exception:
                name = None
            self._kategorien[wert] = name or t(u"<Kategorie %s>", u"<category %s>", u"<categoría %s>") % wert
        return self._kategorien[wert]

    def kategorienamen(self, kategorie_ids):
        return sorted((self.kategoriename(k) for k in kategorie_ids or []),
                      key=lambda n: n.lower())

    def elementname(self, element_id):
        wert = id_wert(element_id)
        if wert is None or wert == -1:
            return t(u"(keine)", u"(none)", u"(ninguno)")
        element = self.doc.GetElement(ElementId(wert))
        if element is None:
            return u"<Id %s>" % wert
        try:
            return u"%s" % element.Name
        except Exception:
            return u"<Id %s>" % wert

    # -- Bearbeitungsbereiche ----------------------------------------------

    def worksets(self):
        """{WorksetId-Zahl: Name} aller Worksets (leer ohne Teamarbeit)."""
        if self._worksets is None:
            self._worksets = {}
            try:
                if self.doc.IsWorkshared:
                    for ws in FilteredWorksetCollector(self.doc):
                        self._worksets[int(ws.Id.IntegerValue)] = ws.Name
            except Exception:
                pass
        return self._worksets

    def benutzer_worksets(self):
        """[(Name, WorksetId-Zahl)] der Benutzer-Worksets, sortiert."""
        eintraege = []
        try:
            if self.doc.IsWorkshared:
                for ws in FilteredWorksetCollector(self.doc).OfKind(
                        WorksetKind.UserWorkset):
                    eintraege.append((ws.Name, int(ws.Id.IntegerValue)))
        except Exception:
            pass
        return sorted(eintraege, key=lambda e: e[0].lower())

    def workset_name(self, wert):
        name = self.worksets().get(int(wert))
        return name if name is not None else \
            t(u"<Bearbeitungsbereich %d nicht vorhanden>", u"<workset %d does not exist>", u"<subproyecto %d no existe>") % int(wert)

    # -- Werte --------------------------------------------------------------

    def _verlustfrei(self, units, spec, text, wert):
        try:
            ergebnis = UnitFormatUtils.TryParse(units, spec, text)
            if isinstance(ergebnis, tuple) and ergebnis[0]:
                return abs(float(ergebnis[1]) - wert) <= TOLERANZ_FORMAT
        except Exception:
            pass
        return False

    def zahl(self, param_id, wert, zum_bearbeiten=False):
        """Double-Regelwert (interne Einheit) in Projekteinheiten - bei
        Bedarf genauer als die Projekteinheiten, damit nichts verloren geht."""
        spec = self.info_mit_spec(param_id).spec
        wert = float(wert)
        try:
            if spec is not None and UnitUtils.IsMeasurableSpec(spec):
                units = self.doc.GetUnits()
                text = UnitFormatUtils.Format(units, spec, wert,
                                              zum_bearbeiten)
                if self._verlustfrei(units, spec, text, wert):
                    return text
                genau = self._zahl_genau(units, spec, wert, zum_bearbeiten)
                return genau if genau is not None else text
        except Exception:
            pass
        text = u"%g" % wert
        return text if zum_bearbeiten else text + t(u" (intern)", u" (internal)", u" (interno)")

    def _zahl_genau(self, units, spec, wert, zum_bearbeiten):
        """Mit steigender Nachkommastellenzahl formatieren, bis der Text den
        Wert exakt wiedergibt. None, wenn die Einheit das nicht erlaubt
        (z.B. Fuß und Zoll mit Brüchen)."""
        basis = units.GetFormatOptions(spec)
        einheit = basis.GetUnitTypeId()
        for stellen in range(1, 10):
            genauigkeit = 10.0 ** -stellen
            try:
                if not FormatOptions.IsValidAccuracy(einheit, genauigkeit):
                    continue
                optionen = FormatOptions(basis)
                optionen.UseDefault = False
                optionen.Accuracy = genauigkeit
                if optionen.CanSuppressTrailingZeros():
                    optionen.SuppressTrailingZeros = True
                wertoptionen = FormatValueOptions()
                wertoptionen.SetFormatOptions(optionen)
                text = UnitFormatUtils.Format(units, spec, wert,
                                              zum_bearbeiten, wertoptionen)
            except Exception:
                continue
            if self._verlustfrei(units, spec, text, wert):
                return text
        return None

    def zahl_lesen(self, param_id, text):
        """Eingabe in Projekteinheiten -> interner Wert. ValueError bei Fehler."""
        text = (text or u"").strip()
        if not text:
            raise ValueError(t(u"Es wurde kein Wert eingegeben.", u"No value was entered.", u"No se introdujo ningún valor."))
        spec = self.info_mit_spec(param_id).spec
        if spec is not None:
            try:
                if UnitUtils.IsMeasurableSpec(spec):
                    ergebnis = UnitFormatUtils.TryParse(
                        self.doc.GetUnits(), spec, text)
                    if isinstance(ergebnis, tuple):
                        ok, wert = ergebnis[0], ergebnis[1]
                        if ok:
                            return float(wert)
                        raise ValueError(t(u"'%s' ist kein gültiger Wert.", u"'%s' is not a valid value.", u"'%s' no es un valor válido.")
                                         % text)
            except ValueError:
                raise
            except Exception:
                pass
        try:
            return float(text.replace(u",", u"."))
        except ValueError:
            raise ValueError(t(u"'%s' ist keine Zahl.", u"'%s' is not a number.", u"'%s' no es un número.") % text)

    def ganzzahl(self, param_id, wert):
        info = self.info_mit_spec(param_id)
        if info.bearbeitungsbereich:
            return self.workset_name(wert)
        if ist_ja_nein(info.spec):
            return t(u"Ja", u"Yes", u"Sí") if wert else t(u"Nein", u"No", u"No")
        return u"%d" % wert
