# -*- coding: utf-8 -*-
"""Bearbeitbares Regelmodell und schreibende Filteraktionen.

Der Regeleditor arbeitet wie der native Dialog mit "Sätzen":

    Satz (UND | ODER)
    ├─ Regel  (Parameter, Operator, Wert)
    └─ Satz   (verschachtelt)

Umwandlung in die Revit-Struktur (baue_elementfilter):
    Regel        -> ElementParameterFilter(eine FilterRule)
    Satz UND     -> LogicalAndFilter(Kinder)
    Satz ODER    -> LogicalOrFilter(Kinder)
    Satz mit genau einem Kind -> das Kind selbst

Die Regeln entstehen über ParameterFilterRuleFactory. Welche Operatoren
angeboten werden, hängt von der Speicherart des Parameters ab
(Überladungen der Factory, per RevitAPI.xml 2026 geprüft):

    String      gleich, ungleich, >, >=, <, <=, enthält (nicht),
                beginnt (nicht) mit, endet (nicht) mit, hat (keinen) Wert
    Double      gleich, ungleich, >, >=, <, <=, hat (keinen) Wert
    Integer     wie Double; Ja/Nein-Parameter nur gleich/ungleich/Wert
    ElementId   wie Double (Vergleich über die Id, z.B. Ebenen, Phasen)

Nicht im Editor bearbeitbar (Filter wird dann nur angezeigt):
kategoriebezogene Zweige (FilterCategoryRule), SharedParameterApplicableRule,
Verknüpfungen mit globalen Parametern, invertierte ElementFilter und negierte
Größer/Kleiner-Vergleiche. Diese Konstrukte erzeugt der native Dialog
selbst nicht; sie stammen aus anderen Werkzeugen oder der API.
"""

from Autodesk.Revit.DB import (
    ElementFilter,
    ElementId,
    ElementMulticategoryFilter,
    ElementParameterFilter,
    FilteredElementCollector,
    FilterElement,
    LogicalAndFilter,
    LogicalOrFilter,
    ParameterFilterElement,
    ParameterFilterRuleFactory,
    ParameterFilterUtilities,
    ParameterValueProvider,
    StorageType,
)
from System.Collections.Generic import HashSet, List

from filter_manager import regelbaum as rb
from filter_manager.aufloesung import id_liste, id_wert, ist_ja_nein
from mlg_sprache import t

# In Filternamen verboten (ParameterFilterElement.Create, API 2026)
VERBOTENE_ZEICHEN = u"{}[]|;<>?`~"

# Toleranz für "ist gleich" bei Double-Werten (interne Einheit, Fuß)
EPSILON = 1e-6

# Operator-Schlüssel -> Anzeigetext (Reihenfolge wie im nativen Dialog)
OPERATOR_TEXTE = [
    ("gleich", t(u"ist gleich", u"equals", u"igual a")),
    ("ungleich", t(u"ist ungleich", u"does not equal", u"no es igual a")),
    ("groesser", t(u"ist größer als", u"is greater than", u"es mayor que")),
    ("groesser_gleich", t(u"ist größer als oder gleich", u"is greater than or equal to", u"es mayor o igual que")),
    ("kleiner", t(u"ist kleiner als", u"is less than", u"es menor que")),
    ("kleiner_gleich", t(u"ist kleiner als oder gleich", u"is less than or equal to", u"es menor o igual que")),
    ("enthaelt", t(u"enthält", u"contains", u"contiene")),
    ("enthaelt_nicht", t(u"enthält nicht", u"does not contain", u"no contiene")),
    ("beginnt", t(u"beginnt mit", u"begins with", u"empieza por")),
    ("beginnt_nicht", t(u"beginnt nicht mit", u"does not begin with", u"no empieza por")),
    ("endet", t(u"endet mit", u"ends with", u"termina por")),
    ("endet_nicht", t(u"endet nicht mit", u"does not end with", u"no termina por")),
    ("hat_wert", t(u"hat einen Wert", u"has a value", u"tiene un valor")),
    ("hat_keinen_wert", t(u"hat keinen Wert", u"has no value", u"no tiene valor")),
]
OPERATOR_TEXT = dict(OPERATOR_TEXTE)
OHNE_WERT = ("hat_wert", "hat_keinen_wert")
NUR_TEXT = ("enthaelt", "enthaelt_nicht", "beginnt", "beginnt_nicht",
            "endet", "endet_nicht")
VERGLEICHE = ("groesser", "groesser_gleich", "kleiner", "kleiner_gleich")

# Negation aus dem Parser (FilterInverseRule) -> Editor-Operator
NEGIERT = {"gleich": "ungleich", "enthaelt": "enthaelt_nicht",
           "beginnt": "beginnt_nicht", "endet": "endet_nicht"}

# Editor-Operator -> Name der Factory-Methode
FABRIK = {
    "gleich": "CreateEqualsRule",
    "ungleich": "CreateNotEqualsRule",
    "groesser": "CreateGreaterRule",
    "groesser_gleich": "CreateGreaterOrEqualRule",
    "kleiner": "CreateLessRule",
    "kleiner_gleich": "CreateLessOrEqualRule",
    "enthaelt": "CreateContainsRule",
    "enthaelt_nicht": "CreateNotContainsRule",
    "beginnt": "CreateBeginsWithRule",
    "beginnt_nicht": "CreateNotBeginsWithRule",
    "endet": "CreateEndsWithRule",
    "endet_nicht": "CreateNotEndsWithRule",
    "hat_wert": "CreateHasValueParameterRule",
    "hat_keinen_wert": "CreateHasNoValueParameterRule",
}

# Höchstens so viele Elemente nach Wertvorschlägen durchsuchen
MAX_ELEMENTE_VORSCHLAG = 3000
MAX_VORSCHLAEGE = 300


class NichtBearbeitbar(Exception):
    """Der Filter enthält Konstrukte, die der Editor nicht abbilden kann."""


class FilterFehler(Exception):
    """Fachlicher Fehler mit verständlicher Meldung für den Nutzer."""


def fehlertext(fehler):
    """Kurze, lesbare Meldung aus Python- oder .NET-Ausnahmen."""
    text = getattr(fehler, "Message", None) or u"%s" % fehler
    zeilen = [z.strip() for z in text.splitlines() if z.strip()]
    return zeilen[0] if zeilen else fehler.__class__.__name__


# ---------------------------------------------------------------------------
# Modell
# ---------------------------------------------------------------------------

class Regel(object):
    def __init__(self, parameter=None, operator="gleich", wert=u"",
                 wert_id=None, original=None):
        self.parameter = parameter      # Zahlenwert der Parameter-Id
        self.operator = operator        # Schlüssel aus OPERATOR_TEXTE
        self.wert = wert                # Texteingabe (Projekteinheiten)
        self.wert_id = wert_id          # Zahlenwert einer ElementId
        # Nur Double-Regeln aus Revit: (Parameter, angezeigter Text,
        # interner Wert, Epsilon). Solange Parameter und Text unverändert
        # sind, wird exakt der ursprüngliche Wert zurückgeschrieben.
        self.original = original

    def kopie(self):
        return Regel(self.parameter, self.operator, self.wert, self.wert_id,
                     self.original)


class Satz(object):
    def __init__(self, oder=False, elemente=None):
        self.oder = oder
        self.elemente = list(elemente or [])

    def kopie(self):
        return Satz(self.oder, [e.kopie() for e in self.elemente])

    def ist_leer(self):
        return not any(isinstance(e, Regel) or not e.ist_leer()
                       for e in self.elemente)


def operatoren_fuer(info):
    """Zulässige Operator-Schlüssel für einen Parameter (ParameterInfo)."""
    if info is None:
        return [k for k, _t in OPERATOR_TEXTE]
    art = info.speicherart
    if art == StorageType.String:
        return [k for k, _t in OPERATOR_TEXTE]
    if getattr(info, "bearbeitungsbereich", False):
        # Wie im nativen Dialog: Auswahl eines Worksets
        return ["gleich", "ungleich"]
    if art == StorageType.Integer and ist_ja_nein(info.spec):
        return ["gleich", "ungleich", "hat_wert", "hat_keinen_wert"]
    if art in (StorageType.Integer, StorageType.Double, StorageType.ElementId):
        return [k for k, _t in OPERATOR_TEXTE if k not in NUR_TEXT]
    # Speicherart unbekannt: alles anbieten, Revit prüft beim Speichern
    return [k for k, _t in OPERATOR_TEXTE]


# ---------------------------------------------------------------------------
# Revit-Struktur -> Modell
# ---------------------------------------------------------------------------

def _regel_aus_knoten(knoten, aufloeser):
    if knoten.art != rb.REGEL:
        raise NichtBearbeitbar(u"nicht lesbare Regel (%s)" % knoten.regeltyp)
    if knoten.regeltyp == "FilterCategoryRule":
        raise NichtBearbeitbar(u"kategoriebezogene Regeln")
    if knoten.regeltyp == "SharedParameterApplicableRule":
        raise NichtBearbeitbar(u"Regel \"Parameter ist vorhanden\"")
    if knoten.regeltyp == "FilterGlobalParameterAssociationRule":
        raise NichtBearbeitbar(u"Verknüpfung mit globalem Parameter")
    if knoten.parameter_id is None:
        raise NichtBearbeitbar(u"Regel ohne Parameter (%s)" % knoten.regeltyp)

    operator = knoten.operator
    if operator in OHNE_WERT:
        positiv = (operator == "hat_wert") != knoten.negiert
        return Regel(knoten.parameter_id,
                     "hat_wert" if positiv else "hat_keinen_wert")
    if knoten.negiert:
        if operator not in NEGIERT:
            raise NichtBearbeitbar(u"negierter Vergleich (%s)"
                                   % knoten.text_operator)
        operator = NEGIERT[operator]
    if operator not in FABRIK:
        raise NichtBearbeitbar(u"unbekannter Operator (%s)" % operator)

    regel = Regel(knoten.parameter_id, operator)
    if knoten.regeltyp == "FilterStringRule":
        regel.wert = knoten.rohwert or u""
    elif knoten.regeltyp == "FilterDoubleRule":
        regel.wert = aufloeser.zahl(knoten.parameter_id, knoten.rohwert,
                                    zum_bearbeiten=True)
        regel.original = (knoten.parameter_id, regel.wert, knoten.rohwert,
                          knoten.epsilon)
    elif knoten.regeltyp == "FilterIntegerRule":
        regel.wert = u"%d" % knoten.rohwert
    elif knoten.regeltyp == "FilterElementIdRule":
        regel.wert_id = knoten.rohwert
        regel.wert = aufloeser.elementname(knoten.rohwert)
    return regel


def satz_aus_baum(baum, aufloeser):
    """Knotenbaum des Parsers -> Satz. Wirft NichtBearbeitbar."""
    if baum is None or baum.art == rb.LEER:
        return Satz(oder=False)
    if baum.art in (rb.FEHLER, rb.UNBEKANNT):
        raise NichtBearbeitbar(u"nicht lesbare Filterstruktur")

    def satz(knoten):
        if knoten.invertiert:
            raise NichtBearbeitbar(u"invertierter Filter")
        ergebnis = Satz(oder=knoten.art == rb.ODER)
        for kind in knoten.kinder:
            if kind.art == rb.REGEL or kind.art in (rb.FEHLER, rb.UNBEKANNT):
                ergebnis.elemente.append(_regel_aus_knoten(kind, aufloeser))
            elif kind.art == rb.REGELN:
                if kind.invertiert:
                    raise NichtBearbeitbar(u"invertierter Filter")
                if not ergebnis.oder or len(kind.kinder) == 1:
                    # Unter UND bzw. als Einzelregel: direkt übernehmen
                    for r in kind.kinder:
                        ergebnis.elemente.append(
                            _regel_aus_knoten(r, aufloeser))
                else:
                    ergebnis.elemente.append(satz(kind))
            elif kind.art in (rb.UND, rb.ODER):
                ergebnis.elemente.append(satz(kind))
        return ergebnis

    return satz(baum)


# ---------------------------------------------------------------------------
# Modell -> Revit-Struktur
# ---------------------------------------------------------------------------

def _regel_beschreibung(regel, aufloeser):
    name = aufloeser.parametername(regel.parameter) if regel.parameter \
        else t(u"(kein Parameter)", u"(no parameter)", u"(sin parámetro)")
    return u"%s %s %s" % (name, OPERATOR_TEXT.get(regel.operator, u"?"),
                          regel.wert or u"")


def double_wert(regel, aufloeser):
    """(interner Wert, Epsilon) einer Double-Regel.

    Unveränderte Regeln aus Revit behalten Wert und Toleranz exakt - sonst
    würde jedes Speichern eines Filters die Zahlen auf die Genauigkeit der
    Projekteinheiten runden und die Toleranz ersetzen. Wurde nur der Wert
    geändert, bleibt die ursprüngliche Toleranz erhalten."""
    text = (regel.wert or u"").strip()
    original = regel.original
    if original is not None and original[0] == regel.parameter:
        _param, alter_text, alter_wert, epsilon = original
        if text == (alter_text or u"").strip():
            return alter_wert, epsilon or EPSILON
        return aufloeser.zahl_lesen(regel.parameter, text), epsilon or EPSILON
    return aufloeser.zahl_lesen(regel.parameter, text), EPSILON


def baue_regel(regel, aufloeser):
    """Regel -> FilterRule. Wirft FilterFehler mit verständlicher Meldung."""
    if regel.parameter is None:
        raise FilterFehler(t(u"Eine Regel hat keinen Parameter.", u"A rule has no parameter.", u"Una regla no tiene parámetro."))
    info = aufloeser.info_mit_spec(regel.parameter)
    beschreibung = _regel_beschreibung(regel, aufloeser)
    if regel.operator not in operatoren_fuer(info):
        raise FilterFehler(t(u"Der Operator \"%s\" ist für den Parameter "
                           u"\"%s\" nicht zulässig.", u"The operator \"%s\" is not allowed for the parameter \"%s\".", u"El operador \"%s\" no está permitido para el parámetro \"%s\".")
                           % (OPERATOR_TEXT.get(regel.operator), info.name))
    methode = getattr(ParameterFilterRuleFactory, FABRIK[regel.operator])
    param_id = ElementId(regel.parameter)

    if regel.operator in OHNE_WERT:
        return methode(param_id)

    art = info.speicherart
    text = (regel.wert or u"").strip()
    try:
        if art == StorageType.Double:
            wert, epsilon = double_wert(regel, aufloeser)
            return methode(param_id, wert, epsilon)
        if art == StorageType.Integer:
            if ist_ja_nein(info.spec):
                wert = {u"ja": 1, u"nein": 0, u"1": 1, u"0": 0,
                        t(u"ja", u"yes", u"sí"): 1,
                        t(u"nein", u"no", u"no"): 0}.get(text.lower())
                if wert is None:
                    raise ValueError(t(u"Bitte \"Ja\" oder \"Nein\" wählen.", u"Please choose \"Yes\" or \"No\".", u"Elija \"Sí\" o \"No\"."))
            elif not text:
                raise ValueError(t(u"Bitte einen Wert eingeben bzw. wählen.", u"Please enter or choose a value.", u"Introduzca o elija un valor."))
            else:
                wert = int(text)
            return methode(param_id, wert)
        if art == StorageType.ElementId:
            if regel.wert_id is None:
                raise ValueError(t(u"Bitte einen Wert aus der Liste wählen.", u"Please choose a value from the list.", u"Elija un valor de la lista."))
            return methode(param_id, ElementId(regel.wert_id))
        # String (oder unbekannt): Leerer Text ist in Revit zulässig
        return methode(param_id, regel.wert or u"")
    except ValueError as fehler:
        raise FilterFehler(t(u"Regel \"%s\": %s", u"Rule \"%s\": %s", u"Regla \"%s\": %s") % (beschreibung, fehler))
    except Exception as fehler:
        raise FilterFehler(t(u"Regel \"%s\" konnte nicht erstellt werden: %s", u"Rule \"%s\" could not be created: %s", u"No se pudo crear la regla \"%s\": %s")
                           % (beschreibung, fehlertext(fehler)))


def baue_elementfilter(satz, aufloeser):
    """Satz -> ElementFilter oder None (keine Regeln)."""
    teile = []
    for element in satz.elemente:
        if isinstance(element, Satz):
            unterfilter = baue_elementfilter(element, aufloeser)
            if unterfilter is not None:
                teile.append(unterfilter)
        else:
            teile.append(ElementParameterFilter(baue_regel(element,
                                                           aufloeser)))
    if not teile:
        return None
    if len(teile) == 1:
        return teile[0]
    liste = List[ElementFilter]()
    for teil in teile:
        liste.Add(teil)
    return LogicalOrFilter(liste) if satz.oder else LogicalAndFilter(liste)


def verwendete_parameter(satz):
    ids = []
    for element in satz.elemente:
        if isinstance(element, Satz):
            ids.extend(verwendete_parameter(element))
        elif element.parameter is not None:
            ids.append(element.parameter)
    return ids


# ---------------------------------------------------------------------------
# Kategorien und Parameter
# ---------------------------------------------------------------------------

def filterbare_kategorien(doc, aufloeser):
    """[(Name, Zahlenwert)] - dieselbe Liste wie im nativen Dialog."""
    eintraege = []
    for kategorie_id in ParameterFilterUtilities.GetAllFilterableCategories():
        name = aufloeser.kategoriename(kategorie_id)
        if not name.startswith(u"<"):
            eintraege.append((name, id_wert(kategorie_id)))
    return sorted(eintraege, key=lambda e: e[0].lower())


def filterbare_parameter(doc, kategorie_werte, aufloeser):
    """[(Name, Zahlenwert)] der für ALLE Kategorien filterbaren Parameter."""
    if not kategorie_werte:
        return []
    ids = ParameterFilterUtilities.GetFilterableParametersInCommon(
        doc, id_liste(kategorie_werte))
    aufloeser.setze_kontext([ElementId(k) for k in kategorie_werte])
    eintraege = []
    for param_id in ids:
        wert = id_wert(param_id)
        name = aufloeser.parametername(wert)
        if name:
            eintraege.append((name, wert))
    return sorted(eintraege, key=lambda e: e[0].lower())


def wertvorschlaege(doc, kategorie_werte, parameter, aufloeser):
    """Vorhandene Werte eines Parameters an Elementen der Kategorien.

    Rückgabe: [(Anzeigetext, Zahlenwert der ElementId oder None)]
    Nur Text- und Verweisparameter; Zahlen tippt man ein.
    """
    info = aufloeser.info(parameter)
    if not kategorie_werte or info.speicherart not in (StorageType.String,
                                                       StorageType.ElementId):
        return []
    anbieter = ParameterValueProvider(ElementId(parameter))
    kategorienfilter = ElementMulticategoryFilter(id_liste(kategorie_werte))
    gefunden = {}
    geprueft = 0
    for nur_typen in (True, False):
        sammler = FilteredElementCollector(doc).WherePasses(kategorienfilter)
        sammler = (sammler.WhereElementIsElementType() if nur_typen
                   else sammler.WhereElementIsNotElementType())
        for element in sammler:
            geprueft += 1
            if geprueft > MAX_ELEMENTE_VORSCHLAG \
                    or len(gefunden) >= MAX_VORSCHLAEGE:
                break
            try:
                if info.speicherart == StorageType.String:
                    if anbieter.IsStringValueSupported(element):
                        text = anbieter.GetStringValue(element)
                        if text:
                            gefunden.setdefault(text, None)
                elif anbieter.IsElementIdValueSupported(element):
                    wert = id_wert(anbieter.GetElementIdValue(element))
                    if wert is not None and wert != -1:
                        gefunden.setdefault(aufloeser.elementname(wert), wert)
            except Exception:
                continue
    return sorted(gefunden.items(), key=lambda e: e[0].lower())


def ungueltige_regeln(doc, kategorie_werte, satz, aufloeser):
    """Parameternamen von Regeln, die für die Kategorien nicht filterbar sind."""
    if not kategorie_werte:
        return sorted(set(aufloeser.parametername(p)
                          for p in verwendete_parameter(satz)))
    zulaessig = set(id_wert(i) for i in
                    ParameterFilterUtilities.GetFilterableParametersInCommon(
                        doc, id_liste(kategorie_werte)))
    return sorted(set(aufloeser.parametername(p)
                      for p in verwendete_parameter(satz)
                      if p not in zulaessig))


# ---------------------------------------------------------------------------
# Namen
# ---------------------------------------------------------------------------

def pruefe_name(doc, name, alter_name=None):
    """Gibt den bereinigten Namen zurück oder wirft FilterFehler."""
    name = (name or u"").strip()
    if not name:
        raise FilterFehler(t(u"Der Name darf nicht leer sein.", u"The name must not be empty.", u"El nombre no puede estar vacío."))
    verboten = sorted(set(z for z in name if z in VERBOTENE_ZEICHEN))
    if verboten:
        raise FilterFehler(
            t(u"Der Name enthält unzulässige Zeichen: %s\n\nIn Filternamen "
            u"nicht erlaubt: %s", u"The name contains invalid characters: %s\n\nNot allowed in filter names: %s", u"El nombre contiene caracteres no válidos: %s\n\nNo permitidos en nombres de filtro: %s") % (u" ".join(verboten),
                                    u" ".join(VERBOTENE_ZEICHEN)))
    if name == alter_name:
        return name
    if not FilterElement.IsNameUnique(doc, name):
        raise FilterFehler(t(u"Es gibt bereits einen Filter mit dem Namen "
                           u"\"%s\".", u"A filter named \"%s\" already exists.", u"Ya existe un filtro llamado \"%s\".") % name)
    return name


def freier_name(doc, basis):
    """'<basis>', sonst '<basis> 2', '<basis> 3', ..."""
    kandidat = basis
    nummer = 2
    while not FilterElement.IsNameUnique(doc, kandidat):
        kandidat = u"%s %d" % (basis, nummer)
        nummer += 1
    return kandidat


# ---------------------------------------------------------------------------
# Schreibende Aktionen (Transaktion übernimmt der Aufrufer)
# ---------------------------------------------------------------------------

def neuer_filter(doc, name, kategorie_werte=()):
    return ParameterFilterElement.Create(doc, name, id_liste(kategorie_werte))


def dupliziere(doc, filter_element, name):
    kategorien = filter_element.GetCategories()
    element_filter = filter_element.GetElementFilter()
    if element_filter is None:
        return ParameterFilterElement.Create(doc, name, kategorien)
    return ParameterFilterElement.Create(doc, name, kategorien,
                                         element_filter)


def speichere_regeln(doc, filter_element, kategorie_werte, satz, aufloeser):
    """Kategorien und Regeln eines Filters setzen. Wirft FilterFehler."""
    element_filter = baue_elementfilter(satz, aufloeser)
    if element_filter is not None:
        menge = HashSet[ElementId]()
        for wert in kategorie_werte:
            menge.Add(ElementId(wert))
        if not ParameterFilterElement.ElementFilterIsAcceptableForParameterFilterElement(  # noqa: E501
                doc, menge, element_filter):
            raise FilterFehler(t(u"Revit akzeptiert diese Regeln für die "
                               u"gewählten Kategorien nicht.", u"Revit does not accept these rules for the selected categories.", u"Revit no acepta estas reglas para las categorías seleccionadas."))
        if not ParameterFilterElement.AllRuleParametersApplicable(
                doc, id_liste(kategorie_werte), element_filter):
            raise FilterFehler(
                t(u"Mindestens ein Parameter ist nicht für alle gewählten "
                u"Kategorien verfügbar: %s", u"At least one parameter is not available for all selected categories: %s", u"Al menos un parámetro no está disponible para todas las categorías seleccionadas: %s") % u", ".join(ungueltige_regeln(
                    doc, kategorie_werte, satz, aufloeser)))
    filter_element.SetCategories(id_liste(kategorie_werte))
    if element_filter is None:
        filter_element.ClearRules()
    else:
        filter_element.SetElementFilter(element_filter)


def ansichten_mit_filtern(doc):
    """Alle Ansichten und Vorlagen, die Anzeigefilter unterstützen."""
    from Autodesk.Revit.DB import View
    ergebnis = []
    for ansicht in FilteredElementCollector(doc).OfClass(View):
        try:
            if ansicht.AreGraphicsOverridesAllowed():
                ergebnis.append(ansicht)
        except Exception:
            continue
    return ergebnis
