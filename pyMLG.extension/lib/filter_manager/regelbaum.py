# -*- coding: utf-8 -*-
"""Zerlegt die Regeln eines ParameterFilterElement in einen lesbaren Knotenbaum.

Aufbau, den ParameterFilterElement.GetElementFilter() liefert (API 2026):

    ElementFilter
    ├─ LogicalAndFilter / LogicalOrFilter   .GetFilters() -> ElementFilter[]
    │                                        (beliebig tief verschachtelt)
    └─ ElementParameterFilter               .GetRules()   -> FilterRule[]
       (mehrere Regeln = implizites UND)

    FilterRule
    ├─ FilterInverseRule                    .GetInnerRule() - Negation
    ├─ FilterCategoryRule                   .GetCategories()
    ├─ SharedParameterApplicableRule        .ParameterName (Text, keine Id!)
    ├─ ParameterValuePresenceRule          .Parameter
    │  └─ HasValueFilterRule / HasNoValueFilterRule
    └─ FilterValueRule
       ├─ FilterStringRule                  .GetEvaluator(), .RuleString
       └─ FilterNumericValueRule            .GetEvaluator()
          ├─ FilterDoubleRule               .RuleValue (interne Einheit), .Epsilon
          ├─ FilterIntegerRule              .RuleValue
          ├─ FilterElementIdRule            .RuleValue (ElementId)
          └─ FilterGlobalParameterAssociationRule  .RuleValue (Id glob. Param.)

Vererbung per Reflection gegen RevitAPI.dll 2026 geprüft - das sind alle
konkreten FilterRule-Klassen. (Achtung: Die Doku zu ParameterFilterElement.Create
nennt nur FilterValueRule/FilterInverseRule/SharedParameterApplicableRule;
die Anwesenheitsregeln erben aber NICHT von FilterValueRule.)

Die Parameter-Id jeder Regel liefert FilterRule.GetRuleParameter()
(InvalidElementId, wenn die Regel keinen Parameter hat).

"Ist ungleich", "beginnt nicht mit" usw. speichert Revit NICHT als eigene
Auswerter, sondern als FilterInverseRule um die positive Regel. Deshalb wird
die Negation hier beim Abstieg mitgeführt.

Das Modul importiert bewusst nichts aus Autodesk.Revit: Die Typen werden über
GetType().Name erkannt, alle Id-Auflösungen erledigt ein übergebener
"Auflöser" (siehe aufloesung.RevitAufloeser). Dadurch lässt sich die
Logik mit Attrappen ohne Revit prüfen (tools/test_filter_regelbaum.py).
"""

from mlg_sprache import t

# Knotenarten
UND = "UND"
ODER = "ODER"
REGELN = "REGELN"        # ElementParameterFilter: Regeln per UND verknüpft
REGEL = "REGEL"
LEER = "LEER"            # Filter ohne Regeln (nur Kategorien)
UNBEKANNT = "UNBEKANNT"
FEHLER = "FEHLER"

# Auswerter-Klasse -> (Schlüssel, Text, negierter Text)
# Texte wie im deutschen Revit-Dialog "Filter".
OPERATOREN = {
    "FilterStringEquals": ("gleich", t(u"ist gleich", u"equals", u"igual a"), t(u"ist ungleich", u"does not equal", u"no es igual a")),
    "FilterStringBeginsWith": ("beginnt", t(u"beginnt mit", u"begins with", u"empieza por"), t(u"beginnt nicht mit", u"does not begin with", u"no empieza por")),
    "FilterStringEndsWith": ("endet", t(u"endet mit", u"ends with", u"termina por"), t(u"endet nicht mit", u"does not end with", u"no termina por")),
    "FilterStringContains": ("enthaelt", t(u"enthält", u"contains", u"contiene"), t(u"enthält nicht", u"does not contain", u"no contiene")),
    "FilterStringGreater": ("groesser", t(u"ist größer als", u"is greater than", u"es mayor que"),
                            t(u"ist nicht größer als", u"is not greater than", u"no es mayor que")),
    "FilterStringGreaterOrEqual": ("groesser_gleich",
                                   t(u"ist größer als oder gleich", u"is greater than or equal to", u"es mayor o igual que"),
                                   t(u"ist nicht größer als oder gleich", u"is not greater than or equal to", u"no es mayor o igual que")),
    "FilterStringLess": ("kleiner", t(u"ist kleiner als", u"is less than", u"es menor que"),
                         t(u"ist nicht kleiner als", u"is not less than", u"no es menor que")),
    "FilterStringLessOrEqual": ("kleiner_gleich",
                                t(u"ist kleiner als oder gleich", u"is less than or equal to", u"es menor o igual que"),
                                t(u"ist nicht kleiner als oder gleich", u"is not less than or equal to", u"no es menor o igual que")),
    "FilterNumericEquals": ("gleich", t(u"ist gleich", u"equals", u"igual a"), t(u"ist ungleich", u"does not equal", u"no es igual a")),
    "FilterNumericGreater": ("groesser", t(u"ist größer als", u"is greater than", u"es mayor que"),
                             t(u"ist nicht größer als", u"is not greater than", u"no es mayor que")),
    "FilterNumericGreaterOrEqual": ("groesser_gleich",
                                    t(u"ist größer als oder gleich", u"is greater than or equal to", u"es mayor o igual que"),
                                    t(u"ist nicht größer als oder gleich", u"is not greater than or equal to", u"no es mayor o igual que")),
    "FilterNumericLess": ("kleiner", t(u"ist kleiner als", u"is less than", u"es menor que"),
                          t(u"ist nicht kleiner als", u"is not less than", u"no es menor que")),
    "FilterNumericLessOrEqual": ("kleiner_gleich",
                                 t(u"ist kleiner als oder gleich", u"is less than or equal to", u"es menor o igual que"),
                                 t(u"ist nicht kleiner als oder gleich", u"is not less than or equal to", u"no es menor o igual que")),
}

WERTREGELN_ZAHL = ("FilterDoubleRule", "FilterIntegerRule",
                   "FilterElementIdRule")


class Knoten(object):
    """Ein Knoten des Regelbaums.

    art          UND | ODER | REGELN | REGEL | LEER | UNBEKANNT | FEHLER
    kinder       Unterknoten (UND/ODER/REGELN)
    invertiert   ElementFilter.Inverted (von Revit-Filtern normalerweise False)

    Nur bei REGEL:
    regeltyp       Klassenname der Revit-Regel (z.B. "FilterStringRule")
    parameter_id   Zahlenwert der Parameter-Id oder None
    parametername  Klartextname oder None (z.B. bei Kategorieregeln)
    operator       Schlüssel ("beginnt", "gleich", "hat_wert", ...)
    negiert        True, wenn die Regel in einer FilterInverseRule steckt
    text_operator  Operator als deutscher Text inkl. Negation
    wert           Vergleichswert als Anzeigetext oder None
    rohwert        Vergleichswert unformatiert (str, float, int oder
                   Zahlenwert der ElementId) - für den Regeleditor
    epsilon        Toleranz einer FilterDoubleRule (interne Einheit) oder None
    """

    def __init__(self, art, **werte):
        self.art = art
        self.kinder = []
        self.invertiert = False
        self.regeltyp = None
        self.parameter_id = None
        self.parametername = None
        self.operator = None
        self.negiert = False
        self.text_operator = None
        self.wert = None
        self.rohwert = None
        self.epsilon = None
        for name, wert in werte.items():
            setattr(self, name, wert)

    def __repr__(self):
        if self.art == REGEL:
            return "<Regel %s>" % regel_als_text(self)
        return "<%s %d Kinder>" % (self.art, len(self.kinder))


# ---------------------------------------------------------------------------
# .NET-Hilfen (funktionieren auch mit Python-Attrappen)
# ---------------------------------------------------------------------------

def _konkret(obj):
    """pythonnet 3 liefert Rückgaben mit Interface-Typ als Interface-Hülle.

    Für Klassen (FilterRule, ElementFilter) ist das normalerweise nicht der
    Fall - die Hülle wird hier trotzdem vorsorglich entfernt.
    """
    return getattr(obj, "__implementation__", obj)


def typname(obj):
    """Laufzeit-Klassenname eines .NET-Objekts (bzw. einer Attrappe)."""
    obj = _konkret(obj)
    try:
        return obj.GetType().Name
    except Exception:
        return type(obj).__name__


def _liste(sammlung):
    return [_konkret(x) for x in (sammlung or [])]


# ---------------------------------------------------------------------------
# Zerlegen
# ---------------------------------------------------------------------------

def zerlege_filter(element_filter, aufloeser):
    """ElementFilter (oder None) -> Knoten."""
    if element_filter is None:
        return Knoten(LEER)
    f = _konkret(element_filter)
    typ = typname(f)
    try:
        if typ in ("LogicalAndFilter", "LogicalOrFilter"):
            knoten = Knoten(UND if typ == "LogicalAndFilter" else ODER)
            for unterfilter in _liste(f.GetFilters()):
                knoten.kinder.append(zerlege_filter(unterfilter, aufloeser))
        elif typ == "ElementParameterFilter":
            knoten = Knoten(REGELN)
            for regel in _liste(f.GetRules()):
                knoten.kinder.append(zerlege_regel(regel, aufloeser))
        else:
            knoten = Knoten(UNBEKANNT, regeltyp=typ)
        knoten.invertiert = bool(getattr(f, "Inverted", False))
    except Exception as fehler:
        knoten = Knoten(FEHLER, regeltyp=typ, wert=u"%s" % fehler)
    return knoten


def _regelparameter(regel, aufloeser):
    """(Zahlenwert, Name) des Regelparameters oder (None, None)."""
    try:
        param_id = regel.GetRuleParameter()
    except Exception:
        # Ersatzweg für ParameterValuePresenceRule
        param_id = getattr(regel, "Parameter", None)
        if param_id is None:
            return None, None
    wert = aufloeser.id_wert(param_id)
    if wert is None or wert == -1:          # InvalidElementId
        return None, None
    return wert, aufloeser.parametername(param_id)


def zerlege_regel(regel, aufloeser, negiert=False):
    """FilterRule -> REGEL-Knoten. Fehler werden als FEHLER-Knoten geliefert."""
    regel = _konkret(regel)
    typ = typname(regel)
    try:
        return _zerlege_regel(regel, typ, aufloeser, negiert)
    except Exception as fehler:
        return Knoten(FEHLER, regeltyp=typ, negiert=negiert,
                      wert=u"%s" % fehler)


def _zerlege_regel(regel, typ, aufloeser, negiert):
    if typ == "FilterInverseRule":
        # Negation weiterreichen; doppelte Negation hebt sich auf
        return zerlege_regel(regel.GetInnerRule(), aufloeser, not negiert)

    knoten = Knoten(REGEL, regeltyp=typ, negiert=negiert)

    if typ == "FilterCategoryRule":
        knoten.operator = "kategorie"
        knoten.text_operator = (t(u"Kategorie ist nicht", u"Category is not", u"La categoría no es") if negiert
                                else t(u"Kategorie ist", u"Category is", u"La categoría es"))
        knoten.wert = u", ".join(aufloeser.kategorienamen(
            regel.GetCategories()))
        return knoten

    if typ == "SharedParameterApplicableRule":
        # Einzige Regel, die den Parameter per Name statt per Id kennt
        knoten.parametername = regel.ParameterName
        knoten.operator = "vorhanden"
        knoten.text_operator = (t(u"ist nicht vorhanden", u"does not exist", u"no existe") if negiert
                                else t(u"ist vorhanden", u"exists", u"existe"))
        return knoten

    knoten.parameter_id, knoten.parametername = _regelparameter(regel,
                                                                aufloeser)

    if typ in ("HasValueFilterRule", "HasNoValueFilterRule"):
        hat_wert = (typ == "HasValueFilterRule") != negiert
        knoten.operator = "hat_wert" if typ == "HasValueFilterRule" \
            else "hat_keinen_wert"
        knoten.text_operator = (t(u"hat einen Wert", u"has a value", u"tiene un valor") if hat_wert
                                else t(u"hat keinen Wert", u"has no value", u"no tiene valor"))
        return knoten

    if typ == "FilterGlobalParameterAssociationRule":
        knoten.operator = "global"
        knoten.text_operator = (
            t(u"ist nicht mit globalem Parameter verknüpft", u"is not associated with global parameter", u"no está asociado a un parámetro global") if negiert
            else t(u"ist mit globalem Parameter verknüpft", u"is associated with global parameter", u"está asociado a un parámetro global"))
        knoten.wert = aufloeser.elementname(regel.RuleValue)
        return knoten

    if typ == "FilterStringRule" or typ in WERTREGELN_ZAHL:
        auswerter = typname(regel.GetEvaluator())
        schluessel, text, text_negiert = OPERATOREN.get(
            auswerter, (auswerter, auswerter, t(u"nicht ", u"not ", u"no ") + auswerter))
        knoten.operator = schluessel
        knoten.text_operator = text_negiert if negiert else text
        if typ == "FilterStringRule":
            knoten.rohwert = regel.RuleString
            knoten.wert = t(u"„%s“", u"\"%s\"", u"«%s»") % regel.RuleString
        elif typ == "FilterDoubleRule":
            knoten.rohwert = float(regel.RuleValue)
            knoten.epsilon = float(getattr(regel, "Epsilon", 0.0) or 0.0) \
                or None
            knoten.wert = aufloeser.zahl(regel.GetRuleParameter(),
                                         regel.RuleValue)
        elif typ == "FilterIntegerRule":
            knoten.rohwert = int(regel.RuleValue)
            knoten.wert = aufloeser.ganzzahl(regel.GetRuleParameter(),
                                             regel.RuleValue)
        else:
            knoten.rohwert = aufloeser.id_wert(regel.RuleValue)
            knoten.wert = aufloeser.elementname(regel.RuleValue)
        return knoten

    # Unbekannte (künftige) Regelart: Parameter trotzdem ausweisen
    knoten.operator = typ
    knoten.text_operator = (t(u"nicht ", u"not ", u"no ") if negiert else u"") + typ
    return knoten


def analysiere(filter_element, aufloeser):
    """ParameterFilterElement -> dict mit allen Angaben für Liste und Suche."""
    ergebnis = {
        "element": filter_element,
        "id": aufloeser.id_wert(filter_element.Id),
        "name": filter_element.Name,
        "kategorien": [],
        "baum": None,
        "parameter": [],
        "fehler": None,
    }
    try:
        kategorie_ids = list(filter_element.GetCategories())
        if hasattr(aufloeser, "setze_kontext"):
            aufloeser.setze_kontext(kategorie_ids)
        ergebnis["kategorien"] = aufloeser.kategorienamen(kategorie_ids)
        ergebnis["baum"] = zerlege_filter(filter_element.GetElementFilter(),
                                          aufloeser)
        ergebnis["parameter"] = parameternamen(ergebnis["baum"])
    except Exception as fehler:
        ergebnis["fehler"] = u"%s" % fehler
    return ergebnis


# ---------------------------------------------------------------------------
# Auswerten und Darstellen
# ---------------------------------------------------------------------------

def regeln(knoten):
    """Alle REGEL-Knoten (Tiefensuche, Reihenfolge wie im Baum)."""
    if knoten is None:
        return []
    if knoten.art == REGEL:
        return [knoten]
    gefunden = []
    for kind in knoten.kinder:
        gefunden.extend(regeln(kind))
    return gefunden


def fehlerknoten(knoten):
    if knoten is None:
        return []
    if knoten.art in (FEHLER, UNBEKANNT):
        return [knoten]
    gefunden = []
    for kind in knoten.kinder:
        gefunden.extend(fehlerknoten(kind))
    return gefunden


def parameternamen(knoten):
    """Eindeutige Parameternamen in Baumreihenfolge (ohne Kategorieregeln)."""
    namen = []
    for regel in regeln(knoten):
        if regel.parametername and regel.parametername not in namen:
            namen.append(regel.parametername)
    return namen


def regel_als_text(knoten):
    teile = []
    if knoten.parametername:
        teile.append(knoten.parametername)
    teile.append(knoten.text_operator or u"?")
    if knoten.wert is not None:
        teile.append(knoten.wert)
    return u" ".join(teile)


def _verknuepfung(knoten):
    return t(u" UND ", u" AND ", u" Y ") if knoten.art in (UND, REGELN) else t(u" ODER ", u" OR ", u" O ")


def als_formel(knoten):
    """Einzeiliger Ausdruck, z.B. '(Typname beginnt mit „WD“ UND …) ODER …'."""
    if knoten is None or knoten.art == LEER:
        return t(u"(keine Regeln)", u"(no rules)", u"(sin reglas)")
    if knoten.art == REGEL:
        text = regel_als_text(knoten)
    elif knoten.art == FEHLER:
        text = t(u"[Fehler: %s]", u"[Error: %s]", u"[Error: %s]") % knoten.wert
    elif knoten.art == UNBEKANNT:
        text = u"[%s]" % knoten.regeltyp
    else:
        teile = []
        for kind in knoten.kinder:
            teil = als_formel(kind)
            if kind.art in (UND, ODER, REGELN) and len(kind.kinder) > 1:
                teil = u"(" + teil + u")"
            teile.append(teil)
        text = _verknuepfung(knoten).join(teile)
    if knoten.invertiert:
        text = t(u"NICHT (", u"NOT (", u"NO (") + text + u")"
    return text


def als_zeilen(knoten, tiefe=0):
    """Mehrzeilige, eingerückte Darstellung für Ausgabefenster/Protokoll."""
    einzug = u"    " * tiefe
    if knoten is None or knoten.art == LEER:
        return [einzug + t(u"(keine Regeln - nur Kategorien)", u"(no rules - categories only)", u"(sin reglas - solo categorías)")]
    nicht = t(u"NICHT ", u"NOT ", u"NO ") if knoten.invertiert else u""
    if knoten.art == REGEL:
        return [einzug + u"- " + nicht + regel_als_text(knoten)
                + u"   [%s%s, Parameter-Id %s]" % (
                    u"FilterInverseRule > " if knoten.negiert else u"",
                    knoten.regeltyp, knoten.parameter_id)]
    if knoten.art == FEHLER:
        return [einzug + t(u"- FEHLER in %s: %s", u"- ERROR in %s: %s", u"- ERROR en %s: %s") % (knoten.regeltyp, knoten.wert)]
    if knoten.art == UNBEKANNT:
        return [einzug + t(u"- UNBEKANNTER Filtertyp: %s", u"- UNKNOWN filter type: %s", u"- Tipo de filtro DESCONOCIDO: %s") % knoten.regeltyp]
    if (knoten.art == REGELN and len(knoten.kinder) == 1
            and not knoten.invertiert):
        # ElementParameterFilter mit nur einer Regel: keine eigene Ebene
        return als_zeilen(knoten.kinder[0], tiefe)
    kopf = {UND: t(u"UND  [LogicalAndFilter]", u"AND  [LogicalAndFilter]", u"Y  [LogicalAndFilter]"),
            ODER: t(u"ODER  [LogicalOrFilter]", u"OR  [LogicalOrFilter]", u"O  [LogicalOrFilter]"),
            REGELN: t(u"UND  [ElementParameterFilter, %d Regel(n)]", u"AND  [ElementParameterFilter, %d rule(s)]", u"Y  [ElementParameterFilter, %d regla(s)]")
                    % len(knoten.kinder)}[knoten.art]
    zeilen = [einzug + nicht + kopf]
    for kind in knoten.kinder:
        zeilen.extend(als_zeilen(kind, tiefe + 1))
    return zeilen
