# -*- coding: utf-8 -*-
"""Offline-Test des Regelbaum-Parsers des Filter-Managers (ohne Revit).

    python tools\\test_filter_regelbaum.py

Die Revit-Klassen werden durch Attrappen mit gleichem Klassennamen und
gleichen Methoden ersetzt. Geprüft wird die Rekursion, die Erkennung der
Regelarten und die Negation über FilterInverseRule. Ob Revit die Struktur
wirklich so liefert, lässt sich nur in Revit prüfen (Filter-Manager öffnen
und die Regeln mit dem nativen Dialog vergleichen).
"""
import os
import sys

# Texte werden auf Deutsch verglichen
os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from filter_manager import regelbaum as rb  # noqa: E402


# ---------------------------------------------------------------------------
# Attrappen
# ---------------------------------------------------------------------------

class _Typ(object):
    def __init__(self, name):
        self.Name = name


class Attrappe(object):
    def GetType(self):
        return _Typ(type(self).__name__)


class ElementId(Attrappe):
    def __init__(self, wert):
        self.Value = wert


UNGUELTIG = ElementId(-1)


class LogicalAndFilter(Attrappe):
    Inverted = False

    def __init__(self, *filter_):
        self._f = list(filter_)

    def GetFilters(self):
        return self._f


class LogicalOrFilter(LogicalAndFilter):
    pass


class ElementParameterFilter(Attrappe):
    Inverted = False

    def __init__(self, *regeln):
        self._r = list(regeln)

    def GetRules(self):
        return self._r


class _Regel(Attrappe):
    param = UNGUELTIG

    def GetRuleParameter(self):
        return self.param


class FilterStringBeginsWith(Attrappe):
    pass


class FilterStringContains(Attrappe):
    pass


class FilterStringEquals(Attrappe):
    pass


class FilterNumericGreater(Attrappe):
    pass


class FilterNumericEquals(Attrappe):
    pass


class FilterStringRule(_Regel):
    def __init__(self, param, auswerter, text):
        self.param, self._a, self.RuleString = ElementId(param), auswerter, text

    def GetEvaluator(self):
        return self._a


class FilterDoubleRule(_Regel):
    def __init__(self, param, auswerter, wert):
        self.param, self._a, self.RuleValue = ElementId(param), auswerter, wert

    def GetEvaluator(self):
        return self._a


class FilterElementIdRule(FilterDoubleRule):
    pass


class FilterInverseRule(_Regel):
    def __init__(self, innen):
        self._i = innen

    def GetInnerRule(self):
        return self._i


class HasValueFilterRule(_Regel):
    def __init__(self, param):
        self.param = ElementId(param)


class HasNoValueFilterRule(HasValueFilterRule):
    pass


class FilterCategoryRule(_Regel):
    def __init__(self, *kategorien):
        self._k = [ElementId(k) for k in kategorien]

    def GetCategories(self):
        return self._k


class SharedParameterApplicableRule(_Regel):
    def __init__(self, name):
        self.ParameterName = name


class KaputteRegel(_Regel):
    def GetRuleParameter(self):
        raise RuntimeError("defekt")


class FilterStringRuleKaputt(FilterStringRule):
    """Meldet sich als FilterStringRule, wirft aber beim Auswerter."""

    def GetType(self):
        return _Typ("FilterStringRule")

    def GetEvaluator(self):
        raise RuntimeError("Auswerter nicht lesbar")


class ParameterFilterElement(Attrappe):
    def __init__(self, name, kategorien, element_filter):
        self.Name, self.Id = name, ElementId(4711)
        self._k, self._f = [ElementId(k) for k in kategorien], element_filter

    def GetCategories(self):
        return self._k

    def GetElementFilter(self):
        return self._f


NAMEN = {-1002002: u"Typname", -1010106: u"Kommentare", -1001203: u"Kennzeichen",
         -1001205: u"Kosten", -1012101: u"Phase erstellt", 123456: u"WD_Klasse"}
KATEGORIEN = {-2000011: u"Wände", -2000023: u"Türen"}


class Aufloeser(object):
    kontext = None

    def setze_kontext(self, ids):
        self.kontext = [i.Value for i in ids]

    def id_wert(self, eid):
        return None if eid is None else eid.Value

    def parametername(self, eid):
        return NAMEN.get(eid.Value, u"<%s>" % eid.Value)

    def kategorienamen(self, ids):
        return sorted(KATEGORIEN[i.Value] for i in ids)

    def elementname(self, eid):
        return u"Phase %s" % eid.Value

    def zahl(self, eid, wert):
        return u"%g €" % wert

    def ganzzahl(self, eid, wert):
        return u"%d" % wert


# ---------------------------------------------------------------------------

def pruefe(name, bedingung, info=None):
    print(("  OK    " if bedingung else "  FEHLER") + "  " + name)
    if not bedingung and info is not None:
        print("          ist: %r" % (info,))
    return bedingung


def main():
    ok = True
    a = Aufloeser()

    # Wie im Screenshot: Typname beginnt mit WD, verschachtelt in UND/ODER
    baum = LogicalOrFilter(
        LogicalAndFilter(
            ElementParameterFilter(FilterStringRule(
                -1002002, FilterStringBeginsWith(), u"WD")),
            ElementParameterFilter(FilterStringRule(
                -1010106, FilterStringContains(), u"Test"))),
        ElementParameterFilter(FilterInverseRule(FilterStringRule(
            -1001203, FilterStringEquals(), u"X"))),
        ElementParameterFilter(
            FilterDoubleRule(-1001205, FilterNumericGreater(), 100.0),
            HasValueFilterRule(-1001203),
            FilterInverseRule(FilterInverseRule(
                HasNoValueFilterRule(123456)))),
        ElementParameterFilter(FilterCategoryRule(-2000011),
                               SharedParameterApplicableRule(u"WD_Klasse")),
        ElementParameterFilter(FilterElementIdRule(
            -1012101, FilterNumericEquals(), ElementId(99))),
    )
    fe = ParameterFilterElement(u"WD Wände", [-2000011, -2000023], baum)
    erg = rb.analysiere(fe, a)
    for zeile in rb.als_zeilen(erg["baum"]):
        print("    " + zeile)
    print("    Formel: " + rb.als_formel(erg["baum"]))

    ok &= pruefe("kein Fehler", erg["fehler"] is None, erg["fehler"])
    ok &= pruefe("Kontext gesetzt", a.kontext == [-2000011, -2000023], a.kontext)
    ok &= pruefe("Kategorien", erg["kategorien"] == [u"Türen", u"Wände"],
                 erg["kategorien"])
    ok &= pruefe("Parameternamen in Reihenfolge, eindeutig",
                 erg["parameter"] == [u"Typname", u"Kommentare", u"Kennzeichen",
                                      u"Kosten", u"WD_Klasse",
                                      u"Phase erstellt"], erg["parameter"])
    texte = [rb.regel_als_text(r) for r in rb.regeln(erg["baum"])]
    erwartet = [u"Typname beginnt mit „WD“",
                u"Kommentare enthält „Test“",
                u"Kennzeichen ist ungleich „X“",
                u"Kosten ist größer als 100 €",
                u"Kennzeichen hat einen Wert",
                u"WD_Klasse hat keinen Wert",
                u"Kategorie ist Wände",
                u"WD_Klasse ist vorhanden",
                u"Phase erstellt ist gleich Phase 99"]
    ok &= pruefe("Regeltexte inkl. Negation", texte == erwartet, texte)
    ok &= pruefe("Wurzel ist ODER", erg["baum"].art == rb.ODER)
    ok &= pruefe("Negation am Knoten markiert",
                 rb.regeln(erg["baum"])[2].negiert
                 and not rb.regeln(erg["baum"])[5].negiert)
    ok &= pruefe("Formel geklammert",
                 rb.als_formel(erg["baum"]).startswith(
                     u"(Typname beginnt mit „WD“ UND Kommentare"))

    # Filter ohne Regeln
    leer = rb.analysiere(ParameterFilterElement(u"Nur Kat", [-2000023], None), a)
    ok &= pruefe("Filter ohne Regeln", leer["baum"].art == rb.LEER
                 and leer["parameter"] == [])

    # Einzelnes ElementParameterFilter direkt als Wurzel
    einzeln = rb.analysiere(ParameterFilterElement(
        u"Einzeln", [-2000011], ElementParameterFilter(FilterStringRule(
            -1002002, FilterStringBeginsWith(), u"WD"))), a)
    ok &= pruefe("einzelnes ElementParameterFilter",
                 einzeln["parameter"] == [u"Typname"], einzeln["parameter"])

    # Defekte Regeln dürfen den Rest nicht verhindern
    defekt = rb.analysiere(ParameterFilterElement(
        u"Defekt", [-2000011], ElementParameterFilter(
            FilterStringRuleKaputt(-1002002, None, u"x"),
            FilterStringRule(-1010106, FilterStringContains(), u"ok"))), a)
    ok &= pruefe("Fehler isoliert, zweite Regel gelesen",
                 len(rb.fehlerknoten(defekt["baum"])) == 1
                 and defekt["parameter"] == [u"Kommentare"],
                 (defekt["parameter"], rb.fehlerknoten(defekt["baum"])))

    # Unbekannte Regelart (z.B. künftige API) mit Parameter
    class FilterZukunftRule(_Regel):
        param = ElementId(-1001203)
    unbekannt = rb.zerlege_regel(FilterZukunftRule(), a)
    ok &= pruefe("unbekannte Regelart behält Parameter",
                 unbekannt.parametername == u"Kennzeichen")

    print("")
    print("ALLE TESTS BESTANDEN" if ok else "ES GIBT FEHLER")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
