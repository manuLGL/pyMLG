# -*- coding: utf-8 -*-
# Dupliziert die ausgewählten Pläne (Projektbrowser-Auswahl oder aktiver
# Plan) samt Ansichten über die Revit-Funktion "Plan duplizieren" und
# vergibt die nächste freie Plannummer (A-101 -> A-102 ...). Die Lage der
# Ansichtsfenster bleibt erhalten. Braucht Revit 2023 oder neuer.
# (Kommentar statt Docstring: pyRevit liest unter IronPython den
#  Docstring als Tooltip und scheitert dabei an Umlauten.)

__title__ = "DuplicateSheet"
__author__ = "Manuel"

from Autodesk.Revit.DB import Transaction
from pyrevit import forms, revit

from mlg_plaene import logik as lg
from mlg_plaene import revit as rv
from mlg_sprache import t

TITEL = t(u"Plan duplizieren", u"Duplicate Sheet", u"Duplicar plano")

try:
    from Autodesk.Revit.DB import SheetDuplicateOption
except ImportError:
    SheetDuplicateOption = None

# (Text, Name der Option) - fehlt eine Option in dieser Revit-Version,
# wird sie nicht angeboten
OPTIONEN = [
    (t(u"Mit Ansichten und Detaillierung", u"With views and detailing",
       u"Con vistas y detalles"), "DuplicateSheetWithViewsAndDetailing"),
    (t(u"Mit Ansichten, ohne Detaillierung", u"With views, without detailing",
       u"Con vistas, sin detalles"), "DuplicateSheetWithViewsOnly"),
    (t(u"Mit Ansichten als abhängige Ansichten", u"With views as dependent views",
       u"Con vistas como vistas dependientes"), "DuplicateSheetWithViewsAsDependent"),
    (t(u"Nur Plan mit Detaillierung (ohne Ansichten)",
       u"Sheet with detailing only (no views)",
       u"Sólo plano con detalles (sin vistas)"), "DuplicateSheetWithDetailing"),
    (t(u"Leerer Plan (nur Plankopf)", u"Empty sheet (title block only)",
       u"Plano vacío (sólo cajetín)"), "DuplicateEmptySheet"),
]


def plaene_waehlen(doc, uidoc):
    gewaehlt = rv.gewaehlte_plaene(doc, uidoc)
    if gewaehlt:
        return gewaehlt
    alle = rv.plaene(doc)
    texte = forms.SelectFromList.show(
        [rv.plan_text(p) for p in alle], multiselect=True,
        title=t(u"Pläne zum Duplizieren wählen", u"Select sheets to duplicate",
                u"Seleccione planos para duplicar"),
        button_name=t(u"Weiter", u"Next", u"Siguiente"))
    if not texte:
        return []
    texte = set(texte)
    return [p for p in alle if rv.plan_text(p) in texte]


def anzahl_fragen():
    text = forms.ask_for_string(
        default="1",
        prompt=t(u"Wie viele Kopien je Plan?", u"How many copies per sheet?",
                 u"¿Cuántas copias por plano?"),
        title=TITEL)
    if text is None:
        return None
    try:
        anzahl = int(text.strip())
    except ValueError:
        anzahl = 0
    if not 1 <= anzahl <= 50:
        forms.alert(t(u"Bitte eine Zahl von 1 bis 50 angeben.",
                      u"Please enter a number from 1 to 50.",
                      u"Indique un número del 1 al 50."), title=TITEL)
        return None
    return anzahl


def main():
    doc, uidoc = revit.doc, revit.uidoc
    if SheetDuplicateOption is None:
        forms.alert(t(u"Plan duplizieren braucht Revit 2023 oder neuer.",
                      u"Duplicating sheets requires Revit 2023 or newer.",
                      u"Duplicar planos requiere Revit 2023 o posterior."), title=TITEL)
        return
    plaene = plaene_waehlen(doc, uidoc)
    if not plaene:
        return

    optionen = [(text, getattr(SheetDuplicateOption, name))
                for text, name in OPTIONEN if hasattr(SheetDuplicateOption, name)]
    if not optionen:
        return
    wahl = forms.CommandSwitchWindow.show(
        [text for text, _ in optionen],
        message=t(u"{} Plan/Pläne duplizieren:", u"Duplicate {} sheet(s):",
                  u"Duplicar {} plano(s):").format(len(plaene)))
    if not wahl:
        return
    option = dict(optionen)[wahl]
    anzahl = anzahl_fragen()
    if not anzahl:
        return

    vergeben = rv.vergebene_nummern(doc)
    neue, hinweise = [], []
    transaktion = Transaction(doc, TITEL)
    transaktion.Start()
    try:
        for plan in plaene:
            if not plan.CanBeDuplicated(option):
                hinweise.append(t(u"{}: lässt sich so nicht duplizieren",
                                  u"{}: cannot be duplicated this way",
                                  u"{}: no se puede duplicar así").format(rv.plan_text(plan)))
                continue
            letzte = plan.SheetNumber
            for _ in range(anzahl):
                kopie = doc.GetElement(plan.Duplicate(option))
                letzte = lg.naechste_freie_nummer(letzte, vergeben)
                kopie.SheetNumber = letzte
                kopie.Name = plan.Name
                neue.append(rv.plan_text(kopie))
        transaktion.Commit()
    except Exception as fehler:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        forms.alert(t(u"Duplizieren fehlgeschlagen, nichts wurde geändert.",
                      u"Duplication failed, nothing was changed.",
                      u"Error al duplicar, no se cambió nada."),
                    sub_msg=u"{}".format(fehler), title=TITEL)
        return

    meldung = t(u"{} Plan/Pläne erstellt.", u"{} sheet(s) created.",
                u"{} plano(s) creados.").format(len(neue))
    forms.alert(meldung, sub_msg=u"\n".join(hinweise + neue[:30]), title=TITEL)


main()
