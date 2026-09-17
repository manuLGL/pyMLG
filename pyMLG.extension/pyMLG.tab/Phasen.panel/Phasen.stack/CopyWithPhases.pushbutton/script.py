# -*- coding: utf-8 -*-
"""Kopiert Elemente von einem Basispunkt zu einem Zielpunkt und behält
"Phase erstellt" und "Phase abgebrochen" der Originale bei (Revit setzt bei
Kopien sonst die Phase der aktiven Ansicht). Abhängige Elemente wie Türen
und Fenster werden mit übernommen."""

__title__ = "Copy With Phases"
__author__ = "Manuel"

from Autodesk.Revit.DB import Transaction
from Autodesk.Revit.Exceptions import (InvalidOperationException,
                                       OperationCanceledException)
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import forms, revit

from phasen import revit_kopie as rk

TITEL = u"Copy With Phases"


def auswahl(uidoc):
    ids = list(uidoc.Selection.GetElementIds())
    if ids:
        return ids
    refs = uidoc.Selection.PickObjects(ObjectType.Element,
                                       u"Elemente zum Kopieren wählen")
    return [r.ElementId for r in refs]


def main():
    doc, uidoc = revit.doc, revit.uidoc
    try:
        ids, uebersprungen = rk.modell_elemente(doc, auswahl(uidoc))
        if not ids:
            forms.alert(u"Keine Modellelemente ausgewählt.", title=TITEL)
            return
        basis = uidoc.Selection.PickPoint(u"Basispunkt wählen")
        ziel = uidoc.Selection.PickPoint(u"Zielpunkt wählen")
    except OperationCanceledException:
        return
    except InvalidOperationException:
        forms.alert(u"In dieser Ansicht kann kein Punkt gewählt werden "
                    u"(keine Arbeitsebene). Bitte in einem Grundriss starten "
                    u"oder eine Arbeitsebene festlegen.", title=TITEL)
        return

    t = Transaction(doc, TITEL)
    t.Start()
    try:
        neue_ids, paare, ohne_original = rk.kopieren(doc, ids, ziel - basis)
        for kopie, original in paare:
            rk.phasen_uebertragen(kopie, original)
        t.Commit()
    except Exception as fehler:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        forms.alert(u"Kopieren fehlgeschlagen, nichts wurde geändert.",
                    sub_msg=u"{}".format(fehler), title=TITEL)
        return

    rk.auswahl_setzen(uidoc, neue_ids)
    hinweise = []
    if ohne_original:
        hinweise.append(u"{} Kopie(n) konnte kein Original zugeordnet werden - "
                        u"deren Phasen bitte prüfen.".format(ohne_original))
    if uebersprungen:
        hinweise.append(u"{} ansichtsspezifische(s) Element(e) (Beschriftungen, "
                        u"Detaillinien ...) wurden nicht kopiert.".format(uebersprungen))
    if hinweise:
        forms.alert(u"{} Element(e) kopiert.".format(len(neue_ids)),
                    sub_msg=u"\n".join(hinweise), title=TITEL)


main()
