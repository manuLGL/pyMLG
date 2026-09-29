# -*- coding: utf-8 -*-
# Verschiebt die ausgewählten Ansichtsfenster des aktiven Plans auf einen
# anderen Plan - an gleicher Position, mit Typ, Drehung, Titellage und
# Detailnummer. Die Ansichten selbst bleiben unverändert, Beschriftungen
# darin bleiben erhalten; Verweise zeigen danach auf den neuen Plan.
# (Kommentar statt Docstring: pyRevit liest unter IronPython den
#  Docstring als Tooltip und scheitert dabei an Umlauten.)

__title__ = "MoveViewport"
__author__ = "Manuel"

from Autodesk.Revit.DB import ElementId, Transaction, Viewport, ViewSheet
from Autodesk.Revit.Exceptions import OperationCanceledException
from Autodesk.Revit.UI.Selection import ObjectType
from System.Collections.Generic import List
from pyrevit import forms, revit

from mlg_plaene import revit as rv
from mlg_sprache import t

TITEL = t(u"Ansicht verschieben", u"Move Viewport", u"Mover vista")


def fenster_waehlen(doc, uidoc):
    gewaehlt = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    gewaehlt = [e for e in gewaehlt if isinstance(e, Viewport)]
    if gewaehlt:
        return gewaehlt
    refs = uidoc.Selection.PickObjects(
        ObjectType.Element, rv.NurAnsichtsfenster(),
        t(u"Ansichtsfenster zum Verschieben wählen, dann 'Fertig stellen'",
          u"Select viewports to move, then 'Finish'",
          u"Seleccione ventanas gráficas para mover y pulse 'Finalizar'"))
    return [doc.GetElement(r.ElementId) for r in refs]


def main():
    doc, uidoc = revit.doc, revit.uidoc
    plan = doc.ActiveView
    if not isinstance(plan, ViewSheet):
        forms.alert(t(u"Bitte den Plan öffnen, auf dem die Ansichten liegen.",
                      u"Please open the sheet containing the views.",
                      u"Abra el plano que contiene las vistas."), title=TITEL)
        return
    try:
        auswahl = fenster_waehlen(doc, uidoc)
    except OperationCanceledException:
        return
    if not auswahl:
        return

    andere = [p for p in rv.plaene(doc) if rv.id_wert(p.Id) != rv.id_wert(plan.Id)]
    text = forms.SelectFromList.show(
        [rv.plan_text(p) for p in andere], multiselect=False,
        title=t(u"Zielplan wählen", u"Select target sheet",
                u"Seleccione el plano de destino"),
        button_name=t(u"Verschieben", u"Move", u"Mover"))
    if not text:
        return
    ziel = [p for p in andere if rv.plan_text(p) == text][0]

    neue, hinweise = [], []
    transaktion = Transaction(doc, TITEL)
    transaktion.Start()
    try:
        for fenster in auswahl:
            ansicht_id = fenster.ViewId
            werte = rv.fenster_eigenschaften(fenster)
            doc.Delete(fenster.Id)
            if not Viewport.CanAddViewToSheet(doc, ziel.Id, ansicht_id):
                raise Exception(t(u"'{}' kann nicht auf '{}' platziert werden",
                                  u"'{}' cannot be placed on '{}'",
                                  u"'{}' no se puede colocar en '{}'").format(
                    doc.GetElement(ansicht_id).Name, rv.plan_text(ziel)))
            neu = Viewport.Create(doc, ziel.Id, ansicht_id, werte["mitte"])
            hinweise.extend(rv.eigenschaften_setzen(doc, neu, werte))
            neue.append(neu.Id)
        transaktion.Commit()
    except Exception as fehler:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        forms.alert(t(u"Verschieben fehlgeschlagen, nichts wurde geändert.",
                      u"Move failed, nothing was changed.",
                      u"Error al mover, no se cambió nada."),
                    sub_msg=u"{}".format(fehler), title=TITEL)
        return

    # Zielplan öffnen und die verschobenen Ansichten markieren
    uidoc.ActiveView = ziel
    liste = List[ElementId]()
    for eid in neue:
        liste.Add(eid)
    uidoc.Selection.SetElementIds(liste)
    if hinweise:
        forms.alert(t(u"{} Ansicht(en) nach '{}' verschoben.",
                      u"{} view(s) moved to '{}'.",
                      u"{} vista(s) movidas a '{}'.").format(len(neue), rv.plan_text(ziel)),
                    sub_msg=u"\n".join(hinweise), title=TITEL)


main()
