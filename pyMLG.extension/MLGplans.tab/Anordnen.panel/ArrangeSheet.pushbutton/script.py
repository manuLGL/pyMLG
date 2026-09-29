# -*- coding: utf-8 -*-
# Zeigt den aktiven Plan verkleinert in einem Vorschaufenster: Ansichten
# per Maus verschieben, bündig ausrichten, gleichmäßig verteilen oder
# automatisch anordnen. Sind Ansichtsfenster markiert, sind nur diese
# verschiebbar, die übrigen bleiben grau stehen. "Übernehmen" verschiebt
# die Ansichtsfenster auf dem Plan.
# (Kommentar statt Docstring: pyRevit liest unter IronPython den
#  Docstring als Tooltip und scheitert dabei an Umlauten.)

__title__ = "ArrangeSheet"
__author__ = "Manuel"

from Autodesk.Revit.DB import Transaction, Viewport, ViewSheet
from pyrevit import forms, revit

from mlg_plaene import logik as lg
from mlg_plaene import revit as rv
from mlg_plaene import vorschau
from mlg_sprache import t

TITEL = t(u"Plan anordnen", u"Arrange Sheet", u"Organizar plano")
RAND_CM = 2.0
ABSTAND_CM = 1.0


def main():
    doc, uidoc = revit.doc, revit.uidoc
    plan = doc.ActiveView
    if not isinstance(plan, ViewSheet):
        forms.alert(t(u"Bitte den Plan öffnen, der angeordnet werden soll.",
                      u"Please open the sheet to arrange.",
                      u"Abra el plano que desea organizar."), title=TITEL)
        return
    rv.messung_zuruecksetzen()
    alle = [doc.GetElement(i) for i in plan.GetAllViewports()]
    if not alle:
        forms.alert(t(u"Auf diesem Plan liegen keine Ansichten.",
                      u"There are no views on this sheet.",
                      u"No hay vistas en este plano."), title=TITEL)
        return

    markiert = set(rv.id_wert(i) for i in uidoc.Selection.GetElementIds()
                   if isinstance(doc.GetElement(i), Viewport))
    beweglich = [f for f in alle if not markiert or rv.id_wert(f.Id) in markiert]
    fest_ids = [f.Id for f in beweglich]

    umrisse = [rv.fenster_rechteck(f, doc) for f in beweglich]
    eintraege = [vorschau.Eintrag(None, name, r, beweglich=False)
                 for name, r in rv.belegte_eintraege(doc, plan, fest_ids)]
    eintraege += [vorschau.Eintrag(i, rv.vorschau_name(doc.GetElement(f.ViewId)), r)
                  for i, (f, r) in enumerate(zip(beweglich, umrisse))]

    kopf = rv.plankopf_rechteck(doc, plan)
    flaeche = lg.zeichenflaeche(kopf, rv.fuss(RAND_CM))
    ergebnis = vorschau.zeigen(
        u"{} - {}".format(TITEL, rv.plan_text(plan)), kopf, flaeche, eintraege,
        rv.fuss(ABSTAND_CM), rv.plankopf_linien(doc, plan))
    if not ergebnis:
        return

    transaktion = Transaction(doc, TITEL)
    transaktion.Start()
    try:
        for i, neu in ergebnis.items():
            if neu != umrisse[i]:
                rv.verschiebe_rechteck_auf(beweglich[i], umrisse[i], neu)
        transaktion.Commit()
    except Exception as fehler:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        forms.alert(t(u"Anordnen fehlgeschlagen, nichts wurde geändert.",
                      u"Arranging failed, nothing was changed.",
                      u"Error al organizar, no se cambió nada."),
                    sub_msg=u"{}".format(fehler), title=TITEL)


main()
