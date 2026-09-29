# -*- coding: utf-8 -*-
# Richtet andere Pläne wie den aktiven Plan aus: jede Ansicht des aktiven
# Plans (oder nur die markierten) bekommt auf jedem Zielplan ihr Gegenstück
# - gleiche Ansichtsart, Vorrang für dieselbe Ansicht, dann gleicher
# Maßstab, dann die nächstliegende. Das Gegenstück übernimmt Lage,
# Ansichtsfenster-Typ und Titellage:
#   - Gebäude an gleicher Stelle: der Projektursprung liegt gleich
#     (Grundrisse EG bis DG übereinander); Legenden und Zeichenansichten
#     nach der Mitte
#   - Ansichtsfenster an gleicher Stelle: die Mitte liegt gleich
# Die Technik steckt in lib/mlg_plaene/ausrichten.py.
# (Kommentar statt Docstring: pyRevit liest unter IronPython den
#  Docstring als Tooltip und scheitert dabei an Umlauten.)

__title__ = "AlignViewports"
__author__ = "Manuel"

from Autodesk.Revit.DB import Transaction, Viewport, ViewSheet
from pyrevit import forms, revit

from mlg_plaene import ausrichten as ar
from mlg_plaene import revit as rv
from mlg_sprache import t

TITEL = t(u"Ansichten ausrichten", u"Align Viewports", u"Alinear vistas")


def vorbild_fenster(doc, uidoc, plan):
    """Markierte Ansichtsfenster des Plans, sonst alle."""
    gewaehlt = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    gewaehlt = [e for e in gewaehlt if isinstance(e, Viewport)
                and rv.id_wert(e.SheetId) == rv.id_wert(plan.Id)]
    if gewaehlt:
        return gewaehlt
    return [doc.GetElement(i) for i in plan.GetAllViewports()]


def plan_ausrichten(doc, vorbilder, ziel, art):
    """Richtet die Ansichten eines Zielplans aus. Liefert (Anzahl, Hinweise)."""
    ziel_fenster = [doc.GetElement(i) for i in ziel.GetAllViewports()]
    anzahl, hinweise = ar.wie_vorbild(doc, vorbilder, ziel_fenster, art)
    return anzahl, [u"{}: {}".format(ziel.SheetNumber, h) for h in hinweise]


def main():
    doc, uidoc = revit.doc, revit.uidoc
    plan = doc.ActiveView
    if not isinstance(plan, ViewSheet):
        forms.alert(t(u"Bitte den Vorbild-Plan öffnen. Ausgerichtet werden alle "
                      u"seine Ansichten oder nur die markierten.",
                      u"Please open the reference sheet. All its views are "
                      u"aligned, or only the selected ones.",
                      u"Abra el plano de referencia. Se alinean todas sus vistas "
                      u"o sólo las seleccionadas."), title=TITEL)
        return
    vorbilder = vorbild_fenster(doc, uidoc, plan)
    if not vorbilder:
        forms.alert(t(u"Auf diesem Plan liegen keine Ansichten.",
                      u"There are no views on this sheet.",
                      u"No hay vistas en este plano."), title=TITEL)
        return

    arten = dict(ar.art_texte())
    modus = forms.CommandSwitchWindow.show(
        [text for text, _ in ar.art_texte()],
        message=t(u"{} Ansicht(en) von '{}' übertragen - wie ausrichten?",
                  u"Transfer {} view(s) of '{}' - how to align?",
                  u"Transferir {} vista(s) de '{}': ¿cómo alinear?").format(
            len(vorbilder), plan.SheetNumber))
    if not modus:
        return
    art = arten[modus]

    andere = [p for p in rv.plaene(doc) if rv.id_wert(p.Id) != rv.id_wert(plan.Id)]
    texte = forms.SelectFromList.show(
        [rv.plan_text(p) for p in andere], multiselect=True,
        title=t(u"Zielpläne wählen", u"Select target sheets",
                u"Seleccione los planos de destino"),
        button_name=t(u"Ausrichten", u"Align", u"Alinear"))
    if not texte:
        return
    texte = set(texte)
    ziele = [p for p in andere if rv.plan_text(p) in texte]

    ausgerichtet, hinweise = 0, []
    transaktion = Transaction(doc, TITEL)
    transaktion.Start()
    try:
        for ziel in ziele:
            anzahl, texte_plan = plan_ausrichten(doc, vorbilder, ziel, art)
            ausgerichtet += anzahl
            hinweise.extend(texte_plan)
        transaktion.Commit()
    except Exception as fehler:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        forms.alert(t(u"Ausrichten fehlgeschlagen, nichts wurde geändert.",
                      u"Alignment failed, nothing was changed.",
                      u"Error al alinear, no se cambió nada."),
                    sub_msg=u"{}".format(fehler), title=TITEL)
        return

    meldung = t(u"{} Ansicht(en) auf {} Plan/Plänen wie '{}' ausgerichtet.",
                u"{} view(s) on {} sheet(s) aligned like '{}'.",
                u"{} vista(s) en {} plano(s) alineadas como '{}'.").format(
        ausgerichtet, len(ziele), plan.SheetNumber)
    if hinweise:
        forms.alert(meldung, sub_msg=u"\n".join(hinweise[:25]), title=TITEL)
    else:
        forms.alert(meldung, title=TITEL)


main()
