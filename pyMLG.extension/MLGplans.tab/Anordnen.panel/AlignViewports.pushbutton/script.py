# -*- coding: utf-8 -*-
# Richtet andere Pläne wie den aktiven Plan aus: jede Ansicht des aktiven
# Plans (oder nur die markierten) bekommt auf jedem Zielplan ihr Gegenstück
# - gleiche Ansichtsart, Vorrang für dieselbe Ansicht, dann gleicher
# Maßstab, dann die nächstliegende. Das Gegenstück übernimmt Lage,
# Ansichtsfenster-Typ und Titellage:
#   - deckungsgleich im Modell: der Projektursprung liegt gleich
#     (Grundrisse EG bis DG übereinander); Legenden und Zeichenansichten
#     nach der Mitte
#   - gleiche Mitte: die Mitte des Ansichtsfensters liegt gleich
# (Kommentar statt Docstring: pyRevit liest unter IronPython den
#  Docstring als Tooltip und scheitert dabei an Umlauten.)

__title__ = "AlignViewports"
__author__ = "Manuel"

from Autodesk.Revit.DB import Transaction, Viewport, ViewSheet, ViewType
from pyrevit import forms, revit

from mlg_plaene import logik as lg
from mlg_plaene import revit as rv
from mlg_sprache import t

TITEL = t(u"Ansichten ausrichten", u"Align Viewports", u"Alinear vistas")
MODELL = t(u"Deckungsgleich im Modell", u"Matching in the model",
           u"Coincidente en el modelo")
MITTE = t(u"Gleiche Mitte auf dem Plan", u"Same centre on the sheet",
          u"Mismo centro en el plano")

OHNE_MODELL = (ViewType.Legend, ViewType.DraftingView)


def vorbild_fenster(doc, uidoc, plan):
    """Markierte Ansichtsfenster des Plans, sonst alle."""
    gewaehlt = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    gewaehlt = [e for e in gewaehlt if isinstance(e, Viewport)
                and rv.id_wert(e.SheetId) == rv.id_wert(plan.Id)]
    if gewaehlt:
        return gewaehlt
    return [doc.GetElement(i) for i in plan.GetAllViewports()]


def info(doc, fenster):
    ansicht = doc.GetElement(fenster.ViewId)
    mitte = fenster.GetBoxCenter()
    return lg.Fenster(rv.id_wert(fenster.Id), str(ansicht.ViewType), ansicht.Scale,
                      rv.id_wert(ansicht.Id), (mitte.X, mitte.Y))


def aussehen_uebernehmen(vorbild, fenster):
    """Ansichtsfenster-Typ und Titellage wie beim Vorbild."""
    if rv.id_wert(fenster.GetTypeId()) != rv.id_wert(vorbild.GetTypeId()):
        fenster.ChangeTypeId(vorbild.GetTypeId())
    try:
        fenster.LabelOffset = vorbild.LabelOffset
        fenster.LabelLineLength = vorbild.LabelLineLength
    except Exception:
        pass


def verschiebung(doc, vorbild, fenster, modus):
    """(Verschiebung, Hinweis oder None) für fenster."""
    ansicht = doc.GetElement(fenster.ViewId)
    vorbild_ansicht = doc.GetElement(vorbild.ViewId)
    if modus == MODELL and ansicht.ViewType not in OHNE_MODELL:
        soll = rv.modell_auf_plan(vorbild, vorbild_ansicht)
        ist = rv.modell_auf_plan(fenster, ansicht)
        if soll is not None and ist is not None:
            hinweis = None
            if ansicht.Scale != vorbild_ansicht.Scale:
                hinweis = t(u"'{}': anderer Maßstab (1:{}) - nur der Ursprung liegt gleich",
                            u"'{}': different scale (1:{}) - only the origin matches",
                            u"'{}': otra escala (1:{}): sólo coincide el origen").format(
                    ansicht.Name, ansicht.Scale)
            return soll - ist, hinweis
    return vorbild.GetBoxCenter() - fenster.GetBoxCenter(), None


def plan_ausrichten(doc, vorbilder, ziel, modus):
    """Richtet die Ansichten eines Zielplans aus. Liefert (Anzahl, Hinweise)."""
    ziel_fenster = [doc.GetElement(i) for i in ziel.GetAllViewports()]
    nach_id = dict((rv.id_wert(f.Id), f) for f in vorbilder + ziel_fenster)
    paare = lg.fenster_zuordnen([info(doc, f) for f in vorbilder],
                                [info(doc, f) for f in ziel_fenster])
    hinweise = []
    for vorbild_id, fenster_id in paare:
        vorbild, fenster = nach_id[vorbild_id], nach_id[fenster_id]
        aussehen_uebernehmen(vorbild, fenster)
        doc.Regenerate()
        weg, hinweis = verschiebung(doc, vorbild, fenster, modus)
        fenster.SetBoxCenter(fenster.GetBoxCenter() + weg)
        if hinweis:
            hinweise.append(hinweis)

    zugeordnet = set(v for v, _ in paare)
    for vorbild in vorbilder:
        if rv.id_wert(vorbild.Id) not in zugeordnet:
            hinweise.append(t(u"kein Gegenstück für '{}'", u"no counterpart for '{}'",
                              u"sin equivalente para '{}'").format(
                doc.GetElement(vorbild.ViewId).Name))
    return len(paare), [u"{}: {}".format(ziel.SheetNumber, h) for h in hinweise]


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

    modus = forms.CommandSwitchWindow.show(
        [MODELL, MITTE],
        message=t(u"{} Ansicht(en) von '{}' übertragen - wie ausrichten?",
                  u"Transfer {} view(s) of '{}' - how to align?",
                  u"Transferir {} vista(s) de '{}': ¿cómo alinear?").format(
            len(vorbilder), plan.SheetNumber))
    if not modus:
        return

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
            anzahl, texte_plan = plan_ausrichten(doc, vorbilder, ziel, modus)
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
