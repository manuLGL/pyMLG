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
from mlg_sprache import t

TITEL = t(u"Kopieren mit Phasen", u"Copy With Phases", u"Copiar con fases")


def auswahl(uidoc):
    ids = list(uidoc.Selection.GetElementIds())
    if ids:
        return ids
    refs = uidoc.Selection.PickObjects(ObjectType.Element,
                                       t(u"Elemente zum Kopieren wählen", u"Select elements to copy", u"Seleccione elementos para copiar"))
    return [r.ElementId for r in refs]


def main():
    doc, uidoc = revit.doc, revit.uidoc
    try:
        ids, uebersprungen = rk.modell_elemente(doc, auswahl(uidoc))
        if not ids:
            forms.alert(t(u"Keine Modellelemente ausgewählt.", u"No model elements selected.", u"No hay elementos de modelo seleccionados."), title=TITEL)
            return
        basis = uidoc.Selection.PickPoint(t(u"Basispunkt wählen", u"Pick base point", u"Designe el punto base"))
        ziel = uidoc.Selection.PickPoint(t(u"Zielpunkt wählen", u"Pick target point", u"Designe el punto de destino"))
    except OperationCanceledException:
        return
    except InvalidOperationException:
        forms.alert(t(u"In dieser Ansicht kann kein Punkt gewählt werden "
                    u"(keine Arbeitsebene). Bitte in einem Grundriss starten "
                    u"oder eine Arbeitsebene festlegen.", u"No point can be picked in this view (no work plane). Please start in a floor plan or set a work plane.", u"No se puede designar un punto en esta vista (sin plano de trabajo). Inicie en una planta o defina un plano de trabajo."), title=TITEL)
        return

    transaktion = Transaction(doc, TITEL)
    transaktion.Start()
    try:
        neue_ids, paare, ohne_original = rk.kopieren(doc, ids, ziel - basis)
        for kopie, original in paare:
            rk.phasen_uebertragen(kopie, original)
        transaktion.Commit()
    except Exception as fehler:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        forms.alert(t(u"Kopieren fehlgeschlagen, nichts wurde geändert.", u"Copy failed, nothing was changed.", u"Error al copiar, no se cambió nada."),
                    sub_msg=u"{}".format(fehler), title=TITEL)
        return

    rk.auswahl_setzen(uidoc, neue_ids)
    hinweise = []
    if ohne_original:
        hinweise.append(t(u"{} Kopie(n) konnte kein Original zugeordnet werden - "
                        u"deren Phasen bitte prüfen.", u"{} copy/copies could not be matched to an original - please check their phases.", u"{} copia(s) no se pudieron asociar a un original: compruebe sus fases.").format(ohne_original))
    if uebersprungen:
        hinweise.append(t(u"{} ansichtsspezifische(s) Element(e) (Beschriftungen, "
                        u"Detaillinien ...) wurden nicht kopiert.", u"{} view-specific element(s) (tags, detail lines ...) were not copied.", u"{} elemento(s) específicos de vista (etiquetas, líneas de detalle ...) no se copiaron.").format(uebersprungen))
    if hinweise:
        forms.alert(t(u"{} Element(e) kopiert.", u"{} element(s) copied.", u"{} elemento(s) copiados.").format(len(neue_ids)),
                    sub_msg=u"\n".join(hinweise), title=TITEL)


main()
