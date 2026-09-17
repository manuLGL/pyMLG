# -*- coding: utf-8 -*-
"""Fügt die ausgewählten Elemente an gleicher Grundrissposition auf der Ebene
der aktiven Ansicht ein (wie "Ausgerichtet an aktueller Ansicht einfügen")
und behält "Phase erstellt" und "Phase abgebrochen" der Originale bei.

Elemente von verschiedenen Ebenen behalten ihren Höhenabstand zueinander.
Versätze zur Ebene bleiben erhalten; obere Abhängigkeiten (z.B. Wand bis
Ebene 2) werden auf die entsprechend höhere Ebene gesetzt - gibt es die
nicht, wird eine Wand "nicht verbunden" mit gleicher Höhe."""

__title__ = "Paste Aligned View"
__author__ = "Manuel"

from Autodesk.Revit.DB import Transaction, XYZ
from pyrevit import forms, revit

from phasen import revit_kopie as rk
from mlg_sprache import t

TITEL = t(u"Paste Aligned View (mit Phasen)", u"Paste Aligned View (with phases)", u"Pegar alineado (con fases)")


def gruppieren(doc, ids, ziel_ebene):
    """{Höhenunterschied: [Ids]} je Ausgangsebene.

    Elemente ohne erkennbare Ebene kommen zur grössten Gruppe.
    """
    gruppen, ohne_ebene = {}, []
    for eid in ids:
        ebene = rk.basis_ebene(doc, doc.GetElement(eid))
        if ebene is None:
            ohne_ebene.append(eid)
            continue
        delta = round(ziel_ebene.Elevation - ebene.Elevation, 6)
        gruppen.setdefault(delta, []).append(eid)
    if ohne_ebene and gruppen:
        groesste = max(gruppen, key=lambda d: len(gruppen[d]))
        gruppen[groesste].extend(ohne_ebene)
    return gruppen


def main():
    doc, uidoc = revit.doc, revit.uidoc
    ids, uebersprungen = rk.modell_elemente(doc, uidoc.Selection.GetElementIds())
    if not ids:
        forms.alert(t(u"Keine Modellelemente ausgewählt.\n\nElemente auswählen, "
                    u"in den Grundriss der Zielebene wechseln und das Werkzeug "
                    u"starten.", u"No model elements selected.\n\nSelect elements, switch to the floor plan of the target level and start the tool.", u"No hay elementos de modelo seleccionados.\n\nSeleccione elementos, cambie a la planta del nivel de destino e inicie la herramienta."), title=TITEL)
        return

    ziel_ebene = getattr(doc.ActiveView, "GenLevel", None)
    if ziel_ebene is None:
        forms.alert(t(u"Die aktive Ansicht hat keine Ebene. Bitte einen Grundriss "
                    u"der Zielebene öffnen.", u"The active view has no level. Please open a floor plan of the target level.", u"La vista activa no tiene nivel. Abra una planta del nivel de destino."), title=TITEL)
        return

    gruppen = gruppieren(doc, ids, ziel_ebene)
    if not gruppen:
        forms.alert(t(u"Für die Auswahl ließ sich keine Ebene ermitteln.", u"No level could be determined for the selection.", u"No se pudo determinar un nivel para la selección."), title=TITEL)
        return
    if all(abs(d) < rk.HOEHEN_TOLERANZ for d in gruppen):
        forms.alert(t(u"Die Elemente liegen bereits auf '{}'.", u"The elements are already on '{}'.", u"Los elementos ya están en '{}'.").format(ziel_ebene.Name),
                    title=TITEL)
        return

    alle_ebenen = rk.ebenen(doc)
    alle_neuen, hinweise, ohne_original = [], [], 0

    transaktion = Transaction(doc, TITEL)
    transaktion.Start()
    try:
        for delta, gruppe in gruppen.items():
            neue_ids, paare, ohne = rk.kopieren(doc, gruppe, XYZ(0, 0, delta))
            alle_neuen.extend(neue_ids)
            ohne_original += ohne
            for kopie, original in paare:
                for text in rk.ebenen_anpassen(doc, kopie, original, delta, alle_ebenen):
                    hinweise.append(u"{} ({}): {}".format(
                        kopie.Name, rk.id_wert(kopie.Id), text))
            rk.hoehen_korrigieren(doc, paare, delta)
            for kopie, original in paare:
                rk.phasen_uebertragen(kopie, original)
        transaktion.Commit()
    except Exception as fehler:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        forms.alert(t(u"Einfügen fehlgeschlagen, nichts wurde geändert.", u"Paste failed, nothing was changed.", u"Error al pegar, no se cambió nada."),
                    sub_msg=u"{}".format(fehler), title=TITEL)
        return

    rk.auswahl_setzen(uidoc, alle_neuen)
    if ohne_original:
        hinweise.insert(0, t(u"{} Kopie(n) konnte kein Original zugeordnet werden - "
                           u"deren Phasen bitte prüfen.", u"{} copy/copies could not be matched to an original - please check their phases.", u"{} copia(s) no se pudieron asociar a un original: compruebe sus fases.").format(ohne_original))
    if uebersprungen:
        hinweise.insert(0, t(u"{} ansichtsspezifische(s) Element(e) wurden nicht "
                           u"kopiert.", u"{} view-specific element(s) were not copied.", u"{} elemento(s) específicos de vista no se copiaron.").format(uebersprungen))
    if hinweise:
        forms.alert(t(u"{} Element(e) auf '{}' eingefügt.", u"{} element(s) pasted on '{}'.", u"{} elemento(s) pegados en '{}'.").format(
                        len(alle_neuen), ziel_ebene.Name),
                    sub_msg=u"\n".join(hinweise[:15]), title=TITEL)


main()
