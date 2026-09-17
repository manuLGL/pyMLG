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

TITEL = u"Paste Aligned View (mit Phasen)"


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
        forms.alert(u"Keine Modellelemente ausgewählt.\n\nElemente auswählen, "
                    u"in den Grundriss der Zielebene wechseln und das Werkzeug "
                    u"starten.", title=TITEL)
        return

    ziel_ebene = getattr(doc.ActiveView, "GenLevel", None)
    if ziel_ebene is None:
        forms.alert(u"Die aktive Ansicht hat keine Ebene. Bitte einen Grundriss "
                    u"der Zielebene öffnen.", title=TITEL)
        return

    gruppen = gruppieren(doc, ids, ziel_ebene)
    if not gruppen:
        forms.alert(u"Für die Auswahl ließ sich keine Ebene ermitteln.", title=TITEL)
        return
    if all(abs(d) < rk.HOEHEN_TOLERANZ for d in gruppen):
        forms.alert(u"Die Elemente liegen bereits auf '{}'.".format(ziel_ebene.Name),
                    title=TITEL)
        return

    alle_ebenen = rk.ebenen(doc)
    alle_neuen, hinweise, ohne_original = [], [], 0

    t = Transaction(doc, TITEL)
    t.Start()
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
        t.Commit()
    except Exception as fehler:
        if t.HasStarted() and not t.HasEnded():
            t.RollBack()
        forms.alert(u"Einfügen fehlgeschlagen, nichts wurde geändert.",
                    sub_msg=u"{}".format(fehler), title=TITEL)
        return

    rk.auswahl_setzen(uidoc, alle_neuen)
    if ohne_original:
        hinweise.insert(0, u"{} Kopie(n) konnte kein Original zugeordnet werden - "
                           u"deren Phasen bitte prüfen.".format(ohne_original))
    if uebersprungen:
        hinweise.insert(0, u"{} ansichtsspezifische(s) Element(e) wurden nicht "
                           u"kopiert.".format(uebersprungen))
    if hinweise:
        forms.alert(u"{} Element(e) auf '{}' eingefügt.".format(
                        len(alle_neuen), ziel_ebene.Name),
                    sub_msg=u"\n".join(hinweise[:15]), title=TITEL)


main()
