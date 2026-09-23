# -*- coding: utf-8 -*-
# Setzt die Schnittbox aus der Zwischenablage in die aktive 3D-Ansicht -
# an dieselbe Stelle wie im Herkunftsprojekt (gemeinsame Koordinaten) und
# zoomt darauf. Fuer den Weg vom Koordinationsmodell ins Bearbeitungsmodell.
# (Kommentar statt Docstring: pyRevit liest unter IronPython den
#  Docstring als Tooltip und scheitert dabei an Umlauten.)

__title__ = "Section Box Paste"
__author__ = "Manuel"

from Autodesk.Revit.DB import View3D
from pyrevit import forms, revit

from mlg_sprache import t
from section_box import logik as lg
from section_box import revit as sb
from section_box import zwischenablage

TITEL = t(u"Schnittbox einfügen", u"Paste Section Box",
          u"Pegar caja de sección")

# Ab diesem Versatz zwischen gemeinsamen und internen Koordinaten lohnt die
# Rueckfrage (Meter)
SCHWELLE = 0.01


def waehle_kasten(doc, box):
    """Kasten in Projektkoordinaten. Weichen gemeinsame und interne Lage
    voneinander ab, entscheidet der Anwender."""
    gemeinsam = sb.baue_kasten(doc, box, gemeinsam=True)
    intern = sb.baue_kasten(doc, box, gemeinsam=False)

    versatz = sb.abstand_m(gemeinsam, intern)
    if versatz < SCHWELLE:
        return gemeinsam

    nach_gemeinsam = t(u"Gemeinsame Koordinaten (empfohlen)",
                       u"Shared coordinates (recommended)",
                       u"Coordenadas compartidas (recomendado)")
    nach_intern = t(u"Projektinterne Koordinaten",
                    u"Project internal coordinates",
                    u"Coordenadas internas del proyecto")
    antwort = forms.alert(
        t(u"Die beiden Projekte sind unterschiedlich verortet - die beiden "
          u"Deutungen liegen %.2f m auseinander.\n\n"
          u"Gemeinsame Koordinaten sind richtig, wenn beide Modelle "
          u"gemeinsame Koordinaten teilen (Vermessungspunkt).\n"
          u"Projektintern ist richtig, wenn die Modelle Ursprung auf "
          u"Ursprung verknüpft sind.",
          u"The two projects are positioned differently - the two readings "
          u"are %.2f m apart.\n\n"
          u"Shared coordinates are right if both models share coordinates "
          u"(survey point).\n"
          u"Project internal is right if the models are linked origin to "
          u"origin.",
          u"Los dos proyectos están ubicados de forma distinta: las dos "
          u"lecturas difieren en %.2f m.\n\n"
          u"Las coordenadas compartidas son correctas si ambos modelos "
          u"comparten coordenadas (punto de topografía).\n"
          u"Las internas lo son si los modelos están vinculados origen con "
          u"origen.") % versatz,
        title=TITEL, options=[nach_gemeinsam, nach_intern])

    if antwort == nach_gemeinsam:
        return gemeinsam
    if antwort == nach_intern:
        return intern
    return None


def main():
    doc = revit.doc
    uidoc = revit.uidoc
    ansicht = revit.active_view

    # Erst die Zwischenablage pruefen - sonst wechselt die Ansicht umsonst
    try:
        box = lg.lies_text(zwischenablage.lies())
    except lg.Fehler as fehler:
        forms.alert(u"%s" % fehler, title=TITEL, exitscript=True)

    if not isinstance(ansicht, View3D):
        # Keine 3D-Ansicht offen: die Standard-3D-Ansicht oeffnen
        ansicht = sb.standard_3d(uidoc, doc)
        if ansicht is None:
            forms.alert(t(u"In diesem Projekt lässt sich keine 3D-Ansicht "
                          u"öffnen.",
                          u"No 3D view can be opened in this project.",
                          u"No se puede abrir ninguna vista 3D en este "
                          u"proyecto."),
                        title=TITEL, exitscript=True)

    kasten = waehle_kasten(doc, box)
    if kasten is None:
        return

    try:
        sb.setze_box(doc, ansicht, kasten, TITEL)
    except Exception as fehler:
        forms.alert(t(u"Die Schnittbox ließ sich nicht setzen.\n\n%s\n\n"
                      u"Steuert eine Ansichtsvorlage die Schnittbox dieser "
                      u"Ansicht, muss sie dort freigegeben werden.",
                      u"The section box could not be set.\n\n%s\n\n"
                      u"If a view template controls the section box of this "
                      u"view, it has to be released there.",
                      u"No se pudo aplicar la caja de sección.\n\n%s\n\n"
                      u"Si una plantilla de vista controla la caja de "
                      u"sección, hay que liberarla allí.") % fehler,
                    title=TITEL, exitscript=True)

    sb.zeige(uidoc, ansicht, kasten)
    uidoc.RefreshActiveView()

    herkunft = box.quelle
    forms.alert(t(u"Schnittbox gesetzt.\n\nGröße: %s\nHerkunft: %s",
                  u"Section box applied.\n\nSize: %s\nSource: %s",
                  u"Caja de sección aplicada.\n\nTamaño: %s\nOrigen: %s")
                % (box.masse_text(),
                   herkunft or t(u"unbekannt", u"unknown", u"desconocido")),
                title=TITEL)


main()
