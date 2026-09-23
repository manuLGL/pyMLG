# -*- coding: utf-8 -*-
# Legt eine Schnittbox um die gewaehlten Elemente - auch um Elemente aus
# Verknuepfungen, die mit Tab angetippt wurden. Ohne Auswahl fragt das
# Werkzeug nach den Elementen in der Verknuepfung.
# (Kommentar statt Docstring: pyRevit liest unter IronPython den
#  Docstring als Tooltip und scheitert dabei an Umlauten.)

__title__ = "Section Box Selection"
__author__ = "Manuel"

from Autodesk.Revit.DB import View3D
from pyrevit import forms, revit

from mlg_sprache import t
from section_box import auswahl
from section_box import revit as sb

TITEL = t(u"Schnittbox um Auswahl", u"Section Box around Selection",
          u"Caja de sección alrededor de la selección")


def hole_ecken(uidoc):
    """Ecken der Auswahl - ohne Auswahl wird gepickt. None: abgebrochen."""
    ecken = auswahl.ecken_der_auswahl(uidoc)
    if ecken:
        return ecken
    return auswahl.picke(uidoc)


def main():
    doc = revit.doc
    uidoc = revit.uidoc
    ansicht = revit.active_view

    ecken = hole_ecken(uidoc)
    if ecken is None:
        return
    if not ecken:
        forms.alert(t(u"Die gewählten Elemente haben keine Ausdehnung im "
                      u"Modell.\n\nElemente in einer Verknüpfung mit Tab "
                      u"antippen, eigene Elemente vorher auswählen.",
                      u"The selected elements have no extent in the "
                      u"model.\n\nTab onto elements inside a link, select "
                      u"your own elements beforehand.",
                      u"Los elementos seleccionados no tienen extensión en "
                      u"el modelo.\n\nSeñale con Tab los elementos de un "
                      u"vínculo; los propios, selecciónelos antes."),
                    title=TITEL, exitscript=True)

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

    kasten = auswahl.kasten_um(ecken)

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


main()
