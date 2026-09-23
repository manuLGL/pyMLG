# -*- coding: utf-8 -*-
# Kopiert die Schnittbox der aktiven 3D-Ansicht als Text in die
# Zwischenablage - fuer "Schnittbox einfuegen" in einem anderen Projekt
# oder zum Weitergeben per Mail.
# (Kommentar statt Docstring: pyRevit liest unter IronPython den
#  Docstring als Tooltip und scheitert dabei an Umlauten.)

__title__ = "Section Box Copy"
__author__ = "Manuel"

from datetime import datetime

from Autodesk.Revit.DB import View3D
from pyrevit import forms, revit

from mlg_sprache import t
from section_box import logik as lg
from section_box import revit as sb
from section_box import zwischenablage

TITEL = t(u"Schnittbox kopieren", u"Copy Section Box",
          u"Copiar caja de sección")


def main():
    doc = revit.doc
    ansicht = revit.active_view

    if not isinstance(ansicht, View3D):
        forms.alert(t(u"Bitte eine 3D-Ansicht öffnen - nur dort gibt es eine "
                      u"Schnittbox.",
                      u"Please open a 3D view - only there is a section box.",
                      u"Abra una vista 3D: sólo allí hay una caja de "
                      u"sección."),
                    title=TITEL, exitscript=True)

    if not ansicht.IsSectionBoxActive:
        forms.alert(t(u"In dieser 3D-Ansicht ist keine Schnittbox aktiv.",
                      u"No section box is active in this 3D view.",
                      u"No hay ninguna caja de sección activa en esta vista "
                      u"3D."),
                    title=TITEL, exitscript=True)

    quelle = u"%s | %s | %s" % (doc.Title, ansicht.Name,
                                datetime.now().strftime("%Y-%m-%d %H:%M"))
    box = sb.lies_box(doc, ansicht, quelle=quelle)
    zwischenablage.schreibe(lg.erzeuge_text(box))

    forms.alert(t(u"Die Schnittbox liegt als Text in der Zwischenablage.\n\n"
                  u"Größe: %s\n\n"
                  u"Im anderen Projekt eine 3D-Ansicht öffnen und "
                  u"\"Schnittbox einfügen\" wählen.",
                  u"The section box is in the clipboard as text.\n\n"
                  u"Size: %s\n\n"
                  u"In the other project, open a 3D view and choose "
                  u"\"Paste Section Box\".",
                  u"La caja de sección está en el portapapeles como texto.\n\n"
                  u"Tamaño: %s\n\n"
                  u"En el otro proyecto, abra una vista 3D y elija "
                  u"\"Pegar caja de sección\".") % box.masse_text(),
                title=TITEL)


main()
