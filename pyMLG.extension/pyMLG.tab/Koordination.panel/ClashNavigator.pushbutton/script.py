# -*- coding: utf-8 -*-
# Oeffnet das andockbare Panel des Clash Navigators. Angemeldet wird es beim
# Start von Revit in startup.py der Extension (lib/clash_navigator/panel.py).
# (Kommentar statt Docstring: pyRevit liest unter IronPython den
#  Docstring als Tooltip und scheitert dabei an Umlauten.)

__title__ = "Clash Navigator"
__author__ = "Manuel"
__context__ = "zero-doc"

from pyrevit import forms

from clash_navigator import PANEL_ID
from mlg_sprache import t


def main():
    try:
        forms.open_dockable_panel(PANEL_ID)
    except Exception as fehler:
        forms.alert(t(u"Das Panel ist noch nicht angemeldet.\n\n"
                      u"Revit einmal neu starten - andockbare Panels "
                      u"meldet Revit nur beim Hochfahren an. Ein "
                      u"pyRevit-Reload genügt nicht.\n\nDetails: %s\n"
                      u"Protokoll: %%LOCALAPPDATA%%\\pyMLG\\"
                      u"ClashNavigator_Fehler.log",
                      u"The panel is not registered yet.\n\n"
                      u"Restart Revit once - Revit only registers dockable "
                      u"panels during start-up. A pyRevit reload is not "
                      u"enough.\n\nDetails: %s\n"
                      u"Log: %%LOCALAPPDATA%%\\pyMLG\\"
                      u"ClashNavigator_Fehler.log",
                      u"El panel todavía no está registrado.\n\n"
                      u"Reinicie Revit una vez: Revit solo registra paneles "
                      u"acoplables al arrancar. No basta con recargar "
                      u"pyRevit.\n\nDetalles: %s\n"
                      u"Registro: %%LOCALAPPDATA%%\\pyMLG\\"
                      u"ClashNavigator_Fehler.log") % fehler,
                    title=u"Clash Navigator")


main()
