# -*- coding: utf-8 -*-
# Laeuft einmal beim Laden der Extension (IronPython, API-Kontext).
#
# Meldet das andockbare Panel des Clash Navigators bei Revit an - das geht
# nur beim Hochfahren. Ein Fehler hier darf den Start von Revit nicht stoeren:
# er landet nur im Protokoll %LOCALAPPDATA%\pyMLG\ClashNavigator_Fehler.log.

import io
import os
import traceback

try:
    from clash_navigator import panel as _clash_panel
    _clash_panel.registriere()
except Exception:
    try:
        _basis = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        _ordner = os.path.join(_basis, "pyMLG")
        if not os.path.isdir(_ordner):
            os.makedirs(_ordner)
        with io.open(os.path.join(_ordner, "ClashNavigator_Fehler.log"), "a",
                     encoding="utf-8") as _datei:
            _datei.write(u"startup.py\n" + traceback.format_exc()
                         + u"\n" + u"-" * 70 + u"\n")
    except Exception:
        pass
