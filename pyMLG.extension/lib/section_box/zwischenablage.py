# -*- coding: utf-8 -*-
# Text aus der Zwischenablage holen und hineinlegen.
#
# System.Windows.Clipboard steckt in PresentationCore - unter IronPython muss
# die Assembly erst geladen werden, sonst: "No module named Windows".

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

import time  # noqa: E402

from System.Windows import Clipboard  # noqa: E402

VERSUCHE = 10
PAUSE = 0.1


def schreibe(text):
    # Andere Programme (Zwischenablage-Verlauf, OneDrive, Teams, RDP ...)
    # halten die Zwischenablage oft kurz offen - dann scheitert OpenClipboard
    # mit CLIPBRD_E_CANT_OPEN. WPF versucht es nicht erneut, also selbst.
    # SetText ruft intern Flush(); scheitert nur das, bleibt als letzter
    # Versuch SetDataObject ohne Flush (Text gilt, solange Revit laeuft).
    # Gibt False zurueck, wenn es gar nicht geklappt hat.
    for _ in range(VERSUCHE):
        try:
            Clipboard.SetText(text)
            return True
        except Exception:
            time.sleep(PAUSE)
    for _ in range(VERSUCHE):
        try:
            Clipboard.SetDataObject(text, False)
            return True
        except Exception:
            time.sleep(PAUSE)
    return False


def lies():
    try:
        if Clipboard.ContainsText():
            return Clipboard.GetText()
    except Exception:
        pass
    return u""
