# -*- coding: utf-8 -*-
# Text aus der Zwischenablage holen und hineinlegen.
#
# System.Windows.Clipboard steckt in PresentationCore - unter IronPython muss
# die Assembly erst geladen werden, sonst: "No module named Windows".

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Clipboard  # noqa: E402


def schreibe(text):
    Clipboard.SetText(text)


def lies():
    try:
        if Clipboard.ContainsText():
            return Clipboard.GetText()
    except Exception:
        pass
    return u""
