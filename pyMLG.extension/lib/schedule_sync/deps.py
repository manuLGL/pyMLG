# -*- coding: utf-8 -*-
"""Lädt die externen Python-Abhängigkeiten für die pyRevit-CPython3-Engine.

openpyxl ist in der pyRevit-CPython-Umgebung standardmässig NICHT enthalten.
Dieses Modul sucht die Bibliothek an mehreren Stellen und liefert andernfalls
eine verständliche deutsche Fehlermeldung mit Installationsanleitung.
"""

import os
import sys

from mlg_sprache import t


class AbhaengigkeitFehlt(Exception):
    """Wird geworfen, wenn openpyxl nicht importiert werden kann."""


INSTALL_HINWEIS = (
    t(u"Die Bibliothek 'openpyxl' wurde in der pyRevit-CPython-Umgebung nicht gefunden.\n\n"
    u"Installation (einmalig, ausserhalb von Revit in einer Eingabeaufforderung):\n\n"
    u"    pip install --target \"{ziel}\" openpyxl\n\n"
    u"Dabei muss 'pip' zu einer CPython-Version gehören, die zur pyRevit-CPython-Engine\n"
    u"passt (pyRevit 5.x = Python 3.12, pyRevit 4.8 = Python 3.8). openpyxl ist reines\n"
    u"Python, daher ist die Nebenversion unkritisch.\n\n"
    u"Alternativ: In den pyRevit-Einstellungen unter 'Custom Search Paths' den Ordner\n"
    u"ergänzen, der openpyxl enthält, und Revit neu starten.", u"The library 'openpyxl' was not found in the pyRevit CPython environment.\n\nInstallation (once, outside Revit in a command prompt):\n\n    pip install --target \"{ziel}\" openpyxl\n\n'pip' must belong to a CPython version matching the pyRevit CPython engine\n(pyRevit 5.x = Python 3.12, pyRevit 4.8 = Python 3.8). openpyxl is pure\nPython, so the minor version does not matter.\n\nAlternatively: add the folder containing openpyxl under 'Custom Search Paths'\nin the pyRevit settings and restart Revit.", u"No se encontró la biblioteca 'openpyxl' en el entorno CPython de pyRevit.\n\nInstalación (una vez, fuera de Revit en un símbolo del sistema):\n\n    pip install --target \"{ziel}\" openpyxl\n\n'pip' debe pertenecer a una versión de CPython compatible con el motor CPython de pyRevit\n(pyRevit 5.x = Python 3.12, pyRevit 4.8 = Python 3.8). openpyxl es Python\npuro, así que la versión menor no importa.\n\nAlternativa: añada la carpeta que contiene openpyxl en 'Custom Search Paths'\nde la configuración de pyRevit y reinicie Revit.")
)


def extension_ordner():
    """Pfad des Extension-Ordners (…/pyMLG.extension)."""
    # deps.py -> schedule_sync -> lib -> <extension>
    hier = os.path.dirname(os.path.abspath(__file__))
    return os.path.dirname(os.path.dirname(hier))


def _suchpfade():
    """Zusätzliche Ordner, in denen openpyxl liegen darf."""
    ext = extension_ordner()
    repo = os.path.dirname(ext)
    return [
        # Empfohlener Ort: wird mit der Extension ausgeliefert und versioniert
        os.path.join(ext, "site-packages"),
        # Entwicklungs-Setup dieses Repos (.venv neben dem Extension-Ordner)
        os.path.join(repo, ".venv", "Lib", "site-packages"),
    ]


def ist_cpython():
    """True, wenn das Skript unter der CPython3-Engine läuft (nicht IronPython)."""
    version = sys.version.lower()
    return not version.startswith("2.") and "ironpython" not in version


def lade_openpyxl():
    """Importiert openpyxl und ergänzt bei Bedarf sys.path.

    Rückgabe: das openpyxl-Modul.
    Wirft AbhaengigkeitFehlt mit Installationsanleitung, wenn nichts gefunden wird.
    """
    try:
        import openpyxl
        return openpyxl
    except ImportError:
        pass

    for pfad in _suchpfade():
        if os.path.isdir(pfad) and pfad not in sys.path:
            sys.path.append(pfad)

    try:
        import openpyxl
        return openpyxl
    except ImportError:
        ziel = os.path.join(extension_ordner(), "site-packages")
        hinweis = INSTALL_HINWEIS.format(ziel=ziel)
        if not ist_cpython():
            hinweis = (
                t(u"Dieses Werkzeug benötigt die CPython3-Engine von pyRevit.\n"
                u"Die erste Zeile von script.py muss '#! python3' lauten.\n\n", u"This tool requires the pyRevit CPython3 engine.\nThe first line of script.py must be '#! python3'.\n\n", u"Esta herramienta requiere el motor CPython3 de pyRevit.\nLa primera línea de script.py debe ser '#! python3'.\n\n")
            ) + hinweis
        raise AbhaengigkeitFehlt(hinweis)
