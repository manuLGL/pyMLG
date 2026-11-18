#! python3
# -*- coding: utf-8 -*-
"""GroupToLink: alle Instanzen gewählter Modellgruppen durch Verknüpfungen
der gleichnamigen .rvt-Dateien aus einem Ordner ersetzen.

Die Logik liegt in lib/gruppen_export (Revit: revit.py, Zuordnung
Gruppe -> Datei: logik.py). Die Dateien erzeugt GroupExport.

CPython-Besonderheiten dieses Repos (siehe ScheduleSync.panel/README.md):
pyrevit.forms ist nicht nutzbar - Dialoge kommen aus schedule_sync.ui.
Kein script.exit() - der Ablauf endet nur über 'return'.
"""

import io
import os
import sys
import traceback

_EXT = os.path.dirname(os.path.abspath(__file__))
while not _EXT.endswith(".extension") and os.path.dirname(_EXT) != _EXT:
    _EXT = os.path.dirname(_EXT)
if os.path.join(_EXT, "lib") not in sys.path:
    sys.path.append(os.path.join(_EXT, "lib"))

from mlg_sprache import t  # noqa: E402

TITEL = t(u"Gruppen durch Links ersetzen", u"Replace Groups with Links",
          u"Reemplazar grupos por vínculos")

# So viele Einträge werden in Meldungen einzeln genannt
MAX_ZEILEN = 8


def _liste(eintraege):
    zeilen = [u"• %s" % e for e in eintraege[:MAX_ZEILEN]]
    if len(eintraege) > MAX_ZEILEN:
        zeilen.append(t(u"… und %d weitere", u"… and %d more",
                        u"… y %d más") % (len(eintraege) - MAX_ZEILEN))
    return u"\n".join(zeilen)


def _warnung(hauptzeile, text):
    from schedule_sync import ui
    ui.meldung(text, titel=TITEL, hauptzeile=hauptzeile, warnung=True)


def _startordner(doc):
    pfad = doc.PathName
    if pfad and os.path.isdir(os.path.dirname(pfad)):
        return os.path.dirname(pfad)
    return u""


def _waehle(passend, instanzen):
    """Liste der Gruppen mit Datei zum Ankreuzen. Rückgabe [Gruppenname]
    oder None."""
    from schedule_sync import ui

    namen = sorted(passend, key=lambda n: n.lower())
    texte = [u"%s   (%d×)" % (n, len(instanzen[n])) for n in namen]
    gewaehlt = ui.waehle_mehrfach(
        texte, titel=TITEL,
        hinweis=t(u"Gruppen wählen - alle ihre Instanzen werden durch die "
                  u"gleichnamige Verknüpfung ersetzt:",
                  u"Choose groups - all their instances are replaced with "
                  u"the link of the same name:",
                  u"Elija grupos: todas sus instancias se reemplazan por el "
                  u"vínculo del mismo nombre:"),
        schaltflaeche=t(u"Weiter", u"Next", u"Siguiente"))
    if not gewaehlt:
        return None
    nach_text = dict(zip(texte, namen))
    return [nach_text[x] for x in gewaehlt]


def main():
    from schedule_sync import ui

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        _warnung(t(u"Kein Projekt geöffnet", u"No project open",
                   u"No hay ningún proyecto abierto"),
                 t(u"Gruppen ersetzen geht nur in Projektdateien.",
                   u"Replacing groups only works in project files.",
                   u"Reemplazar grupos solo funciona en archivos de "
                   u"proyecto."))
        return
    if doc.IsReadOnly:
        _warnung(t(u"Gruppen können nicht ersetzt werden",
                   u"Groups cannot be replaced",
                   u"No se pueden reemplazar los grupos"),
                 t(u"Das Dokument ist schreibgeschützt.",
                   u"The document is read-only.",
                   u"El documento es de solo lectura."))
        return

    from gruppen_export import logik as lg
    from gruppen_export import revit as rv

    instanzen, verschachtelt = rv.instanzen(doc)
    vorauswahl = sorted(rv.gewaehlte_gruppen(
        doc, uidoc.Selection.GetElementIds()))
    if vorauswahl:
        kandidaten = [n for n in vorauswahl if n in instanzen]
    else:
        kandidaten = list(instanzen)
    if not kandidaten:
        _warnung(t(u"Keine Modellgruppen gefunden",
                   u"No model groups found",
                   u"No se han encontrado grupos de modelo"),
                 t(u"Es gibt keine Modellgruppe, die ersetzt werden kann "
                   u"(verschachtelte Gruppen zählen nicht).",
                   u"There is no model group that can be replaced (nested "
                   u"groups do not count).",
                   u"No hay ningún grupo de modelo que se pueda reemplazar "
                   u"(los grupos anidados no cuentan)."))
        return

    ordner = ui.ordner_waehlen(
        _startordner(doc),
        titel=t(u"Ordner mit den .rvt-Dateien der Gruppen wählen",
                u"Choose the folder with the groups' .rvt files",
                u"Elija la carpeta con los archivos .rvt de los grupos"))
    if not ordner:
        return
    passend = lg.zuordnung(kandidaten, os.listdir(ordner))
    if not passend:
        _warnung(t(u"Keine passenden Dateien",
                   u"No matching files",
                   u"No hay archivos correspondientes"),
                 t(u"Im Ordner gibt es zu keiner Gruppe eine .rvt-Datei mit "
                   u"ihrem Namen:",
                   u"The folder has no .rvt file named like any of the "
                   u"groups:",
                   u"La carpeta no tiene ningún archivo .rvt con el nombre "
                   u"de alguno de los grupos:")
                 + u"\n" + ordner)
        return

    namen = (sorted(passend, key=lambda n: n.lower()) if vorauswahl
             else _waehle(passend, instanzen))
    if not namen:
        return
    ohne_datei = [n for n in kandidaten if n not in passend]

    anzahl = sum(len(instanzen[n]) for n in namen)
    zeilen = [t(u"%d Gruppen, %d Instanzen:", u"%d groups, %d instances:",
                u"%d grupos, %d instancias:") % (len(namen), anzahl),
              _liste([u"%s  ->  %s" % (n, passend[n]) for n in namen])]
    if vorauswahl and ohne_datei:
        zeilen += [u"", t(u"Ohne Datei im Ordner, bleiben Gruppe:",
                          u"No file in the folder, stay groups:",
                          u"Sin archivo en la carpeta, siguen como grupo:"),
                   _liste(ohne_datei)]
    innen = sum(verschachtelt.get(n, 0) for n in namen)
    if innen:
        zeilen += [u"", t(u"%d verschachtelte Instanzen (in anderen Gruppen) "
                          u"bleiben Gruppe.",
                          u"%d nested instances (inside other groups) stay "
                          u"groups.",
                          u"%d instancias anidadas (dentro de otros grupos) "
                          u"siguen como grupo.") % innen]
    zeilen += [u"", t(u"Die Gruppeninstanzen werden gelöscht, die "
                      u"Gruppentypen bleiben. Ein Rückgängig-Schritt.",
                      u"The group instances are deleted, the group types "
                      u"stay. One undo step.",
                      u"Las instancias de grupo se eliminan, los tipos de "
                      u"grupo se conservan. Un paso de deshacer.")]
    if not ui.frage(u"\n".join(zeilen), titel=TITEL, standard_ja=True,
                    hauptzeile=t(u"Gruppen durch Verknüpfungen ersetzen?",
                                 u"Replace groups with links?",
                                 u"¿Reemplazar grupos por vínculos?")):
        return

    fertig, fehler = rv.ersetze(
        doc, [(n, os.path.join(ordner, passend[n]), instanzen[n])
              for n in namen])

    zeilen = [t(u"%d Instanzen von %d Gruppen ersetzt.",
                u"%d instances of %d groups replaced.",
                u"%d instancias de %d grupos reemplazadas.")
              % (sum(fertig.values()), len(fertig))]
    if fertig:
        zeilen += [u"", _liste([u"%s (%d×)" % (n, fertig[n])
                                for n in sorted(fertig)])]
    if fehler:
        zeilen += [u"", t(u"Nicht ersetzt:", u"Not replaced:",
                          u"No reemplazados:"), _liste(fehler)]
    ui.meldung(u"\n".join(zeilen), titel=TITEL, warnung=bool(fehler),
               hauptzeile=t(u"Fertig", u"Done", u"Terminado") if not fehler
               else t(u"Fertig mit Fehlern", u"Done with errors",
                      u"Terminado con errores"))


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "GroupToLink_Fehler.log")
    try:
        if not os.path.isdir(os.path.dirname(protokoll)):
            os.makedirs(os.path.dirname(protokoll))
        with io.open(protokoll, "a", encoding="utf-8") as datei:
            datei.write(spur + u"\n" + u"-" * 70 + u"\n")
    except Exception:
        protokoll = None
    try:
        letzte_zeile = (spur.strip().splitlines() or [u""])[-1]
        _warnung(t(u"Gruppen ersetzen ist fehlgeschlagen",
                   u"Replace Groups with Links failed",
                   u"Reemplazar grupos por vínculos ha fallado"),
                 letzte_zeile + (t(u"\n\nProtokoll: %s", u"\n\nLog: %s",
                                   u"\n\nRegistro: %s") % protokoll
                                 if protokoll else u""))
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
