#! python3
# -*- coding: utf-8 -*-
"""StackValues: Parameterwerte gewählter Stützen, Wände und Unterzüge auf
die Elemente an derselben Stelle (XY) in anderen Geschossen übertragen.

Die Logik liegt in lib/werte_stapel (Fenster: fenster.py, Revit:
revit.py, Lage und Plan: logik.py).

CPython-Besonderheiten dieses Repos (siehe ScheduleSync.panel/README.md):
pyrevit.forms ist nicht nutzbar - das Fenster ist eigenes WPF per XAML.
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

from mlg_sprache import sprache, t  # noqa: E402

TITEL = t(u"Werte nach oben/unten übertragen", u"Copy Values Up/Down",
          u"Copiar valores arriba/abajo")

# So viele Einträge werden in Meldungen einzeln genannt
MAX_ZEILEN = 6


def _liste(eintraege):
    zeilen = [u"• %s" % e for e in eintraege[:MAX_ZEILEN]]
    if len(eintraege) > MAX_ZEILEN:
        zeilen.append(t(u"… und %d weitere", u"… and %d more",
                        u"… y %d más") % (len(eintraege) - MAX_ZEILEN))
    return u"\n".join(zeilen)


def _abschnitt(zeilen, titel, eintraege):
    if eintraege:
        zeilen += [u"", u"%s (%d):" % (titel, len(eintraege)),
                   _liste(eintraege)]


def _laenge(fuss):
    text = u"%.2f m" % (fuss * 0.3048)
    return text if sprache() == u"en" else text.replace(u".", u",")


def _ohne_gegenstueck(zeilen, bestand, plan, einst, ebenen_namen):
    """Quellen ohne Gegenstück - mit dem Grund, damit man sieht, ob die
    Toleranz, die Option 'gleiche Achse' oder die Geschosswahl hilft.
    Hängt den Abschnitt an zeilen an und liefert die Tipps (die gehören
    nach oben, lange Meldungen werden am Ende gekürzt)."""
    from werte_stapel import logik as lg
    from werte_stapel import revit as rv

    if not plan.ohne_gegenstueck:
        return []
    gruende = lg.warum(bestand.quellen, bestand.ziele, einst.zuordnung,
                       einst.ebenen, einst.toleranz, plan.ohne_gegenstueck)

    def ebene(zi):
        return ebenen_namen.get(bestand.ziele[zi].ebene, u"?")

    def text(qi):
        grund, zi, weg = gruende[qi]
        if grund == lg.NICHT_GEWAEHLT:
            erklaerung = t(u"nur auf nicht gewähltem Geschoss %s",
                           u"only on unchecked level %s",
                           u"solo en el nivel no marcado %s") % ebene(zi)
        elif grund == lg.ANDERE_LAENGE:
            erklaerung = t(u"gleiche Achse auf %s, Enden weichen %s ab",
                           u"same axis on %s, ends differ by %s",
                           u"mismo eje en %s, los extremos difieren %s")                 % (ebene(zi), _laenge(weg))
        elif grund == lg.DANEBEN:
            erklaerung = t(u"auf %s %s daneben",
                           u"on %s %s off",
                           u"en %s desplazado %s") % (ebene(zi),
                                                       _laenge(weg))
        else:
            erklaerung = t(u"nichts darüber/darunter",
                           u"nothing above/below",
                           u"nada encima/debajo")
        return u"%s: %s" % (rv.beschreibe(bestand.quell_revit[qi]),
                            erklaerung)

    # Häufigster Grund zuerst - wiederholt sich meist
    reihenfolge = (lg.ANDERE_LAENGE, lg.NICHT_GEWAEHLT, lg.DANEBEN,
                   lg.NICHTS)
    sortiert = sorted(plan.ohne_gegenstueck,
                      key=lambda qi: reihenfolge.index(gruende[qi][0]))
    _abschnitt(zeilen, t(u"Ohne Gegenstück auf den gewählten Geschossen",
                         u"No match on the chosen levels",
                         u"Sin coincidencia en los niveles elegidos"),
               [text(qi) for qi in sortiert])
    anzahl = dict((g, sum(1 for qi in sortiert if gruende[qi][0] == g))
                  for g in reihenfolge)
    tipps = []
    if anzahl[lg.ANDERE_LAENGE] and not einst.achse:
        tipps.append(t(u"Tipp: %d ohne Gegenstück liegen auf derselben Achse, nur "
                        u"anders lang - Option 'gleiche Achse genügt' "
                        u"anhaken.",
                        u"Tip: %d unmatched ones lie on the same axis, just with a "
                        u"different length - check 'same axis is enough'.",
                        u"Consejo: %d sin coincidencia están en el mismo eje, solo "
                        u"con otra longitud; marque 'basta el mismo eje'.")
                      % anzahl[lg.ANDERE_LAENGE])
    if anzahl[lg.NICHT_GEWAEHLT]:
        tipps.append(t(u"Tipp: %d haben ein Gegenstück auf einem "
                        u"abgewählten Geschoss (Basisebene des Elements).",
                        u"Tip: %d have a match on an unchecked level (the "
                        u"element's base level).",
                        u"Consejo: %d tienen coincidencia en un nivel no "
                        u"marcado (nivel base del elemento).")
                      % anzahl[lg.NICHT_GEWAEHLT])
    if anzahl[lg.DANEBEN]:
        tipps.append(t(u"Tipp: %d liegen knapp daneben - Lagetoleranz "
                        u"erhöhen?",
                        u"Tip: %d are slightly off - increase the position "
                        u"tolerance?",
                        u"Consejo: %d están ligeramente desplazados; "
                        u"¿aumentar la tolerancia?") % anzahl[lg.DANEBEN])
    return tipps


def _bericht(bestand, plan, fehler, namen, ebenen_namen, einst):
    """Zeilen der Zusammenfassung."""
    from werte_stapel import logik as lg
    from werte_stapel import revit as rv

    def ziel(zi):
        return rv.beschreibe(bestand.ziel_revit[zi])

    def wert(s, w):
        return u"„%s“" % bestand.anzeige(s, w)

    gescheitert = set((zi, s) for zi, s, _ in fehler)
    geschrieben = [(zi, s) for zi, s, _ in plan.auftraege
                   if (zi, s) not in gescheitert]
    elemente = sorted(set(zi for zi, _ in geschrieben))
    je_ebene = {}
    for zi in elemente:
        ebene = bestand.ziele[zi].ebene
        je_ebene[ebene] = je_ebene.get(ebene, 0) + 1

    zeilen = [t(u"%d Werte in %d Elementen geschrieben.",
                u"%d values written to %d elements.",
                u"%d valores escritos en %d elementos.")
              % (len(geschrieben), len(elemente))]
    if je_ebene:
        zeilen.append(u", ".join(
            u"%s: %d" % (ebenen_namen.get(e, u"?"), n)
            for e, n in sorted(je_ebene.items(),
                               key=lambda p: ebenen_namen.get(p[0], u""))))
    zeilen.append(t(u"Modus: vorhandene Werte überschreiben",
                    u"Mode: overwrite existing values",
                    u"Modo: sobrescribir valores existentes")
                  if einst.ueberschreiben else
                  t(u"Modus: nur leere Felder füllen",
                    u"Mode: only fill empty fields",
                    u"Modo: solo rellenar campos vacíos"))
    if plan.behalten and not einst.ueberschreiben:
        zeilen.append(t(u"Tipp: %d Werte wurden nicht überschrieben - dafür "
                        u"im Fenster 'Vorhandene Werte überschreiben' "
                        u"anhaken.",
                        u"Tip: %d values were not overwritten - check "
                        u"'Overwrite existing values' in the window.",
                        u"Consejo: %d valores no se sobrescribieron; marque "
                        u"'Sobrescribir valores existentes' en la ventana.")
                      % len(plan.behalten))
    if plan.gleich:
        zeilen.append(t(u"%d Werte waren schon gleich.",
                        u"%d values were already the same.",
                        u"%d valores ya eran iguales.") % plan.gleich)
    if plan.quelle_leer:
        zeilen.append(t(u"%d Werte übersprungen, weil die Quelle leer ist.",
                        u"%d values skipped because the source is empty.",
                        u"%d valores omitidos porque el origen está vacío.")
                      % plan.quelle_leer)

    kopf = len(zeilen)
    _abschnitt(zeilen, t(u"Nicht überschrieben, anderer Wert vorhanden",
                         u"Not overwritten, other value present",
                         u"No sobrescritos, ya hay otro valor"),
               [u"%s – %s: %s, %s %s" % (
                   ziel(zi), namen.get(s, s), wert(s, alt),
                   t(u"Quelle", u"source", u"origen"), wert(s, neu))
                for zi, s, alt, neu in plan.behalten])
    _abschnitt(zeilen, t(u"Mehrdeutig, Quellen mit verschiedenen Werten",
                         u"Ambiguous, sources with different values",
                         u"Ambiguo, orígenes con valores distintos"),
               [u"%s – %s: %s" % (ziel(zi), namen.get(s, s),
                                  u" / ".join(wert(s, w) for w in werte))
                for zi, s, werte in plan.mehrdeutig])
    _abschnitt(zeilen, t(u"Parameter fehlt oder ist schreibgeschützt",
                         u"Parameter missing or read-only",
                         u"Parámetro inexistente o de solo lectura"),
               [u"%s – %s (%s)" % (
                   ziel(zi), namen.get(s, s),
                   t(u"schreibgeschützt", u"read-only", u"solo lectura")
                   if grund is lg.SCHREIBGESCHUETZT
                   else t(u"fehlt", u"missing", u"no existe"))
                for zi, s, grund in plan.fehlt])
    _abschnitt(zeilen, t(u"Fehler beim Schreiben", u"Errors while writing",
                         u"Errores al escribir"),
               [u"%s – %s: %s" % (ziel(zi), namen.get(s, s), grund)
                for zi, s, grund in fehler])
    tipps = _ohne_gegenstueck(zeilen, bestand, plan, einst, ebenen_namen)
    if tipps:
        zeilen[kopf:kopf] = [u""] + tipps
    _abschnitt(zeilen, t(u"Gewählt, aber ohne Lage im Grundriss",
                         u"Selected, but without a plan position",
                         u"Seleccionados, pero sin posición en planta"),
               [rv.beschreibe(e) for e in bestand.ohne_lage])
    return zeilen


def _quellen(uidoc):
    """Gewählte Stützen/Wände/Unterzüge - ohne passende Auswahl wählen
    lassen."""
    from werte_stapel import revit as rv

    elemente, _ = rv.auswahl(uidoc)
    if elemente:
        return elemente
    return rv.waehle_quellen(uidoc)


def main():
    from schedule_sync import ui

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument:
        ui.meldung(t(u"Das Werkzeug geht nur in Projektdateien.",
                     u"The tool only works in project files.",
                     u"La herramienta solo funciona en archivos de "
                     u"proyecto."),
                   titel=TITEL, hauptzeile=t(u"Kein Projekt geöffnet",
                                             u"No project open",
                                             u"No hay ningún proyecto "
                                             u"abierto"),
                   warnung=True)
        return
    if doc.IsReadOnly:
        ui.meldung(t(u"Das Dokument ist schreibgeschützt.",
                     u"The document is read-only.",
                     u"El documento es de solo lectura."), titel=TITEL,
                   hauptzeile=t(u"Werte können nicht übertragen werden",
                                u"Values cannot be copied",
                                u"No se pueden copiar los valores"),
                   warnung=True)
        return

    from Autodesk.Revit.DB import TransactionGroup
    from level_auto_set.revit import ebenen as alle_ebenen
    from werte_stapel import fenster
    from werte_stapel import logik as lg
    from werte_stapel import revit as rv

    quell_elemente = _quellen(uidoc)
    if not quell_elemente:
        return
    bestand = rv.Bestand(doc, quell_elemente)
    if not bestand.quellen:
        ui.meldung(t(u"Keines der gewählten Elemente hat eine Lage im "
                     u"Grundriss (Einfügepunkt bzw. Achse).",
                     u"None of the selected elements has a plan position "
                     u"(insertion point or axis).",
                     u"Ninguno de los elementos seleccionados tiene "
                     u"posición en planta (punto de inserción o eje)."),
                   titel=TITEL, hauptzeile=t(u"Nichts zu übertragen",
                                             u"Nothing to copy",
                                             u"Nada que copiar"),
                   warnung=True)
        return
    optionen = rv.parameter_optionen(bestand.quell_revit)
    if not optionen:
        ui.meldung(t(u"Die gewählten Elemente haben keine beschreibbaren "
                     u"Parameter.",
                     u"The selected elements have no writable parameters.",
                     u"Los elementos seleccionados no tienen parámetros "
                     u"editables."),
                   titel=TITEL, warnung=True)
        return

    einst = fenster.frage_einstellungen(__revit__, bestand, optionen)
    if einst is None:
        return

    bestand.lies_quellen(einst.schluessel)
    bestand.lies_ziele([zi for zi in einst.zuordnung
                        if bestand.ziele[zi].ebene in einst.ebenen],
                       einst.schluessel)
    plan = lg.plane(bestand.quellen, bestand.ziele, einst.zuordnung,
                    einst.schluessel, einst.ebenen, einst.ueberschreiben)
    namen = dict((o.schluessel, o.name) for o in optionen)
    ebenen_namen = dict((e.id, e.name) for e in alle_ebenen(doc))
    ebenen_namen[rv.OHNE_EBENE] = t(u"(ohne Ebene)", u"(no level)",
                                    u"(sin nivel)")

    if not plan.auftraege:
        zeilen = _bericht(bestand, plan, [], namen, ebenen_namen, einst)
        zeilen[0] = t(u"Es gibt nichts zu schreiben.",
                      u"There is nothing to write.",
                      u"No hay nada que escribir.")
        ui.meldung(u"\n".join(zeilen), titel=TITEL,
                   hauptzeile=t(u"Keine Werte übertragen",
                                u"No values copied",
                                u"No se copiaron valores"),
                   warnung=bool(plan.behalten or plan.mehrdeutig
                                or plan.fehlt))
        return

    gruppe = TransactionGroup(doc, TITEL)
    gruppe.Start()
    try:
        fehler = rv.schreibe(bestand, plan.auftraege)
        zeilen = _bericht(bestand, plan, fehler, namen, ebenen_namen,
                          einst)
        zeilen += [u"", t(u"Nein = alles zurücknehmen.",
                          u"No = undo everything.",
                          u"No = deshacer todo.")]
        behalten = ui.frage(u"\n".join(zeilen), titel=TITEL,
                            standard_ja=True,
                            warnung=bool(fehler or plan.behalten
                                         or plan.mehrdeutig),
                            hauptzeile=t(u"Werte übernehmen?",
                                         u"Keep the values?",
                                         u"¿Conservar los valores?"))
    except Exception:
        gruppe.RollBack()
        raise
    if behalten:
        gruppe.Assimilate()
    else:
        gruppe.RollBack()


def _zeige_fehler(spur):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    protokoll = os.path.join(basis, "pyMLG", "StackValues_Fehler.log")
    try:
        if not os.path.isdir(os.path.dirname(protokoll)):
            os.makedirs(os.path.dirname(protokoll))
        with io.open(protokoll, "a", encoding="utf-8") as datei:
            datei.write(spur + u"\n" + u"-" * 70 + u"\n")
    except Exception:
        protokoll = None
    try:
        from schedule_sync import ui
        letzte_zeile = (spur.strip().splitlines() or [u""])[-1]
        ui.meldung(letzte_zeile + (t(u"\n\nProtokoll: %s", u"\n\nLog: %s",
                                     u"\n\nRegistro: %s") % protokoll
                                   if protokoll else u""),
                   titel=TITEL,
                   hauptzeile=t(u"Werte übertragen ist fehlgeschlagen",
                                u"Copy Values failed",
                                u"Copiar valores ha fallado"),
                   warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
