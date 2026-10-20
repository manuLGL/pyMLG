#! python3
# -*- coding: utf-8 -*-
"""AreaTable: Flächentabelle nach WoFlV auf Wohnungspläne zeichnen -
Wohnfläche, Balkon-/Terrassenfläche mit Anrechnung, Zimmerzahl.

Die Logik liegt in lib/wohnflaeche (logik.py ohne Revit, revit.py).

CPython-Besonderheiten dieses Repos (siehe ScheduleSync.panel/README.md):
pyrevit.forms ist nicht nutzbar - Dialoge aus schedule_sync.ui.
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

TITEL = t(u"Flächentabelle WoFlV", u"WoFlV Area Table",
          u"Tabla de superficies WoFlV")

MAX_ZEILEN = 12


def _protokoll_pfad(name):
    basis = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP", ".")
    ordner = os.path.join(basis, "pyMLG")
    if not os.path.isdir(ordner):
        os.makedirs(ordner)
    return os.path.join(ordner, name)


def _plaene(doc, uidoc):
    from mlg_plaene import revit as plan_rv
    from schedule_sync import ui

    gewaehlt = plan_rv.gewaehlte_plaene(doc, uidoc)
    if gewaehlt:
        return gewaehlt, True
    wahl = ui.waehle_option(
        [(t(u"Alle Pläne mit Räumen im Grundriss",
            u"All sheets with rooms in a floor plan",
            u"Todos los planos con habitaciones en planta"),
          t(u"Jeder Plan bekommt die Tabelle seiner Wohnung.",
            u"Each sheet gets the table of its apartment.",
            u"Cada plano recibe la tabla de su vivienda."))],
        titel=TITEL,
        hauptzeile=t(u"Kein Plan gewählt", u"No sheet selected",
                     u"Ningún plano seleccionado"),
        text=t(u"Tipp: Pläne im Projektbrowser markieren oder einen Plan "
               u"öffnen, dann gilt das nur für diese.",
               u"Tip: select sheets in the project browser or open a sheet "
               u"to process only those.",
               u"Consejo: seleccione planos en el navegador o abra un plano "
               u"para procesar solo esos."))
    if wahl is None:
        return [], False
    return plan_rv.plaene(doc), False


def _raum_zeile(raum):
    from wohnflaeche import logik as lg
    faktor, abzug = raum.faktor()
    zusatz = u""
    if raum.faktor_vorgabe is not None and raum.art == lg.AUSSEN:
        zusatz = t(u" (Parameter %s ignoriert)", u" (parameter %s ignored)",
                   u" (parámetro %s ignorado)") % lg.zahl(raum.faktor_vorgabe)
    elif raum.faktor_vorgabe is not None and raum.art != lg.ALLGEMEIN:
        zusatz = t(u" (Parameter)", u" (parameter)", u" (parámetro)")
    elif abzug > 0.0:
        zusatz = t(u" (Dachschräge)", u" (sloped ceiling)",
                   u" (techo inclinado)")
    hinweis = raum.parameter_hinweis or t(
        u"kein Parameter Faktor Anrechenbarkeit",
        u"no parameter Faktor Anrechenbarkeit",
        u"sin parámetro Faktor Anrechenbarkeit")
    return u"    %-32s %10s  × %-6s = %10s%s  [%s]" % (
        raum.text()[:32], lg.m2(raum.flaeche), lg.zahl(faktor),
        lg.m2(lg.angerechnete_flaeche(raum)), zusatz, hinweis)


def main():
    from schedule_sync import ui

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None or doc.IsFamilyDocument or doc.IsReadOnly:
        ui.meldung(t(u"Das Werkzeug geht nur in beschreibbaren Projektdateien.",
                     u"The tool only works in writable project files.",
                     u"La herramienta solo funciona en proyectos editables."),
                   titel=TITEL, warnung=True)
        return

    from Autodesk.Revit.DB import Transaction
    from wohnflaeche import logik as lg
    from wohnflaeche import revit as rv

    plaene, gewaehlt = _plaene(doc, uidoc)
    if not plaene:
        return

    fertig, ohne_raeume, mehrere, fehler = [], [], [], []
    ohne_sichtbare, ohne_belegung = [], []
    mit_parameter, parameter_namen = [False], []
    bericht = [TITEL, u"=" * len(TITEL), u""]
    transaktion = Transaction(doc, TITEL)
    transaktion.Start()
    try:
        for plan in plaene:
            revit_raeume, ersatz = rv.raeume_auf_plan(doc, plan)
            if revit_raeume:
                ohne_sichtbare += [u"%s: %s" % (plan.SheetNumber, name)
                                   for name in ersatz]
            if not revit_raeume:
                if gewaehlt:
                    ohne_raeume.append(u"%s - %s" % (plan.SheetNumber, plan.Name))
                continue
            einheit = lg.eine_wohnung([rv.raum_daten(r) for r in revit_raeume])
            raeume = einheit.raeume
            if not raeume:
                continue
            if einheit.andere and not gewaehlt:
                # Übersichtsplan mit mehreren Wohnungen - nur auf ausdrücklich
                # gewählten Plänen zählt dann die grösste
                mehrere.append(u"%s - %s" % (plan.SheetNumber, plan.Name))
                continue
            raeume.sort(key=lambda r: (r.nummer, r.name))
            wohnung = lg.wohnungs_name(einheit.name, plan.Name)
            ergebnis = lg.berechne(raeume, wohnung)
            if not einheit.nach_belegung:
                ohne_belegung.append(u"%s - %s" % (plan.SheetNumber, plan.Name))
            try:
                alte = rv.vorhandene_tabellen(doc, plan)
                # von Hand geänderter Titel und Miete bleiben
                titel, miete, lose = rv.alte_texte(doc, plan, alte)
                zeilen = lg.tabelle(ergebnis, miete)
                rv.entferne(doc, alte)
                if lose:
                    rv.loesche(doc, lose)
                rv.zeichne(doc, plan, zeilen)
                rv.zeichne_titel(doc, plan, titel or wohnung)
            except Exception as ausnahme:
                fehler.append(u"%s: %s" % (plan.SheetNumber, ausnahme))
                continue
            fertig.append(u"%s – %s: %s, %s" % (
                plan.SheetNumber, titel or wohnung, lg.m2(ergebnis.wohnflaeche),
                t(u"Gewerbe", u"commercial", u"comercial") if ergebnis.gewerbe
                else u"%d %s" % (ergebnis.zimmer,
                                 t(u"Zimmer", u"rooms", u"habitaciones"))))
            bericht.append(u"%s - %s  |  %s" % (plan.SheetNumber, plan.Name,
                                                lg.titel(wohnung, ergebnis.zimmer)))
            bericht += [_raum_zeile(r) for r in raeume]
            for raum in raeume:
                if raum.parameter_hinweis is not None:
                    mit_parameter[0] = True
                elif not parameter_namen and raum.ref is not None:
                    parameter_namen.extend(sorted(
                        p.Definition.Name for p in raum.ref.Parameters))
            bericht.append(u"    %s: %s" % (t(u"Wohnfläche", u"Living area",
                                              u"Superficie útil"),
                                            lg.m2(ergebnis.wohnflaeche)))
            if ergebnis.allgemein:
                bericht.append(u"    %s: %s" % (
                    t(u"Gemeinschaftsfläche, nicht in der Tabelle",
                      u"common area, not in the table",
                      u"zona común, no figura en la tabla"),
                    u", ".join(u"%s (%s)" % (r.text(), lg.m2(r.flaeche))
                               for r in ergebnis.allgemein)))
            if einheit.andere:
                bericht.append(u"    %s: %s" % (
                    t(u"andere Wohnung, nicht gezählt",
                      u"other apartment, not counted",
                      u"otra vivienda, no contada"),
                    u", ".join(u"%s (%s)" % (
                        r.text(), r.belegung if einheit.nach_belegung
                        else r.abteilung) for r in einheit.andere)))
            if einheit.ohne:
                bericht.append(u"    %s: %s" % (
                    t(u"ohne Belegung, nicht gezählt",
                      u"no occupancy, not counted",
                      u"sin ocupación, no contada"),
                    u", ".join(r.text() for r in einheit.ohne)))
            bericht.append(u"")
        transaktion.Commit()
    except Exception:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        raise

    if parameter_namen:
        bericht += [t(u"Parameter eines Raums ohne Faktor Anrechenbarkeit:",
                      u"Parameters of a room without Faktor Anrechenbarkeit:",
                      u"Parámetros de una habitación sin "
                      u"Faktor Anrechenbarkeit:"),
                    u"    " + u", ".join(parameter_namen)]
    protokoll = None
    try:
        protokoll = _protokoll_pfad(u"Wohnflaeche_Bericht.txt")
        with io.open(protokoll, "w", encoding="utf-8") as datei:
            datei.write(u"\n".join(bericht))
    except Exception:
        protokoll = None

    text = [t(u"%d Tabellen gezeichnet.", u"%d tables drawn.",
              u"%d tablas dibujadas.") % len(fertig)]
    text += [u"• " + z for z in fertig[:MAX_ZEILEN]]
    if len(fertig) > MAX_ZEILEN:
        text.append(t(u"… und %d weitere", u"… and %d more", u"… y %d más")
                    % (len(fertig) - MAX_ZEILEN))
    if ohne_raeume:
        text += [u"", t(u"Ohne Räume im Grundriss:", u"No rooms in a floor plan:",
                        u"Sin habitaciones en planta:")]
        text += [u"• " + z for z in ohne_raeume[:MAX_ZEILEN]]
    if ohne_sichtbare:
        text += [u"", t(u"Kategorie Räume ist ausgeschaltet - dort zählen alle "
                        u"Räume der Ebene im Zuschnitt, auch ausgeblendete:",
                        u"Rooms category is turned off - all rooms of the "
                        u"level inside the crop count there, hidden ones too:",
                        u"La categoría Habitaciones está desactivada; allí "
                        u"cuentan todas las del nivel dentro del recorte, "
                        u"también las ocultas:")]
        text += [u"• " + z for z in ohne_sichtbare[:MAX_ZEILEN]]
    if ohne_belegung:
        text += [u"", t(u"Kein Raum mit Belegung - ersatzweise nach Abteilung "
                        u"bzw. allen sichtbaren Räumen gezählt:",
                        u"No room with occupancy - counted by department or "
                        u"all visible rooms instead:",
                        u"Ninguna habitación con ocupación; contado por "
                        u"departamento o todas las visibles:")]
        text += [u"• " + z for z in ohne_belegung[:MAX_ZEILEN]]
    if mehrere:
        text += [u"", t(u"Übersprungen, mehrere Wohnungen (Belegung) auf dem "
                        u"Plan - Plan markieren, dann zählt die grösste:",
                        u"Skipped, several apartments (occupancy) on the "
                        u"sheet - select the sheet to use the largest:",
                        u"Omitidos, varias viviendas (ocupación) en el "
                        u"plano; selecciónelo para usar la mayor:")]
        text += [u"• " + z for z in mehrere[:MAX_ZEILEN]]
    if fehler:
        text += [u"", t(u"Fehler:", u"Errors:", u"Errores:")]
        text += [u"• " + z for z in fehler[:MAX_ZEILEN]]
    if fertig and not mit_parameter[0]:
        text += [u"", t(u"Achtung: Kein Raum hat einen Parameter "
                        u"\"Faktor Anrechenbarkeit\" - alle Innenräume "
                        u"zählen voll. Die vorhandenen Parameternamen stehen "
                        u"im Bericht.",
                        u"Warning: no room has a parameter "
                        u"\"Faktor Anrechenbarkeit\" - all indoor rooms count "
                        u"fully. The existing parameter names are in the "
                        u"report.",
                        u"Atención: ninguna habitación tiene el parámetro "
                        u"\"Faktor Anrechenbarkeit\"; todos los espacios "
                        u"interiores cuentan al 100 %. Los nombres de "
                        u"parámetros están en el informe.")]
    if fertig and not rv.volumen_an(doc):
        text += [u"", t(u"Hinweis: Volumenberechnung ist aus - Dachschrägen "
                        u"werden nicht abgezogen (Flächen- und "
                        u"Volumenberechnungen > Volumen berechnen).",
                        u"Note: volume computation is off - sloped ceilings "
                        u"are not deducted (Area and Volume Computations > "
                        u"Compute volumes).",
                        u"Nota: el cálculo de volúmenes está desactivado; no "
                        u"se descuentan los techos inclinados (Cálculos de "
                        u"área y volumen > Calcular volúmenes).")]
    if protokoll:
        text += [u"", t(u"Raumliste mit Faktoren: %s", u"Room list with "
                        u"factors: %s", u"Lista de habitaciones con factores: "
                        u"%s") % protokoll]
    ui.meldung(u"\n".join(text), titel=TITEL,
               hauptzeile=t(u"Flächentabellen aktualisiert",
                            u"Area tables updated",
                            u"Tablas de superficies actualizadas")
               if fertig else t(u"Keine Tabelle gezeichnet",
                                u"No table drawn", u"Ninguna tabla dibujada"),
               warnung=bool(fehler or not fertig))


def _zeige_fehler(spur):
    try:
        protokoll = _protokoll_pfad("AreaTable_Fehler.log")
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
                   hauptzeile=t(u"Flächentabelle ist fehlgeschlagen",
                                u"Area table failed",
                                u"La tabla de superficies ha fallado"),
                   warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc())
