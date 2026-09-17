# -*- coding: utf-8 -*-
"""Liest Schedules (Bauteillisten) aus dem aktiven Dokument aus.

Jede Bauteilliste wird doppelt ausgelesen, weil zwei Ziele sich widersprechen:

1. Datenblatt (lese_felder/lese_zeilen) - eine Zeile je Element.
   Die Zuordnung Zeile -> Element wird NICHT über ViewSchedule.GetTableData()
   hergestellt. Aus dieser Struktur lässt sich laut Revit-API keine
   zuverlässige Zeile-zu-Element-Zuordnung ableiten (Gruppenköpfe, Summen,
   zusammengefasste Zeilen verschieben die Indizes). Stattdessen liefert
   FilteredElementCollector(doc, schedule.Id) exakt die Elemente, die in der
   Bauteilliste inklusive ihrer Filter enthalten sind - unabhängig davon, ob
   die Liste aufgeschlüsselt, gruppiert oder zusammengefasst ist.

2. Ansicht (lese_ansicht) - die Tabelle Zelle für Zelle so, wie Revit sie
   anzeigt. Hier ist GetTableData() richtig, denn es wird nichts zugeordnet
   und nichts zurückgeschrieben.
"""

import re

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    SectionType,
    SharedParameterElement,
    ViewSchedule,
)

from schedule_sync.revit_helpers import (
    eid_wert,
    einheit_kuerzel,
    finde_parameter,
    ist_schreibgeschuetzt,
    kategorie_name,
    lese_wert,
    parameter_map,
    typ_von,
)
from mlg_sprache import t

# Spaltenaufbau der Excel-Blätter (1-basiert wie in openpyxl)
SPALTE_UNIQUEID = 1      # A - versteckt, stabile Kennung für den Rückimport
SPALTE_ELEMENTID = 2     # B - nur zur Orientierung, wird beim Import ignoriert
SPALTE_KATEGORIE = 3     # C - nur zur Orientierung
ERSTE_DATENSPALTE = 4    # D - ab hier ein Feld der Bauteilliste je Spalte
KOPFZEILE = 1
ERSTE_DATENZEILE = 2

# Feldtypen, hinter denen ein echter, potenziell beschreibbarer Parameter steht.
# Alles andere (Count, Formula, Percentage, CombinedParameter, ...) ist ein von
# Revit berechnetes Feld und kann nicht zurückgeschrieben werden.
PARAMETER_FELDTYPEN = ("Instance", "ElementType")

# Zellstatus, der die Formatierung in Excel und das Verhalten beim Import steuert
STATUS_OK = "ok"                 # beschreibbarer Instanzparameter
STATUS_TYP = "typ"               # beschreibbarer Typparameter (wirkt auf alle Instanzen)
STATUS_GESPERRT = "gesperrt"     # Parameter.IsReadOnly
STATUS_BERECHNET = "berechnet"   # von Revit berechnetes Feld ohne Parameter
STATUS_FEHLT = "fehlt"           # Parameter am Element nicht gefunden

SCHREIBBAR = (STATUS_OK, STATUS_TYP)

# Tabellenabschnitte, die für das Ansichtsblatt ausgelesen werden
ANSICHT_ABSCHNITTE = ("Header", "Body", "Summary", "Footer")
ABSCHNITT_KOPF = "Header"

HINWEIS_ANSICHT = t(u"Die Darstellung wie in Revit steht im Blatt 'Ansicht'.", u"The layout as in Revit is in the sheet 'View'.", u"La presentación como en Revit está en la hoja 'Vista'.")


def _hat(objekt, name, standard=None):
    """getattr mit Exception-Schutz - manche Revit-Properties werfen beim Zugriff."""
    try:
        wert = getattr(objekt, name, standard)
        return wert
    except Exception:
        return standard


# ---------------------------------------------------------------------------
# Schedules einsammeln und prüfen
# ---------------------------------------------------------------------------

def sammle_schedules(doc):
    """Alle auswählbaren ViewSchedule-Instanzen des Dokuments.

    Ausgeschlossen: Vorlagen (Schedule-Templates), interne Keynote-Listen und
    Revisionslisten von Plankopf-Familien - diese haben keine eigene,
    sinnvoll exportierbare View.
    """
    ergebnis = []
    for schedule in FilteredElementCollector(doc).OfClass(ViewSchedule):
        try:
            if schedule.IsTemplate:
                continue
            if _hat(schedule, "IsInternalKeynoteSchedule", False):
                continue
            if _hat(schedule, "IsTitleblockRevisionSchedule", False):
                continue
            ergebnis.append(schedule)
        except Exception:
            continue
    ergebnis.sort(key=lambda s: s.Name)
    return ergebnis


def pruefe_schedule(schedule):
    """Prüft, ob eine Bauteilliste exportierbar ist.

    Grundsätzlich ist jede Bauteilliste exportierbar. Abgelehnt wird nur, was
    gar keine lesbare Definition hat. Besonderheiten der Darstellung werden als
    Hinweis zurückgegeben, damit der Nutzer versteht, warum das Datenblatt
    anders aussieht als die Tabelle in Revit.

    Rückgabe: (ist_exportierbar, ablehnungsgrund, [hinweise])
    """
    hinweise = []
    try:
        definition = schedule.Definition
    except Exception as fehler:
        return False, t(u"Definition nicht lesbar (%s)", u"Definition not readable (%s)", u"Definición no legible (%s)") % fehler, hinweise

    if definition is None:
        return False, t(u"Bauteilliste hat keine Definition", u"Schedule has no definition", u"La tabla no tiene definición"), hinweise

    if not _hat(definition, "IsItemized", True):
        hinweise.append(
            t(u"Revit fasst in dieser Liste gleiche Elemente zu einer Zeile "
            u"zusammen. Das Datenblatt enthält trotzdem eine Zeile je Element "
            u"- nur so lässt sich jede Änderung eindeutig zurückschreiben. ", u"Revit groups identical elements into one row in this schedule. The data sheet still contains one row per element - only then can every change be written back unambiguously. ", u"Revit agrupa en esta tabla elementos iguales en una fila. La hoja de datos contiene igualmente una fila por elemento: solo así se puede reescribir cada cambio sin ambigüedad. ")
            + HINWEIS_ANSICHT)

    if _hat(definition, "IsMaterialTakeoff", False):
        hinweise.append(
            t(u"Materialauszug: Revit zeigt eine Zeile je Element und Material. "
            u"Im Datenblatt steht eine Zeile je Element, Materialspalten sind "
            u"dort grau. ", u"Material takeoff: Revit shows one row per element and material. The data sheet has one row per element; material columns are grey there. ", u"Cómputo de materiales: Revit muestra una fila por elemento y material. La hoja de datos tiene una fila por elemento; las columnas de material aparecen en gris. ") + HINWEIS_ANSICHT)

    if _hat(definition, "IncludeLinkedFiles", False):
        hinweise.append(
            t(u"Elemente aus verknüpften Modellen stehen nur im Blatt 'Ansicht' "
            u"- sie lassen sich nicht zurückschreiben.", u"Elements from linked models are only in the sheet 'View' - they cannot be written back.", u"Los elementos de modelos vinculados solo están en la hoja 'Vista': no se pueden reescribir."))

    if _hat(definition, "IsKeySchedule", False):
        hinweise.append(
            t(u"Schlüsselliste: Bearbeitet werden die Schlüssel selbst, nicht die "
            u"Bauteile, denen sie zugewiesen sind.", u"Key schedule: the keys themselves are edited, not the elements they are assigned to.", u"Tabla de claves: se editan las claves, no los elementos a los que están asignadas."))

    try:
        for index in range(definition.GetSortGroupFieldCount()):
            sortierfeld = definition.GetSortGroupField(index)
            if sortierfeld.ShowHeader or sortierfeld.ShowFooter:
                hinweise.append(
                    t(u"Gruppenköpfe und Zwischensummen stehen im Blatt "
                    u"'Ansicht'. Das Datenblatt ist genauso sortiert.", u"Group headers and subtotals are in the sheet 'View'. The data sheet is sorted the same way.", u"Los encabezados de grupo y subtotales están en la hoja 'Vista'. La hoja de datos está ordenada igual."))
                break
    except Exception:
        pass

    if _hat(definition, "ShowGrandTotal", False):
        hinweise.append(t(u"Die Gesamtsumme steht im Blatt 'Ansicht'.", u"The grand total is in the sheet 'View'.", u"El total general está en la hoja 'Vista'."))

    return True, None, hinweise


# ---------------------------------------------------------------------------
# Felder
# ---------------------------------------------------------------------------

def _parameter_kennung(doc, feld):
    """(param_id, guid) eines Schedule-Felds.

    param_id < 0  -> BuiltInParameter (Enumwert)
    param_id > 0  -> Projekt-/Shared-Parameter dieses Dokuments
    guid          -> nur bei Shared Parameters, dokumentübergreifend stabil
    """
    try:
        parameter_id = feld.ParameterId
    except Exception:
        return None, None

    wert = eid_wert(parameter_id)
    if wert is None or wert == -1:
        return None, None

    guid = None
    if wert > 0:
        try:
            parameter_element = doc.GetElement(parameter_id)
            if isinstance(parameter_element, SharedParameterElement):
                guid = str(parameter_element.GuidValue)
        except Exception:
            guid = None
    return wert, guid


def _feld_schluessel(feld_id):
    """Vergleichbarer Schlüssel einer ScheduleFieldId (für die Sortierung)."""
    wert = _hat(feld_id, "IntegerValue", None)
    return int(wert) if wert is not None else str(feld_id)


def _spaltentitel(feld):
    """Überschrift wie in der Revit-Tabelle, ersatzweise der Feldname.

    Revit zeigt ColumnHeading an - der Nutzer kann es frei umbenennen
    ('Name' -> 'Raumbezeichnung'). Genau dieser Text soll in Excel stehen.
    """
    for kandidat in (_hat(feld, "ColumnHeading", None), _feldname(feld)):
        if kandidat and kandidat.strip():
            return u" ".join(kandidat.split())
    return u""


def _feldname(feld):
    try:
        return feld.GetName() or u""
    except Exception:
        return u""


def _feld_info(doc, feld, feld_id, name):
    """Beschreibung eines Schedule-Felds als Dict."""
    feldtyp = str(_hat(feld, "FieldType", u""))
    param_id, guid = _parameter_kennung(doc, feld)
    kombiniert = bool(_hat(feld, "IsCombinedParameterField", False))
    return {
        "spalte": None,
        "name": name,
        # Echter Parametername - Rückfallebene, wenn die Überschrift umbenannt ist
        "feldname": _feldname(feld) or name,
        "param_id": param_id,
        "guid": guid,
        "feldtyp": feldtyp,
        "ist_typfeld": feldtyp == "ElementType",
        "berechnet": (feldtyp not in PARAMETER_FELDTYPEN
                      or kombiniert
                      or param_id is None),
        "einheit": u"",
        "feld_schluessel": _feld_schluessel(feld_id),
    }


def lese_felder(doc, schedule):
    """Sichtbare Felder der Bauteilliste in Anzeigereihenfolge.

    Rückgabe: Liste von Dicts mit Spaltennummer, Überschrift, Parameterkennung
    und der Information, ob es sich um ein berechnetes bzw. Typfeld handelt.
    """
    definition = schedule.Definition
    felder = []
    spalte = ERSTE_DATENSPALTE
    vergebene_namen = {}

    for feld_id in definition.GetFieldOrder():
        try:
            feld = definition.GetField(feld_id)
        except Exception:
            continue
        if feld is None:
            continue
        # Ausgeblendete Spalten sind auch in Revit nicht sichtbar
        if _hat(feld, "IsHidden", False):
            continue

        name = _spaltentitel(feld) or t(u"Feld %d", u"Field %d", u"Campo %d") % spalte

        # Excel-Kopfzeilen müssen eindeutig sein, Schedule-Felder sind es nicht zwingend
        if name in vergebene_namen:
            vergebene_namen[name] += 1
            name = u"%s (%d)" % (name, vergebene_namen[name])
        else:
            vergebene_namen[name] = 1

        info = _feld_info(doc, feld, feld_id, name)
        info["spalte"] = spalte
        felder.append(info)
        spalte += 1

    return felder


def lese_sortierung(doc, schedule, felder):
    """Sortier-/Gruppierfelder der Bauteilliste.

    Enthält auch ausgeblendete Felder - Revit darf nach Spalten sortieren, die
    gar nicht angezeigt werden.
    Rückgabe: Liste von (feld, absteigend)
    """
    definition = schedule.Definition
    sichtbare = dict((f["feld_schluessel"], f) for f in felder)
    ergebnis = []

    try:
        anzahl = definition.GetSortGroupFieldCount()
    except Exception:
        return ergebnis

    for index in range(anzahl):
        try:
            sortierfeld = definition.GetSortGroupField(index)
            feld_id = sortierfeld.FieldId
            absteigend = "descending" in str(sortierfeld.SortOrder).lower()
        except Exception:
            continue

        feld = sichtbare.get(_feld_schluessel(feld_id))
        if feld is None:
            try:
                revit_feld = definition.GetField(feld_id)
                feld = _feld_info(doc, revit_feld, feld_id,
                                  _spaltentitel(revit_feld))
            except Exception:
                continue
        ergebnis.append((feld, absteigend))

    return ergebnis


# ---------------------------------------------------------------------------
# Zeilen
# ---------------------------------------------------------------------------

def sortierschluessel(wert):
    """Sortierschlüssel, der Revit nahekommt: leer zuerst, Zahlen numerisch,
    Texte 'natürlich' (EG 2 vor EG 10) und ohne Gross-/Kleinschreibung."""
    if wert is None or (isinstance(wert, str) and not wert.strip()):
        return (0,)
    if isinstance(wert, (int, float)) and not isinstance(wert, bool):
        return (1, float(wert))
    teile = re.split(r"(\d+)", str(wert).lower())
    return (2, tuple((0, int(teil), u"") if teil.isdigit() else (1, 0, teil)
                     for teil in teile if teil))


def _sortierwert(doc, element, feld, zeile, instanz_map, typ_element, typ_map):
    """Wert eines Sortierfelds - aus der Zeile oder, falls ausgeblendet, direkt."""
    if feld.get("spalte") is not None:
        return zeile["werte"].get(feld["spalte"])
    if feld.get("berechnet"):
        return None
    treffer = finde_parameter(doc, element, feld, instanz_map=instanz_map,
                              typ_element=typ_element, typ_map=typ_map)
    if treffer is None:
        return None
    wert, _anzeige = lese_wert(doc, treffer.parameter)
    return wert


def lese_zeilen(doc, schedule, felder, sortierung=None):
    """Liest alle Elemente der Bauteilliste mit ihren Feldwerten.

    Eine Zeile je Element, sortiert wie in Revit (soweit die Sortierfelder
    Werte liefern). Rückgabe: (zeilen, anzahl_uebersprungen)
    Jede Zeile ist ein Dict mit uid/element_id/kategorie sowie je Spalte einem
    Wert und einem Status (siehe STATUS_*).
    """
    sortierung = sortierung or []
    try:
        elemente = list(FilteredElementCollector(doc, schedule.Id))
    except Exception as fehler:
        raise RuntimeError(
            t(u"Elemente der Bauteilliste '%s' nicht lesbar: %s", u"Elements of schedule '%s' not readable: %s", u"Elementos de la tabla '%s' no legibles: %s") % (schedule.Name, fehler))

    zeilen = []
    uebersprungen = 0

    for element in elemente:
        if element is None:
            uebersprungen += 1
            continue
        try:
            unique_id = element.UniqueId
        except Exception:
            uebersprungen += 1
            continue

        instanz_map = parameter_map(element)
        typ_element = typ_von(doc, element)
        typ_map = parameter_map(typ_element) if typ_element is not None else {}

        zeile = {
            "uid": unique_id,
            "element_id": eid_wert(element.Id),
            "kategorie": kategorie_name(element),
            "werte": {},
            "status": {},
        }

        for feld in felder:
            spalte = feld["spalte"]

            if feld["berechnet"]:
                # Eine Zeile je Element: die Anzahl ist hier immer 1
                zeile["werte"][spalte] = 1 if feld["feldtyp"] == "Count" else None
                zeile["status"][spalte] = STATUS_BERECHNET
                continue

            treffer = finde_parameter(doc, element, feld,
                                      instanz_map=instanz_map,
                                      typ_element=typ_element,
                                      typ_map=typ_map)
            if treffer is None:
                zeile["werte"][spalte] = None
                zeile["status"][spalte] = STATUS_FEHLT
                continue

            wert, _anzeige = lese_wert(doc, treffer.parameter)
            zeile["werte"][spalte] = wert

            if ist_schreibgeschuetzt(treffer.parameter):
                zeile["status"][spalte] = STATUS_GESPERRT
            elif treffer.ist_typparameter:
                zeile["status"][spalte] = STATUS_TYP
            else:
                zeile["status"][spalte] = STATUS_OK

            # Einheitenkürzel einmalig für die Kopfzeilen-Notiz merken
            if not feld["einheit"]:
                feld["einheit"] = einheit_kuerzel(treffer.parameter)

        zeile["_sortierung"] = [
            _sortierwert(doc, element, feld, zeile, instanz_map, typ_element,
                         typ_map)
            for feld, _absteigend in sortierung]
        zeilen.append(zeile)

    sortiere_zeilen(zeilen, sortierung)
    for zeile in zeilen:
        zeile.pop("_sortierung", None)

    return zeilen, uebersprungen


def sortiere_zeilen(zeilen, sortierung):
    """Sortiert nach allen Sortierfeldern, das erste Feld hat Vorrang.

    Python sortiert stabil (auch mit reverse=True). Deshalb wird vom letzten
    zum ersten Feld sortiert - so entsteht die mehrstufige Reihenfolge.
    """
    for index in reversed(range(len(sortierung))):
        _feld, absteigend = sortierung[index]
        zeilen.sort(key=lambda z, i=index: sortierschluessel(z["_sortierung"][i]),
                    reverse=absteigend)


def spaltenstatus(felder, zeilen):
    """Ermittelt je Spalte den 'schwächsten' Status über alle Zeilen.

    Eine Spalte gilt nur dann als beschreibbar, wenn mindestens eine Zeile
    beschreibbar ist. So bleibt die Excel-Formatierung ehrlich: alles Graue
    wird beim Import garantiert ignoriert.
    """
    ergebnis = {}
    for feld in felder:
        spalte = feld["spalte"]
        if feld["berechnet"]:
            ergebnis[spalte] = STATUS_BERECHNET
            continue
        stati = set(zeile["status"].get(spalte) for zeile in zeilen)
        if STATUS_OK in stati:
            ergebnis[spalte] = STATUS_OK
        elif STATUS_TYP in stati:
            ergebnis[spalte] = STATUS_TYP
        elif STATUS_GESPERRT in stati:
            ergebnis[spalte] = STATUS_GESPERRT
        elif stati == {STATUS_FEHLT} or not stati:
            ergebnis[spalte] = STATUS_FEHLT
        else:
            ergebnis[spalte] = STATUS_GESPERRT
    return ergebnis


# ---------------------------------------------------------------------------
# Ansicht: die Tabelle wie in Revit
# ---------------------------------------------------------------------------

def lese_ansicht(schedule):
    """Die Tabelle Zelle für Zelle so, wie Revit sie anzeigt.

    Enthält Titel, Spaltenköpfe, Gruppenköpfe, zusammengefasste Zeilen und
    Summen als formatierten Text. Nur zum Lesen - hier ist GetTableData
    richtig, weil keine Zeile einem Element zugeordnet werden muss.

    Rückgabe: Liste von (abschnitt, [zelltexte])
    """
    try:
        tabelle = schedule.GetTableData()
    except Exception:
        return []

    ergebnis = []
    for abschnitt_name in ANSICHT_ABSCHNITTE:
        abschnitt = getattr(SectionType, abschnitt_name, None)
        if abschnitt is None:
            continue
        try:
            daten = tabelle.GetSectionData(abschnitt)
        except Exception:
            continue
        if daten is None or _hat(daten, "HideSection", False):
            continue
        try:
            erste_zeile, letzte_zeile = daten.FirstRowNumber, daten.LastRowNumber
            erste_spalte, letzte_spalte = (daten.FirstColumnNumber,
                                           daten.LastColumnNumber)
        except Exception:
            continue

        for zeile in range(erste_zeile, letzte_zeile + 1):
            texte = []
            for spalte in range(erste_spalte, letzte_spalte + 1):
                try:
                    texte.append(schedule.GetCellText(abschnitt, zeile, spalte) or u"")
                except Exception:
                    texte.append(u"")
            ergebnis.append((abschnitt_name, texte))

    return ergebnis
