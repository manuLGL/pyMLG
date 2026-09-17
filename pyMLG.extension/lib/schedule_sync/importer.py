# -*- coding: utf-8 -*-
"""Rückimport der Excel-Werte in das Revit-Modell.

Der Import läuft zweistufig:
    1. analysiere()  - liest nur, ermittelt alle tatsächlichen Änderungen,
                       sammelt übersprungene Zeilen und Fehler. Keine Transaktion.
    2. wende_an()    - schreibt die ermittelten Änderungen. Muss innerhalb einer
                       offenen Transaktion aufgerufen werden.

Durch die Trennung kann dem Nutzer vor dem Schreiben angezeigt werden, was
passieren wird, und es landen nur echte Abweichungen in der Transaktion.
"""

from collections import OrderedDict

from schedule_sync.revit_helpers import (
    Konvertierungsfehler,
    eid_wert,
    element_name,
    finde_parameter,
    fremder_besitzer,
    ist_schreibgeschuetzt,
    kategorie_name,
    lese_wert,
    parameter_map,
    typ_von,
    unterscheidet_sich,
    zielwert,
)
from schedule_sync.schedule_model import SCHREIBBAR
from mlg_sprache import t


class Aenderung(object):
    """Eine geplante Parameteränderung an genau einem Element."""

    def __init__(self, blatt, excel_zeile, uid, element, feldname, parameter,
                 besitzer, ist_typparameter, neuer_wert, alt_anzeige):
        self.blatt = blatt
        self.excel_zeile = excel_zeile
        self.uid = uid
        self.element = element
        self.element_id = eid_wert(element.Id)
        self.feldname = feldname
        self.parameter = parameter
        self.besitzer = besitzer
        self.ist_typparameter = ist_typparameter
        self.neuer_wert = neuer_wert
        self.alt_anzeige = alt_anzeige
        self.angewendet = False


class Bericht(object):
    """Sammelstelle für Ergebnisse und Meldungen eines Importlaufs."""

    def __init__(self):
        self.aenderungen = []
        self.uebersprungen = []      # Zeilen, die bewusst nicht angefasst wurden
        self.fehler = []             # Zellen, die nicht verarbeitet werden konnten
        self.unveraendert = 0        # Werte, die bereits übereinstimmten
        self.ignoriert_gesperrt = 0  # schreibgeschützte/berechnete Spalten
        self.zeilen_gesamt = 0
        # Diagnose je Spalte: zeigt, was mit jeder Excel-Spalte passiert ist.
        # Unverzichtbar, um "es passiert nichts" von "es gibt nichts zu tun"
        # unterscheiden zu können.
        self.spalten = OrderedDict()

    def notiere_uebersprungen(self, blatt, zeile, uid, grund):
        self.uebersprungen.append({
            "blatt": blatt, "zeile": zeile, "uid": uid, "grund": grund})

    def notiere_fehler(self, blatt, zeile, element_id, bezeichnung, feld, meldung):
        self.fehler.append({
            "blatt": blatt, "zeile": zeile, "element_id": element_id,
            "element": bezeichnung, "feld": feld, "meldung": meldung})

    @property
    def betroffene_elemente(self):
        return len(set(a.uid for a in self.aenderungen if a.angewendet))

    @property
    def angewendete_aenderungen(self):
        return len([a for a in self.aenderungen if a.angewendet])


def _bezeichnung(element):
    """Kurzbeschreibung eines Elements für Meldungen."""
    name = element_name(element)
    kategorie = kategorie_name(element)
    if name and kategorie:
        return u"%s: %s" % (kategorie, name)
    return name or kategorie or t(u"Element", u"Element", u"Elemento")


def _ueberspringbar(feld):
    """True, wenn die Spalte laut Exportmetadaten gar nicht beschreibbar ist."""
    if feld.get("berechnet"):
        return True
    status = feld.get("status")
    # Ohne Metadaten (status None) wird die Spalte versucht; der Schreibschutz
    # wird dann zur Laufzeit am Parameter selbst geprüft.
    return status is not None and status not in SCHREIBBAR


def _spalte(bericht, blattname, feld):
    """Diagnose-Eintrag einer Excel-Spalte (wird bei Bedarf angelegt)."""
    schluessel = (blattname, feld["spalte"])
    eintrag = bericht.spalten.get(schluessel)
    if eintrag is None:
        eintrag = {
            "blatt": blattname,
            "spalte": feld["spalte"],
            "name": feld["name"],
            "quelle": t(u"Metadaten", u"metadata", u"metadatos") if feld.get("aus_meta") else t(u"nur Spaltenname", u"column name only", u"solo nombre de columna"),
            "status": feld.get("status") or u"unbekannt",
            "param_id": feld.get("param_id"),
            "verglichen": 0,
            "gleich": 0,
            "geplant": 0,
            "gesperrt": 0,
            "uebersprungen": 0,
            "fehler": 0,
        }
        bericht.spalten[schluessel] = eintrag
    return eintrag


def analysiere(doc, blatt, bericht, gesehen=None):
    """Ermittelt alle Änderungen eines Excel-Blatts, ohne etwas zu schreiben.

    gesehen: Dict zur Erkennung doppelter Ziele (z. B. derselbe Typparameter
    in mehreren Zeilen). Wird über mehrere Blätter hinweg weitergereicht.
    """
    if gesehen is None:
        gesehen = {}

    blattname = blatt["blattname"]
    felder = [f for f in blatt["felder"] if not _ueberspringbar(f)]
    stumme_felder = [f for f in blatt["felder"] if _ueberspringbar(f)]

    # Jede Spalte taucht in der Diagnose auf - auch die, die nie angefasst wird
    for feld in blatt["felder"]:
        _spalte(bericht, blattname, feld)

    for zeile in blatt["zeilen"]:
        bericht.zeilen_gesamt += 1
        for feld in stumme_felder:
            _spalte(bericht, blattname, feld)["uebersprungen"] += 1
            bericht.ignoriert_gesperrt += 1
        uid = zeile["uid"]

        try:
            element = doc.GetElement(uid)
        except Exception:
            element = None

        if element is None:
            bericht.notiere_uebersprungen(
                blattname, zeile["excel_zeile"], uid,
                t(u"Element existiert nicht mehr (gelöscht oder aus anderem Projekt)", u"Element no longer exists (deleted or from another project)", u"El elemento ya no existe (eliminado o de otro proyecto)"))
            continue

        bezeichnung = _bezeichnung(element)
        instanz_map = parameter_map(element)
        typ_element = typ_von(doc, element)
        typ_map = parameter_map(typ_element) if typ_element is not None else {}

        for feld in felder:
            eintrag = _spalte(bericht, blattname, feld)
            eintrag["verglichen"] += 1
            zellwert = zeile["werte"].get(feld["spalte"])
            treffer = finde_parameter(doc, element, feld,
                                      instanz_map=instanz_map,
                                      typ_element=typ_element,
                                      typ_map=typ_map)
            if treffer is None:
                eintrag["fehler"] += 1
                bericht.notiere_fehler(
                    blattname, zeile["excel_zeile"], eid_wert(element.Id),
                    bezeichnung, feld["name"],
                    t(u"Parameter am Element nicht gefunden", u"Parameter not found on the element", u"Parámetro no encontrado en el elemento"))
                continue

            parameter = treffer.parameter

            # Schreibschutz kann sich seit dem Export geändert haben
            if ist_schreibgeschuetzt(parameter):
                eintrag["gesperrt"] += 1
                bericht.ignoriert_gesperrt += 1
                continue

            try:
                neuer_wert = zielwert(doc, parameter, zellwert)
            except Konvertierungsfehler as fehler:
                eintrag["fehler"] += 1
                bericht.notiere_fehler(
                    blattname, zeile["excel_zeile"], eid_wert(element.Id),
                    bezeichnung, feld["name"], str(fehler))
                continue
            except Exception as fehler:
                eintrag["fehler"] += 1
                bericht.notiere_fehler(
                    blattname, zeile["excel_zeile"], eid_wert(element.Id),
                    bezeichnung, feld["name"],
                    t(u"Wert nicht verwertbar: %s", u"Value not usable: %s", u"Valor no utilizable: %s") % fehler)
                continue

            # Diff-Erkennung: unveränderte Werte erzeugen keine Transaktion
            try:
                if not unterscheidet_sich(parameter, neuer_wert):
                    eintrag["gleich"] += 1
                    bericht.unveraendert += 1
                    continue
            except Exception:
                pass

            # Dasselbe Ziel darf nur einmal beschrieben werden. Relevant vor
            # allem bei Typparametern, die in jeder Instanzzeile auftauchen.
            schluessel = (eid_wert(treffer.besitzer.Id), eid_wert(parameter.Id))
            if schluessel in gesehen:
                frueher = gesehen[schluessel]
                if frueher.neuer_wert != neuer_wert:
                    eintrag["fehler"] += 1
                    bericht.notiere_fehler(
                        blattname, zeile["excel_zeile"], eid_wert(element.Id),
                        bezeichnung, feld["name"],
                        t(u"Widersprüchlicher Wert für denselben Parameter "
                        u"(bereits in Zeile %d gesetzt)", u"Conflicting value for the same parameter (already set in row %d)", u"Valor contradictorio para el mismo parámetro (ya definido en la fila %d)") % frueher.excel_zeile)
                else:
                    eintrag["gleich"] += 1
                    bericht.unveraendert += 1
                continue

            # Worksharing: fremd ausgeliehene Elemente würden die gesamte
            # Transaktion scheitern lassen
            besitzer_name = fremder_besitzer(doc, treffer.besitzer)
            if besitzer_name:
                eintrag["uebersprungen"] += 1
                bericht.notiere_uebersprungen(
                    blattname, zeile["excel_zeile"], uid,
                    t(u"Element ist von '%s' ausgeliehen (%s)", u"Element is borrowed by '%s' (%s)", u"El elemento lo tiene prestado '%s' (%s)")
                    % (besitzer_name, feld["name"]))
                continue

            _alt_zellwert, alt_anzeige = lese_wert(doc, parameter)
            aenderung = Aenderung(
                blattname, zeile["excel_zeile"], uid, element, feld["name"],
                parameter, treffer.besitzer, treffer.ist_typparameter,
                neuer_wert, alt_anzeige)
            gesehen[schluessel] = aenderung
            eintrag["geplant"] += 1
            bericht.aenderungen.append(aenderung)

    return gesehen


def wende_an(doc, aenderungen, bericht):
    """Schreibt die geplanten Änderungen. Erwartet eine offene Transaktion."""
    for aenderung in aenderungen:
        parameter = aenderung.parameter
        try:
            # Zustand kann sich durch vorherige Änderungen verschoben haben
            if ist_schreibgeschuetzt(parameter):
                bericht.ignoriert_gesperrt += 1
                continue
            if not unterscheidet_sich(parameter, aenderung.neuer_wert):
                bericht.unveraendert += 1
                continue
            erfolg = parameter.Set(aenderung.neuer_wert)
            if erfolg is False:
                bericht.notiere_fehler(
                    aenderung.blatt, aenderung.excel_zeile, aenderung.element_id,
                    _bezeichnung(aenderung.element), aenderung.feldname,
                    t(u"Revit hat den Wert abgelehnt", u"Revit rejected the value", u"Revit rechazó el valor"))
                continue
            aenderung.angewendet = True
        except Exception as fehler:
            bericht.notiere_fehler(
                aenderung.blatt, aenderung.excel_zeile, aenderung.element_id,
                _bezeichnung(aenderung.element), aenderung.feldname,
                t(u"Schreiben fehlgeschlagen: %s", u"Writing failed: %s", u"Error al escribir: %s") % fehler)
    return bericht
