# -*- coding: utf-8 -*-
"""Bewertungen der Kollisionen speichern - als JSON neben dem Bericht.

    TGA_vs_TWP.xml
    TGA_vs_TWP.xml.pymlg.json   <- Status, Klassifizierung, Kommentar

Liegt der Bericht in einem gemeinsamen Ordner, sieht das Team denselben
Stand. Vor jedem Schreiben wird die Datei neu gelesen und nur die geänderten
Kollisionen werden überschrieben - so gehen die Bewertungen eines Kollegen,
der gleichzeitig arbeitet, nicht verloren.

Das Revit-Modell bleibt unberührt: keine Transaktion, kein Parameter.
Ist der Ordner schreibgeschützt, landet die Datei unter
%LOCALAPPDATA%\\pyMLG\\ClashNavigator\\.

Einstellungen des Panels: %LOCALAPPDATA%\\pyMLG\\ClashNavigator.json
"""

import copy
import hashlib
import io
import json
import os
import time

ENDUNG = u".pymlg.json"
VERSION = 1
MAX_RUECKGAENGIG = 10


def benutzer():
    return (os.environ.get("USERNAME") or os.environ.get("USER")
            or u"").strip()


def jetzt():
    return time.strftime("%Y-%m-%d %H:%M")


def eigener_ordner():
    basis = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    ordner = os.path.join(basis, "pyMLG")
    try:
        if not os.path.isdir(ordner):
            os.makedirs(ordner)
    except Exception:
        ordner = os.environ.get("TEMP", ".")
    return ordner


def _lies_json(pfad):
    if not pfad or not os.path.isfile(pfad):
        return None
    try:
        with io.open(pfad, "r", encoding="utf-8") as datei:
            return json.loads(datei.read())
    except Exception:
        return None


def _schreibe_json(pfad, daten):
    """Erst in eine Nebendatei, dann austauschen - ein Absturz mitten im
    Schreiben hinterlässt so keine halbe Datei."""
    text = json.dumps(daten, indent=1, sort_keys=True, ensure_ascii=False)
    if not isinstance(text, type(u"")):
        text = text.decode("utf-8")
    zwischen = pfad + u".tmp"
    with io.open(zwischen, "w", encoding="utf-8") as datei:
        datei.write(text)
    if os.path.exists(pfad):
        os.remove(pfad)
    os.rename(zwischen, pfad)


def _schreibbar(ordner):
    probe = os.path.join(ordner, u".pymlg_probe")
    try:
        with io.open(probe, "w", encoding="utf-8") as datei:
            datei.write(u"x")
        os.remove(probe)
        return True
    except Exception:
        return False


def speicherpfad(berichtpfad):
    """Neben dem Bericht - sonst im eigenen Ordner (eindeutig per Hash)."""
    neben = berichtpfad + ENDUNG
    if os.path.isfile(neben) or _schreibbar(os.path.dirname(berichtpfad)):
        return neben
    kennung = hashlib.md5(os.path.abspath(berichtpfad).lower()
                          .encode("utf-8")).hexdigest()[:12]
    ordner = os.path.join(eigener_ordner(), "ClashNavigator")
    if not os.path.isdir(ordner):
        os.makedirs(ordner)
    return os.path.join(ordner, u"%s_%s%s" % (
        os.path.basename(berichtpfad), kennung, ENDUNG))


class Speicher(object):
    """Bewertungen eines Berichts mit Rückgängig-Stapel.

    eintraege  {schluessel: {"status", "klasse", "kommentar", "von", "zeit"}}
    """

    def __init__(self, pfad):
        self.pfad = pfad
        self.eintraege = {}
        self.stapel = []
        self.neu_laden()

    def neu_laden(self):
        daten = _lies_json(self.pfad) or {}
        self.eintraege = dict(daten.get(u"clashes") or {})

    def _schreibe(self, geaendert):
        """Geänderte Schlüssel in die aktuelle Datei einarbeiten.
        geaendert: {schluessel: eintrag oder None (= löschen)}"""
        daten = _lies_json(self.pfad) or {}
        auf_platte = dict(daten.get(u"clashes") or {})
        for schluessel, eintrag in geaendert.items():
            if eintrag is None:
                auf_platte.pop(schluessel, None)
            else:
                auf_platte[schluessel] = eintrag
        daten[u"version"] = VERSION
        daten[u"clashes"] = auf_platte
        _schreibe_json(self.pfad, daten)
        # Stand der Kollegen übernehmen
        self.eintraege = auf_platte

    def anwenden(self, schluessel_liste, aenderung, beschreibung):
        """aenderung: Felder, die gesetzt werden, z.B. {"status": "geprueft",
        "klasse": "CLR"}. Leere Texte löschen das Feld.

        Rückgabe: Zahl der geänderten Kollisionen.
        """
        vorher = {}
        geaendert = {}
        for schluessel in schluessel_liste:
            alt = self.eintraege.get(schluessel)
            neu = dict(alt or {})
            for feld, wert in aenderung.items():
                if wert:
                    neu[feld] = wert
                else:
                    neu.pop(feld, None)
            inhalt = dict((k, v) for k, v in neu.items()
                          if k not in (u"von", u"zeit"))
            if inhalt == dict((k, v) for k, v in (alt or {}).items()
                              if k not in (u"von", u"zeit")):
                continue
            if inhalt:
                neu[u"von"] = benutzer()
                neu[u"zeit"] = jetzt()
            vorher[schluessel] = copy.deepcopy(alt)
            geaendert[schluessel] = neu if inhalt else None
        if not geaendert:
            return 0
        self._schreibe(geaendert)
        self.stapel.append({u"beschreibung": beschreibung,
                            u"zeit": time.strftime("%H:%M"),
                            u"vorher": vorher,
                            u"anzahl": len(geaendert)})
        del self.stapel[:-MAX_RUECKGAENGIG]
        return len(geaendert)

    def rueckgaengig(self):
        """Letzte Aktion zurücknehmen. Rückgabe: ihre Beschreibung oder
        None, wenn nichts zurückzunehmen ist."""
        if not self.stapel:
            return None
        schritt = self.stapel.pop()
        self._schreibe(schritt[u"vorher"])
        return schritt[u"beschreibung"]

    @property
    def letzter_schritt(self):
        return self.stapel[-1] if self.stapel else None


# ---------------------------------------------------------------------------
# Einstellungen des Panels
# ---------------------------------------------------------------------------

VORGABEN = {
    u"rand_cm": 15,
    u"schnitt": True,
    u"nur_offen": False,
    u"nur_aktiv": True,
    u"aktive_ansicht": False,
    u"auswaehlen": True,
    u"weiter": True,
    u"koordinaten": u"auto",
    u"ebene1": u"status",
    u"ebene2": u"",
    u"tabelle": False,
    u"letzte_datei": u"",
    u"klassen": [],
    u"abstand_cm": 5,
    u"winkel45": True,
    u"winkel90": True,
    u"vorrang": [],                 # leer = Standard aus vorrang.py
}


def einstellungspfad():
    return os.path.join(eigener_ordner(), u"ClashNavigator.json")


def lies_einstellungen():
    werte = dict(VORGABEN)
    gelesen = _lies_json(einstellungspfad())
    if isinstance(gelesen, dict):
        for schluessel in VORGABEN:
            if schluessel in gelesen:
                werte[schluessel] = gelesen[schluessel]
    return werte


def schreibe_einstellungen(werte):
    try:
        _schreibe_json(einstellungspfad(), werte)
    except Exception:
        pass
