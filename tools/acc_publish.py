# -*- coding: utf-8 -*-
"""ACC-Publish: Revit-Cloud-Modelle in der ACC automatisch publizieren.

Macht dasselbe wie im Revit-Startbildschirm "Neueste Version publizieren",
nur ueber die Autodesk-APS-Schnittstelle - Revit muss dafuer nicht laufen.
Nur Python-Standardbibliothek, keine Zusatzpakete.

Befehle (CPython 3):

    python tools\\acc_publish.py init          Konfiguration anlegen
    python tools\\acc_publish.py login         einmalig mit Autodesk anmelden
    python tools\\acc_publish.py liste         zeigen, was publiziert wuerde
    python tools\\acc_publish.py publizieren   jetzt publizieren
    python tools\\acc_publish.py aufgabe       taeglichen Windows-Task anlegen
    python tools\\acc_publish.py aufgabe --entfernen

Konfiguration, Anmeldung und Logs liegen in %APPDATA%\\pyMLG\\acc_publish
(nicht im Repo - dort steht das Client Secret).
"""
from __future__ import print_function

import base64
import fnmatch
import json
import logging
import os
import re
import subprocess
import sys
import threading
import time
import webbrowser
from datetime import datetime

try:  # CPython 3
    from http.server import BaseHTTPRequestHandler, HTTPServer
    from urllib.error import HTTPError, URLError
    from urllib.parse import parse_qs, quote, urlencode, urlparse
    from urllib.request import Request, urlopen
except ImportError:  # pragma: no cover - nur CPython 3 wird unterstuetzt
    raise SystemExit("acc_publish braucht CPython 3.")

APS = "https://developer.api.autodesk.com"
SCOPES = "data:read data:write data:create"
TASK_NAME = r"pyMLG\ACC Publish"

TYP_MODELL_ITEM = "items:autodesk.bim360:C4RModel"
TYP_MODELL_VERSION = "versions:autodesk.bim360:C4RModel"
CMD_PUBLISH = "commands:autodesk.bim360:C4RModelPublish"
CMD_JOB = "commands:autodesk.bim360:C4RModelGetPublishJob"

STANDARD_KONFIG = {
    "client_id": "",
    "client_secret": "",
    "redirect_uri": "http://localhost:8080/api/auth/callback",
    "uhrzeit": "20:00",
    "hub": "",
    "projekte": [
        {
            "name": "Projektname genau wie in der ACC",
            "ordner": ["Project Files"],
            "unterordner": True,
        }
    ],
    "ausschliessen": [],
    "nur_bei_aenderungen": False,
}

log = logging.getLogger("acc_publish")


class Fehler(Exception):
    """Fehler mit verstaendlicher Meldung fuer den Benutzer."""


# --------------------------------------------------------------------------
# Pfade
# --------------------------------------------------------------------------

def arbeitsordner():
    basis = os.environ.get("PYMLG_ACC_ORDNER")
    if basis:
        return basis
    appdata = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(appdata, "pyMLG", "acc_publish")


def pfad(name):
    return os.path.join(arbeitsordner(), name)


def schreibe_json(datei, daten):
    """Atomar schreiben, damit ein Abbruch keine halbe Datei hinterlaesst."""
    ordner = os.path.dirname(datei)
    if ordner and not os.path.isdir(ordner):
        os.makedirs(ordner)
    tmp = datei + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(daten, f, indent=2, ensure_ascii=False)
    os.replace(tmp, datei)


# --------------------------------------------------------------------------
# Reine Logik (offline testbar, siehe tools/test_acc_publish.py)
# --------------------------------------------------------------------------

def uhrzeit_gueltig(text):
    m = re.match(r"^(\d{1,2}):(\d{2})$", text or "")
    return bool(m) and int(m.group(1)) < 24 and int(m.group(2)) < 60


def uhrzeit_normal(text):
    std, minute = text.split(":")
    return "%02d:%02d" % (int(std), int(minute))


def pruefe_konfig(konfig):
    """Ergaenzt Standardwerte und liefert die Konfiguration oder wirft Fehler."""
    k = dict(STANDARD_KONFIG)
    k.update(konfig or {})
    probleme = []
    if not k["client_id"] or not k["client_secret"]:
        probleme.append("client_id und client_secret der APS-App fehlen.")
    if not uhrzeit_gueltig(k["uhrzeit"]):
        probleme.append("uhrzeit muss wie 20:00 aussehen.")
    if not urlparse(k["redirect_uri"]).port:
        probleme.append("redirect_uri braucht einen Port, z.B. "
                        "http://localhost:8080/api/auth/callback")
    projekte = []
    for i, p in enumerate(k.get("projekte") or []):
        if isinstance(p, str):
            p = {"name": p}
        if not p.get("name"):
            probleme.append("Projekt Nr. %d hat keinen Namen." % (i + 1))
            continue
        ordner = p.get("ordner") or ["Project Files"]
        if isinstance(ordner, str):
            ordner = [ordner]
        projekte.append({
            "name": p["name"],
            "ordner": ordner,
            "unterordner": p.get("unterordner", True),
        })
    if not projekte:
        probleme.append("Unter projekte ist kein Projekt eingetragen.")
    if probleme:
        raise Fehler("Konfiguration %s:\n  - %s"
                     % (pfad("config.json"), "\n  - ".join(probleme)))
    k["projekte"] = projekte
    k["uhrzeit"] = uhrzeit_normal(k["uhrzeit"])
    k["ausschliessen"] = list(k.get("ausschliessen") or [])
    return k


def pfad_teile(text):
    """'Project Files/01 Modelle\\\\Arch' -> ['Project Files', '01 Modelle', 'Arch']"""
    return [t.strip() for t in re.split(r"[\\/]+", text or "") if t.strip()]


def gleicher_name(a, b):
    return (a or "").strip().lower() == (b or "").strip().lower()


def name_von(eintrag):
    attr = eintrag.get("attributes") or {}
    return attr.get("displayName") or attr.get("name") or ""


def erweiterung(eintrag):
    return ((eintrag.get("attributes") or {}).get("extension") or {}).get("type")


def ist_revit_cloud_modell(item, versionen):
    """Revit-Cloud-Modell (Cloud Worksharing) erkennen.

    versionen: dict id -> Version aus 'included' der Ordner-Antwort.
    """
    if item.get("type") != "items":
        return False
    if erweiterung(item) == TYP_MODELL_ITEM:
        return True
    tip = (((item.get("relationships") or {}).get("tip") or {})
           .get("data") or {}).get("id")
    return erweiterung(versionen.get(tip, {})) == TYP_MODELL_VERSION


def zerlege_inhalt(seiten):
    """Ordner-Inhalt (Liste von API-Seiten) -> (ordner, modelle).

    ordner:  [(id, name)]
    modelle: [(item_id, name)]
    """
    ordner, modelle, eintraege, versionen = [], [], [], {}
    for seite in seiten:
        eintraege.extend(seite.get("data") or [])
        for v in seite.get("included") or []:
            versionen[v.get("id")] = v
    for e in eintraege:
        if e.get("type") == "folders":
            ordner.append((e["id"], name_von(e)))
        elif ist_revit_cloud_modell(e, versionen):
            modelle.append((e["id"], name_von(e)))
    return ordner, modelle


def ausgeschlossen(name, muster):
    n = (name or "").lower()
    return any(fnmatch.fnmatch(n, m.lower()) for m in muster)


def befehl_body(typ, item_ids):
    return {
        "jsonapi": {"version": "1.0"},
        "data": {
            "type": "commands",
            "attributes": {"extension": {"type": typ, "version": "1.0.0"}},
            "relationships": {"resources": {"data": [
                {"type": "items", "id": i} for i in item_ids]}},
        },
    }


def braucht_publish(job_antwort):
    """Auswertung von C4RModelGetPublishJob.

    Ohne neue Synchronisierungen liefert APS kein Job-Objekt (data = null).
    Bei Unklarheit lieber publizieren als still ueberspringen.
    """
    if not isinstance(job_antwort, dict) or "data" not in job_antwort:
        return True
    return job_antwort.get("data") is not None


def token_gueltig(token, jetzt=None):
    jetzt = time.time() if jetzt is None else jetzt
    return bool(token and token.get("access_token")
                and token.get("laeuft_ab", 0) - 120 > jetzt)


def aufgabe_befehl(python, skript, uhrzeit):
    return ["schtasks", "/Create", "/F", "/SC", "DAILY", "/ST", uhrzeit,
            "/TN", TASK_NAME,
            "/TR", '"%s" "%s" publizieren' % (python, skript)]


# --------------------------------------------------------------------------
# APS-Zugriff
# --------------------------------------------------------------------------

class Aps(object):
    def __init__(self, konfig):
        self.k = konfig
        self.token_datei = pfad("token.json")
        self._token = None

    # --- Anmeldung -------------------------------------------------------

    def _basic(self):
        roh = ("%s:%s" % (self.k["client_id"], self.k["client_secret"]))
        return "Basic " + base64.b64encode(roh.encode("utf-8")).decode("ascii")

    def _token_holen(self, felder):
        daten = self._roh("POST", APS + "/authentication/v2/token",
                          body=urlencode(felder).encode("utf-8"),
                          kopf={"Authorization": self._basic(),
                                "Content-Type":
                                    "application/x-www-form-urlencoded"},
                          mit_token=False)
        daten["laeuft_ab"] = time.time() + int(daten.get("expires_in", 3600))
        schreibe_json(self.token_datei, daten)
        self._token = daten
        return daten

    def login(self):
        ziel = urlparse(self.k["redirect_uri"])
        ergebnis = {}

        class Antwort(BaseHTTPRequestHandler):
            def do_GET(self):  # noqa: N802
                url = urlparse(self.path)
                if url.path.rstrip("/") != ziel.path.rstrip("/"):
                    self.send_response(404)
                    self.end_headers()
                    return
                q = parse_qs(url.query)
                ergebnis["code"] = (q.get("code") or [None])[0]
                ergebnis["fehler"] = (q.get("error_description")
                                      or q.get("error") or [None])[0]
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                text = ("Anmeldung erfolgreich - Fenster kann geschlossen "
                        "werden." if ergebnis["code"]
                        else "Anmeldung fehlgeschlagen.")
                self.wfile.write(("<h3>pyMLG ACC-Publish</h3><p>%s</p>"
                                  % text).encode("utf-8"))

            def log_message(self, *args):
                pass

        server = HTTPServer((ziel.hostname or "localhost", ziel.port), Antwort)
        server.timeout = 1
        url = APS + "/authentication/v2/authorize?" + urlencode({
            "response_type": "code",
            "client_id": self.k["client_id"],
            "redirect_uri": self.k["redirect_uri"],
            "scope": SCOPES,
        })
        print("Browser oeffnet sich zur Autodesk-Anmeldung. Falls nicht:\n"
              + url)
        threading.Timer(0.5, webbrowser.open, [url]).start()
        ende = time.time() + 300
        try:
            while "code" not in ergebnis and time.time() < ende:
                server.handle_request()
        finally:
            server.server_close()
        if not ergebnis.get("code"):
            raise Fehler("Keine Anmeldung erhalten (%s)."
                         % (ergebnis.get("fehler") or "Zeitlimit 5 Minuten"))
        self._token_holen({"grant_type": "authorization_code",
                           "code": ergebnis["code"],
                           "redirect_uri": self.k["redirect_uri"]})

    def access_token(self):
        if token_gueltig(self._token):
            return self._token["access_token"]
        if self._token is None and os.path.exists(self.token_datei):
            with open(self.token_datei, encoding="utf-8") as f:
                self._token = json.load(f)
            if token_gueltig(self._token):
                return self._token["access_token"]
        if not (self._token or {}).get("refresh_token"):
            raise Fehler("Nicht angemeldet. Bitte einmal ausfuehren:\n"
                         "  python tools\\acc_publish.py login")
        try:
            # Refresh-Token gilt 15 Tage und wird bei jeder Nutzung erneuert.
            self._token_holen({"grant_type": "refresh_token",
                               "refresh_token": self._token["refresh_token"],
                               "scope": SCOPES})
        except Fehler as e:
            raise Fehler("Anmeldung abgelaufen (%s). Bitte neu anmelden:\n"
                         "  python tools\\acc_publish.py login" % e)
        return self._token["access_token"]

    # --- HTTP ------------------------------------------------------------

    def _roh(self, methode, url, body=None, kopf=None, mit_token=True):
        for versuch in range(6):
            k = dict(kopf or {})
            if mit_token:
                k["Authorization"] = "Bearer " + self.access_token()
            try:
                with urlopen(Request(url, data=body, headers=k,
                                     method=methode), timeout=60) as r:
                    text = r.read().decode("utf-8")
                    return json.loads(text) if text.strip() else {}
            except HTTPError as e:
                text = e.read().decode("utf-8", "replace")
                if e.code == 429 or e.code >= 500:
                    warte = int(e.headers.get("Retry-After") or 0) \
                        or 2 ** (versuch + 1)
                    log.warning("HTTP %s, neuer Versuch in %ss", e.code, warte)
                    time.sleep(min(warte, 60))
                    continue
                raise Fehler("HTTP %s bei %s %s\n%s"
                             % (e.code, methode, url, text[:1000]))
            except URLError as e:
                if versuch < 3:
                    time.sleep(5)
                    continue
                raise Fehler("Keine Verbindung zu %s (%s)" % (url, e.reason))
        raise Fehler("APS antwortet nicht (%s %s)" % (methode, url))

    def get(self, pfad_):
        return self._roh("GET", pfad_ if pfad_.startswith("http")
                         else APS + pfad_)

    def seiten(self, pfad_):
        """Alle Seiten einer JSON:API-Liste."""
        url = pfad_
        while url:
            seite = self.get(url)
            yield seite
            weiter = (seite.get("links") or {}).get("next")
            url = weiter.get("href") if isinstance(weiter, dict) else weiter

    def befehl(self, projekt_id, typ, item_ids):
        return self._roh(
            "POST", APS + "/data/v1/projects/%s/commands" % quote(projekt_id),
            body=json.dumps(befehl_body(typ, item_ids)).encode("utf-8"),
            kopf={"Content-Type": "application/vnd.api+json"})

    # --- Navigation ------------------------------------------------------

    def projekte(self):
        """[(hub_id, hub_name, projekt_id, projekt_name)]"""
        treffer = []
        for seite in self.seiten("/project/v1/hubs"):
            for hub in seite.get("data") or []:
                if self.k["hub"] and not gleicher_name(name_von(hub),
                                                       self.k["hub"]):
                    continue
                for ps in self.seiten("/project/v1/hubs/%s/projects"
                                      % quote(hub["id"])):
                    for p in ps.get("data") or []:
                        treffer.append((hub["id"], name_von(hub),
                                        p["id"], name_von(p)))
        return treffer

    def inhalt(self, projekt_id, ordner_id):
        return zerlege_inhalt(self.seiten(
            "/data/v1/projects/%s/folders/%s/contents"
            % (quote(projekt_id), quote(ordner_id))))

    def ordner_finden(self, hub_id, projekt_id, teile):
        top = self.get("/project/v1/hubs/%s/projects/%s/topFolders"
                       % (quote(hub_id), quote(projekt_id))).get("data") or []
        start = [t for t in top if gleicher_name(name_von(t), teile[0])]
        if start:
            aktuell, rest = start[0]["id"], teile[1:]
        else:
            # Pfad ohne "Project Files" angegeben -> dort beginnen.
            pf = [t for t in top if gleicher_name(name_von(t),
                                                  "Project Files")]
            if not pf:
                return None
            aktuell, rest = pf[0]["id"], teile
        for teil in rest:
            unter, _ = self.inhalt(projekt_id, aktuell)
            passend = [o for o in unter if gleicher_name(o[1], teil)]
            if not passend:
                return None
            aktuell = passend[0][0]
        return aktuell

    def modelle(self, projekt_id, ordner_id, ordner_name, rekursiv):
        """[(item_id, anzeigepfad)]"""
        ergebnis, offen = [], [(ordner_id, ordner_name)]
        while offen:
            oid, oname = offen.pop(0)
            unter, modelle = self.inhalt(projekt_id, oid)
            ergebnis.extend((i, oname + "/" + n) for i, n in modelle)
            if rekursiv:
                offen.extend((u, oname + "/" + n) for u, n in unter)
        return ergebnis


# --------------------------------------------------------------------------
# Ablauf
# --------------------------------------------------------------------------

def sammle(aps, k):
    """[(projekt_name, projekt_id, item_id, pfad)] ohne Doppelte."""
    alle = aps.projekte()
    gesehen, liste = set(), []
    for p in k["projekte"]:
        passend = [a for a in alle if gleicher_name(a[3], p["name"])]
        if not passend:
            log.error("Projekt nicht gefunden: %s", p["name"])
            continue
        hub_id, _, projekt_id, projekt_name = passend[0]
        for o in p["ordner"]:
            teile = pfad_teile(o) or ["Project Files"]
            oid = aps.ordner_finden(hub_id, projekt_id, teile)
            if not oid:
                log.error("%s: Ordner nicht gefunden: %s", projekt_name, o)
                continue
            for item_id, anzeige in aps.modelle(projekt_id, oid,
                                                "/".join(teile),
                                                p["unterordner"]):
                if item_id in gesehen:
                    continue
                gesehen.add(item_id)
                if ausgeschlossen(anzeige.rsplit("/", 1)[-1],
                                  k["ausschliessen"]):
                    log.info("ausgeschlossen: %s", anzeige)
                    continue
                liste.append((projekt_name, projekt_id, item_id, anzeige))
    return liste


def publizieren(aps, k, nur_zeigen=False):
    liste = sammle(aps, k)
    if not liste:
        log.warning("Keine Revit-Cloud-Modelle gefunden.")
    fehler = 0
    for projekt_name, projekt_id, item_id, anzeige in liste:
        titel = "%s | %s" % (projekt_name, anzeige)
        try:
            if k["nur_bei_aenderungen"]:
                job = aps.befehl(projekt_id, CMD_JOB, [item_id])
                log.debug("PublishJob %s: %s", titel, json.dumps(job))
                if not braucht_publish(job):
                    log.info("aktuell, uebersprungen: %s", titel)
                    continue
            if nur_zeigen:
                log.info("wuerde publizieren: %s", titel)
                continue
            antwort = aps.befehl(projekt_id, CMD_PUBLISH, [item_id])
            status = ((antwort.get("data") or {}).get("attributes")
                      or {}).get("status", "?")
            log.info("publiziert (%s): %s", status, titel)
            log.debug("Antwort: %s", json.dumps(antwort))
        except Fehler as e:
            fehler += 1
            log.error("FEHLER bei %s: %s", titel, e)
    log.info("Fertig: %d Modelle, %d Fehler.", len(liste), fehler)
    return fehler


def lade_konfig():
    datei = pfad("config.json")
    if not os.path.exists(datei):
        raise Fehler("Keine Konfiguration. Bitte zuerst ausfuehren:\n"
                     "  python tools\\acc_publish.py init")
    with open(datei, encoding="utf-8") as f:
        try:
            return pruefe_konfig(json.load(f))
        except ValueError as e:
            raise Fehler("config.json ist kein gueltiges JSON: %s" % e)


def logging_einrichten(ausfuehrlich):
    ordner = pfad("logs")
    if not os.path.isdir(ordner):
        os.makedirs(ordner)
    datei = os.path.join(ordner, datetime.now().strftime("%Y-%m") + ".log")
    log.setLevel(logging.DEBUG)
    fh = logging.FileHandler(datei, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    log.addHandler(fh)
    if sys.stdout is not None:  # pythonw hat keine Konsole
        sh = logging.StreamHandler(sys.stdout)
        sh.setLevel(logging.DEBUG if ausfuehrlich else logging.INFO)
        sh.setFormatter(logging.Formatter("%(message)s"))
        log.addHandler(sh)
    return datei


def cmd_init():
    datei = pfad("config.json")
    if os.path.exists(datei):
        print("Konfiguration existiert schon: " + datei)
    else:
        schreibe_json(datei, STANDARD_KONFIG)
        print("Konfiguration angelegt: " + datei)
    print("Bitte client_id, client_secret und die Projekte eintragen.")
    if os.name == "nt":
        os.startfile(datei)  # noqa: S606 - oeffnet im Standard-Editor


def cmd_aufgabe(k, entfernen):
    if entfernen:
        befehl = ["schtasks", "/Delete", "/F", "/TN", TASK_NAME]
    else:
        python = sys.executable
        w = os.path.join(os.path.dirname(python), "pythonw.exe")
        if os.path.exists(w):
            python = w  # ohne Konsolenfenster
        befehl = aufgabe_befehl(python, os.path.abspath(__file__),
                                k["uhrzeit"])
    r = subprocess.run(befehl, capture_output=True, text=True)
    print((r.stdout or "") + (r.stderr or ""))
    if r.returncode == 0 and not entfernen:
        print("Task '%s' laeuft taeglich um %s Uhr (solange dieser "
              "Benutzer angemeldet ist)." % (TASK_NAME, k["uhrzeit"]))
    return r.returncode


def main(argv):
    befehl = argv[1] if len(argv) > 1 else "publizieren"
    ausfuehrlich = "-v" in argv
    try:
        if befehl == "init":
            cmd_init()
            return 0
        k = lade_konfig()
        if befehl == "aufgabe":
            return cmd_aufgabe(k, "--entfernen" in argv)
        logdatei = logging_einrichten(ausfuehrlich)
        aps = Aps(k)
        if befehl == "login":
            aps.login()
            log.info("Angemeldet. Token gespeichert in %s", aps.token_datei)
            return 0
        if befehl in ("liste", "publizieren"):
            log.info("=== ACC-Publish %s (%s) ===", befehl,
                     datetime.now().strftime("%d.%m.%Y %H:%M"))
            fehler = publizieren(aps, k, nur_zeigen=(befehl == "liste"))
            log.info("Log: %s", logdatei)
            return 1 if fehler else 0
        print(__doc__)
        return 2
    except Fehler as e:
        if log.handlers:
            log.error(str(e))
        else:
            print(str(e))
        return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
