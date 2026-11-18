# -*- coding: utf-8 -*-
"""Offline-Test der Logik von acc_publish (ohne Netz, ohne APS).

    python tools\\test_acc_publish.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import acc_publish as ap  # noqa: E402


def _konfig(**extra):
    k = {"client_id": "id", "client_secret": "geheim",
         "projekte": [{"name": "P1"}]}
    k.update(extra)
    return k


def test_konfig_standardwerte():
    k = ap.pruefe_konfig(_konfig(uhrzeit="8:05"))
    assert k["uhrzeit"] == "08:05"
    assert k["projekte"] == [{"name": "P1", "ordner": ["Project Files"],
                              "unterordner": True}]
    assert k["nur_bei_aenderungen"] is False
    k = ap.pruefe_konfig(_konfig(projekte=["P2"]))
    assert k["projekte"][0]["name"] == "P2"
    k = ap.pruefe_konfig(_konfig(projekte=[{"name": "P", "ordner": "A/B"}]))
    assert k["projekte"][0]["ordner"] == ["A/B"]


def test_konfig_fehler():
    for falsch in (_konfig(client_id=""), _konfig(uhrzeit="25:00"),
                   _konfig(uhrzeit="20 Uhr"), _konfig(projekte=[]),
                   _konfig(projekte=[{"ordner": ["x"]}]),
                   _konfig(redirect_uri="http://localhost/cb")):
        try:
            ap.pruefe_konfig(falsch)
        except ap.Fehler:
            continue
        raise AssertionError("kein Fehler fuer %r" % falsch)


def test_pfad_teile():
    assert ap.pfad_teile("Project Files/01 Modelle\\Arch/") == \
        ["Project Files", "01 Modelle", "Arch"]
    assert ap.pfad_teile("  ") == []


def test_zerlege_inhalt():
    seiten = [
        {"data": [
            {"type": "folders", "id": "f1",
             "attributes": {"displayName": "Arch"}},
            {"type": "items", "id": "i1",
             "attributes": {"displayName": "A.rvt",
                            "extension": {"type": ap.TYP_MODELL_ITEM}}},
            {"type": "items", "id": "i2",
             "attributes": {"displayName": "Plan.pdf",
                            "extension": {"type": "items:autodesk.bim360:File"}},
             "relationships": {"tip": {"data": {"id": "v2"}}}},
        ], "included": [
            {"type": "versions", "id": "v2", "attributes": {
                "extension": {"type": "versions:autodesk.bim360:File"}}}]},
        {"data": [
            {"type": "items", "id": "i3",
             "attributes": {"displayName": "B.rvt",
                            "extension": {"type": "items:autodesk.bim360:File"}},
             "relationships": {"tip": {"data": {"id": "v3"}}}},
        ], "included": [
            {"type": "versions", "id": "v3", "attributes": {
                "extension": {"type": ap.TYP_MODELL_VERSION}}}]},
    ]
    ordner, modelle = ap.zerlege_inhalt(seiten)
    assert ordner == [("f1", "Arch")]
    assert modelle == [("i1", "A.rvt"), ("i3", "B.rvt")]


def test_ausschliessen():
    assert ap.ausgeschlossen("TGA_ALT.rvt", ["*_alt*"])
    assert not ap.ausgeschlossen("ARC.rvt", ["*_alt*"])
    assert not ap.ausgeschlossen("ARC.rvt", [])


def test_befehl_body():
    b = ap.befehl_body(ap.CMD_PUBLISH, ["urn:x"])
    assert b["data"]["attributes"]["extension"]["type"] == ap.CMD_PUBLISH
    assert b["data"]["relationships"]["resources"]["data"] == \
        [{"type": "items", "id": "urn:x"}]


def test_braucht_publish():
    assert ap.braucht_publish({"data": {"attributes": {}}})
    assert not ap.braucht_publish({"data": None})
    assert ap.braucht_publish({})
    assert ap.braucht_publish(None)


def test_token_gueltig():
    assert ap.token_gueltig({"access_token": "a", "laeuft_ab": 1000}, 500)
    assert not ap.token_gueltig({"access_token": "a", "laeuft_ab": 1000}, 900)
    assert not ap.token_gueltig(None, 0)


def test_aufgabe_befehl():
    b = ap.aufgabe_befehl(r"C:\Py\pythonw.exe", r"C:\x\acc_publish.py",
                          "20:00")
    assert b[b.index("/ST") + 1] == "20:00"
    assert b[b.index("/SC") + 1] == "DAILY"
    assert b[-1] == r'"C:\Py\pythonw.exe" "C:\x\acc_publish.py" publizieren'


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items())
             if n.startswith("test_") and callable(f)]
    for name, f in tests:
        f()
        print("ok  " + name)
    print("%d Tests bestanden." % len(tests))
