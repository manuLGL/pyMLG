# -*- coding: utf-8 -*-
"""Offline-Test des Clash Navigators: Berichte, Logik, Speicher (ohne Revit).

    python tools\\test_clash_navigator.py
"""
import io
import os
import shutil
import sys
import tempfile

# Texte werden auf Deutsch verglichen
os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from clash_navigator import bericht as br  # noqa: E402
from clash_navigator import logik as lg  # noqa: E402
from clash_navigator import speicher as sp  # noqa: E402
from clash_navigator import umgehung as ug  # noqa: E402
from clash_navigator import vorrang as vr  # noqa: E402

XML = u"""<?xml version="1.0" encoding="UTF-8"?>
<exchange xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
          units="mm" filename="Koordination.nwf">
  <batchtest name="Koordination" units="mm">
    <clashtests>
      <clashtest name="TGA vs TWP" test_type="hard" status="ok">
        <clashresults>
          <clashresult name="Kollision1" guid="aaa-111" status="new"
                       distance="-52.0">
            <description>Hart</description>
            <resultstatus>Neu</resultstatus>
            <clashpoint><pos3f x="12300" y="4500" z="6700"/></clashpoint>
            <gridlocation>C-4 : Ebene 01</gridlocation>
            <createddate><date year="2026" month="9" day="1" hour="10"
                               minute="5" second="0"/></createddate>
            <comments><comment id="1"><user>Ana</user>
              <body>Bitte prüfen</body></comment></comments>
            <clashobjects>
              <clashobject>
                <objectattribute><name>Element ID</name>
                  <value>398254</value></objectattribute>
                <layer>Ebene 01</layer>
                <pathlink><node>Datei</node><node>TGA_EG.nwc</node>
                  <node>Ebene 01</node><node>Rohre</node></pathlink>
                <smarttags>
                  <smarttag><name>Item Name</name><value>Rundrohr</value>
                  </smarttag>
                  <smarttag><name>Item Type</name><value>Rohre</value>
                  </smarttag>
                </smarttags>
              </clashobject>
              <clashobject>
                <objectattribute><name>ID de elemento</name>
                  <value>12034</value></objectattribute>
                <pathlink><node>Datei</node><node>TWP.rvt</node></pathlink>
                <smarttags>
                  <smarttag><name>Item Name</name><value>STB 30</value>
                  </smarttag>
                </smarttags>
              </clashobject>
            </clashobjects>
          </clashresult>
          <clashgroup name="Gruppe A" guid="grp-1" status="active">
            <clashpoint><pos3f x="1000" y="2000" z="3000"/></clashpoint>
            <clashresults>
              <clashresult name="Kollision2" guid="bbb-222"
                           status="reviewed" distance="-10">
                <clashobjects/>
              </clashresult>
              <clashresult name="Kollision3" guid="ccc-333"
                           status="resolved" distance="-1">
                <clashpoint><pos3f x="0" y="0" z="1000"/></clashpoint>
              </clashresult>
            </clashresults>
          </clashgroup>
        </clashresults>
      </clashtest>
      <clashtest name="Leer"><clashresults/></clashtest>
    </clashtests>
  </batchtest>
</exchange>
"""

CSV = (u"Kollisionsname;Status;Abstand (mm);Test;Kollisionspunkt;"
       u"Element 1 ID;Datei 1;Item 1 Name;Element 2 ID;Datei 2;"
       u"Item 2 Name;Kommentar\r\n"
       u"Kollision1;Neu;-52,0;TGA vs TWP;\"x:12300, y:4500, z:6700\";"
       u"398254;TGA_EG.nwc;Rundrohr;12034;TWP.rvt;STB 30;\"Rohr; neu\"\r\n"
       u"Kollision2;Geprüft;-10;TGA vs TWP;\"1000; 2000; 3000\";;;;;;;\r\n")


class Ordner(object):
    def __enter__(self):
        self.pfad = tempfile.mkdtemp(prefix="pymlg_test_")
        return self.pfad

    def __exit__(self, *_):
        shutil.rmtree(self.pfad, ignore_errors=True)


def schreibe(ordner, name, text, kodierung="utf-8"):
    pfad = os.path.join(ordner, name)
    with io.open(pfad, "w", encoding=kodierung, newline="") as datei:
        datei.write(text)
    return pfad


# --- XML --------------------------------------------------------------------

def xml_bericht():
    with Ordner() as ordner:
        return br.lies(schreibe(ordner, "bericht.xml", XML))


def test_xml_grunddaten():
    bericht = xml_bericht()
    assert bericht.format == u"XML"
    assert [c.name for c in bericht.clashes] == [
        u"Kollision1", u"Kollision2", u"Kollision3"]
    assert bericht.tests == [u"TGA vs TWP"]


def test_xml_einheiten_und_punkt():
    clash = xml_bericht().clashes[0]
    assert abs(clash.abstand - (-0.052)) < 1e-9
    assert max(abs(a - b) for a, b in zip(clash.punkt,
                                          (12.3, 4.5, 6.7))) < 1e-9


def test_xml_felder():
    clash = xml_bericht().clashes[0]
    assert clash.schluessel == u"aaa-111"
    assert clash.status_navis == br.NEU
    assert clash.raster == u"C-4 : Ebene 01"
    assert clash.ebene == u"Ebene 01"
    assert clash.datum == u"2026-09-01 10:05"
    assert clash.kommentare_navis == [u"Ana: Bitte prüfen"]


def test_xml_objekte():
    a, b = xml_bericht().clashes[0].objekte
    assert (a.element_id, a.datei, a.name, a.typ) == (
        398254, u"TGA_EG.nwc", u"Rundrohr", u"Rohre")
    # spanischer Eigenschaftsname, Datei aus dem Pfad
    assert (b.element_id, b.datei, b.name) == (12034, u"TWP.rvt", u"STB 30")


def test_xml_gruppen():
    _eins, zwei, drei = xml_bericht().clashes
    assert zwei.gruppe == u"Gruppe A"
    assert zwei.status_navis == br.GEPRUEFT
    assert drei.status_navis == br.BEHOBEN
    # ohne eigenen Punkt gilt der Punkt der Gruppe, umgerechnet in Meter
    assert zwei.punkt == (1.0, 2.0, 3.0)
    assert drei.punkt == (0.0, 0.0, 1.0)


def test_xml_ohne_tests():
    with Ordner() as ordner:
        pfad = schreibe(ordner, "x.xml", u"<irgendwas/>")
        try:
            br.lies(pfad)
        except br.Fehler as fehler:
            assert u"keine Kollisionstests" in u"%s" % fehler
            return
    raise AssertionError(u"Fehler erwartet")


def test_nwd_erklaert():
    try:
        br.lies(u"C:\\modell.nwd")
    except br.Fehler as fehler:
        assert u"ohne Navisworks nicht" in u"%s" % fehler
        return
    raise AssertionError(u"Fehler erwartet")


# --- CSV --------------------------------------------------------------------

def csv_bericht(kodierung="utf-8"):
    with Ordner() as ordner:
        return br.lies(schreibe(ordner, "bericht.csv", CSV, kodierung))


def test_csv_spalten():
    clash_spalten, objekt_spalten, einheit = br.ordne_spalten(
        CSV.split(u"\r\n")[0].split(u";"))
    assert einheit == u"mm"
    assert clash_spalten[u"name"] == 0
    assert clash_spalten[u"punkt"] == 4
    assert objekt_spalten[(1, u"id")] == 5
    assert objekt_spalten[(1, u"datei")] == 6
    assert objekt_spalten[(1, u"name")] == 7
    assert objekt_spalten[(2, u"id")] == 8


def test_csv_werte():
    for kodierung in ("utf-8", "utf-8-sig", "cp1252", "utf-16"):
        eins, zwei = csv_bericht(kodierung).clashes
        assert eins.status_navis == br.NEU
        assert abs(eins.abstand - (-0.052)) < 1e-9
        assert max(abs(a - b) for a, b in zip(eins.punkt,
                                              (12.3, 4.5, 6.7))) < 1e-9
        assert eins.objekte[0].element_id == 398254
        assert eins.objekte[1].datei == u"TWP.rvt"
        assert eins.kommentare_navis == [u"Rohr; neu"]
        assert zwei.status_navis == br.GEPRUEFT, kodierung
        assert zwei.schluessel == u"TGA vs TWP|Kollision2"
        assert zwei.punkt == (1.0, 2.0, 3.0)
        assert zwei.objekte[0].element_id is None


def test_csv_englisch_komma():
    text = (u"Clash Name,Status,Distance,Item 1 Element ID,"
            u"Item 1 Source File,Item 2 Element ID,Clash Point X,"
            u"Clash Point Y,Clash Point Z\n"
            u"Clash1,Active,-0.05,11,A.nwc,22,1.5,2.5,3.5\n")
    with Ordner() as ordner:
        clash = br.lies(schreibe(ordner, "c.csv", text)).clashes[0]
    assert clash.status_navis == br.AKTIV
    assert clash.punkt == (1.5, 2.5, 3.5)
    assert [o.element_id for o in clash.objekte] == [11, 22]
    assert clash.objekte[0].datei == u"A.nwc"


def test_punkt_aus_text():
    assert br.punkt_aus_text(u"x:1.5 m, y:-2 m, z:3,25 m") == (1.5, -2.0,
                                                               3.25)
    assert br.punkt_aus_text(u"1,5; 2,5; 3,5") == (1.5, 2.5, 3.5)
    assert br.punkt_aus_text(u"1.5, 2.5, 3.5") == (1.5, 2.5, 3.5)
    assert br.punkt_aus_text(u"1.5 2.5") is None


def test_eindeutige_schluessel():
    clashes = []
    for _ in range(3):
        clash = br.Clash()
        clash.name, clash.test = u"K", u"T"
        clashes.append(clash)
    assert [c.schluessel for c in br._eindeutig(clashes)] == [
        u"T|K", u"T|K#2", u"T|K#3"]


# --- Logik ------------------------------------------------------------------

def test_modell_passt():
    namen = [u"TGA_EG_manuel", u"TGA_EG.rvt"]
    assert lg.modell_passt(u"TGA_EG.nwc", namen)
    assert lg.modell_passt(u"C:\\x\\TGA_EG.rvt.nwc", namen)
    assert not lg.modell_passt(u"TWP.rvt", namen)
    assert not lg.modell_passt(u"", namen)
    # kurze Namen passen nicht über den Anfang
    assert not lg.modell_passt(u"TG.nwc", [u"TGA_EG"])


def test_status_und_filter():
    clashes = xml_bericht().clashes
    eintraege = {u"aaa-111": {u"status": br.GEPRUEFT, u"klasse": u"CLR"}}
    assert lg.status(clashes[0], eintraege[u"aaa-111"]) == br.GEPRUEFT
    offen = lg.filtere(clashes, eintraege, nur_offen=True)
    assert offen == []           # 1 geprüft (Bewertung), 2 geprüft, 3 behoben
    assert lg.filtere(clashes, {}, nur_offen=True) == clashes[:1]
    assert lg.filtere(clashes, eintraege, suche=u"clr") == clashes[:1]
    assert lg.filtere(clashes, {}, suche=u"12034") == clashes[:1]
    assert lg.filtere(clashes, {}, modellnamen=[u"TWP"]) == clashes
    assert lg.filtere(clashes, {}, modellnamen=[u"Anderes"]) == clashes[1:]


def test_baum():
    clashes = xml_bericht().clashes
    oben = lg.baum(clashes, {}, u"status")
    assert [wert for wert, _ in oben] == [br.NEU, br.GEPRUEFT, br.BEHOBEN]
    zwei = lg.baum(clashes, {}, u"test", u"gruppe")
    assert zwei[0][0] == u"TGA vs TWP"
    assert [wert for wert, _ in zwei[0][1]] == [u"Gruppe A", u"(ohne)"]
    je_element = lg.baum(clashes[:1], {}, u"element")
    assert len(je_element) == 2
    assert lg.baum(clashes, {}, u"") == [(u"", clashes)]


def test_blatt_text():
    clash = xml_bericht().clashes[0]
    text = lg.blatt_text(clash, {u"klasse": u"CLR"})
    assert text == (u"Kollision1  \u00b7  Rundrohr \u2194 STB 30  \u00b7  "
                    u"-5.2 cm  \u00b7  [CLR]")


def test_kasten_schnittmenge():
    rohr = ((0.0, 0.0, 0.0), (40.0, 0.2, 0.2))
    wand = ((10.0, -1.0, -1.0), (10.3, 1.0, 3.0))
    klein, gross = lg.kasten_fuer([rohr, wand], rand=0.1, mindestmass=0.0)
    assert max(abs(a - b) for a, b in zip(klein, (9.9, -0.1, -0.1))) < 1e-9
    assert max(abs(a - b) for a, b in zip(gross, (10.4, 0.3, 0.3))) < 1e-9


def test_kasten_vereinigung():
    a = ((0.0, 0.0, 0.0), (1.0, 1.0, 1.0))
    b = ((5.0, 5.0, 5.0), (6.0, 6.0, 6.0))
    # getrennte Kästen: Schnittmenge leer -> beide ganz
    assert lg.kasten_fuer([a, b], rand=0.0) == ((0.0,) * 3, (6.0,) * 3)
    assert lg.kasten_fuer([a, b], rand=0.0, schnitt=False) == \
        ((0.0,) * 3, (6.0,) * 3)


def test_kasten_ein_langes_element():
    rohr = ((0.0, 0.0, 0.0), (40.0, 0.2, 0.2))
    klein, gross = lg.kasten_fuer([rohr], punkt=(20.0, 0.1, 0.1), rand=0.0,
                                  umkreis=2.0, mindestmass=0.0)
    assert (klein[0], gross[0]) == (18.0, 22.0)


def test_kasten_nur_punkt_und_mindestmass():
    klein, gross = lg.kasten_fuer([], punkt=(1.0, 2.0, 3.0), rand=0.15,
                                  mindestmass=0.5)
    assert max(abs(a - b) for a, b in zip(klein, (0.75, 1.75, 2.75))) < 1e-9
    assert lg.kasten_fuer([], None) is None


# --- Umgehung ---------------------------------------------------------------

def rohr_quer():
    # Rohr ø 32 mm, keine Dämmung
    return ug.Querschnitt(durchmesser=0.032)


def nahe(a, b, toleranz=1e-6):
    return max(abs(x - y) for x, y in zip(a, b)) < toleranz


def test_umgehung_oben_45():
    """Rohr entlang X in 3 m Höhe, Träger 0.3 m breit und bis 3.2 m hoch."""
    traeger = ((4.85, -2.0, 2.8), (5.15, 2.0, 3.2))
    alle = ug.varianten((0.0, 0.0, 3.0), (10.0, 0.0, 3.0), rohr_quer(),
                        traeger, abstand=0.05)
    oben = [v for v in alle if v.richtung == u"oben" and v.winkel == 45][0]
    # Versatz: Oberkante 0.2 über der Achse + 5 cm + halbes Rohr
    assert abs(oben.versatz - (0.2 + 0.05 + 0.016)) < 1e-9
    p1, p1o, p2o, p2 = oben.punkte
    assert nahe(p1o, (4.85 - 0.05 - 0.016, 0.0, 3.266))
    assert nahe(p2o, (5.15 + 0.05 + 0.016, 0.0, 3.266))
    # 45°: Anlauf gleich Versatz
    assert abs((p1o[0] - p1[0]) - oben.versatz) < 1e-9
    assert abs((p2[0] - p2o[0]) - oben.versatz) < 1e-9
    assert p1[2] == 3.0 and p2[2] == 3.0
    erwartet = 2.0 * (oben.versatz * 2 ** 0.5 - oben.versatz)
    assert abs(oben.zusatzlaenge - erwartet) < 1e-9
    assert oben.gueltig


def test_umgehung_richtungen_und_winkel():
    traeger = ((4.85, -2.0, 2.8), (5.15, 2.0, 3.2))
    alle = ug.varianten((0.0, 0.0, 3.0), (10.0, 0.0, 3.0), rohr_quer(),
                        traeger)
    assert sorted(set(v.richtung for v in alle)) == [
        u"links", u"oben", u"rechts", u"unten"]
    assert len(alle) == 8
    # seitlich: der Träger ist 4 m breit -> Versatz > 1.5 m -> verworfen
    seitlich = [v for v in alle if v.richtung in (u"links", u"rechts")]
    assert all(not v.gueltig for v in seitlich)
    links = [v for v in alle if v.richtung == u"links"][0]
    # links vom Blick entlang +X ist +Y
    assert links.punkte[1][1] > 0


def test_umgehung_bewertung():
    traeger = ((4.85, -2.0, 2.8), (5.15, 2.0, 3.2))
    alle = ug.varianten((0.0, 0.0, 3.0), (10.0, 0.0, 3.0), rohr_quer(),
                        traeger)
    drei = ug.beste(alle)
    assert len(drei) == 3
    assert drei[0].name == u"Oben 45°"      # kürzester Weg
    # eine Kollision schiebt die Variante nach hinten
    drei[0].kollisionen = [u"Wände 123"]
    assert ug.beste(alle)[0] is not drei[0]


def test_umgehung_zu_kurz_und_gefaelle():
    traeger = ((0.35, -2.0, 2.8), (0.45, 2.0, 3.2))
    alle = ug.varianten((0.0, 0.0, 3.0), (0.8, 0.0, 3.0), rohr_quer(),
                        traeger)
    unten = [v for v in alle if v.richtung == u"unten" and v.winkel == 45][0]
    assert not unten.gueltig and u"zu kurz" in unten.grund


def test_umgehung_gefaelleleitung():
    """1 % Gefälle: Varianten gibt es, das Mittelstück bleibt parallel,
    nach oben gibt es einen Hochpunkt und rutscht nach hinten."""
    traeger = ((4.85, -2.0, 2.8), (5.15, 2.0, 3.2))
    start, ende = (0.0, 0.0, 3.0), (10.0, 0.0, 2.9)
    assert ug.hat_gefaelle(start, ende)
    alle = ug.varianten(start, ende, rohr_quer(), traeger)
    oben = [v for v in alle if v.richtung == u"oben" and v.winkel == 45][0]
    unten = [v for v in alle if v.richtung == u"unten" and v.winkel == 45][0]
    assert oben.hochpunkt and not unten.hochpunkt
    assert oben.warnungen[0][1] is True          # kritisch
    assert unten.warnungen == [(u"Gefälle bleibt im Mittelstück erhalten",
                                False)]
    # Mittelstück parallel zur Leitung: gleiches Gefälle
    _p1, p1o, p2o, _p2 = unten.punkte
    gefaelle_neu = (p2o[2] - p1o[2]) / (p2o[0] - p1o[0])
    assert abs(gefaelle_neu - (-0.01)) < 1e-9
    assert ug.beste(alle)[0].richtung == u"unten"
    # waagerechte Leitung: keine Warnungen
    waagerecht = ug.varianten((0.0, 0.0, 3.0), (10.0, 0.0, 3.0), rohr_quer(),
                              traeger)
    assert all(not v.warnungen for v in waagerecht)


def test_verschieben_kurzes_stueck():
    """Kurzes Stück (0.8 m) zwischen zwei Bögen, beide Anschlussleitungen
    gehen nach oben: Umgehungen passen nicht, Verschieben nach unten schon."""
    traeger = ((0.35, -2.0, 2.8), (0.45, 2.0, 3.2))
    nach_oben = ((0.0, 0.0, 1.0), 2.0)
    alle = ug.varianten((0.0, 0.0, 3.0), (0.8, 0.0, 3.0), rohr_quer(),
                        traeger, nachbarn=(nach_oben, nach_oben))
    schieben = dict((v.richtung, v) for v in alle
                    if v.art == ug.VERSCHIEBEN)
    assert sorted(schieben) == [u"links", u"oben", u"rechts", u"unten"]
    unten = schieben[u"unten"]
    assert unten.gueltig
    assert unten.name == u"Verschieben: Unten"
    assert u"keine neuen Bögen" in unten.zusammenfassung()
    # nach unten: beide Anschlussleitungen werden um den Versatz länger
    assert abs(unten.zusatzlaenge - 2.0 * unten.versatz) < 1e-9
    assert len(unten.segmente) == 3
    # seitlich: Anschlussleitungen liegen nicht in Verschieberichtung
    assert not schieben[u"links"].gueltig
    assert u"Verschieberichtung" in schieben[u"links"].grund
    # nach oben: Anschlussleitungen würden kürzer - bei 2 m kein Problem
    assert schieben[u"oben"].gueltig
    assert schieben[u"oben"].zusatzlaenge == 0.0
    # die besten drei beginnen mit einer gültigen Verschiebung
    assert ug.beste(alle)[0].art == ug.VERSCHIEBEN


def test_verschieben_grenzen():
    traeger = ((0.35, -2.0, 2.8), (0.45, 2.0, 3.2))
    kurz_oben = ((0.0, 0.0, 1.0), 0.1)
    alle = ug.varianten((0.0, 0.0, 3.0), (0.8, 0.0, 3.0), rohr_quer(),
                        traeger, nachbarn=(kurz_oben, ug.FEST))
    schieben = [v for v in alle if v.art == ug.VERSCHIEBEN]
    assert all(not v.gueltig for v in schieben)
    assert any(u"Abzweig" in v.grund or u"zu kurz" in v.grund
               for v in schieben)
    # freie Enden: Verschieben geht (seitlich ist der Träger 4 m breit ->
    # Versatz zu groß)
    frei = ug.varianten((0.0, 0.0, 3.0), (0.8, 0.0, 3.0), rohr_quer(),
                        traeger, nachbarn=(None, None))
    frei = dict((v.richtung, v) for v in frei if v.art == ug.VERSCHIEBEN)
    assert frei[u"oben"].gueltig and frei[u"unten"].gueltig
    assert u"zu groß" in frei[u"links"].grund
    # ohne nachbarn gibt es keine Verschiebung
    ohne = ug.varianten((0.0, 0.0, 3.0), (0.8, 0.0, 3.0), rohr_quer(),
                        traeger)
    assert not [v for v in ohne if v.art == ug.VERSCHIEBEN]


def test_abstands_runden_und_warnung():
    assert ug.abstands_runden(0.10) == [0.10, 0.05, 0.02]
    assert ug.abstands_runden(0.05) == [0.05, 0.02]
    assert ug.abstands_runden(0.02) == [0.02]
    traeger = ((4.85, -2.0, 2.8), (5.15, 2.0, 3.2))
    voll = ug.varianten((0.0, 0.0, 3.0), (10.0, 0.0, 3.0), rohr_quer(),
                        traeger, abstand=0.10, soll_abstand=0.10)
    knapp = ug.varianten((0.0, 0.0, 3.0), (10.0, 0.0, 3.0), rohr_quer(),
                         traeger, abstand=0.02, soll_abstand=0.10)
    assert all(not v.warnungen for v in voll)
    k = [v for v in knapp if v.richtung == u"unten" and v.winkel == 45][0]
    assert k.warnungen[0] == (u"Nur 2 cm Sicherheitsabstand (Vorgabe 10 cm)",
                              True)
    # gleiche Variante mit vollem Abstand gewinnt trotz längerem Weg
    v = [v for v in voll if v.richtung == u"unten" and v.winkel == 45][0]
    assert v.kosten < k.kosten


def test_beide_weichen_aus():
    """Zwei Rohre kreuzen sich auf gleicher Höhe: A (entlang X) nach unten,
    B (entlang Y) nach oben - je der halbe Versatz."""
    rohr_a = ((0.0, 0.0, 3.0), (10.0, 0.0, 3.0))
    rohr_b = ((5.0, -5.0, 3.0), (5.0, 5.0, 3.0))
    kasten_b = ((4.984, -5.0, 2.984), (5.016, 5.0, 3.016))
    kasten_a = ((0.0, -0.016, 2.984), (10.0, 0.016, 3.016))
    voll = ug.voller_versatz(rohr_a[0], rohr_a[1], rohr_quer(), kasten_b,
                             0.05, u"unten")
    # Unterkante B 1.6 cm unter der Achse + 5 cm + halbes Rohr, mindestens
    # Rohr + Abstand
    assert abs(voll - max(0.016 + 0.05 + 0.016, 0.032 + 0.05)) < 1e-9
    halb = voll / 2.0
    teil_a = ug.varianten(rohr_a[0], rohr_a[1], rohr_quer(), kasten_b,
                          abstand=0.05, nur_richtungen=[u"unten"],
                          fester_versatz=halb)
    teil_b = ug.varianten(rohr_b[0], rohr_b[1], rohr_quer(), kasten_a,
                          abstand=0.05, nur_richtungen=[u"oben"],
                          fester_versatz=halb)
    assert set(v.richtung for v in teil_a) == set([u"unten"])
    assert all(abs(v.versatz - halb) < 1e-12 for v in teil_a + teil_b)
    a, b = ug.beste(teil_a, 1)[0], ug.beste(teil_b, 1)[0]
    # A geht runter, B hoch: zusammen der volle Versatz Abstand
    assert a.punkte[1][2] < 3.0 < b.punkte[1][2]
    assert abs((b.punkte[1][2] - a.punkte[1][2]) - voll) < 1e-9
    kombi = ug.Kombination([a, b])
    assert kombi.gueltig
    assert abs(kombi.versatz - halb) < 1e-12
    assert abs(kombi.kosten - (a.kosten + b.kosten + ug.AUFSCHLAG_BEIDE)) \
        < 1e-12
    assert kombi.name == u"Unten 45° + Oben 45°"
    assert u"Beide je" in kombi.zusammenfassung()
    b.kollisionen = [u"Wände 1"]
    assert kombi.kollisionen == [u"Wände 1"]


def test_zu_kurz_mit_zahlen():
    traeger = ((0.35, -2.0, 2.8), (0.45, 2.0, 3.2))
    alle = ug.varianten((0.0, 0.0, 3.0), (0.8, 0.0, 3.0), rohr_quer(),
                        traeger)
    grund = [v for v in alle if v.richtung == u"unten"
             and v.winkel == 45][0].grund
    assert u"Stück 0.80 m" in grund and u"nötig" in grund


def test_umgehung_kanal_und_steigleitung():
    kanal = ug.Querschnitt(breite=0.6, hoehe=0.3, achse_b=(0.0, 1.0, 0.0),
                           achse_h=(0.0, 0.0, 1.0), daemmung=0.03)
    assert abs(kanal.halb((0.0, 0.0, 1.0)) - 0.18) < 1e-9
    assert abs(kanal.halb((0.0, 1.0, 0.0)) - 0.33) < 1e-9
    assert kanal.text() == u"600 × 300 mm"
    # Steigleitung: vier waagerechte Richtungen
    alle = ug.varianten((0.0, 0.0, 0.0), (0.0, 0.0, 10.0), rohr_quer(),
                        ((-0.1, -0.1, 4.9), (0.1, 0.1, 5.1)))
    assert sorted(set(v.richtung for v in alle)) == [
        u"seite1", u"seite2", u"seite3", u"seite4"]
    assert all(v.gueltig for v in alle if v.winkel == 45)


# --- Vorrang ----------------------------------------------------------------

def test_gewerk_erkennung():
    g = vr.gewerk
    # Kategorie schlägt alles
    assert g(u"OST_DuctCurves", u"Sanitary", u"ABWASSER") == vr.LUEFTUNG
    assert g(u"OST_CableTray") == vr.ELEKTRO
    assert g(u"OST_StructuralFraming") == vr.BAU
    # Systemname vor Klassifizierung (Screenshot: "REFRIG_Retorno",
    # klassifiziert als Rücklauf Heizung)
    assert g(u"OST_PipeCurves", u"ReturnHydronic",
             u"REFRIG_Retorno 1") == vr.KAELTE
    assert g(u"OST_PipeCurves", u"ReturnHydronic", u"") == vr.HEIZUNG
    assert g(u"OST_PipeCurves", u"DomesticColdWater", u"") == vr.TRINKWASSER
    assert g(u"OST_PipeCurves", u"Sanitary", u"") == vr.ABWASSER
    assert g(u"OST_PipeCurves", u"OtherPipe", u"") == vr.SONSTIGE
    # Kaltwasser ist Trinkwasser, nicht Kälte
    assert g(u"OST_PipeCurves", u"", u"Kaltwasser TWK") == vr.TRINKWASSER
    assert g(u"OST_PipeCurves", u"", u"Kälte Vorlauf") == vr.KAELTE
    # Kürzel nur als eigenes Wort: "ACS" ja, "PLACAS" nein
    assert g(u"OST_PipeCurves", u"", u"ACS Impulsión") == vr.TRINKWASSER
    assert g(u"OST_PipeCurves", u"", u"PLACAS") == vr.SONSTIGE
    # "ventilación" ist kein Abwasser-Lüfter
    assert g(u"OST_PipeCurves", u"", u"Ventilación") == vr.SONSTIGE
    assert g(u"OST_PipeCurves", u"", u"Saneamiento fecales") == vr.ABWASSER


def test_vorrang_reihenfolge():
    liste = vr.reihenfolge(None)
    assert liste == list(vr.STANDARD)
    assert vr.weicht_aus(vr.KAELTE, vr.LUEFTUNG, liste)      # Kälte weicht
    assert not vr.weicht_aus(vr.ABWASSER, vr.HEIZUNG, liste)
    assert not vr.weicht_aus(vr.HEIZUNG, vr.HEIZUNG, liste)
    # gespeicherte Reihenfolge: Unbekanntes fliegt raus, Fehlendes kommt
    # hinten dazu
    eigene = vr.reihenfolge([vr.HEIZUNG, u"gibtsnicht", vr.ABWASSER])
    assert eigene[:2] == [vr.HEIZUNG, vr.ABWASSER]
    assert sorted(eigene) == sorted(vr.STANDARD)
    assert vr.weicht_aus(vr.ABWASSER, vr.HEIZUNG, eigene)


# --- Speicher ---------------------------------------------------------------

def test_speicher_anwenden_und_rueckgaengig():
    with Ordner() as ordner:
        bericht = schreibe(ordner, "b.xml", XML)
        pfad = sp.speicherpfad(bericht)
        assert pfad == bericht + sp.ENDUNG
        speicher = sp.Speicher(pfad)
        assert speicher.anwenden([u"a", u"b"], {u"status": br.GEPRUEFT,
                                                u"klasse": u"CLR"},
                                 u"Schritt 1") == 2
        assert speicher.anwenden([u"a"], {u"status": br.GEPRUEFT,
                                          u"klasse": u"CLR"},
                                 u"ohne Wirkung") == 0
        assert speicher.anwenden([u"a"], {u"klasse": u"", u"kommentar": u"x"},
                                 u"Schritt 2") == 1
        assert speicher.eintraege[u"a"][u"kommentar"] == u"x"
        assert u"klasse" not in speicher.eintraege[u"a"]

        # frisch gelesen = derselbe Stand
        assert sp.Speicher(pfad).eintraege[u"a"][u"kommentar"] == u"x"

        assert speicher.rueckgaengig() == u"Schritt 2"
        assert speicher.eintraege[u"a"][u"klasse"] == u"CLR"
        assert speicher.rueckgaengig() == u"Schritt 1"
        assert speicher.eintraege == {}
        assert speicher.rueckgaengig() is None


def test_speicher_behaelt_fremde_aenderungen():
    """Zwei Benutzer am selben Bericht: keiner überschreibt den anderen."""
    with Ordner() as ordner:
        pfad = os.path.join(ordner, "b.xml" + sp.ENDUNG)
        ich = sp.Speicher(pfad)
        kollege = sp.Speicher(pfad)
        ich.anwenden([u"a"], {u"klasse": u"CLR"}, u"ich")
        kollege.anwenden([u"b"], {u"klasse": u"TGA"}, u"kollege")
        ich.anwenden([u"c"], {u"klasse": u"IGN"}, u"ich")
        stand = sp.Speicher(pfad).eintraege
        assert sorted(stand) == [u"a", u"b", u"c"]


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items())
             if n.startswith("test_") and callable(f)]
    for name, funktion in tests:
        funktion()
        print("  ok:", name)
    print("%d Tests bestanden" % len(tests))
