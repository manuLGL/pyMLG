# -*- coding: utf-8 -*-
"""Offline-Test von DxfLegend: DXF-Leser und Legendenplan (ohne Revit).

    python tools\test_dxf_legende.py
"""
import math
import os
import sys

os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from dxf_legende import dxf, geometrie as geo, logik as lg  # noqa: E402


def _dxf(kopf=(), tabellen=(), bloecke=(), objekte=()):
    """Baut ein ASCII-DXF aus Listen von (code, wert)."""
    zeilen = []

    def abschnitt(name, paare):
        zeilen.extend([u"0", u"SECTION", u"2", name])
        for code, wert in paare:
            zeilen.extend([u"%d" % code, u"%s" % wert])
        zeilen.extend([u"0", u"ENDSEC"])

    abschnitt(u"HEADER", kopf)
    abschnitt(u"TABLES", tabellen)
    abschnitt(u"BLOCKS", bloecke)
    abschnitt(u"ENTITIES", objekte)
    zeilen.extend([u"0", u"EOF"])
    return u"\n".join(zeilen) + u"\n"


TABELLEN = [
    (0, u"TABLE"), (2, u"LTYPE"),
    (0, u"LTYPE"), (2, u"DASHED"), (72, 65), (73, 2), (40, 0.75),
    (49, 0.5), (74, 0), (49, -0.25), (74, 0),
    (0, u"ENDTAB"),
    (0, u"TABLE"), (2, u"LAYER"),
    (0, u"LAYER"), (2, u"0"), (70, 0), (62, 7), (6, u"CONTINUOUS"),
    (0, u"LAYER"), (2, u"FECAL"), (70, 0), (62, 30), (6, u"DASHED"), (370, 50),
    (0, u"LAYER"), (2, u"AUS"), (70, 0), (62, -1), (6, u"CONTINUOUS"),
    (0, u"ENDTAB"),
    (0, u"TABLE"), (2, u"STYLE"),
    (0, u"STYLE"), (2, u"ARIAL"), (70, 0), (40, 0.0), (41, 0.8), (3, u"arial.ttf"),
    (0, u"ENDTAB"),
]

KOPF = [(9, u"$ACADVER"), (1, u"AC1027"), (9, u"$INSUNITS"), (70, 6),
        (9, u"$LTSCALE"), (40, 2.0)]


# Nachbau der Saneamiento-Legende (auch für tools-fremde Tests nutzbar)
OBJEKTE = [
    # Tabellenrahmen weiss
    (0, u"LINE"), (8, u"0"), (10, 0), (20, 0), (11, 20), (21, 0),
    (0, u"LINE"), (8, u"0"), (10, 0), (20, 0), (11, 0), (21, 5),
    # Fecal: dicke Polylinie, Farbe vom Layer, gestrichelt vom Layer
    (0, u"LWPOLYLINE"), (8, u"FECAL"), (90, 2), (70, 0), (43, 0.05),
    (10, 1), (20, 4), (10, 4), (20, 4),
    # Pluvial: TrueColor blau
    (0, u"LWPOLYLINE"), (8, u"FECAL"), (420, 0x0080FF), (90, 2), (70, 0),
    (10, 1), (20, 3), (10, 4), (20, 3),
    # unsichtbarer Layer
    (0, u"LINE"), (8, u"AUS"), (10, 0), (20, 0), (11, 1), (21, 1),
    # Text links auf Grundlinie
    (0, u"TEXT"), (8, u"0"), (10, 6), (20, 3.9), (40, 0.25), (7, u"ARIAL"),
    (1, u"RED SANEAMIENTO FECAL"),
    # Titel zentriert
    (0, u"TEXT"), (8, u"0"), (10, 0), (20, 0), (11, 10), (21, 4.6), (40, 0.4),
    (72, 1), (73, 2), (1, u"LEYENDA"),
    # MTEXT mit Formatierung und Durchmesser
    (0, u"MTEXT"), (8, u"0"), (10, 6), (20, 2.2), (40, 0.25), (71, 1),
    (1, u"{\\fArial|b0|i0|c0|p34;BAJANTE %%c 110 MM}"),
    # Sumidero: Kreis
    (0, u"CIRCLE"), (8, u"FECAL"), (62, 1), (10, 2), (20, 1), (40, 0.3),
    # Arqueta: Volltonschraffur mit Polylinienrand
    (0, u"HATCH"), (8, u"0"), (62, 30), (2, u"SOLID"), (70, 1), (71, 0),
    (91, 1), (92, 2), (72, 0), (73, 1), (93, 4),
    (10, 3), (20, 0.5), (10, 4), (20, 0.5), (10, 4), (20, 1.5), (10, 3), (20, 1.5),
    (97, 0), (75, 0), (76, 1), (98, 0),
    # Block mit VONBLOCK-Farbe, gedreht und skaliert
    (0, u"INSERT"), (8, u"0"), (62, 5), (2, u"SYMBOL"), (10, 10), (20, 1),
    (41, 2), (42, 2), (50, 90),
]
BLOECKE = [
    (0, u"BLOCK"), (8, u"0"), (2, u"SYMBOL"), (70, 0), (10, 0), (20, 0),
    (0, u"LINE"), (8, u"0"), (62, 0), (10, 0), (20, 0), (11, 1), (21, 0),
    (0, u"ARC"), (8, u"0"), (62, 0), (10, 0), (20, 0), (40, 1), (50, 0), (51, 90),
    (0, u"ENDBLK"),
]


def legende_dxf():
    return _dxf(KOPF, TABELLEN, BLOECKE, OBJEKTE)


def _legende():
    return dxf.lese_text(legende_dxf())


def _nah(a, b, tol=1e-6):
    return all(abs(x - y) < tol for x, y in zip(a, b))


def test_aci_farben():
    assert dxf.aci_rgb(1) == (255, 0, 0)
    assert dxf.aci_rgb(10) == (255, 0, 0)
    assert dxf.aci_rgb(11) == (255, 127, 127)
    assert dxf.aci_rgb(12) == (204, 0, 0)
    assert dxf.aci_rgb(22) == (204, 51, 0)
    assert dxf.aci_rgb(30) == (255, 127, 0)
    assert dxf.aci_rgb(31) == (255, 191, 127)
    assert dxf.aci_rgb(150) == (0, 127, 255)
    assert dxf.aci_rgb(19) == (76, 38, 38)
    assert dxf.aci_rgb(254) == (190, 190, 190)


def test_papierfarbe():
    assert lg.papierfarbe((255, 255, 255)) == (0, 0, 0)
    assert lg.papierfarbe((240, 240, 240)) == (0, 0, 0)
    assert lg.papierfarbe((255, 255, 0)) == (255, 255, 0)
    dunkler = lg.papierfarbe((255, 255, 0), abdunkeln=True)
    assert lg.helligkeit(dunkler) <= lg.HELL_GRENZE + 0.01
    assert lg.papierfarbe((0, 0, 255), abdunkeln=True) == (0, 0, 255)


def test_mtext_bereinigen():
    text, fmt = dxf.bereinige_mtext(u"{\\fArial|b1;\\H2.5;Zeile 1\\PZeile\\~2 %%c}")
    assert text == u"Zeile 1\nZeile 2 \u00d8", repr(text)
    assert fmt[u"schrift"] == u"Arial" and fmt[u"hoehe"] == 2.5
    text, fmt = dxf.bereinige_mtext(u"\\C1;rot \\S1^2; \\\\pfad \\U+00E9")
    assert text == u"rot 1/2 \\pfad \u00e9", repr(text)
    assert fmt[u"aci"] == 1


def test_schriften():
    assert dxf.schrift_aus_datei(u"arial.ttf") == u"Arial"
    assert dxf.schrift_aus_datei(u"C:\\Fonts\\ARIALN.TTF") == u"Arial Narrow"
    assert dxf.schrift_aus_datei(u"romans.shx") is None
    assert dxf.schrift_aus_datei(u"") is None


def test_bulge_und_bogen():
    (cx, cy), r, start, oeffnung = geo.bogen((0, 0), (2, 0), 1.0)
    assert _nah((cx, cy, r, oeffnung), (1, 0, 1, math.pi))
    assert _nah(geo.bogenmitte((0, 0), (2, 0), 1.0), (1, -1))
    zug = geo.abtasten([(0, 0, 1.0), (2, 0, 1.0)], True)
    assert _nah(zug[0], zug[-1])
    assert all(abs(math.hypot(x - 1, y) - 1) < 1e-9 for x, y in zug)


def test_transformation_spiegeln():
    m = geo.skalierung(-1, 1)
    punkte = geo.transformiere([(0, 0, 0.5), (1, 0, 0)], False, m)
    assert punkte[0][2] == -0.5 and punkte[1][0] == -1
    # Nicht gleichmässig skaliert: Bogen wird abgetastet
    punkte = geo.transformiere([(0, 0, 1.0), (2, 0, 0)], False, geo.skalierung(1, 2))
    assert len(punkte) > 3 and all(p[2] == 0 for p in punkte)


def test_legende_lesen():
    z = _legende()
    assert z.einheiten == 6 and z.ltscale == 2.0
    assert z.linientypen[u"DASHED"] == [0.5, -0.25]
    # 2 Rahmen + 2 Polylinien + Kreis + Block (Linie + Bogen); Layer AUS fehlt
    assert len(z.kurven) == 7, len(z.kurven)
    fecal = z.kurven[2]
    assert fecal.farbe == dxf.aci_rgb(30) and fecal.linientyp == u"DASHED"
    assert fecal.staerke_mm == 0.5 and abs(fecal.breite - 0.05) < 1e-12
    assert z.kurven[3].farbe == (0, 128, 255)
    assert z.kurven[4].farbe == (255, 0, 0) and z.kurven[4].geschlossen
    # Block: VONBLOCK blau, um 90° gedreht und doppelt gross
    linie, bogen = z.kurven[5], z.kurven[6]
    assert linie.farbe == (0, 0, 255)
    assert _nah(linie.punkte[0][:2], (10, 1)) and _nah(linie.punkte[1][:2], (10, 3))
    assert _nah(bogen.punkte[0][:2], (10, 3)) and _nah(bogen.punkte[1][:2], (8, 1))
    assert bogen.punkte[0][2] > 0


def test_texte():
    z = _legende()
    assert len(z.texte) == 3
    fecal, titel, bajante = z.texte
    assert fecal.schrift == u"Arial" and abs(fecal.breitenfaktor - 0.8) < 1e-12
    # Grundlinie 3.9 -> Mitte der Versalhöhe 3.9 + 0.125
    assert _nah((fecal.x, fecal.y), (6, 4.025)) and fecal.h_ausr == u"links"
    assert titel.h_ausr == u"mitte" and _nah((titel.x, titel.y), (10, 4.6))
    assert bajante.inhalt == u"BAJANTE \u00d8 110 MM"
    # MTEXT oben links: Mitte der Zeile liegt eine halbe Höhe tiefer
    assert _nah((bajante.x, bajante.y), (6, 2.075)) and bajante.v_ausr == u"mitte"


def test_flaeche():
    z = _legende()
    assert len(z.flaechen) == 1
    f = z.flaechen[0]
    assert f.muster is None and f.farbe == dxf.aci_rgb(30)
    assert len(f.schleifen[0]) == 4


def test_muster_normalisieren():
    assert lg.linienmuster_mm([], 1.0) == ()
    assert lg.linienmuster_mm([1.0], 1.0) == ()
    assert lg.linienmuster_mm([0.5, -0.25], 10.0) == ((u"dash", 5.0), (u"space", 2.5))
    # Beginnt mit Lücke -> gedreht
    segmente, vorlauf = lg._normalisiere([-1.0, 4.0, -2.0])
    assert segmente == [(u"dash", 4.0), (u"space", 3.0)] and vorlauf == 1.0
    # Punkt und zu kurze Lücke
    assert lg.linienmuster_mm([0.0, -0.01, 2.0, -1.0], 1.0) == (
        (u"dot", 0.0), (u"space", 0.25), (u"dash", 2.0), (u"space", 1.0))


def test_fuellgitter():
    # 45°-Linien, Abstand senkrecht 1 Einheit, Faktor 10
    w = math.radians(45)
    muster = [(w, (0.0, 0.0), (-math.sin(w), math.cos(w)), [])]
    gitter = lg.fuellgitter_mm(muster, 10.0, (0.0, 0.0))
    assert len(gitter) == 1
    assert abs(gitter[0][3] - 10.0) < 1e-9 and abs(gitter[0][4]) < 1e-9
    assert gitter[0][5] == ()


def test_stift():
    assert lg.stift(0.1) == 1
    assert lg.stift(0.25) == 2
    assert lg.stift(0.5) == 4
    assert lg.stift(0.6) in (4, 5)
    assert lg.stift(50) == 16


def test_plan():
    z = _legende()
    faktor, ziel = lg.vorschlag(z)
    assert ziel == 2.5 and abs(faktor - 10.0) < 1e-9
    assert lg.vorschlag_name(z, u"datei") == u"LEYENDA"
    plan = lg.erstelle_plan(z, faktor)
    assert len(plan.linien) == 7 and len(plan.texte) == 3 and len(plan.flaechen) == 1
    # Fecal: Breite 0.05 * 10 = 0.5 mm -> Stift 4, Muster mit LTSCALE 2
    name = plan.linien[2][0]
    stil = plan.stile[name]
    assert stil.stift == 4 and stil.muster == ((u"dash", 10.0), (u"space", 5.0)), stil.muster
    assert name == u"LEY_FF7F00_DASHED_4", name
    # Weisser Rahmen wird schwarz
    assert plan.stile[plan.linien[0][0]].farbe == (0, 0, 0)
    # Texttypen: Arial 2.5 mm mit Breitenfaktor
    namen = sorted(plan.texttypen)
    assert u"LEY_2.5mm_Arial_000000_x0.8" in namen, namen
    assert u"LEY_4mm_Arial_000000" in namen, namen
    # Ursprung unten links: kein Punkt negativ
    for _s, punkte, _g in plan.linien:
        assert all(p[0] >= -1e-9 and p[1] >= -1e-9 for p in punkte)
    assert plan.breite_mm > 190


def test_namen_eindeutig():
    namen = lg._Namen()
    a = namen.name(u"LEY_X", (1,))
    b = namen.name(u"LEY_X", (2,))
    assert (a, b) == (u"LEY_X", u"LEY_X_2")
    assert namen.name(u"LEY_X", (1,)) == u"LEY_X"
    assert lg.revit_name(u"a:b{c}") == u"a_b_c_"


def test_binaer_abgelehnt():
    try:
        dxf.lese_bytes(b"AutoCAD Binary DXF\r\n\x1a\x00")
    except dxf.DxfFehler:
        return
    raise AssertionError("binär nicht erkannt")


def test_schraffur_kanten_und_muster():
    objekte = [
        (0, u"HATCH"), (8, u"0"), (62, 3), (2, u"ANSI31"), (70, 0), (71, 0),
        (91, 1), (92, 1), (93, 2),
        (72, 1), (10, 0), (20, 0), (11, 2), (21, 0),
        (72, 2), (10, 1), (20, 0), (40, 1), (50, 0), (51, 180), (73, 1),
        (97, 0), (75, 0), (76, 1), (52, 0), (41, 1), (77, 0), (78, 1),
        (53, 45), (43, 0), (44, 0), (45, -0.0883883), (46, 0.0883883), (79, 0),
        (98, 0),
    ]
    z = dxf.lese_text(_dxf(objekte=objekte))
    f = z.flaechen[0]
    assert f.muster_name == u"ANSI31" and len(f.muster) == 1
    schleife = f.schleifen[0]
    # Linie (0,0)-(2,0), dann Halbkreis zurück
    assert _nah(schleife[0][:2], (0, 0)) and _nah(schleife[1][:2], (2, 0))
    assert abs(schleife[1][2] - 1.0) < 1e-9
    gitter = lg.fuellgitter_mm(f.muster, 10.0, (0, 0))
    assert abs(gitter[0][3] - 1.25) < 1e-3, gitter


if __name__ == "__main__":
    for name, funktion in sorted(globals().items()):
        if name.startswith("test_") and callable(funktion):
            funktion()
            print("ok ", name)
