# -*- coding: utf-8 -*-
"""Offline-Test der Logik von AreaTable (ohne Revit).

    python tools\test_wohnflaeche.py
"""
import os
import sys

os.environ["PYMLG_SPRACHE"] = "de"

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                "..", "pyMLG.extension", "lib"))

from wohnflaeche import logik as lg  # noqa: E402


def nah(a, b):
    return abs(a - b) < 1e-6


def test_raumarten():
    assert lg.art(u"Balkon") == (lg.AUSSEN, lg.BALKON)
    assert lg.art(u"Dachterrasse") == (lg.AUSSEN, lg.TERRASSE)
    assert lg.art(u"Terraza") == (lg.AUSSEN, lg.TERRASSE)
    assert lg.art(u"Balcón")[0] == lg.AUSSEN
    assert lg.art(u"Kellerraum")[0] == lg.NICHT
    assert lg.art(u"Treppe")[0] == lg.NICHT         # in der Maisonette
    assert lg.art(u"Treppenhaus")[0] == lg.ALLGEMEIN
    assert lg.art(u"Aufzug")[0] == lg.ALLGEMEIN
    assert lg.art(u"Tiefgarage")[0] == lg.ALLGEMEIN
    assert lg.art(u"Abstellraum")[0] == lg.INNEN    # in der Wohnung: voll
    assert lg.art(u"Küche")[0] == lg.INNEN


def test_zimmer():
    for name in (u"Wohnen/Essen", u"Schlafzimmer", u"Kinderzimmer 1",
                 u"Arbeitszimmer", u"Wohnküche", u"Living/Kitchen",
                 u"Dormitorio 2", u"Salón-comedor", u"Bedroom"):
        assert lg.ist_zimmer(name), name
    for name in (u"Badezimmer", u"Bad", u"Küche", u"Flur", u"WC",
                 u"Abstellraum", u"Balkon", u"Baño", u"Cocina",
                 u"Ankleidezimmer", u"Kellerzimmer"):
        assert not lg.ist_zimmer(name), name


def test_faktor_name():
    for name in (u"Faktor Anrechenbarkeit", u"Faktor_Anrechenbarkeit",
                 u"Anrechenbarkeitsfaktor", u"anrechenbarkeitsfaktor ",
                 u"Anrechenbarkeits-Faktor", u"Anrechenbarkeit",
                 u"WoFlV_Faktor", u"woflv faktor"):
        assert lg.ist_faktor_name(name), name
    for name in (u"Faktor", u"Fläche", u"Belegung"):
        assert not lg.ist_faktor_name(name), name


def test_handtexte_erkennen():
    for text in (u"Miete kalt: 640€", u"NK: 170€", u"Nebenkosten 170 €",
                 u"Miete warm:", u"Miete kalt:\rNK:\rMiete warm:"):
        assert lg.ist_miete_text(text), text
    for text in (u"Neubau Souterrain", u"Küche", u"Wohnfläche"):
        assert not lg.ist_miete_text(text), text
    for text in (u"2-Zimmer-Wohnung Fakten:", u"Wohnung Fakten:",
                 u"2.OG_Whng1_Altbau · 1-Zimmer-Wohnung"):
        assert lg.ist_werkzeug_kopf(text), text
    assert not lg.ist_werkzeug_kopf(u"Neubau Souterrain")


def test_faktor_rang():
    """"Faktor Anrechenbarkeit" geht vor "Anrechenbarkeit"."""
    assert lg.faktor_rang(u"Faktor Anrechenbarkeit") == 0
    assert lg.faktor_rang(u"Anrechenbarkeit") > lg.faktor_rang(
        u"Faktor Anrechenbarkeit")
    assert lg.faktor_rang(u"Fläche") is None


def test_faktor_parameter():
    assert nah(lg.faktor_aus_parameter(u"1/2"), 0.5)
    assert nah(lg.faktor_aus_parameter(u"1/4"), 0.25)
    assert lg.faktor_aus_parameter(u"1/0") is None
    assert lg.faktor_aus_parameter(None) is None
    assert nah(lg.faktor_aus_parameter(0.5), 0.5)
    assert nah(lg.faktor_aus_parameter(50), 0.5)
    assert nah(lg.faktor_aus_parameter(u"25 %"), 0.25)
    assert nah(lg.faktor_aus_parameter(u"0,5"), 0.5)
    assert lg.faktor_aus_parameter(u"") is None
    assert lg.faktor_aus_parameter(u"x") is None


def beispiel():
    """Wie im Bild des Nutzers: 80,45 m² innen + 8 m² Balkon -> 82,45 m²."""
    return [lg.Raum(u"Wohnen/Essen", 30.0, u"Wohnung 03"),
            lg.Raum(u"Schlafzimmer", 15.0, u"Wohnung 03"),
            lg.Raum(u"Kinderzimmer", 12.0, u"Wohnung 03"),
            lg.Raum(u"Küche", 8.45, u"Wohnung 03"),
            lg.Raum(u"Bad", 7.0, u"Wohnung 03"),
            lg.Raum(u"Flur", 8.0, u"Wohnung 03"),
            lg.Raum(u"Balkon", 8.0, u"Wohnung 03")]


def test_beispiel():
    e = lg.berechne(beispiel())
    assert nah(e.wohnflaeche, 82.45), e.wohnflaeche
    assert nah(e.aussen, 8.0) and nah(e.aussen_angerechnet, 2.0)
    assert e.zimmer == 3
    zeilen = lg.tabelle(e)
    assert zeilen[0] == (lg.TITEL, u"Parameterwerte", u"", True)
    assert zeilen[2] == (lg.KOPF, u"3-Zimmer-Wohnung", u"", True)
    # unten: ganze Linie, dann Platz für die Miete
    assert [z[0] for z in zeilen[-3:]] == [lg.FUSS, lg.VOLL, lg.MIETE]
    assert zeilen[-1][1].split(u"\r") == [u"Miete kalt:", u"NK:",
                                           u"Miete warm:"]
    texte = [(z[1], z[2]) for z in zeilen if z[0] == lg.ZEILE]
    assert texte == [(u"Wohnen/Essen", u"30,00 m²"),
                     (u"Schlafzimmer", u"15,00 m²"),
                     (u"Kinderzimmer", u"12,00 m²"),
                     (u"Küche", u"8,45 m²"),
                     (u"Bad", u"7,00 m²"),
                     (u"Flur", u"8,00 m²"),
                     (u"Balkon (25 % von 8,00 m²)", u"2,00 m²"),
                     (u"Wohnfläche", u"82,45 m²"),
                     (u"Zimmer", u"3")], texte
    assert u"zu 25 % angerechnet" in zeilen[-3][1]
    # von Hand eingetragene Miete wird übernommen, € mit Leerzeichen
    zeilen = lg.tabelle(e, u"Miete kalt: 640€\rNK: 170 €\rMiete warm: 810  €")
    assert zeilen[-1][1] == u"Miete kalt: 640 €\rNK: 170 €\rMiete warm: 810 €"


def test_gewerbe():
    """Büro: Gewerbefläche statt Wohnfläche, keine Zimmerzahl."""
    raeume = [lg.Raum(u"Archiv", 12.43), lg.Raum(u"Büro2", 16.18),
              lg.Raum(u"EDV", 10.83), lg.Raum(u"Küche/Aufenthalt", 17.22),
              lg.Raum(u"Toiletten Steuerbüro", 16.42),
              lg.Raum(u"Balkon", 14.02), lg.Raum(u"Keller", 10.05)]
    e = lg.berechne(raeume, u"EG_Buero")
    assert e.gewerbe
    zeilen = lg.tabelle(e)
    assert zeilen[2] == (lg.KOPF, u"Gewerbefläche", u"", True)
    texte = [(z[1], z[2]) for z in zeilen if z[0] == lg.ZEILE]
    assert (u"Gewerbefläche", u"76,59 m²") == texte[-1], texte[-1]
    assert not any(z[1] == u"Zimmer" for z in zeilen)
    assert zeilen[-3][1].startswith(u"Summe der angerechneten Raumflächen.")
    # auch ohne Hinweis im Namen: Büroräume und kein Wohn-/Schlafraum
    assert lg.berechne(raeume).gewerbe
    # Wohnung mit Arbeitszimmer bleibt Wohnung
    assert not lg.berechne(beispiel() + [lg.Raum(u"Büro", 9.0)]).gewerbe


def test_innen_links():
    """Innenrahmen als geschlossenes Rechteck, links durch eine Querlinie
    unterbrochen - Werte in m (meter=1)."""
    aussen = [(0.0, 0.0), (0.841, 0.0), (0.841, 0.594), (0.0, 0.594), (0.0, 0.0)]
    innen = [(0.02, 0.005), (0.836, 0.005), (0.836, 0.589), (0.02, 0.589),
             (0.02, 0.30), (0.02, 0.005)]
    quer = [(0.0, 0.30), (0.02, 0.30)]
    x = lg.innen_links([aussen, innen, quer], 0.0, 0.0, 0.594)
    assert nah(x, 0.02), x
    # kurze senkrechte Striche (Faltmarken) zählen nicht -> Heftrand 2 cm
    marke = [(0.01, 0.20), (0.01, 0.21)]
    assert nah(lg.innen_links([aussen, marke], 0.0, 0.0, 0.594), 0.02)
    # in Fuss (Revit)
    m = 1.0 / 0.3048
    innen_fuss = [[(p[0] * m, p[1] * m) for p in innen]]
    assert nah(lg.innen_links(innen_fuss, 0.0, 0.0, 0.594 * m, m), 0.02 * m)


def test_zimmer_nach_namen():
    """Unter der Dachschräge geteilte Räume gleichen Namens: ein Zimmer."""
    raeume = [lg.Raum(u"Schlafen", 10.0), lg.Raum(u"Schlafen", 5.25, faktor=0.5),
              lg.Raum(u"Wohnen", 18.0), lg.Raum(u"Wohnen", 5.9, faktor=0.5),
              lg.Raum(u"Bad", 4.0), lg.Raum(u"Flur", 5.39)]
    assert lg.berechne(raeume).zimmer == 2


def test_summe_der_gerundeten():
    """Die Summe in der Tabelle ist die Summe der angezeigten Raumwerte."""
    raeume = [lg.Raum(u"Wohnen", 10.004), lg.Raum(u"Bad", 10.004),
              lg.Raum(u"Flur", 10.004)]
    e = lg.berechne(raeume)
    assert lg.m2(e.wohnflaeche) == u"30,00 m²", e.wohnflaeche


def test_gleiche_namen():
    """Gleicher Name -> eine Zeile, auch bei unterschiedlicher Anrechnung."""
    raeume = [lg.Raum(u"Bad", 1.38), lg.Raum(u"Wohnen", 18.0),
              lg.Raum(u"Bad ", 2.59), lg.Raum(u"Raum", 2.63),
              lg.Raum(u"Raum", 5.15, faktor=0.5),
              lg.Raum(u"Dachstuhl", 16.6, faktor=0.5),
              lg.Raum(u"Dachstuhl", 4.0, faktor=0.5),
              lg.Raum(u"Kinderzimmer", 10.0), lg.Raum(u"Kinderzimmer", 11.0)]
    e = lg.berechne(raeume)
    texte = [(z[1], z[2]) for z in lg.tabelle(e) if z[0] == lg.ZEILE]
    assert texte[:5] == [
        (u"Bad", u"3,97 m²"), (u"Wohnen", u"18,00 m²"),
        (u"Raum (teils angerechnet, Grundfläche 7,78 m²)", u"5,21 m²"),
        (u"Dachstuhl (50 % von 20,60 m²)", u"10,30 m²"),
        (u"Kinderzimmer", u"21,00 m²")], texte
    # Summe passt zu den Zeilen; Zimmer nach Namen (Wohnen, Kinderzimmer)
    assert texte[5] == (u"Wohnfläche", u"58,48 m²"), texte[5]
    assert e.zimmer == 2


def test_reihenfolge():
    raeume = [lg.Raum(u"Treppe", 4.0), lg.Raum(u"Balkon", 6.0),
              lg.Raum(u"Bad", 5.0), lg.Raum(u"Wohnen", 20.0)]
    assert [r.name for r in lg.reihenfolge(raeume)] == [
        u"Bad", u"Wohnen", u"Balkon", u"Treppe"]
    assert lg.raum_beschriftung(raeume[0]) ==         u"Treppe (nicht angerechnet, 4,00 m²)"


def test_dachschraege():
    # 20 m², davon 60 % ab 2 m, 85 % ab 1 m:
    # 0,6*20 + 0,5*(0,25*20) = 12 + 2,5 = 14,5
    raum = lg.Raum(u"Schlafzimmer", 20.0, anteil_1m=0.85, anteil_2m=0.6)
    faktor, abzug = raum.faktor()
    assert nah(faktor * 20.0, 14.5) and nah(abzug, 5.5)
    e = lg.berechne([raum])
    zeilen = lg.tabelle(e)
    assert (lg.ZEILE, u"Schlafzimmer (Dachschräge, Grundfläche 20,00 m²)",
            u"14,50 m²", False) in zeilen, zeilen
    assert u"Dachschrägen" in zeilen[-3][1]


def test_parameter_uebersteuert():
    # Innenraum: Anrechenbarkeitsfaktor gilt, keine Dachschrägen-Messung
    raum = lg.Raum(u"Abstellraum", 6.0, faktor=0.5, anteil_1m=0.0,
                   anteil_2m=0.0)
    e = lg.berechne([raum])
    assert nah(e.wohnflaeche, 3.0)
    assert (lg.ZEILE, u"Abstellraum (50 % von 6,00 m²)", u"3,00 m²", False)         in lg.tabelle(e)
    # Terrasse: Parameter 0,5 wird ignoriert, immer 25 %
    e = lg.berechne([lg.Raum(u"Terrasse", 20.0, faktor=0.5)])
    assert nah(e.wohnflaeche, 5.0)
    assert (lg.ZEILE, u"Terrasse (25 % von 20,00 m²)", u"5,00 m²", False)         in lg.tabelle(e)
    # Faktor 0 im Parameter: nicht angerechnet
    e = lg.berechne([lg.Raum(u"Abstellraum", 6.0, faktor=0.0)])
    assert nah(e.wohnflaeche, 0.0)


def test_eine_wohnung_nach_abteilung():
    """Ohne Belegung: Abteilung, Räume ohne Abteilung bleiben dabei."""
    raeume = beispiel() + [lg.Raum(u"Bad", 6.0, u"Wohnung 04"),
                           lg.Raum(u"Treppenhaus", 15.0)]
    w = lg.eine_wohnung(raeume)
    assert w.name == u"Wohnung 03" and not w.nach_belegung
    assert len(w.raeume) == 8 and [r.abteilung for r in w.andere] == [u"Wohnung 04"]
    # Treppenhaus ohne Abteilung bleibt, erscheint aber nicht in der Tabelle
    e = lg.berechne(w.raeume)
    assert nah(e.wohnflaeche, 82.45)
    assert [r.name for r in e.allgemein] == [u"Treppenhaus"]
    assert not any(u"Treppenhaus" in z[1] for z in lg.tabelle(e))
    # auch mit Anrechenbarkeitsfaktor nicht dabei
    e = lg.berechne([lg.Raum(u"Treppenhaus", 10.0, faktor=1.0)])
    assert nah(e.wohnflaeche, 0.0) and not e.raeume


def test_eine_wohnung_nach_belegung():
    """Belegung = Wohnungsnummer; ohne Belegung zählt nicht."""
    raeume = [lg.Raum(r.name, r.flaeche, belegung=u"WE 03") for r in beispiel()]
    raeume += [lg.Raum(u"Wohnen", 25.0, belegung=u"WE 04"),
               lg.Raum(u"Abstellraum", 3.0),           # Nachbarfläche
               lg.Raum(u"Treppenhaus", 15.0)]
    w = lg.eine_wohnung(raeume)
    assert w.name == u"WE 03" and w.nach_belegung
    assert len(w.raeume) == 7
    assert [r.name for r in w.andere] == [u"Wohnen"]
    assert [r.name for r in w.ohne] == [u"Abstellraum", u"Treppenhaus"]
    assert nah(lg.berechne(w.raeume).wohnflaeche, 82.45)


def test_ohne_balkon():
    e = lg.berechne([lg.Raum(u"Wohnen", 40.0)])
    zeilen = lg.tabelle(e)
    assert zeilen[2][1] == u"1-Zimmer-Wohnung"
    assert not any(u"Balkon" in z[1] for z in zeilen)


if __name__ == "__main__":
    for name, funktion in sorted(globals().items()):
        if name.startswith("test_") and callable(funktion):
            funktion()
            print("ok", name)
