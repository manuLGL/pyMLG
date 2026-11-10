# -*- coding: utf-8 -*-
"""Revit-freie Logik von ApartmentAreas.

Die Raumkategorie (Schlüssel, z. B. "2_W_SANITÄR") beginnt mit Nummer und
Kürzel: "2" und "W". Danach entscheidet sich, in welchen Parameter des
Eingangspfeils die Raumfläche zählt. Verglichen werden nur Nummer und
Kürzel - so stören kaputte Umlaute im Rest ("3_GW_ERSCHLIE?UNG") nicht.

Kategorien, die hier fehlen (z. B. 2_WA5.2_EIGENGARTEN, 2_WE_EINLAGERUNG,
2_WA_WINTERGARTEN, 4_GW_VERK_...), zählen nirgends.
"""

# Parameter des Eingangspfeils (Namen wie in der Familie)
WNFL = u"TOP_01_WNFL_FLÄ"
TERRASSE = u"TOP_02_TERRASSE_FLÄ"
LOGGIA = u"TOP_03_LOGGIA_FLÄ"
BALKON = u"TOP_04_BALKON_FLÄ"
TERRASSE_EG = u"TOP_07_TERRASSE_EG_FLÄ"
TERRASSE_DT = u"TOP_08_TERRASSE_DT_FLÄ"
GW_NFL = u"TOP_GW_01_NFL_FLÄ"

# In dieser Reihenfolge werden sie geschrieben und gemeldet
ALLE = (WNFL, TERRASSE, LOGGIA, BALKON, TERRASSE_EG, TERRASSE_DT, GW_NFL)

# (Nummer, Kürzel) der Raumkategorie -> Parameter
REGELN = {
    (u"1", u"W"): WNFL,           # 1_W_WOHNEN
    (u"2", u"W"): WNFL,           # 2_W_NEBENRAUM, 2_W_SAN_..., 2_W_SANITÄR,
                                  # 2_W_VERKEHRSFLÄCHE
    (u"2", u"WA1"): LOGGIA,       # 2_WA1_LOGGIA
    (u"2", u"WA2"): BALKON,       # 2_WA2_BALKON
    (u"2", u"WA3"): TERRASSE,     # 2_WA3_TERRASSE
    (u"2", u"WA4"): TERRASSE_EG,  # 2_WA4_TERRASSE (EG)
    (u"2", u"WA5.1"): TERRASSE_DT,  # 2_WA5.1_TERRASSE (DACH)
    (u"3", u"GW"): GW_NFL,        # 3_GW_BESPRECHUNG ... 3_GW_SANITÄR
}

# Interne Revit-Fläche (Quadratfuss) -> m²
QM_JE_QF = 0.09290304


def parameter_fuer(kategorie):
    """Parameter, in den eine Raumkategorie zählt, sonst None.
    '2_W_SANITÄR' -> WNFL, '4_GW_VERK_BÜRO' -> None."""
    teile = (kategorie or u"").strip().upper().split(u"_")
    if len(teile) < 3:
        return None
    return REGELN.get((teile[0].strip(), teile[1].strip()))


def summen(raeume):
    """raeume: [(Raumkategorie, Fläche)] -> ({Parameter: Summe} mit allen
    Parametern, auch 0, {Kategorie: Anzahl} der nicht gezählten)."""
    ergebnis = dict((name, 0.0) for name in ALLE)
    ungezaehlt = {}
    for kategorie, flaeche in raeume:
        name = parameter_fuer(kategorie)
        if name is None:
            schluessel = (kategorie or u"").strip() or u"-"
            ungezaehlt[schluessel] = ungezaehlt.get(schluessel, 0) + 1
        else:
            ergebnis[name] += flaeche or 0.0
    return ergebnis, ungezaehlt


def qm(flaeche_qf):
    return flaeche_qf * QM_JE_QF


def qm_text(flaeche_qf):
    return (u"%.2f m²" % qm(flaeche_qf)).replace(u".", u",")
