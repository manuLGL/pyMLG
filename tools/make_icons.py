# -*- coding: utf-8 -*-
"""Erzeugt einheitliche icon.png (96x96 RGBA) fuer alle pyMLG-Werkzeuge.

Keine externen Abhaengigkeiten: gezeichnet wird mit 4x-Supersampling in einen
Pixelpuffer, PNG-Ausgabe ueber zlib. Bildsprache wie beim Excel-Export:
dunkelblaue Kontur, weisses "Papier", eine Akzentfarbe je Funktion.
Wenige grosse Formen, damit die Symbole auch bei 16-32 px lesbar bleiben.
"""
import math
import os
import struct
import sys
import zlib

S = 4
W = H = 96
BW, BH = W * S, H * S

# --- Palette ---------------------------------------------------------------
DUNKEL = (68, 84, 106)      # Kontur/Struktur
WEISS = (255, 255, 255)     # Papierflaeche
LINIE = (176, 186, 198)     # Hilfslinien
BLAU = (6, 150, 215)        # Ansichten/Plaene
GRUEN = (33, 115, 70)       # hinzufuegen / an
ORANGE = (214, 124, 38)     # aendern / umkehren
ROT = (192, 58, 58)         # aus / entfernen
VIOLETT = (126, 87, 194)    # Phasen
GELB = (232, 178, 46)       # Hervorhebung


# --- Zeichenprimitive ------------------------------------------------------
def leer():
    return [[(0, 0, 0, 0)] * BW for _ in range(BH)]


def _setze(buf, x, y, farbe):
    if 0 <= x < BW and 0 <= y < BH:
        buf[y][x] = (farbe[0], farbe[1], farbe[2], 255)


def rect(buf, x0, y0, x1, y1, farbe):
    for y in range(max(0, int(y0 * S)), min(BH, int(y1 * S))):
        for x in range(max(0, int(x0 * S)), min(BW, int(x1 * S))):
            _setze(buf, x, y, farbe)


def rahmen(buf, x0, y0, x1, y1, dicke, farbe):
    rect(buf, x0, y0, x1, y0 + dicke, farbe)
    rect(buf, x0, y1 - dicke, x1, y1, farbe)
    rect(buf, x0, y0, x0 + dicke, y1, farbe)
    rect(buf, x1 - dicke, y0, x1, y1, farbe)


def _dreieck(buf, p0, p1, p2, farbe):
    pts = [(p[0] * S, p[1] * S) for p in (p0, p1, p2)]
    (x0, y0), (x1, y1), (x2, y2) = pts
    xs = [x0, x1, x2]
    ys = [y0, y1, y2]

    def kante(ax, ay, bx, by, px, py):
        return (bx - ax) * (py - ay) - (by - ay) * (px - ax)

    orient = kante(x0, y0, x1, y1, x2, y2)
    if abs(orient) < 1e-9:
        return
    vz = 1 if orient > 0 else -1
    for y in range(max(0, int(min(ys))), min(BH, int(max(ys)) + 1)):
        for x in range(max(0, int(min(xs))), min(BW, int(max(xs)) + 1)):
            px, py = x + 0.5, y + 0.5
            if (kante(x0, y0, x1, y1, px, py) * vz >= 0
                    and kante(x1, y1, x2, y2, px, py) * vz >= 0
                    and kante(x2, y2, x0, y0, px, py) * vz >= 0):
                _setze(buf, x, y, farbe)


def polygon(buf, punkte, farbe):
    """Gefuelltes konvexes Polygon (Faecher-Triangulierung)."""
    for i in range(1, len(punkte) - 1):
        _dreieck(buf, punkte[0], punkte[i], punkte[i + 1], farbe)


def ellipse(buf, cx, cy, rx, ry, farbe):
    for y in range(max(0, int((cy - ry) * S)), min(BH, int((cy + ry) * S) + 1)):
        for x in range(max(0, int((cx - rx) * S)), min(BW, int((cx + rx) * S) + 1)):
            dx = (x + 0.5) / S - cx
            dy = (y + 0.5) / S - cy
            if (dx / rx) ** 2 + (dy / ry) ** 2 <= 1.0:
                _setze(buf, x, y, farbe)


def kreis(buf, cx, cy, r, farbe):
    ellipse(buf, cx, cy, r, r, farbe)


def linie(buf, p0, p1, dicke, farbe):
    (x0, y0), (x1, y1) = p0, p1
    dx, dy = x1 - x0, y1 - y0
    laenge = math.hypot(dx, dy)
    if laenge < 1e-9:
        return
    nx, ny = -dy / laenge * dicke / 2.0, dx / laenge * dicke / 2.0
    polygon(buf, [(x0 + nx, y0 + ny), (x1 + nx, y1 + ny),
                  (x1 - nx, y1 - ny), (x0 - nx, y0 - ny)], farbe)


def pfeil(buf, p0, p1, dicke, kopf, farbe):
    """Linie mit Dreieckspitze an p1."""
    (x0, y0), (x1, y1) = p0, p1
    dx, dy = x1 - x0, y1 - y0
    laenge = math.hypot(dx, dy)
    if laenge < 1e-9:
        return
    ux, uy = dx / laenge, dy / laenge
    basis = (x1 - ux * kopf, y1 - uy * kopf)
    linie(buf, p0, basis, dicke, farbe)
    px, py = -uy, ux
    polygon(buf, [(x1, y1),
                  (basis[0] + px * kopf * 0.62, basis[1] + py * kopf * 0.62),
                  (basis[0] - px * kopf * 0.62, basis[1] - py * kopf * 0.62)],
            farbe)


def blattform(buf, x0, y0, x1, y1, akzent=None, kopfhoehe=0):
    """Weisses Blatt mit dunkler Kontur, optional farbiger Kopfleiste."""
    rect(buf, x0, y0, x1, y1, DUNKEL)
    rect(buf, x0 + 2, y0 + 2, x1 - 2, y1 - 2, WEISS)
    if akzent and kopfhoehe:
        rect(buf, x0 + 2, y0 + 2, x1 - 2, y0 + 2 + kopfhoehe, akzent)


def verkleinern(buf):
    ergebnis = []
    flaeche = float(S * S)
    for y in range(H):
        zeile = []
        for x in range(W):
            sr = sg = sb = sa = 0
            for dy in range(S):
                quelle = buf[y * S + dy]
                for dx in range(S):
                    r, g, b, a = quelle[x * S + dx]
                    sr += r * a
                    sg += g * a
                    sb += b * a
                    sa += a
            if sa == 0:
                zeile.append((0, 0, 0, 0))
            else:
                zeile.append((int(sr / sa), int(sg / sa), int(sb / sa),
                              int(sa / flaeche)))
        ergebnis.append(zeile)
    return ergebnis


def schreibe_png(pfad, pixel):
    roh = b"".join(b"\x00" + bytes(k for px in zeile for k in px)
                   for zeile in pixel)

    def chunk(typ, daten):
        return (struct.pack(">I", len(daten)) + typ + daten
                + struct.pack(">I", zlib.crc32(typ + daten) & 0xffffffff))

    kopf = struct.pack(">IIBBBBB", W, H, 8, 6, 0, 0, 0)
    with open(pfad, "wb") as datei:
        datei.write(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", kopf)
                    + chunk(b"IDAT", zlib.compress(roh, 9)) + chunk(b"IEND", b""))


# --- Die einzelnen Symbole -------------------------------------------------
def tab_manager(b):
    """Drei Ribbon-Reiter, der aktive hervorgehoben."""
    rect(b, 8, 28, 88, 84, DUNKEL)
    rect(b, 10, 30, 86, 82, WEISS)
    for x0, x1, aktiv in ((8, 32, False), (34, 62, True), (64, 88, False)):
        oben = 12 if aktiv else 18
        rect(b, x0, oben, x1, 32, DUNKEL)
        rect(b, x0 + 2, oben + 2, x1 - 2, 32, BLAU if aktiv else LINIE)
    rect(b, 18, 44, 76, 48, LINIE)
    rect(b, 18, 56, 60, 60, LINIE)


def duplicate_plan(b):
    """Zwei Plaene mit Schriftfeld, plus Zeichen."""
    blattform(b, 10, 8, 64, 62)
    rect(b, 40, 46, 60, 58, BLAU)          # Schriftfeld hinten
    blattform(b, 30, 30, 86, 88)
    rect(b, 60, 70, 82, 84, BLAU)          # Schriftfeld vorne
    rect(b, 36, 40, 56, 44, LINIE)
    rect(b, 36, 50, 50, 54, LINIE)
    kreis(b, 20, 74, 15, WEISS)             # Pluszeichen freistellen
    kreis(b, 20, 74, 13, GRUEN)
    linie(b, (13, 74), (27, 74), 5, WEISS)
    linie(b, (20, 67), (20, 81), 5, WEISS)


def duplicate_view(b):
    """Zwei Ansichtsrahmen mit Blickrichtung, plus Zeichen."""
    rahmen(b, 10, 8, 64, 58, 3, LINIE)
    blattform(b, 30, 28, 88, 84)
    polygon(b, [(59, 42), (75, 34), (75, 58)], BLAU)   # Blickrichtung
    kreis(b, 55, 50, 7, DUNKEL)
    kreis(b, 19, 74, 15, WEISS)
    kreis(b, 19, 74, 13, GRUEN)
    linie(b, (12, 74), (26, 74), 5, WEISS)
    linie(b, (19, 67), (19, 81), 5, WEISS)


def pass_filter_overrides(b):
    """Farbige Filterstreifen wandern von links nach rechts."""
    blattform(b, 6, 16, 38, 80)
    for i, farbe in enumerate((BLAU, ORANGE, GRUEN)):
        rect(b, 10, 26 + i * 14, 34, 36 + i * 14, farbe)
    blattform(b, 58, 16, 90, 80)
    for i, farbe in enumerate((BLAU, ORANGE, GRUEN)):
        rect(b, 62, 26 + i * 14, 86, 36 + i * 14, farbe)
    pfeil(b, (38, 48), (60, 48), 9, 13, DUNKEL)


def tag_distance(b):
    """Wand, Beschriftung und Massangabe dazwischen."""
    rect(b, 8, 10, 24, 86, DUNKEL)          # Wand (bewusst ohne Schraffur:
                                            # feine Linien verschwimmen klein)
    blattform(b, 56, 22, 90, 46, GELB, 6)   # Beschriftung
    rect(b, 62, 36, 84, 40, LINIE)
    linie(b, (24, 58), (56, 58), 5, ORANGE)  # Massline
    linie(b, (26, 50), (26, 66), 5, ORANGE)
    linie(b, (54, 50), (54, 66), 5, ORANGE)
    linie(b, (56, 34), (56, 58), 3, LINIE)


def view_id_visible(b):
    """Ansicht mit Auge: die Id wird sichtbar gemacht."""
    blattform(b, 8, 16, 88, 80, BLAU, 10)
    ellipse(b, 48, 54, 26, 16, DUNKEL)
    ellipse(b, 48, 54, 21, 12, WEISS)
    kreis(b, 48, 54, 9, BLAU)
    kreis(b, 48, 54, 4, DUNKEL)


def view_name_manager(b):
    """Textzeilen und Stift: Ansichtsnamen bearbeiten."""
    blattform(b, 8, 12, 74, 84, BLAU, 10)
    rect(b, 16, 34, 64, 39, LINIE)
    rect(b, 16, 47, 54, 52, LINIE)
    rect(b, 16, 60, 44, 65, LINIE)
    # Stift: Schaft, Spitze, Radiergummi
    polygon(b, [(58, 86), (52, 80), (78, 54), (84, 60)], ORANGE)
    polygon(b, [(52, 80), (50, 88), (58, 86)], DUNKEL)
    polygon(b, [(78, 54), (84, 48), (90, 54), (84, 60)], LINIE)


def view_to_sheet(b):
    """Ansicht wird auf einem Plan platziert."""
    blattform(b, 12, 10, 88, 86)
    rect(b, 62, 68, 84, 82, BLAU)           # Schriftfeld
    rahmen(b, 20, 20, 58, 58, 3, BLAU)      # platzierte Ansicht
    rect(b, 24, 24, 54, 54, (222, 238, 248))
    pfeil(b, (34, 74), (34, 62), 6, 10, GRUEN)


def wall_legend(b):
    """Drei Wandquerschnitte nebeneinander."""
    rect(b, 10, 12, 26, 84, DUNKEL)
    rect(b, 38, 12, 58, 84, DUNKEL)
    rect(b, 42, 16, 54, 80, WEISS)
    for y in range(18, 80, 9):
        linie(b, (42, y + 7), (54, y), 2, LINIE)
    rect(b, 70, 12, 86, 84, BLAU)


def copy_with_phases(b):
    """Kopie plus Phasenkreis."""
    rahmen(b, 10, 10, 58, 58, 3, LINIE)
    blattform(b, 26, 28, 76, 78)
    rect(b, 32, 38, 64, 42, LINIE)
    rect(b, 32, 50, 56, 54, LINIE)
    kreis(b, 74, 70, 18, VIOLETT)
    kreis(b, 74, 70, 12, WEISS)
    polygon(b, [(74, 70), (74, 58), (86, 70)], VIOLETT)


def copy_paste_with_phases(b):
    """Zwischenablage mit Einfuegepfeil und Phasenkreis."""
    blattform(b, 14, 14, 70, 84)
    rect(b, 30, 8, 54, 20, DUNKEL)          # Klemme
    rect(b, 33, 11, 51, 20, LINIE)
    pfeil(b, (42, 34), (42, 62), 9, 14, GRUEN)
    kreis(b, 74, 70, 18, VIOLETT)
    kreis(b, 74, 70, 12, WEISS)
    polygon(b, [(74, 70), (74, 58), (86, 70)], VIOLETT)


def _workset_ebenen(b, farben):
    """Drei gestapelte Bearbeitungssets als Parallelogramme."""
    for i, farbe in enumerate(farben):
        y = 24 + i * 22
        punkte = [(48, y - 12), (86, y), (48, y + 12), (10, y)]
        polygon(b, punkte, DUNKEL)
        punkte_innen = [(48, y - 8), (79, y), (48, y + 8), (17, y)]
        polygon(b, punkte_innen, farbe)


def workset_on(b):
    _workset_ebenen(b, (GRUEN, GRUEN, GRUEN))


def workset_off(b):
    _workset_ebenen(b, (LINIE, LINIE, LINIE))
    linie(b, (14, 84), (84, 14), 11, WEISS)
    linie(b, (14, 84), (84, 14), 7, ROT)


def workset_reverse(b):
    _workset_ebenen(b, (ORANGE, LINIE, ORANGE))
    pfeil(b, (84, 24), (84, 64), 9, 15, DUNKEL)
    pfeil(b, (12, 64), (12, 24), 9, 15, DUNKEL)


def workset_creator(b):
    """Workset-Ebenen mit gruenem Plus: Worksets erstellen."""
    _workset_ebenen(b, (BLAU, BLAU, BLAU))
    kreis(b, 74, 74, 20, WEISS)
    kreis(b, 74, 74, 16, GRUEN)
    rect(b, 71, 63, 77, 85, WEISS)
    rect(b, 63, 71, 85, 77, WEISS)


def join_multiple(b):
    """Stuetze schneidet Decke, Decke schneidet Wand: Rangfolge verbinden."""
    rect(b, 6, 40, 90, 56, DUNKEL)       # Decke
    rect(b, 10, 44, 86, 52, LINIE)
    rect(b, 18, 54, 38, 92, DUNKEL)      # Wand unter der Decke
    rect(b, 22, 54, 34, 88, LINIE)
    rect(b, 58, 4, 80, 92, DUNKEL)       # Stuetze durch die Decke
    rect(b, 62, 8, 76, 88, BLAU)
    kreis(b, 26, 20, 15, WEISS)
    kreis(b, 26, 20, 12, GRUEN)
    rect(b, 24, 12, 28, 28, WEISS)
    rect(b, 18, 18, 34, 22, WEISS)


def filter_more(b):
    """Baum mit Haken, darueber ein Trichter: Auswahl tiefer filtern."""
    polygon(b, [(40, 4), (92, 4), (73, 26), (59, 26)], DUNKEL)
    polygon(b, [(47, 8), (85, 8), (70, 23), (62, 23)], BLAU)
    rect(b, 59, 26, 73, 42, DUNKEL)
    rect(b, 63, 26, 69, 38, BLAU)
    rect(b, 12, 22, 16, 84, DUNKEL)           # Baumlinie
    for i, (x, farbe) in enumerate(((6, GRUEN), (24, GRUEN), (24, LINIE),
                                    (42, GRUEN))):
        y = 16 + i * 20
        if i:
            rect(b, 14, y + 6, x, y + 10, DUNKEL)
        rect(b, x, y, x + 16, y + 16, DUNKEL)
        rect(b, x + 3, y + 3, x + 13, y + 13, farbe if farbe != LINIE
             else WEISS)
        rect(b, x + 22, y + 5, min(x + 60, 92), y + 11, LINIE)


def level_auto_set(b):
    """Ebenenlinien, ein Block bleibt stehen, der Ebenenpfeil springt."""
    for y in (22, 74):
        rect(b, 4, y, 92, y + 4, LINIE)
        polygon(b, [(4, y - 6), (16, y - 6), (10, y)], BLAU)
    rect(b, 30, 30, 66, 70, DUNKEL)                   # Element bleibt
    rect(b, 34, 34, 62, 66, WEISS)
    rect(b, 34, 58, 62, 66, GELB)
    pfeil(b, (82, 70), (82, 30), 7, 13, GRUEN)        # Ebene wechselt


def transfer_single(b):
    """Zwei Projektblätter, ein Element wandert per Pfeil hinüber."""
    blattform(b, 4, 10, 44, 62, BLAU, 8)
    blattform(b, 52, 34, 92, 86, GRUEN, 8)
    rect(b, 12, 30, 26, 44, ORANGE)             # Element in der Quelle
    rect(b, 66, 54, 80, 68, ORANGE)             # übertragene Kopie
    pfeil(b, (28, 50), (62, 62), 7, 13, DUNKEL)


def room_center(b):
    """L-foermiger Raum, Punkt wandert von der Ecke in die Mitte."""
    rect(b, 8, 8, 52, 88, DUNKEL)
    rect(b, 8, 46, 88, 88, DUNKEL)
    rect(b, 14, 14, 46, 82, WEISS)
    rect(b, 14, 52, 82, 82, WEISS)
    kreis(b, 28, 26, 6, LINIE)
    pfeil(b, (29, 33), (35, 55), 6, 11, DUNKEL)
    kreis(b, 38, 66, 11, ORANGE)
    kreis(b, 38, 66, 6, WEISS)
    kreis(b, 38, 66, 3, ORANGE)



def view_template_manager(b):
    """Zwei Vorlagenblaetter uebereinander, Zeilen mit Einschliessen-Haken."""
    blattform(b, 4, 6, 58, 60, BLAU, 8)          # weitere Vorlage dahinter
    blattform(b, 24, 26, 92, 92, BLAU, 10)       # markierte Vorlage
    for y in (50, 66, 82):
        rect(b, 30, y - 2, 62, y + 2, LINIE)     # Parameterzeile
        rahmen(b, 70, y - 7, 84, y + 7, 2, DUNKEL)
        linie(b, (73, y + 1), (76, y + 4), 3, GRUEN)
        linie(b, (76, y + 4), (82, y - 4), 3, GRUEN)


def filter_manager(b):
    """Filtertrichter neben Regelzeilen: Filter und ihre Regeln verwalten."""
    polygon(b, [(4, 14), (54, 14), (36, 40), (22, 40)], DUNKEL)
    polygon(b, [(10, 19), (48, 19), (33, 36), (25, 36)], BLAU)
    rect(b, 22, 40, 36, 70, DUNKEL)
    rect(b, 26, 40, 32, 64, BLAU)
    blattform(b, 46, 40, 92, 88)
    for i, farbe in enumerate((GRUEN, ORANGE, GRUEN)):
        rect(b, 52, 50 + i * 12, 60, 56 + i * 12, farbe)
        rect(b, 64, 51 + i * 12, 86, 55 + i * 12, LINIE)


def linked_ids(b):
    """Kettenglieder ueber dem verknuepften Modell, davor ein Schild mit der Id."""
    blattform(b, 4, 4, 60, 58, BLAU, 8)          # verknuepftes Modell
    rect(b, 10, 28, 54, 44, DUNKEL)              # Rohr in der Verknuepfung
    rect(b, 13, 32, 51, 40, ORANGE)
    kreis(b, 23, 17, 9, DUNKEL)                  # Kettenglieder = Verknuepfung
    kreis(b, 23, 17, 5, WEISS)
    kreis(b, 36, 17, 9, DUNKEL)
    kreis(b, 36, 17, 5, WEISS)
    polygon(b, [(30, 78), (48, 56), (92, 56), (92, 92), (48, 92)], DUNKEL)
    polygon(b, [(38, 78), (52, 61), (87, 61), (87, 87), (52, 87)], WEISS)
    kreis(b, 58, 74, 4, DUNKEL)                  # Loch im Schild
    rect(b, 66, 66, 84, 72, GRUEN)               # Id-Zeilen
    rect(b, 66, 76, 78, 82, GRUEN)


def _schnittkasten(b, akzent):
    """Wuerfel in Isometrie - die Schnittbox."""
    polygon(b, [(28, 28), (52, 44), (28, 60), (4, 44)], WEISS)   # Deckel
    polygon(b, [(4, 44), (28, 60), (28, 82), (4, 66)], LINIE)    # linke Wange
    polygon(b, [(52, 44), (28, 60), (28, 82), (52, 66)], akzent)  # rechte
    kanten = [((28, 28), (52, 44)), ((52, 44), (28, 60)),
              ((28, 60), (4, 44)), ((4, 44), (28, 28)),
              ((4, 44), (4, 66)), ((52, 44), (52, 66)),
              ((28, 60), (28, 82)), ((4, 66), (28, 82)),
              ((28, 82), (52, 66))]
    for p0, p1 in kanten:
        linie(b, p0, p1, 3, DUNKEL)


def section_box_copy(b):
    """Schnittbox, ein Pfeil traegt sie als Text auf ein Blatt."""
    _schnittkasten(b, BLAU)
    blattform(b, 60, 6, 92, 48, GRUEN, 7)        # Text in der Zwischenablage
    rect(b, 64, 24, 88, 28, LINIE)
    rect(b, 64, 32, 80, 36, LINIE)
    pfeil(b, (40, 34), (60, 22), 6, 12, GRUEN)


def section_box_paste(b):
    """Blatt mit der Box, ein Pfeil setzt sie ins zweite Modell."""
    blattform(b, 60, 6, 92, 48, ORANGE, 7)
    rect(b, 64, 24, 88, 28, LINIE)
    rect(b, 64, 32, 80, 36, LINIE)
    _schnittkasten(b, ORANGE)
    pfeil(b, (60, 22), (40, 34), 6, 12, ORANGE)


def section_box_selection(b):
    """Schnittbox um ein markiertes Element, daneben die Kette der
    Verknuepfung."""
    polygon(b, [(28, 28), (52, 44), (28, 60), (4, 44)], WEISS)
    polygon(b, [(4, 44), (28, 60), (28, 82), (4, 66)], LINIE)
    polygon(b, [(52, 44), (28, 60), (28, 82), (52, 66)], BLAU)
    rect(b, 16, 48, 42, 72, DUNKEL)              # Element in der Box
    rect(b, 19, 51, 39, 69, ORANGE)
    kanten = [((28, 28), (52, 44)), ((52, 44), (28, 60)),
              ((28, 60), (4, 44)), ((4, 44), (28, 28)),
              ((4, 44), (4, 66)), ((52, 44), (52, 66)),
              ((28, 60), (28, 82)), ((4, 66), (28, 82)),
              ((28, 82), (52, 66))]
    for p0, p1 in kanten:
        linie(b, p0, p1, 3, DUNKEL)
    kreis(b, 69, 17, 10, DUNKEL)                 # Kette = Verknuepfung
    kreis(b, 69, 17, 6, WEISS)
    kreis(b, 81, 29, 10, DUNKEL)
    kreis(b, 81, 29, 6, WEISS)


SYMBOLE = {
    "Oberflaeche.panel/TabManager.pushbutton": tab_manager,
    "Ansichten.panel/Duplizieren.stack/DuplicatePlan.pushbutton": duplicate_plan,
    "Ansichten.panel/Duplizieren.stack/DuplicateView.pushbutton": duplicate_view,
    "Filter.panel/PassFilterOverrides.pushbutton": pass_filter_overrides,
    "Raeume.panel/Beschriftung.stack/RoomCenter.pushbutton": room_center,
    "Raeume.panel/Beschriftung.stack/TagDistance.pushbutton": tag_distance,
    "Ansichten.panel/Hilfen.stack/ViewIdVisible.pushbutton": view_id_visible,
    "Ansichten.panel/ViewNameManager.pushbutton": view_name_manager,
    "Ansichten.panel/Duplizieren.stack/ViewToSheet.pushbutton": view_to_sheet,
    "Ansichten.panel/Hilfen.stack/WallLegend.pushbutton": wall_legend,
    "Filter.panel/FilterManager.pushbutton": filter_manager,
    "Phasen.panel/Phasen.stack/CopyWithPhases.pushbutton": copy_with_phases,
    "Phasen.panel/Phasen.stack/CopyPasteWithPhases.pushbutton": copy_paste_with_phases,
    "Worksets.panel/Worksets.stack/WorksetON.pushbutton": workset_on,
    "Worksets.panel/Worksets.stack/WorksetOFF.pushbutton": workset_off,
    "Worksets.panel/Worksets.stack/WorksetREVERSE.pushbutton": workset_reverse,
    "Worksets.panel/WorksetCreator.pushbutton": workset_creator,
    "Geometrie.panel/JoinMultiple.pushbutton": join_multiple,
    "Auswahl.panel/FilterMore.pushbutton": filter_more,
    "Auswahl.panel/LinkedIds.pushbutton": linked_ids,
    "Geometrie.panel/LevelAutoSet.pushbutton": level_auto_set,
    "Projekt.panel/TransferSingle.pushbutton": transfer_single,
    "Ansichten.panel/ViewTemplateManager.pushbutton": view_template_manager,
    "Ansichten.panel/SectionBox.stack/SectionBoxCopy.pushbutton": section_box_copy,
    "Ansichten.panel/SectionBox.stack/SectionBoxPaste.pushbutton": section_box_paste,
    "Ansichten.panel/SectionBox.stack/SectionBoxSelection.pushbutton": section_box_selection,
}


if __name__ == "__main__":
    tab_ordner = sys.argv[1]
    nur = sys.argv[2:] if len(sys.argv) > 2 else None
    for rel, zeichner in sorted(SYMBOLE.items()):
        if nur and not any(n in rel for n in nur):
            continue
        ziel = os.path.join(tab_ordner, rel.replace("/", os.sep), "icon.png")
        if not os.path.isdir(os.path.dirname(ziel)):
            print("  fehlt:", ziel)
            continue
        puffer = leer()
        zeichner(puffer)
        schreibe_png(ziel, verkleinern(puffer))
        print("  geschrieben:", rel)
