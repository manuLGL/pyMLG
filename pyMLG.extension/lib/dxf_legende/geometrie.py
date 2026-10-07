# -*- coding: utf-8 -*-
"""Geometrie-Hilfen für DxfLegend (ohne Revit).

Kurven werden überall als Polylinie mit Ausbuchtung beschrieben, wie in
DXF-LWPOLYLINE: eine Liste von Scheitelpunkten (x, y, bulge). bulge gehört
zum Abschnitt vom Punkt zum nächsten Punkt und ist tan(Öffnungswinkel / 4),
positiv gegen den Uhrzeigersinn. Ein Kreis sind zwei Halbbögen mit bulge 1.

Matrizen sind 2D-affin: (a, b, c, d, e, f) mit
    x' = a*x + b*y + e
    y' = c*x + d*y + f
"""

import math

EINHEIT = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

# Ab dieser Abweichung gilt eine Matrix nicht mehr als winkeltreu
TOLERANZ = 1e-9


def verschiebung(dx, dy):
    return (1.0, 0.0, 0.0, 1.0, dx, dy)


def drehung(winkel):
    c, s = math.cos(winkel), math.sin(winkel)
    return (c, -s, s, c, 0.0, 0.0)


def skalierung(sx, sy):
    return (sx, 0.0, 0.0, sy, 0.0, 0.0)


def verkette(m1, m2):
    """m1 nach m2: erst m2 anwenden, dann m1."""
    a1, b1, c1, d1, e1, f1 = m1
    a2, b2, c2, d2, e2, f2 = m2
    return (a1 * a2 + b1 * c2, a1 * b2 + b1 * d2,
            c1 * a2 + d1 * c2, c1 * b2 + d1 * d2,
            a1 * e2 + b1 * f2 + e1, c1 * e2 + d1 * f2 + f1)


def anwenden(m, x, y):
    return (m[0] * x + m[1] * y + m[4], m[2] * x + m[3] * y + m[5])


def richtung(m, dx, dy):
    """Vektor ohne Verschiebung abbilden."""
    return (m[0] * dx + m[1] * dy, m[2] * dx + m[3] * dy)


def determinante(m):
    return m[0] * m[3] - m[1] * m[2]


def mittlerer_faktor(m):
    """Längenfaktor der Matrix (Wurzel der Flächenverzerrung)."""
    return math.sqrt(abs(determinante(m)))


def winkeltreu(m):
    """True, wenn die Matrix Kreise auf Kreise abbildet (gleichmässig
    skaliert, gedreht, verschoben, eventuell gespiegelt)."""
    l1 = math.hypot(m[0], m[2])
    l2 = math.hypot(m[1], m[3])
    if l1 < TOLERANZ or l2 < TOLERANZ:
        return False
    gleich = abs(l1 - l2) <= 1e-7 * max(l1, l2)
    senkrecht = abs(m[0] * m[1] + m[2] * m[3]) <= 1e-7 * l1 * l2
    return gleich and senkrecht


# ---------------------------------------------------------------------------
# Bögen
# ---------------------------------------------------------------------------

def bulge_aus_winkel(oeffnung):
    """Öffnungswinkel (Bogenmass, positiv = gegen Uhrzeiger) -> bulge."""
    return math.tan(oeffnung / 4.0)


def bogen(p1, p2, bulge):
    """Mittelpunkt, Radius, Startwinkel und Öffnungswinkel eines
    Bogenabschnitts. Öffnung positiv = gegen den Uhrzeigersinn."""
    oeffnung = 4.0 * math.atan(bulge)
    dx, dy = p2[0] - p1[0], p2[1] - p1[1]
    sehne = math.hypot(dx, dy)
    radius = sehne / (2.0 * abs(math.sin(oeffnung / 2.0)))
    # Abstand Sehnenmitte -> Mittelpunkt, vorzeichenbehaftet
    h = (sehne / 2.0) / math.tan(oeffnung / 2.0)
    mx, my = (p1[0] + p2[0]) / 2.0, (p1[1] + p2[1]) / 2.0
    nx, ny = -dy / sehne, dx / sehne
    cx, cy = mx + nx * h, my + ny * h
    start = math.atan2(p1[1] - cy, p1[0] - cx)
    return (cx, cy), radius, start, oeffnung


def bogenmitte(p1, p2, bulge):
    """Punkt in der Mitte des Bogens (für Revit Arc.Create mit 3 Punkten)."""
    dx, dy = p2[0] - p1[0], p2[1] - p1[1]
    # Pfeilhöhe = bulge * halbe Sehne, rechts der Laufrichtung bei bulge > 0
    mx, my = (p1[0] + p2[0]) / 2.0, (p1[1] + p2[1]) / 2.0
    return (mx + dy * bulge / 2.0, my - dx * bulge / 2.0)


def abschnitte(punkte, geschlossen):
    """Paare (p1, p2, bulge) einer Polylinie."""
    anzahl = len(punkte)
    ende = anzahl if geschlossen else anzahl - 1
    for i in range(max(0, ende)):
        p1 = punkte[i]
        p2 = punkte[(i + 1) % anzahl]
        yield (p1[0], p1[1]), (p2[0], p2[1]), p1[2] if len(p1) > 2 else 0.0


def abtasten(punkte, geschlossen, grad_je_schritt=10.0):
    """Polylinie mit Bögen -> Punktliste [(x, y)] (Bögen als Sehnenzug).
    Bei geschlossenen Linien ist der erste Punkt am Ende wiederholt."""
    if not punkte:
        return []
    ergebnis = [(punkte[0][0], punkte[0][1])]
    for p1, p2, bulge in abschnitte(punkte, geschlossen):
        if abs(bulge) > 1e-12 and (p1[0] != p2[0] or p1[1] != p2[1]):
            (cx, cy), r, start, oeffnung = bogen(p1, p2, bulge)
            schritte = max(2, int(math.ceil(abs(math.degrees(oeffnung))
                                            / grad_je_schritt)))
            for k in range(1, schritte):
                w = start + oeffnung * k / float(schritte)
                ergebnis.append((cx + r * math.cos(w), cy + r * math.sin(w)))
        ergebnis.append(p2)
    return ergebnis


def transformiere(punkte, geschlossen, m):
    """Polylinie mit Bögen abbilden. Bei winkeltreuen Matrizen bleiben Bögen
    Bögen (gespiegelt kehrt sich der Drehsinn um), sonst werden sie
    abgetastet (aus einem Kreis wird eine Ellipse)."""
    if winkeltreu(m):
        vorzeichen = -1.0 if determinante(m) < 0 else 1.0
        ergebnis = []
        for p in punkte:
            x, y = anwenden(m, p[0], p[1])
            ergebnis.append((x, y, vorzeichen * (p[2] if len(p) > 2 else 0.0)))
        return ergebnis
    zug = abtasten(punkte, geschlossen)
    if geschlossen and len(zug) > 1:
        zug = zug[:-1]
    return [anwenden(m, x, y) + (0.0,) for x, y in zug]


def grenzen(punktlisten):
    """(xmin, ymin, xmax, ymax) mehrerer Punktlisten oder None."""
    xs, ys = [], []
    for liste in punktlisten:
        for p in liste:
            xs.append(p[0])
            ys.append(p[1])
    if not xs:
        return None
    return (min(xs), min(ys), max(xs), max(ys))


# ---------------------------------------------------------------------------
# Splines und Ellipsen
# ---------------------------------------------------------------------------

def ellipse_punkte(mitte, hauptachse, verhaeltnis, start, ende, schritte=None):
    """Punkte einer Ellipse(nbogen). start/ende: Parameter im Bogenmass."""
    while ende <= start:
        ende += 2.0 * math.pi
    oeffnung = ende - start
    if schritte is None:
        schritte = max(8, int(math.ceil(64 * oeffnung / (2.0 * math.pi))))
    ax, ay = hauptachse
    bx, by = -ay * verhaeltnis, ax * verhaeltnis
    punkte = []
    for k in range(schritte + 1):
        w = start + oeffnung * k / float(schritte)
        c, s = math.cos(w), math.sin(w)
        punkte.append((mitte[0] + ax * c + bx * s, mitte[1] + ay * c + by * s))
    return punkte


def spline_punkte(grad, knoten, kontrollpunkte, gewichte=None, je_spanne=12):
    """NURBS mit de Boor abtasten -> [(x, y)]."""
    n = len(kontrollpunkte)
    if n == 0:
        return []
    if n <= grad or len(knoten) != n + grad + 1:
        return [(p[0], p[1]) for p in kontrollpunkte]
    if not gewichte or len(gewichte) != n:
        gewichte = [1.0] * n
    t0, t1 = knoten[grad], knoten[n]
    schritte = max(16, je_spanne * (n - grad))
    punkte = []
    for k in range(schritte + 1):
        t = t0 + (t1 - t0) * k / float(schritte)
        punkte.append(_de_boor(grad, knoten, kontrollpunkte, gewichte, t))
    return punkte


def _de_boor(p, u, kp, w, t):
    n = len(kp)
    # Spanne suchen
    if t >= u[n]:
        k = n - 1
        while k > p and u[k] == u[k + 1]:
            k -= 1
    else:
        k = p
        while k < n - 1 and not (u[k] <= t < u[k + 1]):
            k += 1
    d = []
    for j in range(p + 1):
        i = j + k - p
        d.append([kp[i][0] * w[i], kp[i][1] * w[i], w[i]])
    for r in range(1, p + 1):
        for j in range(p, r - 1, -1):
            i = j + k - p
            nenner = u[i + p - r + 1] - u[i]
            alpha = 0.0 if nenner == 0 else (t - u[i]) / nenner
            d[j] = [(1.0 - alpha) * d[j - 1][q] + alpha * d[j][q]
                    for q in range(3)]
    gw = d[p][2] or 1.0
    return (d[p][0] / gw, d[p][1] / gw)
