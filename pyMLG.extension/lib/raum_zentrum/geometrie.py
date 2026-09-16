# -*- coding: utf-8 -*-
"""Flächenschwerpunkt und Ersatzpunkte für Raumumrisse.

Bewusst ohne Revit-API: Punkte sind (x, y)-Tupel in beliebiger Einheit
(in Revit: interne Fuss). Dadurch lässt sich die Mathematik ausserhalb von
Revit prüfen (siehe tools/test_raum_zentrum.py).
"""

import math

# Punkte, die näher beieinander liegen, gelten als identisch (ca. 0,3 mm).
PUNKT_TOLERANZ = 1e-3


def _abstand2(a, b):
    return (a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2


def verkette_kurvenpunkte(kurven):
    """Setzt die tessellierten Punkte einzelner Kurven zu einer Schleife zusammen.

    kurven: Liste von Punktlisten (je Kurve das Ergebnis von Curve.Tessellate()).
    Revit liefert die Segmente einer Schleife bereits Kopf an Fuss; zur
    Sicherheit wird eine verkehrt herum laufende Kurve umgedreht. Doppelte
    Punkte an den Stössen und der Schlusspunkt werden entfernt.
    """
    tol2 = PUNKT_TOLERANZ ** 2
    punkte = []
    for kurve in kurven:
        kurve = list(kurve)
        if not kurve:
            continue
        if punkte:
            letzter = punkte[-1]
            if _abstand2(kurve[-1], letzter) < _abstand2(kurve[0], letzter):
                kurve.reverse()
            if _abstand2(kurve[0], letzter) < tol2:
                kurve = kurve[1:]
        punkte.extend(kurve)
    if len(punkte) > 1 and _abstand2(punkte[0], punkte[-1]) < tol2:
        punkte.pop()
    return punkte


def flaeche_und_schwerpunkt(punkte):
    """Vorzeichenbehaftete Fläche und Schwerpunkt eines einfachen Polygons.

    Standardformel (Gausssche Trapezformel / Shoelace):
        A  = 1/2 * Σ (x_i*y_i+1 - x_i+1*y_i)
        Cx = 1/(6A) * Σ (x_i + x_i+1) * (x_i*y_i+1 - x_i+1*y_i)
        Cy = 1/(6A) * Σ (y_i + y_i+1) * (x_i*y_i+1 - x_i+1*y_i)
    A > 0 bei Umlauf gegen den Uhrzeigersinn.

    Gerechnet wird relativ zum ersten Punkt: Projektkoordinaten können weit
    vom Ursprung entfernt liegen, und die Kreuzprodukte grosser Zahlen
    verlieren sonst Genauigkeit.
    Rückgabe: (A, cx, cy); cx/cy sind None bei Fläche 0.
    """
    n = len(punkte)
    if n < 3:
        return 0.0, None, None
    ox, oy = punkte[0]
    a2 = 0.0
    sx = 0.0
    sy = 0.0
    for i in range(n):
        x0 = punkte[i][0] - ox
        y0 = punkte[i][1] - oy
        x1 = punkte[(i + 1) % n][0] - ox
        y1 = punkte[(i + 1) % n][1] - oy
        kreuz = x0 * y1 - x1 * y0
        a2 += kreuz
        sx += (x0 + x1) * kreuz
        sy += (y0 + y1) * kreuz
    flaeche = a2 / 2.0
    if abs(flaeche) < 1e-12:
        return 0.0, None, None
    return flaeche, sx / (3.0 * a2) + ox, sy / (3.0 * a2) + oy


def punkt_in_polygon(x, y, punkte):
    """Strahlverfahren (gerade/ungerade Anzahl Schnitte)."""
    innen = False
    n = len(punkte)
    j = n - 1
    for i in range(n):
        xi, yi = punkte[i]
        xj, yj = punkte[j]
        if (yi > y) != (yj > y):
            x_schnitt = xi + (y - yi) * (xj - xi) / (yj - yi)
            if x < x_schnitt:
                innen = not innen
        j = i
    return innen


def schwerpunkt_mit_loechern(schleifen):
    """Flächenschwerpunkt eines Raumumrisses aus einer oder mehreren Schleifen.

    Die Hauptschleife ist die mit der grössten Fläche. Jede Schleife wird nach
    ihrer Verschachtelungstiefe gewichtet: liegt sie in einer ungeraden Anzahl
    anderer Schleifen, ist sie ein Loch (Stütze, Schacht) und zählt negativ,
    sonst positiv. Die Umlaufrichtung aus Revit spielt dadurch keine Rolle.

        Cx = Σ(±|A_i| * Cx_i) / Σ(±|A_i|)

    Rückgabe: dict mit flaeche, cx, cy, schleifen, loecher, punkte.
    Wirft ValueError, wenn keine auswertbare Fläche übrig bleibt.
    """
    auswertbar = []
    for punkte in schleifen:
        flaeche, cx, cy = flaeche_und_schwerpunkt(punkte)
        if cx is not None:
            auswertbar.append((abs(flaeche), cx, cy, punkte))
    if not auswertbar:
        raise ValueError(u"Der Umriss hat keine auswertbare Fläche.")

    # Grösste zuerst - das ist die äussere Hauptschleife
    auswertbar.sort(key=lambda eintrag: eintrag[0], reverse=True)

    summe_a = 0.0
    summe_x = 0.0
    summe_y = 0.0
    loecher = 0
    for index, (flaeche, cx, cy, punkte) in enumerate(auswertbar):
        px, py = punkte[0]
        tiefe = sum(1 for anderer, (_a, _x, _y, andere_punkte)
                    in enumerate(auswertbar)
                    if anderer != index
                    and punkt_in_polygon(px, py, andere_punkte))
        vorzeichen = -1.0 if tiefe % 2 else 1.0
        if vorzeichen < 0:
            loecher += 1
        summe_a += vorzeichen * flaeche
        summe_x += vorzeichen * flaeche * cx
        summe_y += vorzeichen * flaeche * cy

    if summe_a <= 1e-12:
        raise ValueError(u"Die Löcher sind grösser als der Umriss.")

    return {
        "flaeche": summe_a,
        "cx": summe_x / summe_a,
        "cy": summe_y / summe_a,
        "schleifen": len(auswertbar),
        "loecher": loecher,
        "punkte": sum(len(eintrag[3]) for eintrag in auswertbar),
    }


def punkt_in_umriss(x, y, schleifen):
    """Gerade/ungerade-Regel über alle Schleifen: Löcher zählen als aussen."""
    innen = False
    for punkte in schleifen:
        if punkt_in_polygon(x, y, punkte):
            innen = not innen
    return innen


# ---------------------------------------------------------------------------
# Ersatzpunkte für den Fall, dass der Schwerpunkt ausserhalb liegt
# ---------------------------------------------------------------------------

def _intervalle(schleifen, wert, senkrecht):
    """Innenliegende Abschnitte einer waagrechten (y = wert) bzw. senkrechten
    (x = wert) Linie durch den Umriss, gerade/ungerade über alle Schleifen."""
    schnitte = []
    for punkte in schleifen:
        n = len(punkte)
        for i in range(n):
            a = punkte[i]
            b = punkte[(i + 1) % n]
            if senkrecht:
                a = (a[1], a[0])
                b = (b[1], b[0])
            # Halboffene Regel: ein Eckpunkt auf der Linie zählt nur einmal
            if (a[1] > wert) != (b[1] > wert):
                schnitte.append(a[0] + (wert - a[1]) * (b[0] - a[0])
                                / (b[1] - a[1]))
    schnitte.sort()
    return [(schnitte[i], schnitte[i + 1])
            for i in range(0, len(schnitte) - 1, 2)
            if schnitte[i + 1] - schnitte[i] > PUNKT_TOLERANZ]


def _zentriere_quer(schleifen, x, y, senkrecht):
    """Schiebt einen Punkt quer zur Suchlinie in die Mitte seines Abschnitts."""
    wert = x if not senkrecht else y
    lage = y if not senkrecht else x
    for von, bis in _intervalle(schleifen, wert, not senkrecht):
        if von <= lage <= bis:
            mitte = (von + bis) / 2.0
            return (x, mitte) if not senkrecht else (mitte, y)
    return x, y


def ersatzpunkte(schleifen, cx, cy, linien_je_richtung=24):
    """Liefert Innenpunkte, sortiert nach Abstand zum Schwerpunkt (cx, cy).

    Vorgehen ("Achsmittellinien"):
    1. Waagrechte und senkrechte Suchlinien - durch den Schwerpunkt und
       gleichmässig verteilt über die Umrisshülle - mit dem Umriss schneiden.
    2. Jeder innenliegende Abschnitt liefert seinen Mittelpunkt. Er liegt auf
       der Mittelachse des jeweiligen Teilstücks (z. B. eines U-Schenkels).
    3. Diesen Punkt einmal quer zur Suchlinie mittig setzen, damit er nicht
       dicht an einer Wand klebt, sondern im Teilstück zentriert ist.
    4. Nach Abstand zum Schwerpunkt sortieren: der erste Punkt ist der
       nächstgelegene sinnvolle Innenpunkt.
    Alle Punkte liegen nach der gerade/ungerade-Regel im Umriss (nicht in
    Löchern); die endgültige Prüfung macht trotzdem Room.IsPointInRoom().
    """
    alle = [p for punkte in schleifen for p in punkte]
    if not alle:
        return []
    min_x = min(p[0] for p in alle)
    max_x = max(p[0] for p in alle)
    min_y = min(p[1] for p in alle)
    max_y = max(p[1] for p in alle)

    kandidaten = []
    for senkrecht, lage, von, bis in ((False, cy, min_y, max_y),
                                      (True, cx, min_x, max_x)):
        werte = [lage]
        schritt = (bis - von) / float(linien_je_richtung + 1)
        werte += [von + schritt * (k + 1) for k in range(linien_je_richtung)]
        for wert in werte:
            for a, b in _intervalle(schleifen, wert, senkrecht):
                mitte = (a + b) / 2.0
                x, y = (wert, mitte) if senkrecht else (mitte, wert)
                x, y = _zentriere_quer(schleifen, x, y, senkrecht)
                if punkt_in_umriss(x, y, schleifen):
                    kandidaten.append((x, y))

    # Doppelte entfernen, nach Abstand zum Schwerpunkt sortieren
    eindeutig = {}
    for x, y in kandidaten:
        schluessel = (round(x / PUNKT_TOLERANZ), round(y / PUNKT_TOLERANZ))
        eindeutig.setdefault(schluessel, (x, y))
    return sorted(eindeutig.values(),
                  key=lambda p: math.hypot(p[0] - cx, p[1] - cy))
