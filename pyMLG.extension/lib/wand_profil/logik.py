# -*- coding: utf-8 -*-
"""Revit-freie Logik von WallProfileCopy.

Das Profil der Quellwand wird so bewegt, wie man die Quellwand auf die
Zielwand legen würde: Drehung um die Senkrechte, damit die Wandachsen
gleich laufen, Verschiebung Anfang auf Anfang und in der Höhe um den
Unterschied der Wandunterkanten (Ebenenhöhe + Basisversatz).

Die Zielwand muss gleich lang sein - ein Profil lässt sich nicht sinnvoll
strecken.

Punkte als (x, y, z), alle Längen in Revit-internen Einheiten (Fuß).
"""

import math

# Längenunterschied, ab dem eine Wand nicht mehr als gleich gilt (ca. 1 mm)
LAENGEN_TOLERANZ = 1.0 / 304.8


def _laenge_2d(start, ende):
    return math.hypot(ende[0] - start[0], ende[1] - start[1])


def gleich_lang(quelle, ziel, toleranz=LAENGEN_TOLERANZ):
    """quelle/ziel: (start, ende) der Wandachse."""
    return abs(_laenge_2d(*quelle) - _laenge_2d(*ziel)) <= toleranz


def lage(quelle, quelle_basis, ziel, ziel_basis):
    """Bewegung von der Quell- auf die Zielwand.

    quelle, ziel   (start, ende) der Wandachse
    *_basis        absolute Höhe der Wandunterkante
    Rückgabe (winkel, verschiebung): erst um winkel (Bogenmaß) um die
    Z-Achse durch den Ursprung drehen, dann um verschiebung (x, y, z)
    verschieben.
    """
    (qs, qe), (zs, ze) = quelle, ziel
    winkel = (math.atan2(ze[1] - zs[1], ze[0] - zs[0])
              - math.atan2(qe[1] - qs[1], qe[0] - qs[0]))
    gedreht = drehe(winkel, qs)
    return winkel, (zs[0] - gedreht[0], zs[1] - gedreht[1],
                    ziel_basis - quelle_basis)


def drehe(winkel, punkt):
    c, s = math.cos(winkel), math.sin(winkel)
    return (c * punkt[0] - s * punkt[1], s * punkt[0] + c * punkt[1],
            punkt[2])


def wende_an(winkel, verschiebung, punkt):
    gedreht = drehe(winkel, punkt)
    return tuple(g + v for g, v in zip(gedreht, verschiebung))
