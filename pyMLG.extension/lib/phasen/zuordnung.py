# -*- coding: utf-8 -*-
"""Ordnet kopierte Elemente ihren Originalen zu.

ElementTransformUtils.CopyElements() liefert die neuen Ids ohne Bezug zu den
Originalen - Reihenfolge und Anzahl stimmen nicht mit der Eingabe überein
(abhängige Elemente wie Türen in Wänden kommen hinzu). Die Zuordnung erfolgt
deshalb über Kategorie, Typ und Lage: eine Kopie liegt dort, wo das Original
liegt, plus Verschiebung.

Bewusst ohne Revit-API, damit die Logik ausserhalb von Revit prüfbar ist
(siehe tools/test_phasen_zuordnung.py). Einheiten: interne Fuss.
"""

# Grösste zulässige Abweichung eines Referenzpunkts (ca. 15 cm). Lagepunkte
# stimmen exakt überein; Luft für Begrenzungsrahmen, die sich durch
# Verbindungen mit Nachbarbauteilen am neuen Ort leicht ändern.
MAX_ABWEICHUNG = 0.5


class Eintrag(object):
    """Ein Element für die Zuordnung.

    element_id: beliebige Kennung (in Revit die ElementId)
    schluessel: z.B. (Kategorie, Typ) - nur gleiche Schlüssel passen zusammen
    punkte:     Liste von (x, y, z) - Lagepunkt(e) oder Begrenzungsrahmen
    """

    def __init__(self, element_id, schluessel, punkte):
        self.element_id = element_id
        self.schluessel = schluessel
        self.punkte = list(punkte)


def abweichung(original, kopie, verschiebung):
    """Grösste Koordinatenabweichung zwischen verschobenem Original und Kopie."""
    if len(original.punkte) != len(kopie.punkte) or not original.punkte:
        return None
    dx, dy, dz = verschiebung
    groesste = 0.0
    for (ox, oy, oz), (kx, ky, kz) in zip(original.punkte, kopie.punkte):
        groesste = max(groesste, abs(ox + dx - kx), abs(oy + dy - ky),
                       abs(oz + dz - kz))
    return groesste


def zuordnen(originale, kopien, verschiebung, max_abweichung=MAX_ABWEICHUNG):
    """Liefert {kopie_id: original_id}.

    Jedes Original und jede Kopie wird höchstens einmal verwendet. Die besten
    Paare (kleinste Abweichung) werden zuerst vergeben, damit nah beieinander
    liegende gleiche Elemente nicht vertauscht werden.
    """
    paare = []
    for ki, kopie in enumerate(kopien):
        for oi, original in enumerate(originale):
            if original.schluessel != kopie.schluessel:
                continue
            a = abweichung(original, kopie, verschiebung)
            if a is not None and a <= max_abweichung:
                paare.append((a, ki, oi))
    paare.sort(key=lambda p: p[0])

    ergebnis = {}
    vergeben = set()
    for _, ki, oi in paare:
        kid = kopien[ki].element_id
        if kid in ergebnis or oi in vergeben:
            continue
        ergebnis[kid] = originale[oi].element_id
        vergeben.add(oi)
    return ergebnis
