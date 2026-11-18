# -*- coding: utf-8 -*-
"""Revit-freie Regeln von EmptyTags.

Revit zeichnet ein "?" in eine Beschriftung, wenn das Bauteil den Parameter
eines Labels gar nicht hat - etwa weil eine Fensterbeschriftung an einer
anderen Fensterfamilie hängt. IndependentTag.TagText liefert an dieser
Stelle aber kein "?", sondern nur Präfix und Suffix ohne Wert
("AS OK= üFOK"). Ein "?" lässt sich über die API also nicht direkt finden.

Deshalb arbeitet das Werkzeug mit einem Muster: der Nutzer wählt eine (oder
mehrere) Beschriftungen mit "?". Gewählt werden dann alle Beschriftungen der
Ansicht, die zum selben Schlüssel (Typ, mit Shift die ganze Familie) gehören
und genau denselben Text zeigen wie eines der Muster.
"""


def normiere(text):
    """Tag-Text vergleichbar machen.

    Zeilenenden vereinheitlichen, Leerraum am Zeilenende und leere
    Schlusszeilen weglassen - Revit hängt bei leeren Labels gern ein "\\n" an.
    """
    if text is None:
        return None
    zeilen = text.replace(u"\r\n", u"\n").replace(u"\r", u"\n").split(u"\n")
    zeilen = [zeile.rstrip() for zeile in zeilen]
    while zeilen and not zeilen[-1]:
        zeilen.pop()
    return u"\n".join(zeilen)


def muster(beispiele):
    """[(schluessel, text), ...] -> {schluessel: set(normierte Texte)}.

    Beispiele ohne lesbaren Text (None) fallen weg.
    """
    ergebnis = {}
    for schluessel, text in beispiele:
        norm = normiere(text)
        if norm is None:
            continue
        ergebnis.setdefault(schluessel, set()).add(norm)
    return ergebnis


def treffer(kandidaten, vorlagen):
    """IDs der Kandidaten, deren Text zu einem Muster ihres Schlüssels passt.

    kandidaten: [(id, schluessel, text), ...]
    vorlagen:   Ergebnis von muster()
    Reihenfolge wie in kandidaten, jede ID nur einmal.
    """
    gefunden = []
    gesehen = set()
    for element_id, schluessel, text in kandidaten:
        texte = vorlagen.get(schluessel)
        if not texte or element_id in gesehen:
            continue
        if normiere(text) in texte:
            gesehen.add(element_id)
            gefunden.append(element_id)
    return gefunden
