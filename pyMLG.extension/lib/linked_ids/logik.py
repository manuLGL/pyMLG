# -*- coding: utf-8 -*-
"""Revit-freie Regeln von LinkedIds.

Ein Eintrag beschreibt ein gefundenes Element - aus einer Verknüpfung oder
aus dem geöffneten Modell.

Eine ID gilt nur in ihrem eigenen Dokument. Die Einträge werden deshalb nach
Dokument gruppiert ausgegeben; die IDs eines Dokuments stehen mit Komma
getrennt in einer Zeile, genau so nimmt Revit sie in "Auswählen nach ID" an.
"""

from mlg_sprache import t, tt

SPALTEN = (
    (u"Dokument", u"Document", u"Documento"),
    (u"Verknüpfung", u"Link", u"Vínculo"),
    (u"Kategorie", u"Category", u"Categoría"),
    (u"Familie: Typ", u"Family: type", u"Familia: tipo"),
    (u"Name", u"Name", u"Nombre"),
    (u"ID", u"ID", u"ID"),
    (u"Eindeutige ID", u"Unique ID", u"ID única"),
)


class Eintrag(object):
    """Ein gefundenes Element mit seiner ID im eigenen Dokument."""

    def __init__(self, dokument, element_id, kategorie=u"", typ=u"",
                 name=u"", unique_id=u"", verknuepfung=u"",
                 ist_verknuepfung=False, geladen=True):
        self.dokument = dokument or u""
        self.element_id = int(element_id)
        self.kategorie = kategorie or u""
        self.typ = typ or u""
        self.name = name or u""
        self.unique_id = unique_id or u""
        self.verknuepfung = verknuepfung or u""
        self.ist_verknuepfung = bool(ist_verknuepfung)
        # Nicht geladene Verknüpfung: nur die ID ist bekannt
        self.geladen = bool(geladen)

    @property
    def schluessel(self):
        return (self.dokument, self.element_id)

    @property
    def beschreibung(self):
        """Kategorie, Typ und Name in einer Zeile."""
        teile = [teil for teil in (self.kategorie, self.typ) if teil]
        if self.name and self.name not in self.typ:
            teile.append(self.name)
        return u" · ".join(teile) or t(
            u"Element aus einer nicht geladenen Verknüpfung",
            u"Element from a link that is not loaded",
            u"Elemento de un vínculo no cargado")

    @property
    def herkunft(self):
        """Zeile unter der Beschreibung: woher das Element stammt."""
        if not self.ist_verknuepfung:
            return t(u"Aktuelles Modell", u"Current model", u"Modelo actual")
        text = self.dokument
        if self.verknuepfung and self.verknuepfung != self.dokument:
            text = u"%s  (%s)" % (text, self.verknuepfung)
        if not self.geladen:
            text += t(u"  - nicht geladen", u"  - not loaded",
                      u"  - no cargado")
        return text


def ergaenze(eintraege, neue):
    """Neue Einträge anhängen, schon vorhandene (Dokument + ID) auslassen.

    Rückgabe: Anzahl der angehängten Einträge.
    """
    vorhanden = set(eintrag.schluessel for eintrag in eintraege)
    angehaengt = 0
    for eintrag in neue:
        if eintrag.schluessel in vorhanden:
            continue
        vorhanden.add(eintrag.schluessel)
        eintraege.append(eintrag)
        angehaengt += 1
    return angehaengt


def gruppiere(eintraege):
    """[(Dokumentname, [Eintrag, ...]), ...] in der Reihenfolge des Fundes."""
    ordnung = []
    gruppen = {}
    for eintrag in eintraege:
        if eintrag.dokument not in gruppen:
            gruppen[eintrag.dokument] = []
            ordnung.append(eintrag.dokument)
        gruppen[eintrag.dokument].append(eintrag)
    return [(name, gruppen[name]) for name in ordnung]


def ids_text(eintraege):
    """IDs für die Zwischenablage - je Dokument eine Zeile mit Komma.

    Bei mehreren Dokumenten steht der Dokumentname über seinen IDs, sonst
    wüsste man beim Einfügen nicht mehr, in welcher Datei sie gelten.
    """
    gruppen = gruppiere(eintraege)
    teile = []
    for name, liste in gruppen:
        ids = []
        for eintrag in liste:
            if eintrag.element_id not in ids:
                ids.append(eintrag.element_id)
        zeile = u", ".join(u"%d" % wert for wert in ids)
        teile.append(zeile if len(gruppen) == 1
                     else u"%s:\r\n%s" % (name, zeile))
    return u"\r\n\r\n".join(teile)


def tabelle(eintraege):
    """Alle Angaben als Tabelle (Tabulator getrennt, für Excel)."""
    zeilen = [u"\t".join(tt(spalte) for spalte in SPALTEN)]
    for eintrag in eintraege:
        zeilen.append(u"\t".join((
            eintrag.dokument, eintrag.verknuepfung, eintrag.kategorie,
            eintrag.typ, eintrag.name, u"%d" % eintrag.element_id,
            eintrag.unique_id)))
    return u"\r\n".join(zeilen)


def zusammenfassung(eintraege):
    """Kopfzeile des Fensters: wie viele Elemente aus wie vielen Quellen."""
    verknuepft = [eintrag for eintrag in eintraege if eintrag.ist_verknuepfung]
    eigene = len(eintraege) - len(verknuepft)
    if not verknuepft:
        return t(u"%d Element(e) aus dem aktuellen Modell",
                 u"%d element(s) from the current model",
                 u"%d elemento(s) del modelo actual") % eigene
    dokumente = len(set(eintrag.dokument for eintrag in verknuepft))
    text = t(u"%d Element(e) aus %d Verknüpfung(en)",
             u"%d element(s) from %d link(s)",
             u"%d elemento(s) de %d vínculo(s)") % (len(verknuepft), dokumente)
    if eigene:
        text += t(u", %d aus dem aktuellen Modell",
                  u", %d from the current model",
                  u", %d del modelo actual") % eigene
    return text
