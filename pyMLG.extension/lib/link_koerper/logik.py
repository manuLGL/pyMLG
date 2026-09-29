# -*- coding: utf-8 -*-
"""Revit-freie Regeln von LinkKoerper.

Jedes Element der Verknüpfung wird eine eigene Familie "Allgemeines Modell"
mit genau einer Instanz. Mit Volumenkörpern aus einer Familie lassen sich
Elemente des Hauptmodells verschneiden - mit einem DirectShape nicht.

Wiedererkennen ohne Parameter im Modell:

    Familienname  pyMLG Link TGA.rvt Rohre (123456-98765)
                                           Verknüpfung(en)-Element-ID
    Typname       Form 3f9a0c1b7e2d
                  Formkennung: Geometrie relativ zu ihrer Mitte

Die Instanz steht in der Mitte des Körpers. Beim erneuten Ausführen gilt
darum:

    gleiche Form, gleiche Lage   unverändert lassen (Verschneidungen bleiben)
    gleiche Form, andere Lage    Instanz verschieben (Verschneidungen bleiben)
    andere Form                  Familie ersetzen (auch nach Drehung)
    noch nicht vorhanden         neu erstellen
"""

import hashlib
import re

from mlg_sprache import t

PRAEFIX = u"pyMLG Link"
TYP_PRAEFIX = u"Form "

# Zeichen, die Revit in Familiennamen bzw. Windows in Dateinamen ablehnt
_VERBOTEN = u'\\/:*?"<>|{}[];`~()'

# Lage und Form werden auf 1 mm verglichen
TOLERANZ_M = 0.001
# Ebenen gleicher Höhe in Verknüpfung und Hauptmodell
TOLERANZ_EBENE_M = 0.05

NEU = "neu"
GLEICH = "gleich"
VERSCHIEBEN = "verschieben"
ERSETZEN = "ersetzen"

# Ab so vielen neuen Familien fragt das Werkzeug nach
VIELE = 300

_KENNUNG = re.compile(r"\(([0-9.]+)-([0-9]+)\)\s*$")


def bereinige(text):
    """Text für Familien- und Dateinamen tauglich machen."""
    text = u"".join(u"_" if zeichen in _VERBOTEN else zeichen
                    for zeichen in (text or u""))
    return u" ".join(text.split())


def schluessel(link_ids, element_id):
    """Eindeutiger Schlüssel: Kette der Verknüpfungs-IDs + Element-ID."""
    return (u".".join(u"%d" % wert for wert in link_ids), int(element_id))


def familienname(link_ids, element_id, link_titel, kategorie):
    teile = [PRAEFIX, bereinige(link_titel)[:40], bereinige(kategorie)[:30]]
    kette, wert = schluessel(link_ids, element_id)
    return u"%s (%s-%d)" % (u" ".join(teil for teil in teile if teil),
                            kette, wert)


def schluessel_aus_name(name):
    """Familienname -> Schlüssel, oder None bei fremden Familien."""
    if not name or not name.startswith(PRAEFIX + u" "):
        return None
    treffer = _KENNUNG.search(name)
    if treffer is None:
        return None
    return (treffer.group(1), int(treffer.group(2)))


def typname(form):
    return TYP_PRAEFIX + form


def form_aus_typname(name):
    if name and name.startswith(TYP_PRAEFIX):
        return name[len(TYP_PRAEFIX):]
    return None


def mitte(punkte):
    """Mitte des achsparallelen Kastens um die Punkte [(x, y, z), ...]."""
    if not punkte:
        return None
    return tuple((min(p[i] for p in punkte) + max(p[i] for p in punkte)) / 2.0
                 for i in range(3))


def formkennung(punkte_relativ, volumen_m3):
    """Kurze Kennung der Form - Punkte in Metern relativ zur Mitte.

    Auf Millimeter gerundet, damit Rechenrauschen nach einer Verschiebung
    der Verknüpfung dieselbe Kennung liefert. Eine Drehung ändert die
    Kennung, weil die Punkte achsbezogen bleiben.
    """
    gerundet = sorted(set(tuple(int(round(wert / TOLERANZ_M)) for wert in p)
                          for p in punkte_relativ))
    text = u"%r|%d" % (gerundet, int(round(volumen_m3 * 1e6)))
    return hashlib.md5(text.encode("utf-8")).hexdigest()[:12]


def waehle_ebene(ebenen, quell_name, quell_hoehe, koerper_z):
    """Index der Ebene des Hauptmodells für einen Körper, oder None.

    ebenen       [(Name, Höhe in m), ...] des Hauptmodells
    quell_name   Name der Ebene des Elements in der Verknüpfung (oder None)
    quell_hoehe  Höhe dieser Ebene im Hauptmodell in m (oder None)
    koerper_z    Höhe der Körpermitte in m

    Reihenfolge: gleicher Name, dann gleiche Höhe (TOLERANZ_EBENE_M), dann
    die höchste Ebene unter der Quellebene bzw. unter dem Körper.
    """
    if not ebenen:
        return None
    if quell_name:
        gesucht = quell_name.strip().lower()
        for index, (name, _hoehe) in enumerate(ebenen):
            if (name or u"").strip().lower() == gesucht:
                return index
    if quell_hoehe is not None:
        naechste = min(range(len(ebenen)),
                       key=lambda i: abs(ebenen[i][1] - quell_hoehe))
        if abs(ebenen[naechste][1] - quell_hoehe) <= TOLERANZ_EBENE_M:
            return naechste
    bezug = quell_hoehe if quell_hoehe is not None else koerper_z
    darunter = [index for index, (_name, hoehe) in enumerate(ebenen)
                if hoehe <= bezug + 1e-6]
    if darunter:
        return max(darunter, key=lambda i: ebenen[i][1])
    return min(range(len(ebenen)), key=lambda i: ebenen[i][1])


def abstand(a, b):
    return sum((a[i] - b[i]) ** 2 for i in range(3)) ** 0.5


def entscheide(vorhanden_form, vorhanden_mitte, form, neue_mitte):
    """Was mit einem Element geschehen soll (NEU, GLEICH, ...).

    vorhanden_form/-mitte: None, wenn es noch keinen Körper gibt bzw. die
    Familie ohne Instanz im Modell liegt.
    """
    if vorhanden_form is None and vorhanden_mitte is None:
        return NEU
    if vorhanden_form != form or vorhanden_mitte is None:
        return ERSETZEN
    if abstand(vorhanden_mitte, neue_mitte) <= TOLERANZ_M:
        return GLEICH
    return VERSCHIEBEN


class Ergebnis(object):
    """Zähler und Meldungen eines Laufs."""

    def __init__(self):
        self.anzahl = {NEU: 0, GLEICH: 0, VERSCHIEBEN: 0, ERSETZEN: 0}
        self.ohne_geometrie = 0
        self.fehler = []            # Texte
        self.verwaist = 0           # Körper, deren Element im Link fehlt
        self.ebene_angepasst = 0    # vorhandene Körper auf andere Ebene
        self.neue_ids = []          # ElementIds der erstellten Instanzen

    def zaehle(self, aktion):
        self.anzahl[aktion] += 1

    def text(self):
        zeilen = [
            t(u"Neu erstellt: %d", u"Created: %d",
              u"Creados: %d") % self.anzahl[NEU],
            t(u"Verschoben (Lage geändert): %d",
              u"Moved (position changed): %d",
              u"Desplazados (posición cambiada): %d") % self.anzahl[
                VERSCHIEBEN],
            t(u"Ersetzt (Form geändert): %d",
              u"Replaced (shape changed): %d",
              u"Sustituidos (forma cambiada): %d") % self.anzahl[ERSETZEN],
            t(u"Unverändert: %d", u"Unchanged: %d",
              u"Sin cambios: %d") % self.anzahl[GLEICH],
        ]
        if self.ebene_angepasst:
            zeilen.append(t(u"Ebene angepasst: %d", u"Level adjusted: %d",
                            u"Nivel ajustado: %d") % self.ebene_angepasst)
        if self.ohne_geometrie:
            zeilen.append(t(u"Ohne Volumengeometrie übersprungen: %d",
                            u"Skipped without solid geometry: %d",
                            u"Omitidos sin geometría sólida: %d")
                          % self.ohne_geometrie)
        if self.verwaist:
            zeilen.append(t(
                u"Körper, deren Element es in der Verknüpfung nicht mehr "
                u"gibt: %d (nicht gelöscht)",
                u"Bodies whose element no longer exists in the link: %d "
                u"(not deleted)",
                u"Cuerpos cuyo elemento ya no existe en el vínculo: %d "
                u"(no borrados)") % self.verwaist)
        if self.fehler:
            zeilen.append(t(u"Fehler: %d", u"Errors: %d",
                            u"Errores: %d") % len(self.fehler))
        return u"\n".join(zeilen)
