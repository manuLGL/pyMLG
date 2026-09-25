# -*- coding: utf-8 -*-
"""Regeln des Clash Navigators - ohne Revit, offline testbar.

Bewertung ("Eintrag") je Kollision, gespeichert von speicher.py:

    {"status": "geprueft", "klasse": "CLR", "kommentar": "...",
     "von": "manuel", "zeit": "2026-09-24 12:25"}

Fehlt ein Feld, gilt der Wert aus dem Navisworks-Bericht.
"""

import os
import re

from mlg_sprache import t, tt

from clash_navigator import bericht as br

STATUS_TEXTE = {
    br.NEU: (u"Neu", u"New", u"Nuevo"),
    br.AKTIV: (u"Aktiv", u"Active", u"Activo"),
    br.GEPRUEFT: (u"Geprüft", u"Reviewed", u"Revisado"),
    br.GENEHMIGT: (u"Genehmigt", u"Approved", u"Aprobado"),
    br.BEHOBEN: (u"Behoben", u"Resolved", u"Resuelto"),
}

# Diese Status gelten als "offen"
OFFEN = (br.NEU, br.AKTIV)

# Vorschläge für die Klassifizierung - eigene kommen beim Benutzen dazu
KLASSEN = (u"CLR", u"TGA", u"TWP", u"ARC", u"IGN", u"MANUELL")

# Klasse für Kollisionen, die das automatische Lösen nicht schafft
KLASSE_MANUELL = u"MANUELL"

# Gruppierungen des Baums: Schlüssel -> Beschriftung
GRUPPIERUNGEN = (
    (u"status", (u"Status", u"Status", u"Estado")),
    (u"test", (u"Test", u"Test", u"Prueba")),
    (u"gruppe", (u"Navisworks-Gruppe", u"Navisworks group",
                 u"Grupo de Navisworks")),
    (u"ebene", (u"Ebene", u"Level", u"Nivel")),
    (u"klasse", (u"Klassifizierung", u"Classification", u"Clasificación")),
    (u"element", (u"Element (je Element)", u"Element (per element)",
                  u"Elemento (por elemento)")),
    (u"modelle", (u"Modellpaar", u"Model pair", u"Par de modelos")),
    (u"", (u"(keine)", u"(none)", u"(ninguno)")),
)

M_PRO_FUSS = 0.3048


def status_text(status):
    return tt(STATUS_TEXTE.get(status, (status, status, status)))


def status(clash, eintrag):
    return (eintrag or {}).get(u"status") or clash.status_navis


def klasse(eintrag):
    return (eintrag or {}).get(u"klasse") or u""


def ist_offen(clash, eintrag):
    return status(clash, eintrag) in OFFEN


def abstand_text(abstand):
    if abstand is None:
        return u""
    return u"%.1f cm" % (abstand * 100.0)


def objekt_text(objekt):
    teile = [objekt.beschreibung]
    if objekt.element_id is not None:
        teile.append(u"[%d]" % objekt.element_id)
    return u" ".join(teile)


def paar_text(clash):
    namen = [objekt.name or objekt.typ for objekt in clash.objekte]
    namen = [name for name in namen if name]
    return u" \u2194 ".join(namen)


def blatt_text(clash, eintrag):
    """Beschriftung eines Blattes im Baum."""
    teile = [clash.name]
    paar = paar_text(clash)
    if paar:
        teile.append(paar)
    if clash.abstand is not None:
        teile.append(abstand_text(clash.abstand))
    if klasse(eintrag):
        teile.append(u"[%s]" % klasse(eintrag))
    return u"  \u00b7  ".join(teile)


def ohne_wert():
    return t(u"(ohne)", u"(none)", u"(sin valor)")


def gruppen_werte(clash, eintrag, art):
    """Die Zweige, unter denen die Kollision steht - meist einer, bei
    'element' einer je beteiligtem Element."""
    if art == u"status":
        return [status(clash, eintrag)]
    if art == u"test":
        return [clash.test or ohne_wert()]
    if art == u"gruppe":
        return [clash.gruppe or ohne_wert()]
    if art == u"ebene":
        return [clash.ebene or ohne_wert()]
    if art == u"klasse":
        return [klasse(eintrag) or ohne_wert()]
    if art == u"element":
        werte = []
        for objekt in clash.objekte:
            wert = objekt_text(objekt)
            if objekt.datei:
                wert = u"%s: %s" % (modell_kurz(objekt.datei), wert)
            if wert not in werte:
                werte.append(wert)
        return werte or [ohne_wert()]
    if art == u"modelle":
        dateien = sorted(set(modell_kurz(o.datei) for o in clash.objekte
                             if o.datei))
        return [u" \u2194 ".join(dateien) or ohne_wert()]
    return [u""]


def zweig_text(art, wert):
    return status_text(wert) if art == u"status" else wert


def _sortierschluessel(art, wert):
    if art == u"status":
        try:
            return (0, br.STATUS.index(wert), u"")
        except ValueError:
            return (1, 0, wert)
    return (1 if wert == ohne_wert() else 0, 0, wert.lower())


def baum(clashes, eintraege, ebene1, ebene2=u""):
    """Gruppiert nach bis zu zwei Merkmalen.

    Rückgabe: [(wert, [clash ...] oder [(wert2, [clash ...]), ...]), ...]
    Ohne ebene1 ein einziger Zweig mit allen Kollisionen.
    """
    def gruppiere(liste, art):
        zweige = {}
        for clash in liste:
            for wert in gruppen_werte(clash, eintraege.get(clash.schluessel),
                                      art):
                zweige.setdefault(wert, []).append(clash)
        return sorted(zweige.items(),
                      key=lambda paar: _sortierschluessel(art, paar[0]))

    if not ebene1:
        return [(u"", list(clashes))]
    oben = gruppiere(clashes, ebene1)
    if not ebene2 or ebene2 == ebene1:
        return oben
    return [(wert, gruppiere(liste, ebene2)) for wert, liste in oben]


def passt_suche(clash, eintrag, suche):
    """Suchbegriff in Name, Test, Ebene, Elementen, IDs und Kommentaren."""
    suche = (suche or u"").strip().lower()
    if not suche:
        return True
    texte = [clash.name, clash.test, clash.gruppe, clash.raster,
             clash.beschreibung, klasse(eintrag),
             (eintrag or {}).get(u"kommentar") or u""]
    texte.extend(clash.kommentare_navis)
    for objekt in clash.objekte:
        texte.extend([objekt.name, objekt.typ, objekt.datei])
        if objekt.element_id is not None:
            texte.append(u"%d" % objekt.element_id)
    gesamt = u" | ".join(texte).lower()
    return all(wort in gesamt for wort in suche.split())


# ---------------------------------------------------------------------------
# Modelle: Dateien aus Navisworks den Revit-Modellen zuordnen
# ---------------------------------------------------------------------------

_ENDUNGEN = (u".nwc", u".nwd", u".nwf", u".rvt", u".ifc")


def modell_kurz(datei):
    """'C:\\...\\TGA_EG.rvt.nwc' -> 'TGA_EG'."""
    name = os.path.basename((datei or u"").replace(u"\\", u"/")).strip()
    geaendert = True
    while geaendert:
        geaendert = False
        for endung in _ENDUNGEN:
            if name.lower().endswith(endung):
                name = name[:-len(endung)]
                geaendert = True
    return name


def _vergleichsform(name):
    return re.sub(r"[^a-z0-9]", u"", modell_kurz(name).lower())


def modell_passt(datei, namen):
    """Gehört die Datei aus Navisworks zu einem der Modellnamen?

    Namen sind Titel und Dateinamen eines Revit-Modells. Eine lokale Kopie
    heisst 'Modell_benutzer' - deshalb genügt es, wenn einer mit dem anderen
    beginnt (ab 6 Zeichen, damit 'A' nicht auf alles passt).
    """
    gesucht = _vergleichsform(datei)
    if not gesucht:
        return False
    for name in namen:
        form = _vergleichsform(name)
        if not form:
            continue
        if form == gesucht:
            return True
        kurz, lang = sorted((form, gesucht), key=len)
        if len(kurz) >= 6 and lang.startswith(kurz):
            return True
    return False


def betrifft_modell(clash, namen):
    """Mindestens ein Element stammt aus dem Modell - oder die Herkunft ist
    unbekannt (dann lieber anzeigen)."""
    dateien = [objekt.datei for objekt in clash.objekte if objekt.datei]
    if not dateien:
        return True
    return any(modell_passt(datei, namen) for datei in dateien)


def filtere(clashes, eintraege, nur_offen=False, suche=u"",
            modellnamen=None):
    ergebnis = []
    for clash in clashes:
        eintrag = eintraege.get(clash.schluessel)
        if nur_offen and not ist_offen(clash, eintrag):
            continue
        if modellnamen is not None and not betrifft_modell(clash,
                                                            modellnamen):
            continue
        if not passt_suche(clash, eintrag, suche):
            continue
        ergebnis.append(clash)
    return ergebnis


# ---------------------------------------------------------------------------
# Kasten um die Kollision - alles in Metern, Wirtskoordinaten
# ---------------------------------------------------------------------------

def vereinigung(kaesten):
    klein = tuple(min(k[0][a] for k in kaesten) for a in range(3))
    gross = tuple(max(k[1][a] for k in kaesten) for a in range(3))
    return klein, gross


def schnittmenge(kasten_a, kasten_b):
    """Überlappung zweier Kästen - None, wenn sie sich nicht berühren."""
    klein = tuple(max(kasten_a[0][a], kasten_b[0][a]) for a in range(3))
    gross = tuple(min(kasten_a[1][a], kasten_b[1][a]) for a in range(3))
    if any(klein[a] > gross[a] + 1e-6 for a in range(3)):
        return None
    return klein, gross


def um_punkt(punkt, halb):
    return (tuple(w - halb for w in punkt), tuple(w + halb for w in punkt))


def enthaelt(kasten, punkt):
    return all(kasten[0][a] <= punkt[a] <= kasten[1][a] for a in range(3))


def abstand_zu_kasten(punkt, kasten):
    """Abstand eines Punktes zum Kasten (0 innerhalb)."""
    summe = 0.0
    for a in range(3):
        if punkt[a] < kasten[0][a]:
            summe += (kasten[0][a] - punkt[a]) ** 2
        elif punkt[a] > kasten[1][a]:
            summe += (punkt[a] - kasten[1][a]) ** 2
    return summe ** 0.5


def kasten_fuer(kaesten, punkt=None, rand=0.15, schnitt=True,
                mindestmass=0.5, umkreis=2.0):
    """Der Kasten für die Schnittbox.

    kaesten   Hüllkästen der gefundenen Elemente, je (min, max)
    punkt     Kollisionspunkt (Wirtskoordinaten) oder None
    schnitt   True: nur die Überlappung der zwei Elemente, sonst beide ganz
    umkreis   Ist nur ein Element gefunden (etwa ein 40 m langes Rohr),
              wird sein Kasten auf diesen Umkreis um den Punkt beschnitten.

    Rückgabe (min, max) oder None, wenn weder Element noch Punkt da sind.
    """
    kasten = None
    if len(kaesten) >= 2 and schnitt:
        kasten = schnittmenge(kaesten[0], kaesten[1])
    if kasten is None and kaesten:
        kasten = vereinigung(kaesten)
        if len(kaesten) == 1 and punkt is not None:
            kasten = schnittmenge(kasten, um_punkt(punkt, umkreis)) or kasten
    if punkt is not None:
        if kasten is None:
            kasten = um_punkt(punkt, 0.0)
        elif not enthaelt(kasten, punkt) and \
                abstand_zu_kasten(punkt, kasten) < umkreis:
            kasten = vereinigung([kasten, um_punkt(punkt, 0.0)])
    if kasten is None:
        return None

    klein = [kasten[0][a] - rand for a in range(3)]
    gross = [kasten[1][a] + rand for a in range(3)]
    for a in range(3):
        fehlt = mindestmass - (gross[a] - klein[a])
        if fehlt > 0:
            klein[a] -= fehlt / 2.0
            gross[a] += fehlt / 2.0
    return tuple(klein), tuple(gross)
