# -*- coding: utf-8 -*-
# Textformat der Schnittbox: erzeugen und lesen. Ohne Revit - offline testbar.
#
# Der Text geht in die Zwischenablage und ist absichtlich lesbar, damit er
# auch per Mail oder Chat weitergereicht werden kann:
#
#     PYMLG-SECTIONBOX v1
#     # Koordination.rvt | {3D} | 2026-09-23 14:12
#     # Einheit: Meter | SHARED = gemeinsame Koordinaten, INTERN = intern
#     SHARED-O  12.345678 6.789012 0.000000
#     SHARED-X  1.000000 0.000000 0.000000
#     ...
#     MIN       -5.000000 -3.000000 0.000000
#     MAX       5.000000 3.000000 4.000000
#
# O/X/Y/Z beschreiben die Lage des Kastens (Ursprung und seine drei Achsen),
# MIN/MAX seine Ausdehnung in diesem Kastensystem. SHARED wird beim Einfuegen
# bevorzugt: nur damit landet die Box in einem anderen Projekt an derselben
# Stelle. INTERN ist der Rueckfall, wenn beide Projekte dieselbe interne
# Nullpunktlage haben, aber keine gemeinsamen Koordinaten.

from mlg_sprache import t

KOPF = u"PYMLG-SECTIONBOX v1"
M_PRO_FUSS = 0.3048

LAGE = (u"O", u"X", u"Y", u"Z")
SCHLUESSEL = tuple([u"SHARED-" + n for n in LAGE]
                   + [u"INTERN-" + n for n in LAGE]
                   + [u"MIN", u"MAX"])

# Kleinste sinnvolle Kantenlaenge in Metern - darunter ist die Box unbrauchbar
MINDESTMASS = 0.001


class Fehler(Exception):
    """Der Text ist keine (brauchbare) Schnittbox."""


class Box(object):
    """Eine Schnittbox in Metern.

    gemeinsam/intern  je (Ursprung, AchseX, AchseY, AchseZ), jeweils (x, y, z);
                      die Achsen sind Richtungen ohne Einheit
    min/max           Ausdehnung im Kastensystem
    quelle            Herkunftszeile fuer den Kommentar im Text
    """

    def __init__(self, gemeinsam, intern=None, min_punkt=None, max_punkt=None,
                 quelle=u""):
        self.gemeinsam = gemeinsam
        self.intern = intern or gemeinsam
        self.min = min_punkt
        self.max = max_punkt
        self.quelle = quelle

    @property
    def masse(self):
        """Kantenlaengen (Breite, Tiefe, Hoehe) in Metern."""
        return tuple(self.max[i] - self.min[i] for i in range(3))

    def masse_text(self, nachkomma=2):
        muster = u" \u00d7 ".join([u"%%.%df" % nachkomma] * 3) + u" m"
        return muster % self.masse


def _zahl(wert, nachkomma=6):
    return u"%.*f" % (nachkomma, wert)


def _zeile(schluessel, punkt, nachkomma=6):
    return u"%-9s %s" % (schluessel,
                         u" ".join(_zahl(w, nachkomma) for w in punkt))


def erzeuge_text(box):
    """Die Box als Text fuer die Zwischenablage."""
    zeilen = [KOPF]
    if box.quelle:
        zeilen.append(u"# " + box.quelle)
    zeilen.append(u"# " + t(
        u"Einheit: Meter | SHARED = gemeinsame Koordinaten, "
        u"INTERN = Projektkoordinaten",
        u"Unit: metres | SHARED = shared coordinates, "
        u"INTERN = project coordinates",
        u"Unidad: metros | SHARED = coordenadas compartidas, "
        u"INTERN = coordenadas del proyecto"))
    for vorsatz, lage in ((u"SHARED-", box.gemeinsam),
                          (u"INTERN-", box.intern)):
        # Ursprung in Metern, die drei Achsen als feine Richtungen
        zeilen.append(_zeile(vorsatz + LAGE[0], lage[0]))
        for name, achse in zip(LAGE[1:], lage[1:]):
            zeilen.append(_zeile(vorsatz + name, achse, 9))
    zeilen.append(_zeile(u"MIN", box.min))
    zeilen.append(_zeile(u"MAX", box.max))
    return u"\r\n".join(zeilen)


def _saubere_zeile(zeile):
    """Zitatzeichen aus Mails ('> > SHARED-O ...') und Rand entfernen."""
    zeile = zeile.strip()
    while zeile[:1] == u">":
        zeile = zeile[1:].strip()
    return zeile


def _werte(text):
    gefunden = {}
    for rohzeile in (text or u"").splitlines():
        zeile = _saubere_zeile(rohzeile)
        if not zeile or zeile[:1] == u"#":
            continue
        teile = zeile.replace(u"=", u" ").split()
        schluessel = teile[0].upper().rstrip(u":")
        if schluessel not in SCHLUESSEL or len(teile) < 4:
            continue
        try:
            gefunden[schluessel] = tuple(float(w) for w in teile[1:4])
        except ValueError:
            continue
    return gefunden


def _lage(werte, vorsatz):
    punkte = []
    for name in LAGE:
        punkt = werte.get(vorsatz + u"-" + name)
        if punkt is None:
            return None
        punkte.append(punkt)
    return tuple(punkte)


def _laenge(vektor):
    return (vektor[0] ** 2 + vektor[1] ** 2 + vektor[2] ** 2) ** 0.5


def _pruefe_achsen(lage):
    for achse in lage[1:]:
        if abs(_laenge(achse) - 1.0) > 0.01:
            raise Fehler(t(u"Die Achsen der Schnittbox sind fehlerhaft.",
                           u"The axes of the section box are invalid.",
                           u"Los ejes de la caja de sección no son válidos."))


def lies_text(text):
    """Box aus dem Text der Zwischenablage. Wirft Fehler mit Klartext."""
    werte = _werte(text)
    gemeinsam = _lage(werte, u"SHARED")
    if gemeinsam is None or u"MIN" not in werte or u"MAX" not in werte:
        raise Fehler(t(
            u"Die Zwischenablage enthält keine pyMLG-Schnittbox.\n"
            u"Erwartet wird der Text von \"Schnittbox kopieren\".",
            u"The clipboard does not contain a pyMLG section box.\n"
            u"Expected the text produced by \"Copy section box\".",
            u"El portapapeles no contiene una caja de sección de pyMLG.\n"
            u"Se espera el texto de \"Copiar caja de sección\"."))
    intern = _lage(werte, u"INTERN") or gemeinsam
    _pruefe_achsen(gemeinsam)
    _pruefe_achsen(intern)

    box = Box(gemeinsam, intern, werte[u"MIN"], werte[u"MAX"],
              quelle=quelle_aus_text(text))
    if min(box.masse) < MINDESTMASS:
        raise Fehler(t(u"Die Schnittbox aus der Zwischenablage ist leer.",
                       u"The section box from the clipboard is empty.",
                       u"La caja de sección del portapapeles está vacía."))
    return box


def quelle_aus_text(text):
    """Die erste Kommentarzeile nach dem Kopf - Herkunft der Box."""
    zeilen = [_saubere_zeile(z) for z in (text or u"").splitlines()]
    for nummer, zeile in enumerate(zeilen):
        if zeile.upper().startswith(u"PYMLG-SECTIONBOX"):
            naechste = zeilen[nummer + 1] if nummer + 1 < len(zeilen) else u""
            if naechste[:1] == u"#":
                return naechste[1:].strip()
            return u""
    return u""


def huelle(punkte, rand=0.0, mindestmass=0.0):
    """Achsparalleler Kasten um alle Punkte, ringsum um 'rand' erweitert.

    Punkte und Rand in Metern. Bleibt eine Kante unter 'mindestmass' - etwa
    bei einer ebenen Flaeche -, wird sie um ihre Mitte aufgezogen.
    """
    if not punkte:
        return None
    klein = [min(p[achse] for p in punkte) - rand for achse in range(3)]
    gross = [max(p[achse] for p in punkte) + rand for achse in range(3)]
    for achse in range(3):
        fehlt = mindestmass - (gross[achse] - klein[achse])
        if fehlt > 0:
            klein[achse] -= fehlt / 2.0
            gross[achse] += fehlt / 2.0
    return tuple(klein), tuple(gross)
