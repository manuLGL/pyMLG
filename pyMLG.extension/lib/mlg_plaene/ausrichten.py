# -*- coding: utf-8 -*-
"""Ansichten wie auf einem Vorbild-Plan ausrichten.

Jedes Ziel-Ansichtsfenster bekommt sein Gegenstück auf dem Vorbild
(logik.fenster_zuordnen: gleiche Ansichtsart, Vorrang für dieselbe Ansicht,
dann gleicher Maßstab, dann die nächstliegende) und übernimmt Lage,
Titellage und auf Wunsch den Ansichtsfenster-Typ.

Zwei Arten:
    GEBAEUDE  Gebäude an gleicher Stelle - der Projektursprung liegt auf dem
              Plan wie beim Vorbild (Grundrisse EG bis DG übereinander);
              Legenden und Zeichenansichten nach der Mitte
    FENSTER   Ansichtsfenster an gleicher Stelle - die Mitte liegt gleich

Genutzt von "Ansichten ausrichten" und "Ansicht auf Plan". Muss in einer
offenen Transaktion laufen. Läuft unter IronPython.
"""

from Autodesk.Revit.DB import ViewType

from mlg_plaene import logik as lg
from mlg_plaene import revit as rv
from mlg_sprache import t

GEBAEUDE, FENSTER = "gebaeude", "fenster"

OHNE_MODELL = (ViewType.Legend, ViewType.DraftingView)


def art_texte():
    """[(Text, Art)] für Auswahlfelder, in der aktuellen Sprache."""
    return [
        (t(u"Gebäude an gleicher Stelle (Grundrisse, Schnitte)",
           u"Building at the same place (plans, sections)",
           u"Edificio en el mismo lugar (plantas, secciones)"), GEBAEUDE),
        (t(u"Ansichtsfenster an gleicher Stelle (Details, Legenden)",
           u"Viewport at the same place (details, legends)",
           u"Ventana en el mismo lugar (detalles, leyendas)"), FENSTER),
    ]


def _info(doc, fenster):
    ansicht = doc.GetElement(fenster.ViewId)
    mitte = fenster.GetBoxCenter()
    return lg.Fenster(rv.id_wert(fenster.Id), str(ansicht.ViewType), ansicht.Scale,
                      rv.id_wert(ansicht.Id), (mitte.X, mitte.Y))


def aussehen_uebernehmen(vorbild, fenster, typ=True):
    """Titellage und (mit typ) Ansichtsfenster-Typ wie beim Vorbild."""
    if typ and rv.id_wert(fenster.GetTypeId()) != rv.id_wert(vorbild.GetTypeId()):
        fenster.ChangeTypeId(vorbild.GetTypeId())
    try:
        fenster.LabelOffset = vorbild.LabelOffset
        fenster.LabelLineLength = vorbild.LabelLineLength
    except Exception:
        pass


def verschiebung(doc, vorbild, fenster, art):
    """(Verschiebung, Hinweis oder None) für fenster."""
    ansicht = doc.GetElement(fenster.ViewId)
    vorbild_ansicht = doc.GetElement(vorbild.ViewId)
    if art == GEBAEUDE and ansicht.ViewType not in OHNE_MODELL:
        soll = rv.modell_auf_plan(vorbild, vorbild_ansicht)
        ist = rv.modell_auf_plan(fenster, ansicht)
        if soll is not None and ist is not None:
            hinweis = None
            if ansicht.Scale != vorbild_ansicht.Scale:
                hinweis = t(u"'{}': anderer Maßstab (1:{}) - nur der Ursprung liegt gleich",
                            u"'{}': different scale (1:{}) - only the origin matches",
                            u"'{}': otra escala (1:{}): sólo coincide el origen").format(
                    ansicht.Name, ansicht.Scale)
            return soll - ist, hinweis
    return vorbild.GetBoxCenter() - fenster.GetBoxCenter(), None


def wie_vorbild(doc, vorbilder, ziele, art=GEBAEUDE, typ=True, melden="vorbilder"):
    """Richtet die Ziel-Ansichtsfenster wie ihre Gegenstücke aus.

    melden: "vorbilder" nennt Vorbild-Ansichten ohne Gegenstück (ganzer
    Plan wird übertragen), "ziele" nennt Ziel-Ansichten ohne Vorbild (neue
    Ansichten werden eingepasst). Liefert (Anzahl ausgerichtet, Hinweise).
    """
    nach_id = dict((rv.id_wert(f.Id), f) for f in list(vorbilder) + list(ziele))
    paare = lg.fenster_zuordnen([_info(doc, f) for f in vorbilder],
                                [_info(doc, f) for f in ziele])
    hinweise = []
    for vorbild_id, fenster_id in paare:
        vorbild, fenster = nach_id[vorbild_id], nach_id[fenster_id]
        aussehen_uebernehmen(vorbild, fenster, typ)
        doc.Regenerate()
        weg, hinweis = verschiebung(doc, vorbild, fenster, art)
        fenster.SetBoxCenter(fenster.GetBoxCenter() + weg)
        if hinweis:
            hinweise.append(hinweis)

    if melden == "vorbilder":
        zugeordnet = set(v for v, _ in paare)
        for vorbild in vorbilder:
            if rv.id_wert(vorbild.Id) not in zugeordnet:
                hinweise.append(t(u"kein Gegenstück für '{}'", u"no counterpart for '{}'",
                                  u"sin equivalente para '{}'").format(
                    doc.GetElement(vorbild.ViewId).Name))
    else:
        zugeordnet = set(z for _, z in paare)
        for fenster in ziele:
            if rv.id_wert(fenster.Id) not in zugeordnet:
                hinweise.append(t(u"'{}': keine passende Ansicht auf dem Vorbild - "
                                  u"automatische Lage",
                                  u"'{}': no matching view on the reference - "
                                  u"automatic position",
                                  u"'{}': ninguna vista adecuada en la referencia: "
                                  u"posición automática").format(
                    doc.GetElement(fenster.ViewId).Name))
    return len(paare), hinweise
