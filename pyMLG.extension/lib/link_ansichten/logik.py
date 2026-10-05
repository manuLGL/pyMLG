# -*- coding: utf-8 -*-
"""Revit-freie Regeln von LinkedViews - ohne Revit testbar
(tools/test_link_ansichten.py)."""

from mlg_sprache import t

# ViewType -> Gruppe. Revit bietet als verknüpfte Ansicht nur Ansichten an,
# die zur Art der Hauptansicht passen (im Grundriss z.B. auch Flächenpläne).
GRUPPEN = {
    "FloorPlan": "plan",
    "CeilingPlan": "plan",
    "EngineeringPlan": "plan",
    "AreaPlan": "plan",
    "Section": "schnitt",
    "Detail": "schnitt",
    "Elevation": "ansicht",
    "ThreeD": "3d",
}

GRUPPEN_REIHENFOLGE = ("plan", "schnitt", "ansicht", "3d")


def gruppen_namen():
    return {
        "plan": t(u"Grundrisse", u"Plans", u"Planos"),
        "schnitt": t(u"Schnitte", u"Sections", u"Secciones"),
        "ansicht": t(u"Ansichten", u"Elevations", u"Alzados"),
        "3d": t(u"3D-Ansichten", u"3D views", u"Vistas 3D"),
    }


def ansichtstypen():
    """ViewType -> Bezeichnung wie im Projektbrowser."""
    return {
        "FloorPlan": t(u"Grundriss", u"Floor Plan", u"Plano de planta"),
        "CeilingPlan": t(u"Deckenplan", u"Ceiling Plan", u"Plano de techo"),
        "EngineeringPlan": t(u"Tragwerksplan", u"Structural Plan",
                             u"Plano estructural"),
        "AreaPlan": t(u"Flächenplan", u"Area Plan", u"Plano de área"),
        "Section": t(u"Schnitt", u"Section", u"Sección"),
        "Elevation": t(u"Ansicht", u"Elevation", u"Alzado"),
        "Detail": t(u"Detail", u"Detail View", u"Vista de detalle"),
        "ThreeD": t(u"3D-Ansicht", u"3D View", u"Vista 3D"),
    }


def gruppe(ansichtstyp):
    """Gruppe eines ViewType (Text) oder None, wenn Revit dort keine
    verknüpfte Ansicht zulässt."""
    return GRUPPEN.get(u"%s" % ansichtstyp)


def passt(host_typ, link_typ):
    g = gruppe(host_typ)
    return g is not None and g == gruppe(link_typ)


def suchwoerter(text):
    return (text or u"").lower().split()


def trifft(suchtext, woerter):
    """Alle Wörter müssen vorkommen (wie im Filter-Manager)."""
    return all(w in suchtext for w in woerter)


def doppelte(paare):
    """paare: [(Schlüssel, Name)] -> Namen, die mehr als einmal vorkommen.
    Pro Verknüpfung lässt sich nur eine verknüpfte Ansicht setzen."""
    gesehen = {}
    doppelt = []
    for schluessel, name in paare:
        gesehen[schluessel] = gesehen.get(schluessel, 0) + 1
        if gesehen[schluessel] == 2:
            doppelt.append(name)
    return doppelt


def ziele(ansichten, steuernde_vorlage, schluessel):
    """Ansichten, deren RVT-Verknüpfungen eine Vorlage steuert, werden durch
    diese Vorlage ersetzt - jedes Ziel nur einmal.

    steuernde_vorlage(ansicht) -> Vorlage oder None
    schluessel(element)        -> eindeutiger Wert (Id)
    Rückgabe: (Ziele, {Vorlagen-Schlüssel: (Vorlage, [Ansichten])})
    """
    ergebnis = []
    gesehen = set()
    umgeleitet = {}
    for ansicht in ansichten:
        vorlage = steuernde_vorlage(ansicht)
        ziel = vorlage if vorlage is not None else ansicht
        if vorlage is not None:
            eintrag = umgeleitet.setdefault(schluessel(vorlage),
                                            (vorlage, []))
            eintrag[1].append(ansicht)
        k = schluessel(ziel)
        if k not in gesehen:
            gesehen.add(k)
            ergebnis.append(ziel)
    return ergebnis, umgeleitet


def kurzliste(namen, hoechstens=8):
    """'a, b, c' - bei mehr Namen mit '… und n weitere'."""
    namen = list(namen)
    text = u", ".join(namen[:hoechstens])
    if len(namen) > hoechstens:
        text += t(u" … und %d weitere", u" … and %d more",
                  u" … y %d más") % (len(namen) - hoechstens)
    return text
