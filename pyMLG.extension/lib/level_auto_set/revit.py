# -*- coding: utf-8 -*-
"""Revit-Zugriffe von LevelAutoSet.

Jedes Element wird über "Anker" beschrieben: ein Ebenenparameter und die
Versatzparameter, die sich auf diese Ebene beziehen. Der erste passende
Anker aus BASIS bzw. OBEN (Parameter vorhanden und beschreibbar) gilt.

    Wände            Basisbeschränkung + Basisversatz
                     Obere Abhängigkeit + Versatz oben (nicht verbunden:
                     Höhe über Basis - bleibt beim Basiswechsel gleich)
    Stützen          Basisebene + Versatz, Abhängigkeit oben + Versatz
    Treppen          Basisebene + Versatz, obere Ebene + Versatz
    Dächer           Basisebene + Versatz / Referenzebene (Extrusion)
    Träger           Referenzebene + Versatz Anfang und Ende
    Decken, Böden    Ebene + Höhenversatz
    Familien         Ebene + Höhenversatz von Ebene

Eingefügte Elemente (Türen, Fenster) folgen ihrem Wirt; ihr Ebenenparameter
ist schreibgeschützt und sie werden übersprungen.

Kontrolle: Vor und nach dem Umsetzen wird die Bounding-Box verglichen.
Hat sich ein Element doch bewegt, wird es gemeldet.
"""

import clr

clr.AddReference("System")
from System.Collections.Generic import List  # noqa: E402

from Autodesk.Revit.DB import (  # noqa: E402
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    Level,
    StorageType,
    SubTransaction,
    Transaction,
)

from level_auto_set import logik as lg  # noqa: E402
from mlg_sprache import t  # noqa: E402

# Ab dieser Abweichung (Fuß, ca. 0,3 mm) gilt ein Element als verschoben
BEWEGT_TOLERANZ = 1e-3


def id_wert(element_id):
    """Zahlenwert einer ElementId (Revit 2024+: .Value)."""
    if element_id is None:
        return -1
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def net_liste(typ, werte):
    """Python-Iterable -> List[typ] (pythonnet nimmt keine Python-Liste im
    Konstruktor an)."""
    liste = List[typ]()
    for wert in werte:
        liste.Add(wert)
    return liste


def _bip(name):
    """BuiltInParameter per Name - fehlt er in dieser Revit-Version: None."""
    return getattr(BuiltInParameter, name, None)


class Anker(object):
    def __init__(self, ebene, versaetze, freie_hoehe=None, oberkante=False):
        self.ebene = ebene                  # BIP-Name des Ebenenparameters
        self.versaetze = versaetze          # BIP-Namen der Versätze
        self.freie_hoehe = freie_hoehe      # nur Wand oben: Höhe, wenn
        #                                     "nicht verbunden"
        # Bezugspunkt ist die Oberkante (Geschossdecke): folgt der
        # Einstellung "Oben", solange die nicht auf Ignorieren steht
        self.oberkante = oberkante


BASIS = (
    Anker("WALL_BASE_CONSTRAINT", ["WALL_BASE_OFFSET"]),
    Anker("FAMILY_BASE_LEVEL_PARAM", ["FAMILY_BASE_LEVEL_OFFSET_PARAM"]),
    Anker("STAIRS_BASE_LEVEL_PARAM", ["STAIRS_BASE_OFFSET"]),
    Anker("ROOF_BASE_LEVEL_PARAM", ["ROOF_LEVEL_OFFSET_PARAM"]),
    Anker("ROOF_CONSTRAINT_LEVEL_PARAM", ["ROOF_CONSTRAINT_OFFSET_PARAM"]),
    Anker("INSTANCE_REFERENCE_LEVEL_PARAM",
          ["STRUCTURAL_BEAM_END0_ELEVATION", "STRUCTURAL_BEAM_END1_ELEVATION"]),
    Anker("LEVEL_PARAM", ["FLOOR_HEIGHTABOVELEVEL_PARAM"], oberkante=True),
    Anker("LEVEL_PARAM", ["CEILING_HEIGHTABOVELEVEL_PARAM"]),
    Anker("FAMILY_LEVEL_PARAM", ["INSTANCE_ELEVATION_PARAM"]),
    Anker("FAMILY_LEVEL_PARAM", ["INSTANCE_FREE_HOST_OFFSET_PARAM"]),
)

OBEN = (
    Anker("WALL_HEIGHT_TYPE", ["WALL_TOP_OFFSET"], "WALL_USER_HEIGHT_PARAM"),
    Anker("FAMILY_TOP_LEVEL_PARAM", ["FAMILY_TOP_LEVEL_OFFSET_PARAM"]),
    Anker("STAIRS_TOP_LEVEL_PARAM", ["STAIRS_TOP_OFFSET"]),
)


# ---------------------------------------------------------------------------
# Ebenen
# ---------------------------------------------------------------------------

class EbenenInfo(object):
    def __init__(self, ebene, anzeige):
        self.ebene = ebene
        self.id = id_wert(ebene.Id)
        self.name = ebene.Name
        self.hoehe = ebene.Elevation          # intern, wie die Versätze
        self.anzeige = anzeige                # Höhe wie im Revit-Dialog
        self.geschoss = _ja(ebene, "LEVEL_IS_BUILDING_STORY")
        self.tragwerk = _ja(ebene, "LEVEL_IS_STRUCTURAL")


def _ja(element, bip_name):
    bip = _bip(bip_name)
    if bip is None:
        return True
    try:
        parameter = element.get_Parameter(bip)
        return parameter is None or parameter.AsInteger() == 1
    except Exception:
        return True


def ebenen(doc):
    """Alle Ebenen, von oben nach unten."""
    ergebnis = []
    for ebene in FilteredElementCollector(doc).OfClass(Level):
        anzeige = u""
        try:
            anzeige = ebene.get_Parameter(
                BuiltInParameter.LEVEL_ELEV).AsValueString() or u""
        except Exception:
            pass
        ergebnis.append(EbenenInfo(ebene, anzeige))
    return sorted(ergebnis, key=lambda e: -e.hoehe)


# ---------------------------------------------------------------------------
# Elemente lesen
# ---------------------------------------------------------------------------

def _parameter(element, bip_name, beschreibbar=True):
    bip = _bip(bip_name)
    if bip is None:
        return None
    try:
        parameter = element.get_Parameter(bip)
    except Exception:
        return None
    if parameter is None or (beschreibbar and parameter.IsReadOnly):
        return None
    return parameter


def _finde_anker(element, kandidaten):
    """(Anker, Ebenenparameter, [Versatzparameter]) oder None.

    Anker mit freier Höhe (Wand oben): Bei "Nicht verbunden" ist der Versatz
    oben in Revit schreibgeschützt - er wird erst nach dem Setzen der Ebene
    beschreibbar und darf deshalb hier schreibgeschützt sein."""
    for anker in kandidaten:
        ebene = _parameter(element, anker.ebene)
        if ebene is None or ebene.StorageType != StorageType.ElementId:
            continue
        beschreibbar = not anker.freie_hoehe
        versaetze = [_parameter(element, n, beschreibbar)
                     for n in anker.versaetze]
        if any(v is None or v.StorageType != StorageType.Double
               for v in versaetze):
            continue
        return anker, ebene, versaetze
    return None


class Seite(object):
    """Basis oder Oberkante eines Elements im Ausgangszustand."""

    def __init__(self, anker, ebene_param, versatz_params, ebene_hoehe,
                 verbunden, hoehe):
        self.anker = anker
        self.ebene_param = ebene_param
        self.versatz_params = versatz_params
        self.ebene_hoehe = ebene_hoehe        # None, wenn nicht verbunden
        self.versaetze = [p.AsDouble() for p in versatz_params]
        self.verbunden = verbunden
        self.hoehe = hoehe                    # absolute Bezugshöhe


def _ebene_hoehe(doc, parameter):
    ebene = doc.GetElement(parameter.AsElementId())
    return ebene.Elevation if isinstance(ebene, Level) else None


def lies_basis(doc, element):
    gefunden = _finde_anker(element, BASIS)
    if gefunden is None:
        return None
    anker, ebene_param, versatz_params = gefunden
    ebene_hoehe = _ebene_hoehe(doc, ebene_param)
    if ebene_hoehe is None:
        return None
    seite = Seite(anker, ebene_param, versatz_params, ebene_hoehe, True, 0.0)
    seite.hoehe = lg.bezugshoehe(ebene_hoehe, seite.versaetze)
    return seite


def lies_oben(doc, element, basis):
    gefunden = _finde_anker(element, OBEN)
    if gefunden is None:
        return None
    anker, ebene_param, versatz_params = gefunden
    ebene_hoehe = _ebene_hoehe(doc, ebene_param)
    if ebene_hoehe is not None:
        seite = Seite(anker, ebene_param, versatz_params, ebene_hoehe, True,
                      0.0)
        seite.hoehe = lg.bezugshoehe(ebene_hoehe, seite.versaetze)
        return seite
    # Nicht verbunden: Oberkante nur bestimmbar über eine freie Höhe
    freie = _parameter(element, anker.freie_hoehe, beschreibbar=False) \
        if anker.freie_hoehe else None
    if freie is None or basis is None:
        return Seite(anker, ebene_param, versatz_params, None, False, None)
    return Seite(anker, ebene_param, versatz_params, None, False,
                 basis.hoehe + freie.AsDouble())


def _ist_ja(element, bip_name):
    parameter = _parameter(element, bip_name, beschreibbar=False)
    try:
        return parameter is not None and parameter.AsInteger() == 1
    except Exception:
        return False


def profil_bearbeitet(element):
    try:
        return id_wert(getattr(element, "SketchId", None)) != -1
    except Exception:
        return False


def _box_z(element):
    try:
        box = element.get_BoundingBox(None)
        return (box.Min.Z, box.Max.Z) if box is not None else None
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Umsetzen
# ---------------------------------------------------------------------------

class Ergebnis(object):
    def __init__(self):
        self.basis = 0             # Basis-Ebene geändert
        self.oben = 0              # obere Ebene geändert
        self.unveraendert = []     # schon auf der passenden Ebene
        self.uebersprungen = []    # (Element, Grund)
        self.fehler = []           # (Element, Text)
        self.verschoben = []       # Element
        self.geaendert = []        # Element

    def ignorierte(self):
        return ([e for e, _g in self.uebersprungen]
                + [e for e, _t in self.fehler] + self.unveraendert)


def beschreibung(element):
    try:
        kategorie = element.Category.Name if element.Category else u"?"
    except Exception:
        kategorie = u"?"
    try:
        name = element.Name
    except Exception:
        name = u""
    return u"%s '%s' [%d]" % (kategorie, name, id_wert(element.Id))


def _set(parameter, wert):
    """Parameter.Set meldet viele Ablehnungen nur über den Rückgabewert."""
    if not parameter.Set(wert):
        raise ValueError(t(u"Revit lehnt '%s' ab", u"Revit rejects '%s'", u"Revit rechaza '%s'") % parameter.Definition.Name)


def _setze_seite(seite, ebene):
    """Ebene tauschen und Versätze so umrechnen, dass die Höhe bleibt."""
    neue = [lg.neuer_versatz(seite.ebene_hoehe, v, ebene.hoehe)
            for v in seite.versaetze]
    _set(seite.ebene_param, ebene.ebene.Id)
    for parameter, wert in zip(seite.versatz_params, neue):
        _set(parameter, wert)


def _verbinde_oben(element, seite, ebene):
    """Nicht verbundene Oberkante an eine Ebene hängen. Der Versatz wird
    erst danach beschreibbar - deshalb hier neu holen."""
    _set(seite.ebene_param, ebene.ebene.Id)
    for name in seite.anker.versaetze:
        parameter = _parameter(element, name)
        if parameter is None:
            raise ValueError(t(u"Versatz oben ist schreibgeschützt",
                               u"Top offset is read-only",
                               u"El desfase superior es de solo lectura"))
        _set(parameter, seite.hoehe - ebene.hoehe)


def setze_ebenen(doc, elemente, ebenen_liste, modus_basis, modus_oben,
                 oben_unverbunden_ignorieren=False, profil_ignorieren=False,
                 angehaengte_stuetzen_ignorieren=False,
                 fehler_ignorieren=False):
    """Elemente auf Ebenen setzen. Eine Transaktion; jedes Element in einer
    Untertransaktion, damit ein Fehler nur dieses Element zurücknimmt.

    Ohne fehler_ignorieren bricht der erste Fehler alles ab (RuntimeError).
    """
    ergebnis = Ergebnis()
    auswahl = [(e.id, e.hoehe) for e in ebenen_liste]
    nach_id = dict((e.id, e) for e in ebenen_liste)
    boxen = {}

    transaktion = Transaction(doc, t(u"pyMLG Ebenen setzen", u"pyMLG Set levels", u"pyMLG Asignar niveles"))
    transaktion.Start()
    try:
        for element in elemente:
            grund = _pruefe(element, profil_ignorieren,
                            angehaengte_stuetzen_ignorieren)
            if grund:
                ergebnis.uebersprungen.append((element, grund))
                continue
            basis = lies_basis(doc, element)
            if basis is None:
                ergebnis.uebersprungen.append(
                    (element, t(u"keine änderbare Ebene", u"no editable level", u"sin nivel editable")))
                continue
            oben = lies_oben(doc, element, basis)

            modus = lg.modus_fuer_basis(basis.anker.oberkante, modus_basis,
                                        modus_oben)
            hinweise = []
            plan_basis = None
            if _ist_ja(element, "WALL_BOTTOM_IS_ATTACHED"):
                if modus != lg.IGNORIEREN:
                    hinweise.append(t(u"Unterkante angehängt", u"base attached", u"base enlazada"))
            else:
                ziel = lg.waehle_ebene(basis.hoehe, auswahl, modus)
                plan_basis = nach_id.get(ziel)
                if plan_basis is None and modus != lg.IGNORIEREN:
                    hinweise.append(
                        t(u"keine passende Ebene für die Oberkante", u"no suitable level for the top", u"ningún nivel adecuado para la parte superior")
                        if basis.anker.oberkante
                        else t(u"keine passende Ebene für die Basis", u"no suitable level for the base", u"ningún nivel adecuado para la base"))
            plan_oben = None
            if oben is not None and oben.hoehe is not None \
                    and modus_oben != lg.IGNORIEREN \
                    and not (oben_unverbunden_ignorieren
                             and not oben.verbunden):
                if _ist_ja(element, "WALL_TOP_IS_ATTACHED"):
                    hinweise.append(t(u"Oberkante angehängt", u"top attached", u"parte superior enlazada"))
                else:
                    ziel = lg.waehle_ebene(oben.hoehe, auswahl, modus_oben)
                    plan_oben = nach_id.get(ziel)
                    if plan_oben is None:
                        hinweise.append(
                            t(u"keine passende Ebene für die Oberkante", u"no suitable level for the top", u"ningún nivel adecuado para la parte superior"))

            aendern_basis = plan_basis is not None and (
                id_wert(basis.ebene_param.AsElementId()) != plan_basis.id)
            aendern_oben = plan_oben is not None and (
                not oben.verbunden
                or id_wert(oben.ebene_param.AsElementId()) != plan_oben.id)
            if not aendern_basis and not aendern_oben:
                if hinweise:
                    ergebnis.uebersprungen.append(
                        (element, u", ".join(hinweise)))
                else:
                    ergebnis.unveraendert.append(element)
                continue

            boxen[id_wert(element.Id)] = _box_z(element)
            schritte = []
            if aendern_basis:
                schritte.append(lambda: _setze_seite(basis, plan_basis))
            if aendern_oben:
                schritte.append(
                    (lambda: _setze_seite(oben, plan_oben)) if oben.verbunden
                    else (lambda: _verbinde_oben(element, oben, plan_oben)))
            # Revit lehnt z.B. eine Stützenbasis über der aktuellen oberen
            # Ebene ab - dann zuerst die Oberkante setzen
            text = _ausfuehren(doc, schritte) \
                and _ausfuehren(doc, list(reversed(schritte)))
            if text:
                if not fehler_ignorieren:
                    raise RuntimeError(u"%s: %s" % (beschreibung(element),
                                                    text))
                ergebnis.fehler.append((element, text))
                continue
            ergebnis.basis += 1 if aendern_basis else 0
            ergebnis.oben += 1 if aendern_oben else 0
            ergebnis.geaendert.append(element)

        # Kontrolle: Hat sich etwas bewegt?
        doc.Regenerate()
        for element in ergebnis.geaendert:
            vorher = boxen.get(id_wert(element.Id))
            nachher = _box_z(element)
            if vorher and nachher and any(
                    abs(a - b) > BEWEGT_TOLERANZ
                    for a, b in zip(vorher, nachher)):
                ergebnis.verschoben.append(element)
        transaktion.Commit()
    except Exception:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        raise
    return ergebnis


def _ausfuehren(doc, schritte):
    """Schritte in einer Untertransaktion. Rückgabe: None bei Erfolg, sonst
    der Fehlertext (Untertransaktion zurückgenommen)."""
    st = SubTransaction(doc)
    st.Start()
    try:
        for schritt in schritte:
            schritt()
        st.Commit()
        return None
    except Exception as ausnahme:
        if st.HasStarted() and not st.HasEnded():
            st.RollBack()
        text = getattr(ausnahme, "Message", None) or u"%s" % ausnahme
        return (text.strip().splitlines() or [t(u"Fehler", u"Error", u"Error")])[0]


def _pruefe(element, profil_ignorieren, angehaengte_stuetzen_ignorieren):
    """Grund zum Überspringen oder None."""
    if profil_ignorieren and _parameter(element, "WALL_BASE_CONSTRAINT") \
            is not None and profil_bearbeitet(element):
        return t(u"Wand mit bearbeitetem Profil", u"wall with edited profile", u"muro con perfil editado")
    if angehaengte_stuetzen_ignorieren and (
            _ist_ja(element, "COLUMN_TOP_ATTACHED_PARAM")
            or _ist_ja(element, "COLUMN_BASE_ATTACHED_PARAM")):
        return t(u"angehängte Stütze", u"attached column", u"pilar enlazado")
    return None


def waehle_aus(uidoc, elemente):
    uidoc.Selection.SetElementIds(net_liste(ElementId,
                                            [e.Id for e in elemente]))
