# -*- coding: utf-8 -*-
"""Revit-Zugriffe des Ansichtsvorlagen-Managers.

Eine Ansichtsvorlage ist eine View mit IsTemplate == True. Welche Parameter
sie steuern kann, sagt GetTemplateParameterIds(); welche davon nicht
angehakt sind ("Einschließen"), sagt GetNonControlledTemplateParameterIds().
Geschrieben wird der Haken mit SetNonControlledTemplateParameterIds().

Die Werte selbst sind gewöhnliche Parameter der View und werden mit
Parameter.Set() bzw. SetValueString() geschrieben. Drei Fälle sind nicht als
Wert lesbar und heißen hier "Block" - im nativen Dialog steht dort
"Bearbeiten…":

    * der Parameter existiert an der View gar nicht (z.B. die
      V/G-Überschreibungen VIS_GRAPHICS_*),
    * er ist schreibgeschützt,
    * seine Speicherart ist None.

Für Blöcke zeigt das Fenster eine kurze Zusammenfassung (Anzahl Filter,
ausgeblendete und überschriebene Kategorien). Bearbeitet werden sie in den
Editoren von vorlagen_manager.bloecke, soweit die API sie freigibt.

Wo sie das nicht tut (Modelldarstellung, Schatten, Skizzenlinien,
Beleuchtung, Fotobelichtung), bleibt nur das Übertragen aus einer anderen
Vorlage - das steht in vorlagen_manager.uebertragen.

Auswahllisten für Aufzählungsparameter: Die API nennt die möglichen Werte
nicht. Deshalb werden die in den Vorlagen des Projekts tatsächlich
vorkommenden Werte gesammelt - deren Beschriftung liefert Revit selbst über
AsValueString() in der Landessprache. Für die geläufigen Aufzählungen
(Detaillierungsgrad, Disziplin, Teilesichtbarkeit, Unterlage,
Modelldarstellung) kommen die fehlenden Werte aus AUFZAEHLUNGEN dazu.
"""

import clr

from Autodesk.Revit.DB import (
    BuiltInParameter,
    CategoryType,
    CopyPasteOptions,
    ElementId,
    ElementTransformUtils,
    FilteredElementCollector,
    LabelUtils,
    StorageType,
    Transform,
    View,
    ViewDetailLevel,
    ViewDuplicateOption,
    ViewSheet,
    ViewType,
)
from System import Enum

from filter_manager.aufloesung import id_liste, id_wert, ist_ja_nein
from mlg_sprache import t
from vorlagen_manager import logik as lg

# Art einer Parameterzeile - bestimmt das Bedienelement im Fenster
ART_TEXT = "text"
ART_ZAHL = "zahl"            # Double, geschrieben über SetValueString
ART_GANZZAHL = "ganzzahl"
ART_JANEIN = "janein"
ART_LISTE = "liste"          # Aufzählung, Auswahlfeld
ART_ELEMENT = "element"      # Verweis auf ein Element, Auswahlfeld
ART_BLOCK = "block"          # nicht als Wert lesbar ("Bearbeiten…")

# ViewType -> Bezeichnung wie im Projektbrowser
ANSICHTSTYPEN = {
    "FloorPlan": t(u"Grundriss", u"Floor Plan", u"Plano de planta"),
    "CeilingPlan": t(u"Deckenplan", u"Ceiling Plan", u"Plano de techo"),
    "EngineeringPlan": t(u"Tragwerksplan", u"Structural Plan",
                         u"Plano estructural"),
    "AreaPlan": t(u"Flächenplan", u"Area Plan", u"Plano de área"),
    "Section": t(u"Schnitt", u"Section", u"Sección"),
    "Elevation": t(u"Ansicht", u"Elevation", u"Alzado"),
    "Detail": t(u"Detail", u"Detail View", u"Vista de detalle"),
    "ThreeD": t(u"3D-Ansicht", u"3D View", u"Vista 3D"),
    "DraftingView": t(u"Zeichenansicht", u"Drafting View", u"Vista de diseño"),
    "Legend": t(u"Legende", u"Legend", u"Leyenda"),
    "Walkthrough": t(u"Walkthrough", u"Walkthrough", u"Recorrido"),
    "Rendering": t(u"Rendering", u"Rendering", u"Renderización"),
    "Schedule": t(u"Bauteilliste", u"Schedule", u"Tabla de planificación"),
}

# Ganzzahlparameter, die trotz sprechender Beschriftung frei eingegeben
# werden (Maßstab 1:x)
FREIE_ZAHLEN = ("VIEW_SCALE", "VIEW_SCALE_PULLDOWN_METRIC",
                "VIEW_SCALE_PULLDOWN_IMPERIAL")

# Kategorieblöcke: BuiltInParameter -> CategoryType für die Zusammenfassung
KATEGORIE_BLOECKE = {
    "VIS_GRAPHICS_MODEL": CategoryType.Model,
    "VIS_GRAPHICS_ANNOTATION": CategoryType.Annotation,
    "VIS_GRAPHICS_ANALYTICAL_MODEL": CategoryType.AnalyticalModel,
}


class VorlagenFehler(Exception):
    """Fachlicher Fehler - wird als Text angezeigt, nicht protokolliert."""


def fehlertext(fehler):
    text = getattr(fehler, "Message", None) or u"%s" % fehler
    zeilen = [z.strip() for z in text.splitlines() if z.strip()]
    return zeilen[0] if zeilen else fehler.__class__.__name__


# ---------------------------------------------------------------------------
# Aufzählungen, deren Werte die API nicht nennt
# ---------------------------------------------------------------------------

def _aufzaehlung(typ, beschriftungen):
    """[(Zahlenwert, Ersatzbeschriftung)] einer .NET-Aufzählung."""
    eintraege = []
    try:
        clr_typ = clr.GetClrType(typ)
        for name in Enum.GetNames(clr_typ):
            if name not in beschriftungen:
                continue
            eintraege.append((int(Enum.Parse(clr_typ, name)),
                              beschriftungen[name]))
    except Exception:
        return []
    return eintraege


def _baue_aufzaehlungen():
    """BuiltInParameter-Name -> [(Wert, Beschriftung)]."""
    tabelle = {}

    def eintragen(bip_name, typ, beschriftungen):
        werte = _aufzaehlung(typ, beschriftungen)
        if werte:
            tabelle[bip_name] = werte

    eintragen("VIEW_DETAIL_LEVEL", ViewDetailLevel, {
        "Coarse": t(u"Grob", u"Coarse", u"Bajo"),
        "Medium": t(u"Mittel", u"Medium", u"Medio"),
        "Fine": t(u"Fein", u"Fine", u"Alto")})
    try:
        from Autodesk.Revit.DB import ViewDiscipline
        eintragen("VIEW_DISCIPLINE", ViewDiscipline, {
            "Architectural": t(u"Architektur", u"Architectural",
                               u"Arquitectura"),
            "Structural": t(u"Tragwerk", u"Structural", u"Estructura"),
            "Mechanical": t(u"Gebäudetechnik", u"Mechanical", u"Mecánica"),
            "Electrical": t(u"Elektro", u"Electrical", u"Electricidad"),
            "Plumbing": t(u"Sanitär", u"Plumbing", u"Fontanería"),
            "Coordination": t(u"Koordination", u"Coordination",
                              u"Coordinación")})
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import PartsVisibility
        eintragen("VIEW_PARTS_VISIBILITY", PartsVisibility, {
            "ShowPartsOnly": t(u"Teile anzeigen", u"Show Parts",
                               u"Mostrar piezas"),
            "ShowOriginalOnly": t(u"Original anzeigen", u"Show Original",
                                  u"Mostrar original"),
            "ShowPartsAndOriginal": t(u"Beides anzeigen", u"Show Both",
                                      u"Mostrar ambos")})
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import UnderlayOrientation
        eintragen("VIEW_UNDERLAY_ORIENTATION", UnderlayOrientation, {
            "LookingDown": t(u"Nach unten schauen", u"Look down",
                             u"Mirar abajo"),
            "LookingUp": t(u"Nach oben schauen", u"Look up",
                           u"Mirar arriba")})
    except Exception:
        pass
    try:
        from Autodesk.Revit.DB import DisplayStyle
        eintragen("MODEL_GRAPHICS_STYLE", DisplayStyle, {
            "Wireframe": t(u"Drahtmodell", u"Wireframe",
                           u"Estructura alámbrica"),
            "HLR": t(u"Verdeckte Linie", u"Hidden Line", u"Línea oculta"),
            "Shading": t(u"Schattiert", u"Shaded", u"Sombreado"),
            "ShadingWithEdges": t(u"Schattiert mit Kanten",
                                  u"Shaded with Edges",
                                  u"Sombreado con bordes"),
            "FlatColors": t(u"Einheitliche Farben", u"Consistent Colors",
                            u"Colores coherentes"),
            "Realistic": t(u"Realistisch", u"Realistic", u"Realista"),
            "RealisticWithEdges": t(u"Realistisch mit Kanten",
                                    u"Realistic with Edges",
                                    u"Realista con bordes"),
            "Rendering": t(u"Rendering", u"Rendering", u"Renderización")})
    except Exception:
        pass
    return tabelle


AUFZAEHLUNGEN = _baue_aufzaehlungen()


# ---------------------------------------------------------------------------
# Kleine Helfer
# ---------------------------------------------------------------------------

def _bip(wert):
    """Zahlenwert einer Parameter-Id -> BuiltInParameter (oder None)."""
    if wert is None or wert >= 0:
        return None
    try:
        return Enum.ToObject(clr.GetClrType(BuiltInParameter), wert)
    except Exception:
        return None


def _bip_name(bip):
    if bip is None:
        return None
    try:
        return u"%s" % Enum.GetName(clr.GetClrType(BuiltInParameter), bip)
    except Exception:
        return None


def elementname(doc, element_id):
    """Name eines Elements, notfalls "Id 12345"."""
    wert = id_wert(element_id)
    if wert in (None, -1):
        return t(u"<keine>", u"<none>", u"<ninguno>")
    element = doc.GetElement(ElementId(wert))
    if element is None:
        return u"Id %s" % wert
    try:
        return u"%s" % element.Name
    except Exception:
        return u"Id %s" % wert


def ansichtstyp(ansicht):
    schluessel = u"%s" % ansicht.ViewType
    return ANSICHTSTYPEN.get(schluessel, schluessel)


def _janein_text(wert):
    return t(u"Ja", u"Yes", u"Sí") if wert else t(u"Nein", u"No", u"No")


def _ist_janein(parameter):
    try:
        return ist_ja_nein(parameter.Definition.GetDataType())
    except Exception:
        return False


def _hat_ueberschreibung(ogs):
    """True, wenn an dieser Kategorie irgendetwas überschrieben ist."""
    if ogs is None:
        return False
    try:
        if ogs.Halftone or ogs.Transparency:
            return True
        if ogs.ProjectionLineWeight > 0 or ogs.CutLineWeight > 0:
            return True
        for farbe in (ogs.ProjectionLineColor, ogs.CutLineColor,
                      ogs.SurfaceForegroundPatternColor,
                      ogs.SurfaceBackgroundPatternColor,
                      ogs.CutForegroundPatternColor,
                      ogs.CutBackgroundPatternColor):
            if farbe is not None and farbe.IsValid:
                return True
        for muster in (ogs.ProjectionLinePatternId, ogs.CutLinePatternId,
                       ogs.SurfaceForegroundPatternId,
                       ogs.SurfaceBackgroundPatternId,
                       ogs.CutForegroundPatternId,
                       ogs.CutBackgroundPatternId):
            if id_wert(muster) not in (None, -1):
                return True
        if ogs.DetailLevel != ViewDetailLevel.Undefined:
            return True
    except Exception:
        return False
    return False


# ---------------------------------------------------------------------------
# Eine Parameterzeile einer Vorlage
# ---------------------------------------------------------------------------

class Eintrag(object):
    """Ein Parameter einer Vorlage: Name, Art, Wert, Einschließen-Haken."""

    def __init__(self, pid, name, art, bip_name, parameter):
        self.pid = pid                 # Zahlenwert der Parameter-Id
        self.name = name
        self.art = art
        self.bip_name = bip_name       # Name des BuiltInParameter oder None
        self.parameter = parameter     # Parameter oder None (Block)
        self.schluessel = None         # vergleichbarer Rohwert
        self.text = u""                # angezeigter Wert
        self.zusatz = u""              # Klartext zum Wert (für den Tooltip)
        self.enthalten = True          # "Einschließen"-Haken
        gruppe, rang = lg.gruppe_und_rang(bip_name, pid < 0)
        self.gruppe = gruppe
        self.rang = rang

    def sortierung(self):
        return (self.gruppe, self.rang, lg.natuerlich(self.name))

    def suchtext(self):
        return (u"%s %s" % (self.name, self.text)).lower()


class Modell(object):
    """Liest und schreibt die Ansichtsvorlagen eines Dokuments."""

    def __init__(self, doc):
        self.doc = doc
        self._karten = {}        # Vorlagen-Id -> {pid: Parameter}
        self._optionen = {}      # pid -> [(Text, Wert)]
        self._kategorien = None

    # -- Vorlagen -------------------------------------------------------

    def vorlagen(self):
        """Alle Ansichtsvorlagen, nach Namen sortiert."""
        gefunden = [v for v in FilteredElementCollector(self.doc).OfClass(View)
                    if v.IsTemplate]
        gefunden.sort(key=lambda v: lg.natuerlich(v.Name))
        return gefunden

    def namen(self):
        return [v.Name for v in self.vorlagen()]

    def verwendung(self):
        """Vorlagen-Id (Zahl) -> Anzahl Ansichten, die sie zugewiesen haben."""
        zaehler = {}
        for ansicht in FilteredElementCollector(self.doc).OfClass(View):
            if ansicht.IsTemplate:
                continue
            try:
                wert = id_wert(ansicht.ViewTemplateId)
            except Exception:
                continue
            if wert not in (None, -1):
                zaehler[wert] = zaehler.get(wert, 0) + 1
        return zaehler

    def ansichten_mit(self, vorlage):
        """Ansichten, denen diese Vorlage zugewiesen ist."""
        ziel = id_wert(vorlage.Id)
        treffer = []
        for ansicht in FilteredElementCollector(self.doc).OfClass(View):
            if ansicht.IsTemplate:
                continue
            try:
                if id_wert(ansicht.ViewTemplateId) == ziel:
                    treffer.append(ansicht)
            except Exception:
                pass
        treffer.sort(key=lambda a: lg.natuerlich(a.Name))
        return treffer

    def zuweisbare_ansichten(self, vorlage):
        """Ansichten, die diese Vorlage annehmen (ohne Pläne und Vorlagen)."""
        treffer = []
        for ansicht in FilteredElementCollector(self.doc).OfClass(View):
            if ansicht.IsTemplate or isinstance(ansicht, ViewSheet):
                continue
            if ansicht.ViewType == ViewType.Internal:
                continue
            try:
                if ansicht.IsValidViewTemplate(vorlage.Id):
                    treffer.append(ansicht)
            except Exception:
                pass
        treffer.sort(key=lambda a: lg.natuerlich(a.Name))
        return treffer

    def vorlagenfaehige_ansichten(self):
        """Ansichten, aus denen sich eine Vorlage erzeugen lässt."""
        treffer = []
        for ansicht in FilteredElementCollector(self.doc).OfClass(View):
            if ansicht.IsTemplate or isinstance(ansicht, ViewSheet):
                continue
            if ansicht.ViewType == ViewType.Internal:
                continue
            treffer.append(ansicht)
        treffer.sort(key=lambda a: lg.natuerlich(a.Name))
        return treffer

    # -- Parameter ------------------------------------------------------

    def _karte(self, vorlage):
        """{pid: Parameter} einer Vorlage, zwischengespeichert."""
        schluessel = id_wert(vorlage.Id)
        karte = self._karten.get(schluessel)
        if karte is None:
            karte = {}
            for parameter in vorlage.Parameters:
                try:
                    karte[id_wert(parameter.Id)] = parameter
                except Exception:
                    pass
            self._karten[schluessel] = karte
        return karte

    def parameter(self, vorlage, pid):
        """Parameter einer Vorlage über den Zahlenwert seiner Id."""
        return self._karte(vorlage).get(pid)

    def vergiss(self, vorlage=None):
        """Zwischenspeicher leeren (nach Änderungen an einer Vorlage)."""
        if vorlage is None:
            self._karten = {}
            self._optionen = {}
        else:
            self._karten.pop(id_wert(vorlage.Id), None)

    def eintraege(self, vorlage, mit_bloecken=True):
        """{pid: Eintrag} aller Parameter, die diese Vorlage steuern kann.

        mit_bloecken=False lässt die Zusammenfassung der nicht lesbaren
        Einstellungen weg - die zählt Kategorien durch und lohnt sich nur
        für die gerade angezeigte Vorlage.
        """
        karte = self._karte(vorlage)
        try:
            nicht_enthalten = set(
                id_wert(i)
                for i in vorlage.GetNonControlledTemplateParameterIds())
        except Exception:
            nicht_enthalten = set()
        ergebnis = {}
        for param_id in vorlage.GetTemplateParameterIds():
            pid = id_wert(param_id)
            parameter = karte.get(pid)
            bip_name = _bip_name(_bip(pid))
            eintrag = Eintrag(pid, self._name(pid, parameter, bip_name),
                              self._art(parameter, bip_name), bip_name,
                              parameter)
            eintrag.enthalten = pid not in nicht_enthalten
            self._lies_wert(vorlage, eintrag, mit_bloecken)
            ergebnis[pid] = eintrag
        return ergebnis

    def _name(self, pid, parameter, bip_name):
        if parameter is not None:
            try:
                return u"%s" % parameter.Definition.Name
            except Exception:
                pass
        bip = _bip(pid)
        if bip is not None:
            try:
                return u"%s" % LabelUtils.GetLabelFor(bip)
            except Exception:
                return bip_name or (u"Id %s" % pid)
        element = self.doc.GetElement(ElementId(pid))
        if element is not None:
            try:
                return u"%s" % element.Name
            except Exception:
                pass
        return u"Id %s" % pid

    def _art(self, parameter, bip_name):
        if parameter is None:
            return ART_BLOCK
        try:
            if parameter.IsReadOnly:
                return ART_BLOCK
            speicherart = parameter.StorageType
        except Exception:
            return ART_BLOCK
        if speicherart == StorageType.String:
            return ART_TEXT
        if speicherart == StorageType.Double:
            return ART_ZAHL
        if speicherart == StorageType.ElementId:
            return ART_ELEMENT
        if speicherart != StorageType.Integer:
            return ART_BLOCK
        if _ist_janein(parameter):
            return ART_JANEIN
        if bip_name in FREIE_ZAHLEN:
            return ART_GANZZAHL
        if bip_name in AUFZAEHLUNGEN:
            return ART_LISTE
        # Beschriftet Revit den Wert sprechend, ist es eine Aufzählung
        try:
            beschriftung = parameter.AsValueString()
            if beschriftung and beschriftung.strip() != u"%d" % \
                    parameter.AsInteger():
                return ART_LISTE
        except Exception:
            pass
        return ART_GANZZAHL

    def _lies_wert(self, vorlage, eintrag, mit_bloecken=True):
        parameter = eintrag.parameter
        if eintrag.art == ART_BLOCK:
            eintrag.schluessel = None
            eintrag.text = (self.blocktext(vorlage, eintrag.bip_name)
                            if mit_bloecken
                            else t(u"Bearbeiten…", u"Edit…", u"Editar…"))
            return
        try:
            if eintrag.art == ART_TEXT:
                eintrag.schluessel = parameter.AsString() or u""
                eintrag.text = eintrag.schluessel
            elif eintrag.art == ART_ZAHL:
                eintrag.schluessel = round(parameter.AsDouble(), 9)
                eintrag.text = parameter.AsValueString() or u""
            elif eintrag.art == ART_ELEMENT:
                eintrag.schluessel = id_wert(parameter.AsElementId())
                eintrag.text = elementname(self.doc, parameter.AsElementId())
            else:
                eintrag.schluessel = parameter.AsInteger()
                if eintrag.art == ART_JANEIN:
                    eintrag.text = _janein_text(eintrag.schluessel)
                elif eintrag.art == ART_GANZZAHL:
                    # Die reine Zahl, denn genau die wird zurückgeschrieben.
                    # AsValueString() wäre beim Maßstab "1 : 100" - damit
                    # käme keine Eingabe durch int() zurück.
                    eintrag.text = u"%d" % eintrag.schluessel
                    eintrag.zusatz = parameter.AsValueString() or u""
                else:
                    eintrag.text = (parameter.AsValueString()
                                    or u"%d" % eintrag.schluessel)
        except Exception as fehler:
            eintrag.schluessel = None
            eintrag.text = fehlertext(fehler)

    # -- Auswahllisten --------------------------------------------------

    def optionen(self, eintrag, alle_vorlagen):
        """[(Text, Wert)] für Ja/Nein-, Aufzählungs- und Elementparameter."""
        vorhanden = self._optionen.get(eintrag.pid)
        if vorhanden is not None:
            return vorhanden
        if eintrag.art == ART_JANEIN:
            werte = [(_janein_text(1), 1), (_janein_text(0), 0)]
        elif eintrag.art == ART_LISTE:
            werte = self._listen_optionen(eintrag, alle_vorlagen)
        elif eintrag.art == ART_ELEMENT:
            werte = self._element_optionen(eintrag, alle_vorlagen)
        else:
            werte = []
        self._optionen[eintrag.pid] = werte
        return werte

    def _beobachtet(self, pid, alle_vorlagen):
        """{Zahlenwert: Beschriftung} aus allen Vorlagen des Projekts."""
        gesehen = {}
        for vorlage in alle_vorlagen:
            parameter = self._karte(vorlage).get(pid)
            if parameter is None:
                continue
            try:
                wert = parameter.AsInteger()
                gesehen.setdefault(wert, parameter.AsValueString()
                                   or u"%d" % wert)
            except Exception:
                pass
        return gesehen

    def _listen_optionen(self, eintrag, alle_vorlagen):
        gesehen = self._beobachtet(eintrag.pid, alle_vorlagen)
        werte = []
        for wert, ersatz in AUFZAEHLUNGEN.get(eintrag.bip_name, []):
            werte.append((gesehen.pop(wert, None) or ersatz, wert))
        for wert in sorted(gesehen):
            werte.append((gesehen[wert], wert))
        return werte

    def _element_optionen(self, eintrag, alle_vorlagen):
        werte = [(t(u"<keine>", u"<none>", u"<ninguno>"), -1)]
        gesehen = set()
        klasse = None
        try:
            aktuell = self.doc.GetElement(eintrag.parameter.AsElementId())
            if aktuell is not None:
                klasse = aktuell.GetType()
        except Exception:
            klasse = None
        if klasse is not None:
            try:
                for element in FilteredElementCollector(
                        self.doc).OfClass(klasse):
                    wert = id_wert(element.Id)
                    if wert not in gesehen:
                        gesehen.add(wert)
                        werte.append((elementname(self.doc, element.Id), wert))
            except Exception:
                pass
        if len(werte) == 1:
            # Klasse nicht sammelbar: wenigstens die vorkommenden Werte
            for vorlage in alle_vorlagen:
                parameter = self._karte(vorlage).get(eintrag.pid)
                if parameter is None:
                    continue
                try:
                    wert = id_wert(parameter.AsElementId())
                except Exception:
                    continue
                if wert not in gesehen and wert not in (None, -1):
                    gesehen.add(wert)
                    werte.append((elementname(self.doc, wert), wert))
        werte[1:] = sorted(werte[1:], key=lambda e: lg.natuerlich(e[0]))
        return werte

    # -- Blöcke ---------------------------------------------------------

    def _alle_kategorien(self):
        if self._kategorien is None:
            gesammelt = []
            for kategorie in self.doc.Settings.Categories:
                gesammelt.append(kategorie)
                try:
                    for unter in kategorie.SubCategories:
                        gesammelt.append(unter)
                except Exception:
                    pass
            self._kategorien = gesammelt
        return self._kategorien

    def blocktext(self, vorlage, bip_name):
        """Kurzfassung einer nicht lesbaren Einstellung ("Bearbeiten…")."""
        bearbeiten = t(u"Bearbeiten…", u"Edit…", u"Editar…")
        if bip_name == "VIS_GRAPHICS_FILTERS":
            try:
                anzahl = len(list(vorlage.GetFilters()))
            except Exception:
                return bearbeiten
            return t(u"%d Filter", u"%d filters", u"%d filtros") % anzahl
        typ = KATEGORIE_BLOECKE.get(bip_name)
        if typ is None:
            return bearbeiten
        versteckt, ueberschrieben = 0, 0
        for kategorie in self._alle_kategorien():
            try:
                if kategorie.CategoryType != typ:
                    continue
                if not kategorie.get_AllowsVisibilityControl(vorlage):
                    continue
                if vorlage.GetCategoryHidden(kategorie.Id):
                    versteckt += 1
                elif _hat_ueberschreibung(
                        vorlage.GetCategoryOverrides(kategorie.Id)):
                    ueberschrieben += 1
            except Exception:
                continue
        if not versteckt and not ueberschrieben:
            return t(u"ohne Änderung", u"unchanged", u"sin cambios")
        return t(u"%d ausgeblendet, %d überschrieben",
                 u"%d hidden, %d overridden",
                 u"%d ocultas, %d modificadas") % (versteckt, ueberschrieben)

    # -- Schreiben ------------------------------------------------------

    def setze_wert(self, eintrag, wert):
        """Wert einer Parameterzeile schreiben. Wirft VorlagenFehler."""
        parameter = eintrag.parameter
        if parameter is None or eintrag.art == ART_BLOCK:
            raise VorlagenFehler(
                t(u"\"%s\" lässt sich hier nicht setzen.",
                  u"\"%s\" cannot be set here.",
                  u"\"%s\" no se puede establecer aquí.") % eintrag.name)
        try:
            if eintrag.art == ART_TEXT:
                erfolg = parameter.Set(u"%s" % wert)
            elif eintrag.art == ART_ZAHL:
                erfolg = parameter.SetValueString(u"%s" % wert)
            elif eintrag.art == ART_ELEMENT:
                erfolg = parameter.Set(ElementId(int(wert)))
            else:
                erfolg = parameter.Set(int(wert))
        except Exception as fehler:
            raise VorlagenFehler(u"%s: %s" % (eintrag.name,
                                              fehlertext(fehler)))
        if erfolg is False:
            raise VorlagenFehler(
                t(u"Revit hat den Wert für \"%s\" nicht angenommen.",
                  u"Revit did not accept the value for \"%s\".",
                  u"Revit no aceptó el valor de \"%s\".") % eintrag.name)

    def setze_enthalten(self, vorlage, pids, enthalten):
        """"Einschließen"-Haken mehrerer Parameter einer Vorlage setzen."""
        nicht = set(id_wert(i)
                    for i in vorlage.GetNonControlledTemplateParameterIds())
        steuerbar = set(id_wert(i) for i in vorlage.GetTemplateParameterIds())
        for pid in pids:
            if pid not in steuerbar:
                continue
            if enthalten:
                nicht.discard(pid)
            else:
                nicht.add(pid)
        vorlage.SetNonControlledTemplateParameterIds(id_liste(sorted(nicht)))

    # -- Vorlagen verwalten ---------------------------------------------

    def aus_ansicht(self, ansicht, name):
        """Neue Vorlage aus einer Ansicht erzeugen."""
        try:
            vorlage = ansicht.CreateViewTemplate()
        except Exception as fehler:
            raise VorlagenFehler(
                t(u"Aus \"%s\" lässt sich keine Vorlage erzeugen: %s",
                  u"No template can be created from \"%s\": %s",
                  u"No se puede crear una plantilla a partir de \"%s\": %s")
                % (ansicht.Name, fehlertext(fehler)))
        try:
            vorlage.Name = name
        except Exception:
            pass
        return vorlage

    def dupliziere(self, vorlage, name):
        """Kopie einer Vorlage anlegen."""
        neue_id = None
        try:
            if vorlage.CanViewBeDuplicated(ViewDuplicateOption.Duplicate):
                neue_id = vorlage.Duplicate(ViewDuplicateOption.Duplicate)
        except Exception:
            neue_id = None
        if neue_id is None or id_wert(neue_id) == -1:
            try:
                kopien = list(ElementTransformUtils.CopyElements(
                    self.doc, id_liste([vorlage.Id]), self.doc,
                    Transform.Identity, CopyPasteOptions()))
                neue_id = kopien[0] if kopien else None
            except Exception as fehler:
                raise VorlagenFehler(
                    t(u"\"%s\" lässt sich nicht duplizieren: %s",
                      u"\"%s\" cannot be duplicated: %s",
                      u"\"%s\" no se puede duplicar: %s")
                    % (vorlage.Name, fehlertext(fehler)))
        kopie = self.doc.GetElement(neue_id) if neue_id is not None else None
        if kopie is None:
            raise VorlagenFehler(
                t(u"\"%s\" lässt sich nicht duplizieren.",
                  u"\"%s\" cannot be duplicated.",
                  u"\"%s\" no se puede duplicar.") % vorlage.Name)
        try:
            kopie.Name = name
        except Exception:
            pass
        return kopie

    # -- Ansichten ------------------------------------------------------

    def weise_zu(self, ansichten, vorlage):
        """Rückgabe: (Anzahl ok, [(Ansicht, Grund)])"""
        erfolgreich, fehler = 0, []
        for ansicht in ansichten:
            try:
                if not ansicht.IsValidViewTemplate(vorlage.Id):
                    fehler.append(
                        (ansicht.Name,
                         t(u"Vorlage passt nicht zum Ansichtstyp",
                           u"template does not fit the view type",
                           u"la plantilla no encaja con el tipo de vista")))
                    continue
                ansicht.ViewTemplateId = vorlage.Id
                erfolgreich += 1
            except Exception as ausnahme:
                fehler.append((ansicht.Name, fehlertext(ausnahme)))
        return erfolgreich, fehler

    def loese_zuweisung(self, ansichten):
        """Zuweisung der Vorlage aufheben. Rückgabe wie weise_zu()."""
        erfolgreich, fehler = 0, []
        for ansicht in ansichten:
            try:
                ansicht.ViewTemplateId = ElementId.InvalidElementId
                erfolgreich += 1
            except Exception as ausnahme:
                fehler.append((ansicht.Name, fehlertext(ausnahme)))
        return erfolgreich, fehler
