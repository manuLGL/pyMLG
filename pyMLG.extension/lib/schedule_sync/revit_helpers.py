# -*- coding: utf-8 -*-
"""Dünne Hilfsschicht über der Revit-API.

Kapselt alles, was Export und Import gemeinsam brauchen: Parameter finden,
Werte lesen und schreiben, Einheiten umrechnen, Worksharing-Status prüfen.
Bewusst frei von UI- und Excel-Code, damit beide Skripte dieselbe Logik nutzen.
"""

from Autodesk.Revit.DB import (
    Element,
    ElementId,
    FilteredElementCollector,
    SpecTypeId,
    StorageType,
    UnitUtils,
)

from mlg_sprache import t

try:
    # Nur in Zentraldatei-/Worksharing-Projekten relevant
    from Autodesk.Revit.DB import CheckoutStatus, WorksharingUtils
except ImportError:  # pragma: no cover - in Revit 2026 nicht zu erwarten
    CheckoutStatus = None
    WorksharingUtils = None


# Toleranz für den Vergleich von Fliesskommawerten (in internen Einheiten, i. d. R. Fuss).
# Grosszügig gewählt, weil der Excel-Wert auf NACHKOMMASTELLEN gerundet exportiert wird
# und beim Rückimport nie bitgenau dem Originalwert entspricht.
TOLERANZ = 1e-7

# Rundung der in Excel geschriebenen Zahlenwerte (in Projekteinheiten).
NACHKOMMASTELLEN = 6

# Texte, die beim Import als Ja/Nein interpretiert werden.
JA_WERTE = {"ja", "yes", u"sí", "si", "wahr", "true", "verdadero", "x",
            "1", "1.0"}
NEIN_WERTE = {"nein", "no", "falsch", "false", "falso", "", "0", "0.0"}


# ---------------------------------------------------------------------------
# Elementare Helfer
# ---------------------------------------------------------------------------

def eid_wert(element_id):
    """Zahlenwert einer ElementId - API-versionsunabhängig.

    Revit 2024+ nutzt ElementId.Value (Int64), ältere Versionen IntegerValue.
    """
    if element_id is None:
        return None
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def ist_gueltige_id(element_id):
    """True, wenn die ElementId auf ein echtes Element zeigt."""
    wert = eid_wert(element_id)
    return wert is not None and wert > 0


def kategorie_name(element):
    """Kategoriename eines Elements, leer wenn keine Kategorie vorhanden."""
    try:
        kategorie = element.Category
        return kategorie.Name if kategorie is not None else u""
    except Exception:
        return u""


def element_name(element):
    """Robuster Elementname (manche Elemente werfen beim Zugriff auf Name)."""
    if element is None:
        return u""
    try:
        return Element.Name.__get__(element)
    except Exception:
        try:
            return element.Name
        except Exception:
            return u""


# ---------------------------------------------------------------------------
# Parameter finden
# ---------------------------------------------------------------------------

def parameter_map(element):
    """Nachschlagetabelle aller Parameter eines Elements.

    Schlüssel:
        ("id",   <int>)  - ElementId des Parameters. Negative Werte entsprechen
                           exakt dem BuiltInParameter-Enumwert, positive Werte
                           einem Projekt-/Shared-Parameter dieses Dokuments.
        ("guid", <str>)  - GUID eines Shared Parameters (dokumentübergreifend stabil).
        ("name", <str>)  - Parametername als letzter Ausweg.

    Der Weg über die Map vermeidet Enum-Konvertierungen (System.Enum.ToObject),
    die unter pythonnet fehleranfällig sind, und ist pro Element nur einmal nötig.
    """
    eintraege = {}
    if element is None:
        return eintraege
    try:
        parameter = list(element.Parameters)
    except Exception:
        return eintraege

    for param in parameter:
        try:
            eintraege.setdefault(("id", eid_wert(param.Id)), param)
            if param.IsShared and param.GUID is not None:
                eintraege.setdefault(("guid", str(param.GUID)), param)
            definition = param.Definition
            if definition is not None:
                eintraege.setdefault(("name", definition.Name), param)
        except Exception:
            # Einzelne defekte Parameter dürfen die Map nicht sprengen
            continue
    return eintraege


def _aus_map(map_, feld):
    """Sucht einen Parameter anhand der beim Export gespeicherten Kennung."""
    if not map_:
        return None
    guid = feld.get("guid")
    if guid:
        treffer = map_.get(("guid", guid))
        if treffer is not None:
            return treffer
    param_id = feld.get("param_id")
    if param_id is not None:
        treffer = map_.get(("id", param_id))
        if treffer is not None:
            return treffer
    # Fallback über den Namen (z. B. wenn die Metadaten fehlen). Zuerst der echte
    # Parametername, dann die Überschrift - die kann in Revit umbenannt sein.
    for name in (feld.get("feldname"), feld.get("name")):
        if name:
            treffer = map_.get(("name", name))
            if treffer is not None:
                return treffer
    return None


class ParameterTreffer(object):
    """Ergebnis der Parametersuche: Parameter plus Element, an dem er hängt."""

    def __init__(self, parameter, besitzer, ist_typparameter):
        self.parameter = parameter
        self.besitzer = besitzer
        self.ist_typparameter = ist_typparameter


def typ_von(doc, element):
    """Typ-Element (ElementType) einer Instanz, sonst None."""
    try:
        typ_id = element.GetTypeId()
    except Exception:
        return None
    if not ist_gueltige_id(typ_id):
        return None
    try:
        return doc.GetElement(typ_id)
    except Exception:
        return None


def finde_parameter(doc, element, feld, instanz_map=None, typ_element=None,
                    typ_map=None):
    """Sucht den zum Schedule-Feld gehörenden Parameter an Element oder Typ.

    instanz_map/typ_map können vorberechnet übergeben werden, damit die
    Parameter-Map pro Element nur einmal aufgebaut werden muss.
    Rückgabe: ParameterTreffer oder None.
    """
    if element is None:
        return None

    if instanz_map is None:
        instanz_map = parameter_map(element)
    if typ_element is None:
        typ_element = typ_von(doc, element)
    if typ_map is None:
        typ_map = parameter_map(typ_element) if typ_element is not None else {}

    # Typparameter-Felder zuerst am Typ suchen, alle anderen zuerst an der Instanz
    reihenfolge = [(typ_map, typ_element, True), (instanz_map, element, False)]
    if not feld.get("ist_typfeld"):
        reihenfolge.reverse()

    for map_, besitzer, ist_typ in reihenfolge:
        if besitzer is None:
            continue
        param = _aus_map(map_, feld)
        if param is not None:
            return ParameterTreffer(param, besitzer, ist_typ)
    return None


# ---------------------------------------------------------------------------
# Einheiten und Datentypen
# ---------------------------------------------------------------------------

def _spec_id(param):
    """ForgeTypeId der Spezifikation eines Parameters (oder None)."""
    try:
        return param.Definition.GetDataType()
    except Exception:
        return None


def einheit_von(param):
    """Anzeigeeinheit (ForgeTypeId) eines Double-Parameters, sonst None.

    Revit speichert Längen, Flächen, Volumen usw. intern in Fuss-basierten
    Einheiten. Für Excel wird in die Projekteinheit umgerechnet und beim Import
    zurück - ohne das wären alle Zahlen um den Umrechnungsfaktor falsch.
    """
    spec = _spec_id(param)
    if spec is None:
        return None
    try:
        if not UnitUtils.IsMeasurableSpec(spec):
            return None
        return param.GetUnitTypeId()
    except Exception:
        return None


def ist_janein(param):
    """True für Ja/Nein-Parameter (intern Integer 0/1)."""
    spec = _spec_id(param)
    if spec is None:
        return False
    try:
        return spec.TypeId == SpecTypeId.Boolean.YesNo.TypeId
    except Exception:
        return False


def nach_projekteinheit(param, roh):
    """Interner Wert -> Projekteinheit (für die Anzeige in Excel)."""
    einheit = einheit_von(param)
    if einheit is None:
        return roh
    try:
        return UnitUtils.ConvertFromInternalUnits(roh, einheit)
    except Exception:
        return roh


def nach_interner_einheit(param, wert):
    """Projekteinheit -> interner Wert (für Parameter.Set beim Import)."""
    einheit = einheit_von(param)
    if einheit is None:
        return wert
    try:
        return UnitUtils.ConvertToInternalUnits(wert, einheit)
    except Exception:
        return wert


def einheit_kuerzel(param):
    """Lesbares Einheitenkürzel für die Excel-Kopfzeile (z. B. 'mm', 'm²')."""
    einheit = einheit_von(param)
    if einheit is None:
        return u""
    try:
        from Autodesk.Revit.DB import LabelUtils
        bezeichnung = LabelUtils.GetLabelForUnit(einheit)
        if bezeichnung:
            return bezeichnung
    except Exception:
        pass
    try:
        # Fallback: letzter Abschnitt der ForgeTypeId,
        # z. B. 'autodesk.unit.unit:millimeters-1.0.0' -> 'millimeters'
        return einheit.TypeId.split(":")[-1].split("-")[0]
    except Exception:
        return u""


# ---------------------------------------------------------------------------
# Werte lesen
# ---------------------------------------------------------------------------

def lese_wert(doc, param):
    """Liest einen Parameter für den Export.

    Rückgabe: (zellwert, anzeigetext)
        zellwert    - typisierter Wert für Excel (str/int/float), beim Import
                      wieder verarbeitbar
        anzeigetext - formatierter Text laut Revit (AsValueString), nur zur Info
    """
    if param is None:
        return None, u""

    try:
        anzeige = param.AsValueString()
    except Exception:
        anzeige = None
    anzeige = anzeige if anzeige else u""

    try:
        speichertyp = param.StorageType
    except Exception:
        return None, anzeige

    try:
        if speichertyp == StorageType.String:
            text = param.AsString()
            text = text if text is not None else u""
            return text, (anzeige or text)

        if speichertyp == StorageType.Integer:
            roh = param.AsInteger()
            if ist_janein(param):
                # Ja/Nein bewusst als Text, damit der Nutzer in Excel nicht mit 0/1 hantiert
                text = anzeige if anzeige else (t(u"Ja", u"Yes", u"Sí") if roh else t(u"Nein", u"No", u"No"))
                return text, text
            return roh, (anzeige or str(roh))

        if speichertyp == StorageType.Double:
            roh = param.AsDouble()
            wert = nach_projekteinheit(param, roh)
            try:
                wert = round(float(wert), NACHKOMMASTELLEN)
            except Exception:
                pass
            return wert, (anzeige or str(wert))

        if speichertyp == StorageType.ElementId:
            ziel_id = param.AsElementId()
            if not ist_gueltige_id(ziel_id):
                return u"", anzeige
            name = anzeige or element_name(doc.GetElement(ziel_id))
            return name, (anzeige or name)
    except Exception:
        return None, anzeige

    return None, anzeige


def ist_schreibgeschuetzt(param):
    """True, wenn der Parameter nicht gesetzt werden kann."""
    if param is None:
        return True
    try:
        return bool(param.IsReadOnly)
    except Exception:
        return True


# ---------------------------------------------------------------------------
# Werte schreiben
# ---------------------------------------------------------------------------

class Konvertierungsfehler(Exception):
    """Der Excel-Wert passt nicht zum Datentyp des Parameters."""


def _als_zahl(wert):
    """Excel-Zelle -> float, akzeptiert auch Text mit Komma als Dezimaltrenner."""
    if isinstance(wert, bool):
        return 1.0 if wert else 0.0
    if isinstance(wert, (int, float)):
        return float(wert)
    text = str(wert).strip().replace(u" ", u"").replace(u" ", u"")
    if not text:
        raise Konvertierungsfehler(t(u"leerer Wert", u"empty value", u"valor vacío"))
    # Deutsche Schreibweise 1.234,56 ebenso wie 1234.56 zulassen
    if u"," in text:
        text = text.replace(u".", u"").replace(u",", u".")
    try:
        return float(text)
    except ValueError:
        raise Konvertierungsfehler(t(u"'%s' ist keine Zahl", u"'%s' is not a number", u"'%s' no es un número") % wert)


def _als_janein(wert):
    """Excel-Zelle -> 0/1 für Ja/Nein-Parameter."""
    if isinstance(wert, bool):
        return 1 if wert else 0
    text = str(wert).strip().lower()
    if text in JA_WERTE:
        return 1
    if text in NEIN_WERTE:
        return 0
    raise Konvertierungsfehler(t(u"'%s' ist kein Ja/Nein-Wert", u"'%s' is not a yes/no value", u"'%s' no es un valor sí/no") % wert)


def _finde_element_nach_name(doc, param, name):
    """Sucht für einen ElementId-Parameter ein Element mit passendem Namen.

    Strategie: Die Klasse des aktuell gesetzten Elements bestimmt den Suchraum.
    Nur bei genau einem Treffer wird geschrieben - alles andere ist zu riskant.
    """
    name = str(name).strip()
    if not name:
        return ElementId.InvalidElementId

    aktuelle_id = param.AsElementId()
    if not ist_gueltige_id(aktuelle_id):
        raise Konvertierungsfehler(
            t(u"Zielelement nicht bestimmbar (Parameter ist aktuell leer)", u"Target element cannot be determined (parameter is currently empty)", u"No se puede determinar el elemento de destino (el parámetro está vacío)"))

    aktuelles = doc.GetElement(aktuelle_id)
    if aktuelles is None:
        raise Konvertierungsfehler(t(u"Zielelement nicht bestimmbar", u"Target element cannot be determined", u"No se puede determinar el elemento de destino"))

    treffer = []
    try:
        for kandidat in FilteredElementCollector(doc).OfClass(aktuelles.GetType()):
            if element_name(kandidat) == name:
                treffer.append(kandidat.Id)
    except Exception as fehler:
        raise Konvertierungsfehler(t(u"Suche fehlgeschlagen: %s", u"Search failed: %s", u"Error en la búsqueda: %s") % fehler)

    if not treffer:
        raise Konvertierungsfehler(t(u"kein Element namens '%s' gefunden", u"no element named '%s' found", u"no se encontró ningún elemento llamado '%s'") % name)
    if len(treffer) > 1:
        raise Konvertierungsfehler(t(u"Name '%s' ist nicht eindeutig", u"name '%s' is not unique", u"el nombre '%s' no es único") % name)
    return treffer[0]


def zielwert(doc, param, zellwert):
    """Excel-Zelle -> Wert in interner Revit-Darstellung.

    Wirft Konvertierungsfehler mit deutscher Begründung.
    """
    speichertyp = param.StorageType

    if speichertyp == StorageType.String:
        return u"" if zellwert is None else str(zellwert)

    if speichertyp == StorageType.Integer:
        if ist_janein(param):
            return _als_janein(u"" if zellwert is None else zellwert)
        if zellwert is None or str(zellwert).strip() == u"":
            raise Konvertierungsfehler(t(u"leerer Wert für Ganzzahl-Parameter", u"empty value for integer parameter", u"valor vacío para parámetro entero"))
        return int(round(_als_zahl(zellwert)))

    if speichertyp == StorageType.Double:
        if zellwert is None or str(zellwert).strip() == u"":
            raise Konvertierungsfehler(t(u"leerer Wert für Zahl-Parameter", u"empty value for number parameter", u"valor vacío para parámetro numérico"))
        return nach_interner_einheit(param, _als_zahl(zellwert))

    if speichertyp == StorageType.ElementId:
        if zellwert is None or str(zellwert).strip() == u"":
            return ElementId.InvalidElementId
        return _finde_element_nach_name(doc, param, zellwert)

    raise Konvertierungsfehler(t(u"nicht unterstützter Datentyp", u"unsupported data type", u"tipo de datos no admitido"))


def unterscheidet_sich(param, neuer_wert):
    """True, wenn der neue Wert vom aktuellen Parameterwert abweicht.

    Verhindert überflüssige Parameter.Set-Aufrufe und damit unnötige
    Revit-Warnungen, Regenerationen und Änderungsmarkierungen in der Zentraldatei.
    """
    speichertyp = param.StorageType
    try:
        if speichertyp == StorageType.String:
            aktuell = param.AsString()
            return (aktuell if aktuell is not None else u"") != neuer_wert
        if speichertyp == StorageType.Integer:
            return param.AsInteger() != neuer_wert
        if speichertyp == StorageType.Double:
            return abs(param.AsDouble() - neuer_wert) > TOLERANZ
        if speichertyp == StorageType.ElementId:
            return eid_wert(param.AsElementId()) != eid_wert(neuer_wert)
    except Exception:
        # Im Zweifel schreiben - Parameter.Set meldet einen Fehler selbst
        return True
    return True


# ---------------------------------------------------------------------------
# Worksharing
# ---------------------------------------------------------------------------

def fremder_besitzer(doc, element):
    """Name des anderen Benutzers, falls das Element ausgeliehen ist - sonst None.

    In Zentraldatei-Projekten würde ein Schreibversuch auf ein fremd
    ausgeliehenes Element die gesamte Transaktion zum Scheitern bringen.
    Deshalb wird vorher geprüft und die Zeile sauber übersprungen.
    """
    if WorksharingUtils is None or element is None:
        return None
    try:
        if not doc.IsWorkshared:
            return None
        status = WorksharingUtils.GetCheckoutStatus(doc, element.Id)
        if status == CheckoutStatus.OwnedByOtherUser:
            info = WorksharingUtils.GetWorksharingTooltipInfo(doc, element.Id)
            return info.Owner or t(u"anderer Benutzer", u"another user", u"otro usuario")
    except Exception:
        return None
    return None
