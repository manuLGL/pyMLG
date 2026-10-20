# -*- coding: utf-8 -*-
"""Revit-Teil der Flächentabelle.

Räume eines Plans: aus jedem Grundriss-Ansichtsfenster die Räume, die in der
Ansicht sichtbar sind (Revit entscheidet: ausgeblendete Elemente, Filter,
Phase, Ansichtsbereich), auf der Ebene der Ansicht und mit dem Raumpunkt im
Zuschnitt - so zählen am Rand angeschnittene Nachbarräume nicht. Ist die
Kategorie Räume in der Ansicht ganz aus, zählen ersatzweise alle Räume der
Ebene in ihrer Phase mit dem Raumpunkt im Zuschnitt.

Dachschrägen: Querschnitt der Raumgeometrie (Room.ClosedShell) in 1 m und
2 m über dem Boden im Verhältnis zum Querschnitt am Boden. Das wirkt nur,
wenn in den Flächen- und Volumenberechnungen "Volumen berechnen" an ist -
sonst ist die Raumgeometrie ein gerades Prisma.

Die Tabelle ist eine Detailgruppe aus Textnotizen und Linien auf dem Plan,
Gruppentyp "pyMLG Wohnfläche <Plannummer>". Sie sitzt auf jedem Plan rechts
oben im Plankopf und wird beim nächsten Lauf dort ersetzt - von Hand
geänderte Überschrift und Mietblock bleiben.

Läuft unter CPython (List[T] wird leer angelegt und per Add gefüllt).
"""

import clr

from Autodesk.Revit.DB import (
    AreaVolumeSettings,
    BooleanOperationsUtils,
    BuiltInCategory,
    BuiltInParameter as BIP,
    ElementId,
    ElementTypeGroup,
    FilteredElementCollector,
    Group,
    GroupType,
    HorizontalTextAlignment,
    Line,
    PlanarFace,
    Plane,
    Solid,
    StorageType,
    TextNote,
    TextNoteOptions,
    TextNoteType,
    UnitTypeId,
    UnitUtils,
    ViewType,
    XYZ,
)

clr.AddReference("System")
from System.Collections.Generic import List  # noqa: E402

from wohnflaeche import logik as lg  # noqa: E402

GRUPPEN_PREFIX = u"pyMLG Wohnfläche"
METER = 1.0 / 0.3048
# Abstand der Tabelle vom Plankopfrand, wenn sie zum ersten Mal entsteht
RAND = 0.01 * METER


def m2(fuss2):
    return UnitUtils.ConvertFromInternalUnits(fuss2, UnitTypeId.SquareMeters)


def volumen_an(doc):
    try:
        return bool(AreaVolumeSettings.GetAreaVolumeSettings(doc).ComputeVolumes)
    except Exception:
        return False


def _id_liste(ids):
    liste = List[ElementId]()
    for i in ids:
        liste.Add(i)
    return liste


# ---------------------------------------------------------------- Räume

def _in_polygon(x, y, punkte):
    innen = False
    j = len(punkte) - 1
    for i in range(len(punkte)):
        xi, yi = punkte[i]
        xj, yj = punkte[j]
        if (yi > y) != (yj > y) and x < (xj - xi) * (y - yi) / (yj - yi) + xi:
            innen = not innen
        j = i
    return innen


def _zuschnitt_test(ansicht):
    """Funktion (x, y) -> liegt im Zuschnitt? Ohne aktiven Zuschnitt alles."""
    if not ansicht.CropBoxActive:
        return lambda x, y: True
    try:
        formen = ansicht.GetCropRegionShapeManager().GetCropShape()
        schleife = formen[0] if formen.Count else None
    except Exception:
        schleife = None
    if schleife is not None:
        punkte = []
        for kurve in schleife:
            punkte.extend((p.X, p.Y) for p in kurve.Tessellate())
        if len(punkte) >= 3:
            return lambda x, y: _in_polygon(x, y, punkte)
    box = ansicht.CropBox
    umkehr = box.Transform.Inverse

    def im_kasten(x, y):
        p = umkehr.OfPoint(XYZ(x, y, 0.0))
        return (box.Min.X <= p.X <= box.Max.X and box.Min.Y <= p.Y <= box.Max.Y)
    return im_kasten


def grundrisse(doc, plan):
    ansichten = []
    for vid in plan.GetAllViewports():
        fenster = doc.GetElement(vid)
        ansicht = doc.GetElement(fenster.ViewId) if fenster else None
        if ansicht is not None and ansicht.ViewType == ViewType.FloorPlan \
                and ansicht.GenLevel is not None:
            ansichten.append(ansicht)
    return ansichten


def _raeume_ausgeblendet(ansicht):
    try:
        return ansicht.GetCategoryHidden(ElementId(BuiltInCategory.OST_Rooms))
    except Exception:
        return False


def _kandidaten(doc, ansicht):
    """(Räume, ersatzweise?) - sichtbare Räume der Ansicht, bei
    ausgeschalteter Kategorie alle Räume der Phase."""
    if not _raeume_ausgeblendet(ansicht):
        return list(FilteredElementCollector(doc, ansicht.Id)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType()), False
    phase = ansicht.get_Parameter(BIP.VIEW_PHASE)
    phase_id = phase.AsElementId() if phase else None
    raeume = []
    for raum in (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_Rooms)
                 .WhereElementIsNotElementType()):
        raum_phase = raum.get_Parameter(BIP.ROOM_PHASE)
        if phase_id is None or raum_phase is None \
                or raum_phase.AsElementId() == phase_id:
            raeume.append(raum)
    return raeume, True


def raeume_auf_plan(doc, plan):
    """Revit-Räume aller Grundrisse auf dem Plan (ohne Doppelte).
    Rückgabe (Räume, Namen der Ansichten mit Ersatzsuche)."""
    gefunden = {}
    ersatz = []
    for ansicht in grundrisse(doc, plan):
        im_zuschnitt = _zuschnitt_test(ansicht)
        kandidaten, ersatzweise = _kandidaten(doc, ansicht)
        if ersatzweise:
            ersatz.append(ansicht.Name)
        for raum in kandidaten:
            if raum.Area <= 0 or raum.Location is None \
                    or raum.LevelId != ansicht.GenLevel.Id:
                continue
            punkt = raum.Location.Point
            if im_zuschnitt(punkt.X, punkt.Y):
                gefunden[raum.Id] = raum
    return list(gefunden.values()), ersatz


def _text(raum, bip):
    p = raum.get_Parameter(bip)
    return (p.AsString() or u"") if p is not None else u""


def _faktor_parameter_liste(element):
    """Parameter, deren Name nach Anrechenbarkeit aussieht, nach Vorrang
    (lg.faktor_rang): "Faktor Anrechenbarkeit" zuerst."""
    gefunden = []
    for p in element.Parameters:
        try:
            rang = lg.faktor_rang(p.Definition.Name)
        except Exception:
            continue
        if rang is not None:
            gefunden.append((rang, p.Definition.Name, p))
    gefunden.sort(key=lambda e: (e[0], e[1]))
    return [p for _, _, p in gefunden]


def _parameter_wert(doc, p, tiefe=0):
    """Rohwert eines Parameters: Zahl, Text oder None. Bei einem Schlüssel
    (ElementId, Schlüsselplan) der Wert des passenden Parameters im
    Schlüssel."""
    if not p.HasValue:
        return None
    if p.StorageType == StorageType.Double:
        return p.AsDouble()
    if p.StorageType == StorageType.Integer:
        return p.AsInteger()
    if p.StorageType == StorageType.String:
        return p.AsString()
    if p.StorageType == StorageType.ElementId and tiefe == 0:
        schluessel = doc.GetElement(p.AsElementId())
        if schluessel is not None:
            for q in _faktor_parameter_liste(schluessel):
                wert = _parameter_wert(doc, q, 1)
                if wert not in (None, u""):
                    return wert
            return schluessel.Name
    return None


def faktor_parameter(raum):
    """(Faktor 0..1 oder None, Hinweis für den Bericht).

    Der erste Parameter nach Vorrang mit lesbarem Wert gilt - ist "Faktor
    Anrechenbarkeit" leer, zählt z. B. "Anrechenbarkeit". Hinweis None = der
    Raum hat keinen solchen Parameter."""
    liste = _faktor_parameter_liste(raum)
    if not liste:
        return None, None
    hinweise = []
    for p in liste:
        name = p.Definition.Name
        roh = _parameter_wert(raum.Document, p)
        if roh is None or roh == u"":
            hinweise.append(u"%s leer" % name)
            continue
        anzeige = p.AsValueString() or (u"%s" % roh)
        faktor = lg.faktor_aus_parameter(roh)
        if faktor is None:
            hinweise.append(u'%s = "%s" nicht lesbar' % (name, anzeige))
            continue
        return faktor, u"; ".join(hinweise + [u'%s = "%s"' % (name, anzeige)])
    return None, u"; ".join(hinweise)


def _querschnitt(solids, z):
    """Fläche des Raumkörpers in Höhe z (Fuss²)."""
    ebene = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0.0, 0.0, z))
    summe = 0.0
    for solid in solids:
        oben = BooleanOperationsUtils.CutWithHalfSpace(solid, ebene)
        if oben is None or oben.Volume <= 1e-9:
            continue
        for flaeche in oben.Faces:
            if isinstance(flaeche, PlanarFace) and flaeche.FaceNormal.Z < -0.999 \
                    and abs(flaeche.Origin.Z - z) < 1e-4:
                summe += flaeche.Area
    return summe


def hoehen_anteile(raum):
    """(Anteil ab 1 m, Anteil ab 2 m lichter Höhe) oder (None, None)."""
    try:
        schale = raum.ClosedShell
        solids = [g for g in schale if isinstance(g, Solid) and g.Volume > 1e-9] \
            if schale is not None else []
        rahmen = raum.get_BoundingBox(None)
        if not solids or rahmen is None:
            return None, None
        boden = rahmen.Min.Z
        a0 = _querschnitt(solids, boden + 0.01 * METER)
        if a0 <= 1e-9:
            return None, None
        return (_querschnitt(solids, boden + 1.0 * METER) / a0,
                _querschnitt(solids, boden + 2.0 * METER) / a0)
    except Exception:
        return None, None


def raum_daten(raum):
    """lg.Raum aus einem Revit-Raum."""
    name = _text(raum, BIP.ROOM_NAME)
    faktor, hinweis = faktor_parameter(raum)
    anteil_1m = anteil_2m = None
    if faktor is None and lg.art(name)[0] == lg.INNEN:
        anteil_1m, anteil_2m = hoehen_anteile(raum)
    return lg.Raum(name, m2(raum.Area), abteilung=_text(raum, BIP.ROOM_DEPARTMENT),
                   belegung=_text(raum, BIP.ROOM_OCCUPANCY),
                   nummer=_text(raum, BIP.ROOM_NUMBER), faktor=faktor,
                   anteil_1m=anteil_1m, anteil_2m=anteil_2m, ref=raum,
                   parameter_hinweis=hinweis)


# ---------------------------------------------------------------- Tabelle

def vorhandene_tabellen(doc, plan):
    return [g for g in FilteredElementCollector(doc, plan.Id).OfClass(Group)
            if g.GroupType is not None
            and g.GroupType.Name.startswith(GRUPPEN_PREFIX)]


def loesche(doc, ids):
    doc.Delete(_id_liste(ids))


def entferne(doc, gruppen):
    """Löscht alte Tabellen samt Inhalt und Gruppentyp."""
    for gruppe in gruppen:
        typ_id = gruppe.GetTypeId()
        mitglieder = list(gruppe.GetMemberIds())
        gruppe.UngroupMembers()
        if mitglieder:
            doc.Delete(_id_liste(mitglieder))
        typ = doc.GetElement(typ_id)
        if typ is not None and typ.Groups.IsEmpty:
            doc.Delete(typ_id)


def _notiz_text(notiz):
    return (notiz.Text or u"").rstrip(u"\r\n ")


# So weit um eine alte Tabelle herum zählen lose Textnotizen noch zu ihr
UMGEBUNG = 0.02 * METER


def _mitte_in(notiz, plan, rahmen):
    box = notiz.get_BoundingBox(plan)
    if box is None:
        return False
    x = (box.Min.X + box.Max.X) / 2.0
    y = (box.Min.Y + box.Max.Y) / 2.0
    return (rahmen.Min.X - UMGEBUNG <= x <= rahmen.Max.X + UMGEBUNG
            and rahmen.Min.Y - UMGEBUNG <= y <= rahmen.Max.Y + UMGEBUNG)


def alte_texte(doc, plan, alte):
    """Von Hand eingetragene Texte einer vorhandenen Tabelle, damit sie
    beim Neuzeichnen bleiben: (Überschrift, Mietblock, lose Notizen).

    Gesucht wird in der Gruppe und in losen Textnotizen auf dem Plan
    rund um die Gruppe (auch Tabellen der älteren Fassung, in die Titel und
    Miete daneben geschrieben wurden).
    Mietblock: alle Notizen mit Miete/NK/Nebenkosten, von oben nach unten.
    Überschrift: die mittige Notiz, sonst die oberste, wenn sie nicht vom
    Werkzeug stammt ("3-Zimmer-Wohnung ...").
    Lose Notizen, die übernommen werden, gibt die Funktion zurück - sie
    werden gelöscht, damit nichts doppelt steht.

    Der grosse Titel links oben (Texttyp TITEL_TYP) geht allem vor; fehlt
    er, kommt der Titel aus der Tabelle (Fassung mit mittiger Überschrift)."""
    titel_notizen = [n for n in FilteredElementCollector(doc, plan.Id)
                     .OfClass(TextNote)
                     if n.GroupId == ElementId.InvalidElementId
                     and doc.GetElement(n.GetTypeId()) is not None
                     and doc.GetElement(n.GetTypeId()).Name == TITEL_TYP]
    gross = None
    if titel_notizen:
        gross = _notiz_text(max(titel_notizen, key=lambda n: n.Coord.Y)) or None
    titel_ids = [n.Id for n in titel_notizen]
    titel, miete, lose = _tabellen_texte(doc, plan, alte)
    return gross or titel, miete, titel_ids + lose


def _tabellen_texte(doc, plan, alte):
    """(Titel, Mietblock, lose Notizen) aus der alten Tabelle."""
    if not alte:
        return None, None, []
    mitglieder = set()
    notizen = []
    for gruppe in alte:
        for i in gruppe.GetMemberIds():
            mitglieder.add(i)
            element = doc.GetElement(i)
            if isinstance(element, TextNote):
                notizen.append(element)
    rahmen = [g.get_BoundingBox(plan) for g in alte]
    rahmen = [r for r in rahmen if r is not None]
    lose = []
    for notiz in FilteredElementCollector(doc, plan.Id).OfClass(TextNote):
        if notiz.Id in mitglieder or notiz.GroupId != ElementId.InvalidElementId:
            continue
        typ = doc.GetElement(notiz.GetTypeId())
        if typ is not None and typ.Name == TITEL_TYP:
            continue                    # grosser Titel - eigene Behandlung
        if any(_mitte_in(notiz, plan, r) for r in rahmen):
            lose.append(notiz)
            notizen.append(notiz)
    if not notizen:
        return None, None, []
    notizen.sort(key=lambda n: -n.Coord.Y)

    miete = [n for n in notizen if lg.ist_miete_text(_notiz_text(n))]
    rest = [n for n in notizen if n not in miete]
    mittig = [n for n in rest
              if n.HorizontalAlignment == HorizontalTextAlignment.Center]
    titel = mittig[0] if mittig else (rest[0] if rest else None)
    if titel is not None and (lg.ist_werkzeug_kopf(_notiz_text(titel))
                              or not _notiz_text(titel)):
        titel = None
    if titel is not None and miete and titel.Coord.Y < miete[0].Coord.Y:
        titel = None                    # Titel steht nie unter der Miete

    uebernommen = set(n.Id for n in miete)
    if titel is not None:
        uebernommen.add(titel.Id)
    miete_text = u"\r".join(_notiz_text(n) for n in miete) or None
    return ((_notiz_text(titel) if titel is not None else None), miete_text,
            [n.Id for n in lose if n.Id in uebernommen])


def standard_ecke(doc, plan, breite):
    """Linke obere Ecke der Tabelle: auf jedem Plan gleich - rechts oben im
    Plankopf, RAND nach innen (die rechte Kante fluchtet, egal wie breit)."""
    from mlg_plaene.revit import plankopf_rechteck
    x0, y0, x1, y1 = plankopf_rechteck(doc, plan)
    return XYZ(x1 - breite - RAND, y1 - RAND, 0.0)


TEXT_TYP = u"pyMLG Flächentabelle"
TEXT_GROESSE = 3.5 / 304.8      # 3,5 mm
TITEL_TYP = u"pyMLG Flächentabelle Titel"
TITEL_GROESSE = 7.0 / 304.8     # 7 mm


def text_typ(doc, name=TEXT_TYP, vorgabe=TEXT_GROESSE):
    """Texttyp name - fehlt er, als Kopie des Standard-Texttyps mit der
    Grösse vorgabe. Danach gilt die Grösse, die im Typ steht (in Revit
    änderbar, die Tabelle richtet Abstände und Breite danach)."""
    typ = None
    for kandidat in FilteredElementCollector(doc).OfClass(TextNoteType):
        if kandidat.Name == name:
            typ = kandidat
            break
    if typ is None:
        vorlage = doc.GetElement(
            doc.GetDefaultElementTypeId(ElementTypeGroup.TextNoteType))
        if vorlage is None:
            vorlage = FilteredElementCollector(doc).OfClass(TextNoteType).FirstElement()
        typ = vorlage.Duplicate(name)
        typ.get_Parameter(BIP.TEXT_SIZE).Set(vorgabe)
    groesse = typ.get_Parameter(BIP.TEXT_SIZE).AsDouble()
    return typ.Id, groesse if groesse > 0 else vorgabe


def _innen_links(doc, plan, x0, y0, y1):
    """Linke Kante des inneren Plankopfrahmens (Heftrand), sonst x0.

    Jedes Teilstück der Plankopflinien zählt einzeln (der Rahmen ist oft
    ein geschlossenes Rechteck oder durch Querlinien unterbrochen): die
    senkrechten Stücke 3 mm bis 6 cm rechts vom Aussenrand werden je x
    aufsummiert, die am weitesten rechts liegende Kante mit zusammen mehr
    als 30 % der Planhöhe ist der Innenrand."""
    from mlg_plaene.revit import plankopf_linien
    return lg.innen_links(plankopf_linien(doc, plan), x0, y0, y1, METER)


def zeichne_titel(doc, plan, text):
    """Grosser Titel links oben im Plankopf (Texttyp TITEL_TYP, fett) -
    gleiche Höhe wie die Tabelle."""
    from mlg_plaene.revit import plankopf_rechteck
    x0, y0, x1, y1 = plankopf_rechteck(doc, plan)
    typ_id, _ = text_typ(doc, TITEL_TYP, TITEL_GROESSE)
    links = _innen_links(doc, plan, x0, y0, y1)
    return _text_notiz(doc, plan, XYZ(links + RAND, y1 - RAND, 0), text,
                       typ_id, fett=True)


def tabellen_breite(zeilen, groesse):
    """Breite aus der längsten Zeile geschätzt (mittlere Zeichenbreite)."""
    zeichen = 0
    for art, links, rechts, _ in zeilen:
        if art in (lg.TITEL, lg.KOPF, lg.ZEILE, lg.MIETE):
            for teil in links.split(u"\r"):
                zeichen = max(zeichen, len(teil) + len(rechts) + 6)
    return max(44.0, zeichen * 0.62 + 3.2) * groesse


def _text_notiz(doc, plan, punkt, text, typ_id, fett=False, rechts=False,
                breite=None, mittig=False):
    optionen = TextNoteOptions(typ_id)
    optionen.HorizontalAlignment = (
        HorizontalTextAlignment.Center if mittig else
        HorizontalTextAlignment.Right if rechts else
        HorizontalTextAlignment.Left)
    if breite:
        notiz = TextNote.Create(doc, plan.Id, punkt, breite, text, optionen)
    else:
        notiz = TextNote.Create(doc, plan.Id, punkt, text, optionen)
    if fett:
        formatiert = notiz.GetFormattedText()
        formatiert.SetBoldStatus(True)
        notiz.SetFormattedText(formatiert)
    return notiz


def _linie(doc, plan, a, b):
    return doc.Create.NewDetailCurve(plan, Line.CreateBound(a, b))


def _freier_name(doc, basis):
    vergeben = set(g.Name for g in FilteredElementCollector(doc).OfClass(GroupType))
    if basis not in vergeben:
        return basis
    n = 2
    while u"%s (%d)" % (basis, n) in vergeben:
        n += 1
    return u"%s (%d)" % (basis, n)


def zeichne(doc, plan, zeilen, ecke=None):
    """Zeichnet die Tabelle als Detailgruppe; ecke = linke obere Ecke,
    ohne Angabe immer rechts oben im Plankopf (standard_ecke)."""
    typ_id, s = text_typ(doc)
    breite = tabellen_breite(zeilen, s)
    if ecke is None:
        ecke = standard_ecke(doc, plan, breite)
    x0, y = ecke.X, ecke.Y
    rand = 1.6 * s
    links_x, rechts_x = x0 + rand, x0 + breite - rand
    y -= rand
    ids = []
    def unterkante(notiz, ersatz):
        doc.Regenerate()
        rahmen = notiz.get_BoundingBox(plan)
        return rahmen.Min.Y if rahmen is not None else ersatz

    for art, links, rechts, fett in zeilen:
        if art == lg.TITEL:
            notiz = _text_notiz(doc, plan, XYZ(x0 + breite / 2.0, y, 0), links,
                                typ_id, fett=True, mittig=True)
            ids.append(notiz.Id)
            y = unterkante(notiz, y - 1.4 * s) - 1.2 * s
        elif art == lg.VOLL:
            ids.append(_linie(doc, plan, XYZ(x0, y, 0), XYZ(x0 + breite, y, 0)).Id)
            y -= 1.4 * s
        elif art == lg.MIETE:
            notiz = _text_notiz(doc, plan, XYZ(links_x, y, 0), links, typ_id)
            ids.append(notiz.Id)
            y = unterkante(notiz, y - 2.2 * s * (links.count(u"\r") + 1))
        elif art == lg.KOPF:
            ids.append(_text_notiz(doc, plan, XYZ(links_x, y, 0), links, typ_id,
                                   fett=True).Id)
            y -= 2.4 * s
        elif art == lg.TRENNER:
            y -= 0.2 * s
            ids.append(_linie(doc, plan, XYZ(links_x, y, 0), XYZ(rechts_x, y, 0)).Id)
            y -= 1.2 * s
        elif art == lg.ZEILE:
            ids.append(_text_notiz(doc, plan, XYZ(links_x, y, 0), links, typ_id,
                                   fett=fett).Id)
            ids.append(_text_notiz(doc, plan, XYZ(rechts_x, y, 0), rechts, typ_id,
                                   fett=fett, rechts=True).Id)
            y -= 2.2 * s
        elif art == lg.FUSS:
            notiz = _text_notiz(doc, plan, XYZ(links_x, y, 0), links, typ_id,
                                breite=rechts_x - links_x)
            ids.append(notiz.Id)
            y = unterkante(notiz, y - 2.2 * s) - 1.0 * s
    unten = y - rand
    ecken = [XYZ(x0, ecke.Y, 0), XYZ(x0 + breite, ecke.Y, 0),
             XYZ(x0 + breite, unten, 0), XYZ(x0, unten, 0)]
    for i in range(4):
        ids.append(_linie(doc, plan, ecken[i], ecken[(i + 1) % 4]).Id)
    gruppe = doc.Create.NewGroup(_id_liste(ids))
    gruppe.GroupType.Name = _freier_name(doc, u"%s %s" % (GRUPPEN_PREFIX,
                                                         plan.SheetNumber))
    return gruppe
