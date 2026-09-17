#! python3
# -*- coding: utf-8 -*-
"""Setzt den Rauminhaltspunkt (Room.Location) auf den Flächenschwerpunkt.

Ablauf je Raum:
1. Umriss über Room.GetBoundarySegments() lesen, jede Kurve per
   Curve.Tessellate() in Punkte auflösen (Bögen werden mit erfasst).
2. Flächenschwerpunkt mit der Polygon-Schwerpunktformel (Shoelace) berechnen.
   Mehrere Schleifen: grösste = Aussenumriss, Löcher zählen negativ
   (siehe raum_zentrum.geometrie).
3. Ergebnis mit Room.IsPointInRoom() prüfen.
4. Nur X/Y verschieben (ElementTransformUtils.MoveElement), Z bleibt.
5. Optional die Raumbeschriftungen (RoomTags) auf den neuen Punkt setzen -
   in allen Grundrissen oder nur in der aktiven Ansicht. Beschriftungen mit
   Führungslinie wurden bewusst abgesetzt und bleiben unverändert.

Ersatzstrategie, wenn der Schwerpunkt nicht im Raum liegt (U-Form, stark
konkave Räume) - beide Varianten sind umgesetzt, gewählt über
ERSATZSTRATEGIE:

  INNENPUNKT  (gewählt)
      Waagrechte und senkrechte Suchlinien durch den Umriss legen, die
      Mittelpunkte der innenliegenden Abschnitte (= Mittelachsen der
      Teilstücke) quer dazu zentrieren und den dem Schwerpunkt nächstgelegenen
      Punkt nehmen, der IsPointInRoom() besteht. Bei einem U-Raum landet der
      Punkt so in der Mitte des Verbindungsstegs statt in der Öffnung.
      Findet sich kein gültiger Punkt, gilt automatisch BEIBEHALTEN.

  BEIBEHALTEN
      Standort unverändert lassen, Raum als "Zentrum liegt außerhalb, manuell
      prüfen" melden.

  Warum INNENPUNKT: Sinn des Werkzeugs ist, dass Raumbeschriftungen mittig im
  Raum stehen. Mit BEIBEHALTEN blieben ausgerechnet die verwinkelten Räume
  unbearbeitet, bei denen der Punkt am häufigsten schlecht liegt. Der
  Ersatzpunkt wird genauso mit IsPointInRoom() abgesichert und in der
  Auswertung eigens aufgeführt, damit er kontrolliert werden kann.

CPython-Besonderheiten dieses Repos (siehe ScheduleSync.panel/README.md):
pyrevit.forms ist unter CPython nicht nutzbar - die Dialoge kommen aus
schedule_sync.ui (TaskDialog). script.exit()/sys.exit() beenden den
CPython-Host; der Ablauf endet deshalb nur über 'return' in main().
"""

import io
import os
import sys
import traceback

_EXT = os.path.dirname(os.path.abspath(__file__))
while not _EXT.endswith(".extension") and os.path.dirname(_EXT) != _EXT:
    _EXT = os.path.dirname(_EXT)
if os.path.join(_EXT, "lib") not in sys.path:
    sys.path.append(os.path.join(_EXT, "lib"))

from mlg_sprache import t  # noqa: E402

from pyrevit import script

output = script.get_output()

TITEL = t(u"Raumpunkt zentrieren", u"Center Room Point", u"Centrar punto de habitación")

# Ersatzstrategie, siehe Moduldokumentation oben
ERSATZ_INNENPUNKT = "INNENPUNKT"
ERSATZ_BEIBEHALTEN = "BEIBEHALTEN"
ERSATZSTRATEGIE = ERSATZ_INNENPUNKT

# Höchstens so viele Ersatzpunkte mit IsPointInRoom() prüfen (nächste zuerst)
MAX_ERSATZ_PRUEFUNGEN = 40

FUSS_IN_M = 0.3048
# Kleinere Verschiebungen (1 mm) gelten als "bereits zentriert"
MIN_VERSCHIEBUNG = 0.001 / FUSS_IN_M

GRUND_NICHT_PLATZIERT = t(u"Nicht platziert (kein Standortpunkt)", u"Not placed (no location point)", u"No colocada (sin punto de ubicación)")
GRUND_NICHT_UMSCHLOSSEN = t(u"Nicht umschlossen oder redundant (Fläche 0)", u"Not enclosed or redundant (area 0)", u"No delimitada o redundante (área 0)")
GRUND_KEIN_UMRISS = t(u"Kein Umriss", u"No outline", u"Sin contorno")
GRUND_AUSSERHALB = t(u"Zentrum liegt außerhalb, manuell prüfen", u"Centre lies outside, check manually", u"El centro queda fuera, compruébelo manualmente")
GRUND_FEHLER = t(u"Fehler beim Verschieben", u"Error while moving", u"Error al mover")

STATUS_VERSCHIEBEN = "verschieben"
STATUS_ZENTRIERT = "zentriert"
STATUS_UEBERSPRUNGEN = "uebersprungen"

BESCHRIFTUNG_ALLE = 0
BESCHRIFTUNG_ANSICHT = 1
BESCHRIFTUNG_KEINE = 2


def main():
    from Autodesk.Revit.DB import (
        BuiltInCategory,
        BuiltInParameter,
        ElementTransformUtils,
        FilteredElementCollector,
        SpatialElementBoundaryLocation,
        SpatialElementBoundaryOptions,
        SubTransaction,
        Transaction,
        ViewPlan,
        XYZ,
    )
    from Autodesk.Revit.DB.Architecture import Room, RoomTag

    from raum_zentrum import geometrie
    from schedule_sync import ui

    uidoc = getattr(__revit__, "ActiveUIDocument", None)
    doc = uidoc.Document if uidoc is not None else None
    if doc is None:
        ui.meldung(t(u"Es ist kein Projekt geöffnet.", u"No project is open.", u"No hay ningún proyecto abierto."), titel=TITEL,
                   hauptzeile=t(u"Kein aktives Dokument", u"No active document", u"No hay documento activo"), warnung=True)
        return
    if doc.IsFamilyDocument:
        ui.meldung(t(u"Räume gibt es nur in Projektdateien.", u"Rooms only exist in project files.", u"Las habitaciones solo existen en archivos de proyecto."), titel=TITEL,
                   hauptzeile=t(u"Familiendokument wird nicht unterstützt", u"Family document not supported", u"No se admiten documentos de familia"),
                   warnung=True)
        return

    # -----------------------------------------------------------------------
    # Hilfsfunktionen
    # -----------------------------------------------------------------------
    def id_wert(element_id):
        try:
            return int(element_id.Value)
        except AttributeError:
            return int(element_id.IntegerValue)

    def link(element_id):
        try:
            return output.linkify(element_id)
        except Exception:
            return str(id_wert(element_id))

    def raumname(raum):
        parameter = raum.get_Parameter(BuiltInParameter.ROOM_NAME)
        name = parameter.AsString() if parameter is not None else None
        return name or t(u"(ohne Name)", u"(no name)", u"(sin nombre)")

    def nur_raeume(elemente):
        return [e for e in elemente if isinstance(e, Room)]

    def raeume_im_projekt():
        return nur_raeume(FilteredElementCollector(doc)
                          .OfCategory(BuiltInCategory.OST_Rooms)
                          .WhereElementIsNotElementType())

    def raeume_in_ansicht():
        """None, wenn die aktive Ansicht keine Elemente enthalten kann."""
        ansicht = doc.ActiveView
        if ansicht is None:
            return None
        try:
            return nur_raeume(FilteredElementCollector(doc, ansicht.Id)
                              .OfCategory(BuiltInCategory.OST_Rooms)
                              .WhereElementIsNotElementType())
        except Exception:
            return None

    def raeume_in_auswahl():
        return nur_raeume(doc.GetElement(i)
                          for i in uidoc.Selection.GetElementIds())

    def pruefhoehen(raum, z_standort):
        """Z-Werte für IsPointInRoom().

        Die Methode prüft gegen das Raumvolumen. Ein Punkt genau auf der
        Unterkante (Ebene) oder bei Dachschrägen in halber Höhe kann knapp
        daneben liegen, deshalb mehrere Höhen - eine Übereinstimmung genügt.
        """
        hoehen = [z_standort + 0.01]
        box = raum.get_BoundingBox(None)
        if box is not None:
            hoehe = box.Max.Z - box.Min.Z
            if hoehe > 1e-6:
                hoehen.append(box.Min.Z + min(1.0, hoehe * 0.25))
                hoehen.append(box.Min.Z + hoehe * 0.5)
        return hoehen

    def im_raum(raum, x, y, hoehen):
        return any(raum.IsPointInRoom(XYZ(x, y, z)) for z in hoehen)

    boundary_optionen = SpatialElementBoundaryOptions()
    boundary_optionen.SpatialElementBoundaryLocation = \
        SpatialElementBoundaryLocation.Finish

    def lies_umriss(raum):
        """Liste der Schleifen als (x, y)-Punktlisten oder None."""
        segmente = raum.GetBoundarySegments(boundary_optionen)
        if segmente is None:
            return None
        schleifen = []
        for schleife in segmente:
            kurven = []
            for segment in schleife:
                kurve = segment.GetCurve()
                if kurve is not None:
                    kurven.append([(p.X, p.Y) for p in kurve.Tessellate()])
            punkte = geometrie.verkette_kurvenpunkte(kurven)
            if len(punkte) >= 3:
                schleifen.append(punkte)
        return schleifen or None

    def analysiere(raum):
        e = {
            "raum": raum,
            "id": raum.Id,
            "name": raumname(raum),
            "nummer": raum.Number or u"",
            "status": STATUS_UEBERSPRUNGEN,
            "grund": None,
            "strategie": u"-",
            "schwerpunkt_im_raum": None,
            "schleifen": 0,
            "loecher": 0,
            "punkte": 0,
            "verschiebung": None,
            "ziel": None,
        }

        standort = raum.Location
        punkt = getattr(standort, "Point", None) if standort else None
        if punkt is None:
            e["grund"] = GRUND_NICHT_PLATZIERT
            return e
        if raum.Area <= 0:
            e["grund"] = GRUND_NICHT_UMSCHLOSSEN
            return e

        schleifen = lies_umriss(raum)
        if not schleifen:
            e["grund"] = GRUND_KEIN_UMRISS
            return e
        try:
            sp = geometrie.schwerpunkt_mit_loechern(schleifen)
        except ValueError as fehler:
            e["grund"] = u"%s (%s)" % (GRUND_KEIN_UMRISS, fehler)
            return e
        e["schleifen"] = sp["schleifen"]
        e["loecher"] = sp["loecher"]
        e["punkte"] = sp["punkte"]

        hoehen = pruefhoehen(raum, punkt.Z)
        ziel = None
        e["schwerpunkt_im_raum"] = im_raum(raum, sp["cx"], sp["cy"], hoehen)
        if e["schwerpunkt_im_raum"]:
            ziel = (sp["cx"], sp["cy"])
            e["strategie"] = t(u"Schwerpunkt", u"Centroid", u"Centroide")
        elif ERSATZSTRATEGIE == ERSATZ_INNENPUNKT:
            kandidaten = geometrie.ersatzpunkte(schleifen, sp["cx"], sp["cy"])
            for x, y in kandidaten[:MAX_ERSATZ_PRUEFUNGEN]:
                if im_raum(raum, x, y, hoehen):
                    ziel = (x, y)
                    e["strategie"] = t(u"Ersatzpunkt", u"Fallback point", u"Punto alternativo")
                    break
        # ERSATZ_BEIBEHALTEN oder kein gültiger Ersatzpunkt gefunden
        if ziel is None:
            e["strategie"] = u"beibehalten"
            e["grund"] = GRUND_AUSSERHALB
            return e

        # Nur X/Y - Z = 0 hält den Raum auf seiner Ebene
        e["ziel"] = ziel
        e["verschiebung"] = XYZ(ziel[0] - punkt.X, ziel[1] - punkt.Y, 0.0)
        if e["verschiebung"].GetLength() < MIN_VERSCHIEBUNG:
            e["status"] = STATUS_ZENTRIERT
        else:
            e["status"] = STATUS_VERSCHIEBEN
        return e

    def sammle_beschriftungen(nur_aktive_ansicht):
        """Raumbeschriftungen je Raum (Schlüssel: Id-Wert des Raums)."""
        if nur_aktive_ansicht:
            if not isinstance(doc.ActiveView, ViewPlan):
                return {}
            sammler = FilteredElementCollector(doc, doc.ActiveView.Id)
        else:
            sammler = FilteredElementCollector(doc)
        nach_raum = {}
        for tag in (sammler.OfCategory(BuiltInCategory.OST_RoomTags)
                    .WhereElementIsNotElementType()):
            if not isinstance(tag, RoomTag):
                continue
            # Beschriftungen von Räumen aus verknüpften Modellen liefern hier
            # InvalidElementId (-1) und werden ignoriert
            raum_id = tag.TaggedLocalRoomId
            if raum_id is None or id_wert(raum_id) < 0:
                continue
            nach_raum.setdefault(id_wert(raum_id), []).append(tag)
        return nach_raum

    # -----------------------------------------------------------------------
    # 1. Umfang wählen
    # -----------------------------------------------------------------------
    alle = raeume_im_projekt()
    ansicht = raeume_in_ansicht()
    auswahl = raeume_in_auswahl()

    optionen = [
        (t(u"Alle Räume im Projekt", u"All rooms in the project", u"Todas las habitaciones del proyecto"), t(u"%d Räume", u"%d rooms", u"%d habitaciones") % len(alle)),
        (t(u"Nur ausgewählte Räume", u"Selected rooms only", u"Solo habitaciones seleccionadas"), t(u"%d Räume ausgewählt", u"%d rooms selected", u"%d habitaciones seleccionadas") % len(auswahl)),
        (t(u"Nur Räume in aktiver Ansicht", u"Only rooms in active view", u"Solo habitaciones de la vista activa"),
         t(u"%d Räume in \"%s\"", u"%d rooms in \"%s\"", u"%d habitaciones en \"%s\"") % (len(ansicht), doc.ActiveView.Name)
         if ansicht is not None
         else t(u"Die aktive Ansicht kann keine Räume enthalten", u"The active view cannot contain rooms", u"La vista activa no puede contener habitaciones")),
    ]
    umfang = ui.waehle_option(optionen, titel=TITEL,
                              hauptzeile=t(u"Welche Räume sollen bearbeitet "
                                         u"werden?", u"Which rooms should be processed?", u"¿Qué habitaciones se deben procesar?"))
    if umfang is None:
        return
    raeume = [alle, auswahl, ansicht or []][umfang]
    if not raeume:
        ui.meldung(t(u"Im gewählten Umfang wurden keine Räume gefunden.", u"No rooms were found in the chosen scope.", u"No se encontraron habitaciones en el ámbito elegido."),
                   titel=TITEL, hauptzeile=t(u"Keine Räume", u"No rooms", u"No hay habitaciones"), warnung=True)
        return

    # -----------------------------------------------------------------------
    # 2. Beschriftungen: überall, nur aktive Ansicht oder gar nicht
    # -----------------------------------------------------------------------
    aktive_ist_grundriss = isinstance(doc.ActiveView, ViewPlan)
    beschriftung = ui.waehle_option(
        [(t(u"In allen Grundrissen zentrieren", u"Centre in all floor plans", u"Centrar en todas las plantas"),
          t(u"Raumbeschriftungen in allen Grundriss- und Deckenplänen auf den "
          u"neuen Raumpunkt setzen.", u"Move room tags in all floor and ceiling plans to the new room point.", u"Mover las etiquetas de habitación de todas las plantas y planos de techo al nuevo punto.")),
         (t(u"Nur in der aktiven Ansicht zentrieren", u"Centre in the active view only", u"Centrar solo en la vista activa"),
          t(u"Nur Beschriftungen in \"%s\".", u"Only tags in \"%s\".", u"Solo las etiquetas de \"%s\".") % doc.ActiveView.Name
          if aktive_ist_grundriss
          else t(u"Die aktive Ansicht ist kein Grundriss - es würde keine "
               u"Beschriftung geändert.", u"The active view is not a plan - no tag would be changed.", u"La vista activa no es una planta: no se cambiaría ninguna etiqueta.")),
         (t(u"Beschriftungen nicht verändern", u"Do not change tags", u"No cambiar las etiquetas"),
          t(u"Nur der Rauminhaltspunkt wird verschoben.", u"Only the room point is moved.", u"Solo se mueve el punto de la habitación."))],
        titel=TITEL,
        hauptzeile=t(u"Sollen die Raumbeschriftungen mit zentriert werden?", u"Should the room tags be centred as well?", u"¿Centrar también las etiquetas de habitación?"),
        text=t(u"Beschriftungen mit Führungslinie bleiben immer unverändert.", u"Tags with a leader always stay unchanged.", u"Las etiquetas con directriz nunca se modifican."))
    if beschriftung is None:
        return

    # -----------------------------------------------------------------------
    # 3. Probelauf oder Verschieben
    # -----------------------------------------------------------------------
    modus = ui.waehle_option(
        [(t(u"Probelauf - nur prüfen", u"Dry run - check only", u"Prueba - solo comprobar"),
          t(u"Berechnet Schwerpunkt und IsPointInRoom() und zeigt das Ergebnis. "
          u"Das Modell wird nicht verändert.", u"Calculates the centroid and IsPointInRoom() and shows the result. The model is not changed.", u"Calcula el centroide e IsPointInRoom() y muestra el resultado. El modelo no se modifica.")),
         (t(u"Rauminhaltspunkte verschieben", u"Move room points", u"Mover puntos de habitación"),
          t(u"Verschiebt %d Räume in einer Transaktion (rückgängig mit "
          u"Strg+Z).", u"Moves %d rooms in one transaction (undo with Ctrl+Z).", u"Mueve %d habitaciones en una transacción (deshacer con Ctrl+Z).") % len(raeume))],
        titel=TITEL,
        hauptzeile=t(u"%d Räume - wie soll vorgegangen werden?", u"%d rooms - how to proceed?", u"%d habitaciones: ¿cómo continuar?") % len(raeume),
        text=t(u"Empfehlung: zuerst einige Testräume (rechteckig, L-förmig, "
             u"mit Aussparung) auswählen und einen Probelauf starten.", u"Recommendation: first select a few test rooms (rectangular, L-shaped, with a cut-out) and run a dry run.", u"Recomendación: seleccione primero algunas habitaciones de prueba (rectangular, en L, con hueco) y ejecute una prueba."))
    if modus is None:
        return
    probelauf = modus == 0

    if not probelauf and doc.IsReadOnly:
        ui.meldung(t(u"Das Dokument ist schreibgeschützt.", u"The document is read-only.", u"El documento es de solo lectura."), titel=TITEL,
                   hauptzeile=t(u"Verschieben nicht möglich", u"Moving not possible", u"No se puede mover"), warnung=True)
        return
    if not probelauf and umfang == 0:
        if not ui.frage(t(u"Es werden alle %d Räume des Projekts bearbeitet.", u"All %d rooms of the project will be processed.", u"Se procesarán las %d habitaciones del proyecto.")
                        % len(raeume), titel=TITEL,
                        hauptzeile=t(u"Wirklich alle Räume verschieben?", u"Really move all rooms?", u"¿Mover realmente todas las habitaciones?"),
                        warnung=True):
            return

    # -----------------------------------------------------------------------
    # 4. Analysieren
    # -----------------------------------------------------------------------
    ergebnisse = []
    with ui.Fortschritt(output, len(raeume)) as fortschritt:
        for index, raum in enumerate(raeume):
            fortschritt.aktualisiere(index + 1)
            try:
                ergebnisse.append(analysiere(raum))
            except Exception as fehler:
                ergebnisse.append({
                    "raum": raum, "id": raum.Id, "name": raumname(raum),
                    "nummer": raum.Number or u"",
                    "status": STATUS_UEBERSPRUNGEN,
                    "grund": t(u"Fehler bei der Berechnung: %s", u"Calculation error: %s", u"Error de cálculo: %s") % fehler,
                    "strategie": u"-", "schwerpunkt_im_raum": None,
                    "schleifen": 0, "loecher": 0, "punkte": 0,
                    "verschiebung": None, "ziel": None,
                })

    # Beschriftungen der Räume mit gültigem Zielpunkt einordnen
    tag_kandidaten = []      # (Raumergebnis, Beschriftung)
    tag_mit_fuehrung = []
    tag_kein_grundriss = []
    if beschriftung != BESCHRIFTUNG_KEINE:
        tags_nach_raum = sammle_beschriftungen(
            beschriftung == BESCHRIFTUNG_ANSICHT)
        ist_grundriss = {}
        for e in ergebnisse:
            if e["ziel"] is None:
                continue
            for tag in tags_nach_raum.get(id_wert(e["id"]), []):
                ansicht_id = id_wert(tag.OwnerViewId)
                if ansicht_id not in ist_grundriss:
                    ist_grundriss[ansicht_id] = isinstance(
                        doc.GetElement(tag.OwnerViewId), ViewPlan)
                # In Schnitten liegt die Beschriftung nicht in der XY-Ebene
                if not ist_grundriss[ansicht_id]:
                    tag_kein_grundriss.append((e, tag))
                elif tag.HasLeader:
                    tag_mit_fuehrung.append((e, tag))
                else:
                    tag_kandidaten.append((e, tag))

    def tag_verschiebung(e, tag):
        kopf = tag.TagHeadPosition
        return XYZ(e["ziel"][0] - kopf.X, e["ziel"][1] - kopf.Y, 0.0)

    # -----------------------------------------------------------------------
    # 5. Verschieben - eine Transaktion, je Element eine SubTransaction, damit
    #    ein fehlgeschlagenes Element sauber zurückgerollt wird und der Rest
    #    bleibt
    # -----------------------------------------------------------------------
    zu_verschieben = [e for e in ergebnisse if e["status"] == STATUS_VERSCHIEBEN]
    verschoben = []
    tags_verschoben = []
    tags_fehler = []
    if probelauf:
        tags_verschoben = [(e, tag) for e, tag in tag_kandidaten
                           if tag_verschiebung(e, tag).GetLength()
                           >= MIN_VERSCHIEBUNG]
    elif zu_verschieben or tag_kandidaten:
        transaktion = Transaction(doc, t(u"Rauminhaltspunkte zentrieren", u"Center room points", u"Centrar puntos de habitación"))
        transaktion.Start()
        try:
            for e in zu_verschieben:
                teil = SubTransaction(doc)
                teil.Start()
                try:
                    ElementTransformUtils.MoveElement(doc, e["id"],
                                                      e["verschiebung"])
                    teil.Commit()
                    verschoben.append(e)
                except Exception as fehler:
                    if teil.HasStarted() and not teil.HasEnded():
                        teil.RollBack()
                    e["status"] = STATUS_UEBERSPRUNGEN
                    e["grund"] = u"%s: %s" % (GRUND_FEHLER, fehler)

            if tag_kandidaten:
                # Beschriftungspositionen erst nach dem Verschieben der Räume
                # lesen - falls Revit sie schon mitgezogen hat, stimmt der
                # Differenzvektor trotzdem
                doc.Regenerate()
                for e, tag in tag_kandidaten:
                    if e["status"] == STATUS_UEBERSPRUNGEN:
                        continue
                    vektor = tag_verschiebung(e, tag)
                    if vektor.GetLength() < MIN_VERSCHIEBUNG:
                        continue
                    teil = SubTransaction(doc)
                    teil.Start()
                    try:
                        ElementTransformUtils.MoveElement(doc, tag.Id, vektor)
                        teil.Commit()
                        tags_verschoben.append((e, tag))
                    except Exception as fehler:
                        if teil.HasStarted() and not teil.HasEnded():
                            teil.RollBack()
                        tags_fehler.append((e, tag, fehler))
            transaktion.Commit()
        except Exception:
            if transaktion.HasStarted() and not transaktion.HasEnded():
                transaktion.RollBack()
            raise

    # -----------------------------------------------------------------------
    # 6. Auswertung
    # -----------------------------------------------------------------------
    def meter(e):
        v = e["verschiebung"]
        return u"%.3f" % (v.GetLength() * FUSS_IN_M) if v is not None else u"-"

    def ja_nein(wert):
        return u"-" if wert is None else (
            t(u"Ja", u"Yes", u"Sí") if wert else t(u"Nein", u"No", u"No"))

    uebersprungen = [e for e in ergebnisse
                     if e["status"] == STATUS_UEBERSPRUNGEN]
    zentriert = [e for e in ergebnisse if e["status"] == STATUS_ZENTRIERT]
    ersatz = [e for e in ergebnisse if e["strategie"] == t(u"Ersatzpunkt", u"Fallback point", u"Punto alternativo")
              and e["status"] != STATUS_UEBERSPRUNGEN]

    output.print_md(u"# %s - %s" % (TITEL, t(u"Probelauf (nichts verändert)", u"Dry run (nothing changed)", u"Prueba (sin cambios)")
                                    if probelauf else t(u"Ergebnis", u"Result", u"Resultado")))
    if probelauf:
        output.print_md(t(u"- Würden verschoben: **%d**", u"- Would be moved: **%d**", u"- Se moverían: **%d**") % len(zu_verschieben))
    else:
        output.print_md(t(u"- Verschoben: **%d**", u"- Moved: **%d**", u"- Movidas: **%d**") % len(verschoben))
    output.print_md(t(u"- Bereits zentriert (< 1 mm): **%d**", u"- Already centred (< 1 mm): **%d**", u"- Ya centradas (< 1 mm): **%d**") % len(zentriert))
    output.print_md(t(u"- Übersprungen: **%d**", u"- Skipped: **%d**", u"- Omitidas: **%d**") % len(uebersprungen))
    output.print_md(t(u"- davon mit Ersatzpunkt statt Schwerpunkt: **%d**", u"- of which with fallback point instead of centroid: **%d**", u"- de ellas con punto alternativo en lugar del centroide: **%d**")
                    % len(ersatz))
    output.print_md(t(u"- Ersatzstrategie: `%s`", u"- Fallback strategy: `%s`", u"- Estrategia alternativa: `%s`") % ERSATZSTRATEGIE)
    if beschriftung != BESCHRIFTUNG_KEINE:
        output.print_md(t(u"- Beschriftungen %s: **%d** (%s)", u"- Tags %s: **%d** (%s)", u"- Etiquetas %s: **%d** (%s)")
                        % (t(u"würden zentriert", u"would be centred", u"se centrarían") if probelauf
                           else t(u"zentriert", u"centred", u"centradas"), len(tags_verschoben),
                           t(u"alle Grundrisse", u"all plans", u"todas las plantas")
                           if beschriftung == BESCHRIFTUNG_ALLE
                           else t(u"nur aktive Ansicht", u"active view only", u"solo vista activa")))
        output.print_md(t(u"- Beschriftungen mit Führungslinie (unverändert): "
                        u"**%d**", u"- Tags with leader (unchanged): **%d**", u"- Etiquetas con directriz (sin cambios): **%d**") % len(tag_mit_fuehrung))
        if tag_kein_grundriss:
            output.print_md(t(u"- Beschriftungen außerhalb von Grundrissen "
                            u"(unverändert): **%d**", u"- Tags outside plans (unchanged): **%d**", u"- Etiquetas fuera de plantas (sin cambios): **%d**") % len(tag_kein_grundriss))

    if probelauf:
        output.print_table(
            table_data=[[link(e["id"]), e["nummer"], e["name"],
                         u"%d / %d" % (e["schleifen"], e["loecher"]),
                         e["punkte"], ja_nein(e["schwerpunkt_im_raum"]),
                         e["strategie"], meter(e),
                         e["grund"] or e["status"]]
                        for e in ergebnisse],
            title=t(u"Prüfung je Raum", u"Check per room", u"Comprobación por habitación"),
            columns=[u"Id", t(u"Nummer", u"Number", u"Número"), t(u"Name", u"Name", u"Nombre"), t(u"Schleifen / Löcher", u"Loops / holes", u"Bucles / huecos"),
                     t(u"Punkte", u"Points", u"Puntos"), t(u"Schwerpunkt in Raum", u"Centroid in room", u"Centroide en habitación"), t(u"Zielpunkt", u"Target point", u"Punto destino"),
                     t(u"Verschiebung [m]", u"Offset [m]", u"Desplazamiento [m]"), t(u"Status", u"Status", u"Estado")],
        )

    if ersatz:
        output.print_md(t(u"## Ersatzpunkt verwendet - Lage bitte prüfen", u"## Fallback point used - please check location", u"## Se usó un punto alternativo: compruebe la ubicación"))
        for e in ersatz:
            output.print_md(u"- %s %s (%s m) - %s"
                            % (e["nummer"], e["name"], meter(e), link(e["id"])))

    if uebersprungen:
        output.print_md(t(u"## Übersprungen", u"## Skipped", u"## Omitidas"))
        nach_grund = {}
        for e in uebersprungen:
            nach_grund.setdefault(e["grund"], []).append(e)
        for grund in sorted(nach_grund):
            output.print_md(u"### %s (%d)" % (grund, len(nach_grund[grund])))
            for e in nach_grund[grund]:
                output.print_md(u"- %s %s - %s"
                                % (e["nummer"], e["name"], link(e["id"])))

    if tags_fehler:
        output.print_md(t(u"## Beschriftungen nicht verschoben (Fehler)", u"## Tags not moved (error)", u"## Etiquetas no movidas (error)"))
        for e, tag, fehler in tags_fehler:
            output.print_md(t(u"- %s %s - Beschriftung %s: %s", u"- %s %s - tag %s: %s", u"- %s %s - etiqueta %s: %s")
                            % (e["nummer"], e["name"], link(tag.Id), fehler))

    if tag_mit_fuehrung:
        output.print_md(t(u"## Beschriftungen mit Führungslinie - unverändert", u"## Tags with leader - unchanged", u"## Etiquetas con directriz - sin cambios"))
        for e, tag in tag_mit_fuehrung:
            output.print_md(t(u"- %s %s - Beschriftung %s", u"- %s %s - tag %s", u"- %s %s - etiqueta %s")
                            % (e["nummer"], e["name"], link(tag.Id)))


def _zeige_fehler(spur, titel):
    """Fehler sichtbar machen - und dabei selbst nicht scheitern können."""
    protokoll = os.path.join(os.environ.get("TEMP", "."),
                             "pyMLG_RoomCenter_Fehler.log")
    try:
        with io.open(protokoll, "w", encoding="utf-8") as datei:
            datei.write(spur)
    except Exception:
        protokoll = None

    try:
        output.print_md(u"# " + titel)
        print(spur)
        if protokoll:
            print(t(u"Protokoll: ", u"Log: ", u"Registro: ") + protokoll)
    except Exception:
        pass

    try:
        from schedule_sync import ui as _ui
        letzte_zeile = (spur.strip().splitlines() or [u""])[-1]
        _ui.meldung(letzte_zeile + t(u"\n\nEinzelheiten im pyRevit-"
                                   u"Ausgabefenster.", u"\n\nDetails in the pyRevit output window.", u"\n\nDetalles en la ventana de salida de pyRevit."),
                    titel=TITEL, hauptzeile=titel, warnung=True)
    except Exception:
        pass


try:
    main()
except Exception:
    _zeige_fehler(traceback.format_exc(),
                  t(u"Raumpunkt zentrieren ist fehlgeschlagen", u"Center Room Point failed", u"Centrar punto de habitación ha fallado"))
