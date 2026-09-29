# -*- coding: utf-8 -*-
# Setzt die ausgewählten Ansichten auf Pläne: ein Plan je Ansicht oder
# mehrere Ansichten je Plan (bei Platzmangel entstehen weitere Pläne),
# wahlweise auf einen vorhandenen Plan. Plankopf, Ansichtsvorlage,
# Ansichtsfenster-Typ, Position und Nummern-/Namensmuster im Dialog.
# (Kommentar statt Docstring: pyRevit liest unter IronPython den
#  Docstring als Tooltip und scheitert dabei an Umlauten.)

__title__ = "ViewToSheet"
__author__ = "Manuel"

import io
import os

from Autodesk.Revit.DB import (FilteredElementCollector, Transaction,
                               TransactionGroup, View, ViewSheet, ViewType,
                               Viewport)
from pyrevit import forms, revit, script

from mlg_plaene import logik as lg
from mlg_plaene import revit as rv
from mlg_plaene import vorschau
from mlg_sprache import t, tt, uebersetze_xaml

TITEL = t(u"Ansichten auf Pläne", u"Views to Sheets", u"Vistas a planos")
KEINE_VORLAGE = t(u"<Keine Vorlage – Ansicht unverändert>",
                  u"<No template – view unchanged>",
                  u"<Sin plantilla – vista sin cambios>")
NEUER_PLAN = t(u"<Neue Pläne erstellen>", u"<Create new sheets>",
               u"<Crear planos nuevos>")
STANDARD_TYP = t(u"<Standard>", u"<Default>", u"<Predeterminado>")

XAML_TEXTE = {
    "titel": (u"Ansichten auf Pläne", u"Views to Sheets", u"Vistas a planos"),
    "anordnung": (u"Anordnung", u"Arrangement", u"Disposición"),
    "einzeln": (u"Ein Plan je Ansicht", u"One sheet per view",
                u"Un plano por vista"),
    "zusammen": (u"Mehrere Ansichten je Plan (neue Pläne: weitere bei Platzmangel)",
                 u"Several views per sheet (new sheets: more if space runs out)",
                 u"Varias vistas por plano (planos nuevos: más si falta espacio)"),
    "ueberlappen": (u"Überschneidungen erlauben - alle auf einen Plan, nie ein weiterer",
                    u"Allow overlaps - all on one sheet, never another one",
                    u"Permitir solapes: todas en un plano, nunca otro"),
    "zielplan": (u"Zielplan", u"Target sheet", u"Plano de destino"),
    "planfamilie": (u"Plankopf für neue Pläne", u"Title block for new sheets",
                    u"Cajetín para planos nuevos"),
    "vorlage": (u"Ansichtsvorlage", u"View template", u"Plantilla de vista"),
    "fenstertyp": (u"Ansichtsfenster-Typ", u"Viewport type",
                   u"Tipo de ventana gráfica"),
    "position": (u"Position", u"Position", u"Posición"),
    "mitte": (u"Mitte der Zeichenfläche", u"Centre of the drawing area",
              u"Centro del área de dibujo"),
    "xy": (u"Fester Punkt X/Y (Mitte des Ansichtsfensters, vom Plan-Ursprung)",
           u"Fixed point X/Y (viewport centre, from the sheet origin)",
           u"Punto fijo X/Y (centro de la ventana, desde el origen del plano)"),
    "rand": (u"Rand [cm]", u"Margin [cm]", u"Margen [cm]"),
    "rechts": (u"rechts frei [cm]", u"free right [cm]", u"libre dcha. [cm]"),
    "abstand": (u"Abstand [cm]", u"Gap [cm]", u"Separación [cm]"),
    "position_hinweis": (
        u"Zeichenfläche = Plankopf minus Rand, rechts zusätzlich frei für das "
        u"Schriftfeld. Mehrere Ansichten je Plan werden in Leserichtung "
        u"angeordnet, vorhandene Zeichnungen bleiben frei. Gemessen wird die "
        u"Zeichnung (Zuschneidebereich bzw. sichtbare Elemente), nicht der "
        u"größere Revit-Umriss. Auf einem vorhandenen Plan entsteht nie ein "
        u"neuer Plan.",
        u"Drawing area = title block minus margin, with extra free space on "
        u"the right for the title strip. Several views per sheet are arranged "
        u"in reading order around existing drawings. The drawing itself is "
        u"measured (crop region or visible elements), not the larger Revit "
        u"outline. An existing sheet never spawns a new sheet.",
        u"Área de dibujo = cajetín menos margen, con espacio libre adicional a "
        u"la derecha para el rótulo. Varias vistas por plano se ordenan en "
        u"orden de lectura, respetando los dibujos existentes. Se mide el "
        u"dibujo (región de recorte o elementos visibles), no el contorno "
        u"mayor de Revit. Un plano existente nunca genera un plano nuevo."),
    "nummer": (u"Plannummer (Muster)", u"Sheet number (pattern)",
               u"Número de plano (patrón)"),
    "name": (u"Planname (Muster)", u"Sheet name (pattern)",
             u"Nombre de plano (patrón)"),
    "muster_hinweis": (
        u"### = Zähler mit 3 Stellen, {Ansicht}, {Ebene}, {Maßstab} oder "
        u"{Parametername} der Ansicht. Bei mehreren Ansichten je Plan zählt "
        u"die erste.",
        u"### = 3-digit counter, {View}, {Level}, {Scale} or {parameter name} "
        u"of the view. With several views per sheet the first one counts.",
        u"### = contador de 3 cifras, {Vista}, {Nivel}, {Escala} o "
        u"{nombre de parámetro} de la vista. Con varias vistas por plano "
        u"cuenta la primera."),
    "ok": (u"Pläne erstellen", u"Create sheets", u"Crear planos"),
    "abbrechen": (u"Abbrechen", u"Cancel", u"Cancelar"),
}


#_________________________________________________________________________
# Daten sammeln
#_________________________________________________________________________

def platzierbar(el):
    return (isinstance(el, View)
            and not el.IsTemplate
            and el.CanBePrinted
            and not isinstance(el, ViewSheet)
            and el.ViewType != ViewType.Schedule)


def gewaehlte_ansichten(doc, uidoc):
    views = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
    views = [v for v in views if platzierbar(v)]
    if views:
        return views
    return forms.select_views(title=t(u"Ansichten für Pläne wählen",
                                      u"Select views for sheets",
                                      u"Seleccione vistas para planos"),
                              multiple=True,
                              filterfunc=platzierbar) or []


def ansichtsvorlagen(doc):
    return {v.Name: v for v in FilteredElementCollector(doc).OfClass(View)
            if v.IsTemplate}


#_________________________________________________________________________
# Dialog
#_________________________________________________________________________

class ViewToSheetDialog(forms.WPFWindow):
    def __init__(self, anzahl, plaene, vorlagen, koepfe, fenstertypen, cfg):
        pfad = os.path.join(os.path.dirname(__file__), "ui.xaml")
        with io.open(pfad, encoding="utf-8") as datei:
            xaml = uebersetze_xaml(datei.read(), XAML_TEXTE)
        forms.WPFWindow.__init__(self, xaml, literal_string=True)
        self.ergebnis = None
        self.lbl_info.Text = t(u"{} Ansicht(en) ausgewählt.",
                               u"{} view(s) selected.",
                               u"{} vista(s) seleccionadas.").format(anzahl)

        self._auswahl(self.cb_ziel, [NEUER_PLAN] + plaene, NEUER_PLAN)
        self._auswahl(self.cb_template, [KEINE_VORLAGE] + sorted(vorlagen),
                      cfg.get_option("vorlage", KEINE_VORLAGE))
        self._auswahl(self.cb_titleblock, sorted(koepfe),
                      cfg.get_option("planfamilie", ""))
        self._auswahl(self.cb_vptyp, [STANDARD_TYP] + sorted(fenstertypen),
                      cfg.get_option("fenstertyp", STANDARD_TYP))

        zusammen = cfg.get_option("anordnung", "einzeln") == "zusammen"
        self.rb_zusammen.IsChecked = zusammen
        self.rb_einzeln.IsChecked = not zusammen
        self.cb_ueberlappen.IsChecked = cfg.get_option("ueberlappen", "nein") == "ja"
        xy = cfg.get_option("position", "mitte") == "xy"
        self.rb_xy.IsChecked = xy
        self.rb_mitte.IsChecked = not xy

        self.tb_x.Text = cfg.get_option("x_cm", "-57")
        self.tb_y.Text = cfg.get_option("y_cm", "40")
        self.tb_rand.Text = cfg.get_option("rand_cm", "2")
        self.tb_rechts.Text = cfg.get_option("rechts_cm", "0")
        self.tb_abstand.Text = cfg.get_option("abstand_cm", "1")
        # frühere Version speicherte nur ein Präfix
        praefix = cfg.get_option("praefix", "AP")
        self.tb_nummer.Text = cfg.get_option(
            "nummer_muster", u"{}-###".format(praefix) if praefix else u"###")
        self.tb_name.Text = cfg.get_option("name_muster", u"{Ansicht}")

        for element in (self.rb_einzeln, self.rb_zusammen, self.rb_mitte, self.rb_xy):
            element.Checked += self.aktualisieren
        self.cb_ziel.SelectionChanged += self.aktualisieren
        self.btn_ok.Click += self.ok_click
        self.btn_cancel.Click += self.cancel_click
        self.aktualisieren(None, None)

    @staticmethod
    def _auswahl(box, eintraege, vorgabe):
        box.ItemsSource = eintraege
        box.SelectedItem = vorgabe
        if box.SelectedIndex < 0:
            box.SelectedIndex = 0

    def aktualisieren(self, sender, args):
        neu = self.cb_ziel.SelectedIndex <= 0
        if not neu and not self.rb_zusammen.IsChecked:
            self.rb_zusammen.IsChecked = True
        self.rb_einzeln.IsEnabled = neu
        einzeln = bool(self.rb_einzeln.IsChecked)
        xy = einzeln and bool(self.rb_xy.IsChecked)
        self.rb_mitte.IsEnabled = self.rb_xy.IsEnabled = einzeln
        self.tb_x.IsEnabled = self.tb_y.IsEnabled = xy
        self.tb_rand.IsEnabled = self.tb_rechts.IsEnabled = not xy
        self.tb_abstand.IsEnabled = not einzeln
        self.cb_ueberlappen.IsEnabled = not einzeln
        if not einzeln:
            # mehrere je Plan: nach dem Platzieren öffnet sich die Vorschau
            self.btn_ok.Content = t(u"Weiter zur Vorschau …", u"Continue to preview …",
                                    u"Continuar a vista previa …")
            self.btn_ok.Width = 170
        else:
            self.btn_ok.Content = tt(XAML_TEXTE["ok"])
            self.btn_ok.Width = 120

    def ok_click(self, sender, args):
        try:
            werte = [rv.zahl(feld.Text) for feld in
                     (self.tb_x, self.tb_y, self.tb_rand, self.tb_rechts, self.tb_abstand)]
        except ValueError:
            forms.alert(t(u"X, Y, Rand, rechts frei und Abstand müssen Zahlen "
                          u"sein (in cm).",
                          u"X, Y, margin, free right and gap must be numbers "
                          u"(in cm).",
                          u"X, Y, margen, libre dcha. y separación deben ser "
                          u"números (en cm)."),
                        title=t(u"Eingabe prüfen", u"Check input", u"Revise la entrada"))
            return
        if not self.tb_nummer.Text.strip():
            forms.alert(t(u"Bitte ein Muster für die Plannummer angeben.",
                          u"Please enter a sheet number pattern.",
                          u"Indique un patrón para el número de plano."),
                        title=t(u"Eingabe prüfen", u"Check input", u"Revise la entrada"))
            return
        self.ergebnis = {
            "zusammen": bool(self.rb_zusammen.IsChecked),
            "ueberlappen": bool(self.cb_ueberlappen.IsChecked),
            "ziel": self.cb_ziel.SelectedItem,
            "vorlage": self.cb_template.SelectedItem,
            "planfamilie": self.cb_titleblock.SelectedItem,
            "fenstertyp": self.cb_vptyp.SelectedItem,
            "xy": bool(self.rb_xy.IsChecked),
            "x_cm": werte[0], "y_cm": werte[1], "rand_cm": werte[2],
            "rechts_cm": werte[3], "abstand_cm": werte[4],
            "nummer_muster": self.tb_nummer.Text.strip(),
            "name_muster": self.tb_name.Text.strip(),
        }
        self.Close()

    def cancel_click(self, sender, args):
        self.Close()


def einstellungen_speichern(cfg, wahl):
    cfg.anordnung = "zusammen" if wahl["zusammen"] else "einzeln"
    cfg.ueberlappen = "ja" if wahl["ueberlappen"] else "nein"
    cfg.vorlage = wahl["vorlage"]
    cfg.planfamilie = wahl["planfamilie"]
    cfg.fenstertyp = wahl["fenstertyp"]
    cfg.position = "xy" if wahl["xy"] else "mitte"
    for schluessel in ("x_cm", "y_cm", "rand_cm", "rechts_cm", "abstand_cm"):
        setattr(cfg, schluessel, str(wahl[schluessel]))
    cfg.nummer_muster = wahl["nummer_muster"]
    cfg.name_muster = wahl["name_muster"]
    script.save_config()


#_________________________________________________________________________
# Pläne erstellen
#_________________________________________________________________________

class Erstellung(object):
    """Setzt Ansichten auf Pläne - läuft in einer offenen Transaktion."""

    def __init__(self, doc, wahl, vorlage, kopf, typ, zielplan):
        self.doc = doc
        self.wahl = wahl
        self.vorlage = vorlage
        self.kopf = kopf
        self.typ = typ
        self.zielplan = zielplan
        self.vergeben = rv.vergebene_nummern(doc)
        self.neue_plaene = []
        self.platziert = 0
        self.hinweise = []
        self.fehlende_platzhalter = set()
        self.masse = []
        # (Plan, Ansichtsfenster, Umriss) aller neu gesetzten - für die Vorschau
        self.gesetzt = []

    # ---- Hilfen

    def flaeche(self, plan):
        flaeche = lg.zeichenflaeche(rv.plankopf_rechteck(self.doc, plan),
                                    rv.fuss(self.wahl["rand_cm"]),
                                    rv.fuss(self.wahl["rechts_cm"]))
        self.masse.append(t(u"Zeichenfläche {}: {}", u"Drawing area {}: {}",
                            u"Área de dibujo {}: {}").format(
            plan.SheetNumber, rv.masse_text(flaeche)))
        return flaeche

    def messung_merken(self, ansicht, fenster):
        try:
            self.masse.append(u"{}: {}".format(ansicht.Name,
                                               rv.messung_text(self.doc, fenster)))
        except Exception:
            pass

    def neuer_plan(self):
        plan = ViewSheet.Create(self.doc, self.kopf.Id)
        self.doc.Regenerate()
        # vorläufige Nummer von Revit nicht per Muster doppelt vergeben
        self.vergeben.add(lg.norm(plan.SheetNumber))
        self.neue_plaene.append(plan)
        return plan

    def benennen(self, plan, ansicht):
        wert = rv.ansichts_werte(ansicht)
        nummer, fehlend = lg.nummer_aus_muster(self.wahl["nummer_muster"], wert,
                                               self.vergeben)
        name, fehlend_name = lg.muster_anwenden(self.wahl["name_muster"], wert)
        self.fehlende_platzhalter.update(fehlend + fehlend_name)
        plan.SheetNumber = nummer
        plan.Name = name.strip() or ansicht.Name

    def vorlage_anwenden(self, ansicht):
        if self.vorlage is None:
            return
        try:
            ansicht.ViewTemplateId = self.vorlage.Id
        except Exception as fehler:
            self.hinweise.append(t(
                u"{}: Vorlage nicht anwendbar, ohne Vorlage platziert ({})",
                u"{}: template not applicable, placed without template ({})",
                u"{}: plantilla no aplicable, colocada sin plantilla ({})")
                .format(ansicht.Name, fehler))

    def schon_platziert(self, ansicht):
        self.hinweise.append(t(u"{}: bereits auf einem Plan platziert",
                               u"{}: already placed on a sheet",
                               u"{}: ya colocada en un plano").format(ansicht.Name))

    def pruefe_groesse(self, ansicht, umriss, flaeche):
        """Hinweis mit Maßen, wenn die Ansicht grösser als die Zeichenfläche ist."""
        if lg.passt_in(lg.breite(umriss), lg.hoehe(umriss), flaeche):
            return
        self.hinweise.append(t(
            u"{}: größer als die Zeichenfläche (Ansicht {:.0f} x {:.0f} cm, "
            u"Zeichenfläche {:.0f} x {:.0f} cm) - mittig platziert",
            u"{}: larger than the drawing area (view {:.0f} x {:.0f} cm, "
            u"drawing area {:.0f} x {:.0f} cm) - placed in the centre",
            u"{}: mayor que el área de dibujo (vista {:.0f} x {:.0f} cm, "
            u"área {:.0f} x {:.0f} cm): colocada en el centro").format(
            ansicht.Name, rv.cm(lg.breite(umriss)), rv.cm(lg.hoehe(umriss)),
            rv.cm(lg.breite(flaeche)), rv.cm(lg.hoehe(flaeche))))

    def weiterer_plan(self, ansicht, voll, plan):
        voller_plan, umriss, flaeche = voll
        self.hinweise.append(t(
            u"{}: passte nicht mehr auf {} (Ansicht {:.0f} x {:.0f} cm, "
            u"Zeichenfläche {:.0f} x {:.0f} cm) - neuer Plan {}",
            u"{}: did not fit on {} any more (view {:.0f} x {:.0f} cm, "
            u"drawing area {:.0f} x {:.0f} cm) - new sheet {}",
            u"{}: ya no cabía en {} (vista {:.0f} x {:.0f} cm, área "
            u"{:.0f} x {:.0f} cm): plano nuevo {}").format(
            ansicht.Name, voller_plan,
            rv.cm(lg.breite(umriss)), rv.cm(lg.hoehe(umriss)),
            rv.cm(lg.breite(flaeche)), rv.cm(lg.hoehe(flaeche)), plan.SheetNumber))

    def kein_platz(self, ansicht, umriss, flaeche, plan):
        self.hinweise.append(t(
            u"{}: kein freier Platz auf {} (Ansicht {:.0f} x {:.0f} cm, "
            u"Zeichenfläche {:.0f} x {:.0f} cm) - rechts neben den Plankopf "
            u"gelegt, bitte von Hand verschieben",
            u"{}: no free space on {} (view {:.0f} x {:.0f} cm, drawing area "
            u"{:.0f} x {:.0f} cm) - placed to the right of the title block, "
            u"please move it by hand",
            u"{}: no hay espacio libre en {} (vista {:.0f} x {:.0f} cm, área "
            u"{:.0f} x {:.0f} cm): colocada a la derecha del cajetín, muévala "
            u"a mano").format(
            ansicht.Name, plan.SheetNumber,
            rv.cm(lg.breite(umriss)), rv.cm(lg.hoehe(umriss)),
            rv.cm(lg.breite(flaeche)), rv.cm(lg.hoehe(flaeche))))

    # ---- Ein Plan je Ansicht

    def einzeln(self, ansichten):
        for ansicht in ansichten:
            self.vorlage_anwenden(ansicht)
            plan = self.neuer_plan()
            try:
                if not Viewport.CanAddViewToSheet(self.doc, plan.Id, ansicht.Id):
                    self.schon_platziert(ansicht)
                    continue
                if self.wahl["xy"]:
                    ziel = (rv.fuss(self.wahl["x_cm"]), rv.fuss(self.wahl["y_cm"]))
                else:
                    ziel = self.flaeche(plan)
                rv.platzieren(self.doc, plan, ansicht, ziel, self.typ)
                self.benennen(plan, ansicht)
                self.platziert += 1
            except Exception as fehler:
                self.hinweise.append(u"{}: {}".format(ansicht.Name, fehler))

    # ---- Mehrere Ansichten je Plan

    def zusammen(self, ansichten):
        if self.zielplan is not None:
            self.auf_vorhandenen_plan(ansichten)
            return
        plan, flaeche, belegt = None, None, None
        for ansicht in ansichten:
            try:
                self.vorlage_anwenden(ansicht)
                if plan is None:
                    plan = self.neuer_plan()
                    flaeche, belegt = self.flaeche(plan), []
                if not Viewport.CanAddViewToSheet(self.doc, plan.Id, ansicht.Id):
                    self.schon_platziert(ansicht)
                    continue
                # Auf einem leeren Plan immer platzieren - ein weiterer Plan
                # hilft nur, wenn schon etwas darauf liegt
                fenster, umriss = rv.platzieren(self.doc, plan, ansicht, flaeche,
                                                self.typ, belegt,
                                                rv.fuss(self.wahl["abstand_cm"]),
                                                erzwingen=not belegt,
                                                ueberlappen=self.wahl["ueberlappen"])
                if fenster is not None and not belegt:
                    self.pruefe_groesse(ansicht, umriss, flaeche)
                voll = None
                if fenster is None:
                    # Plan voll: nächster Plan
                    voll = (plan.SheetNumber, umriss, flaeche)
                    plan = self.neuer_plan()
                    flaeche, belegt = self.flaeche(plan), []
                    fenster, umriss = rv.platzieren(self.doc, plan, ansicht, flaeche,
                                                    self.typ, belegt, erzwingen=True)
                    self.pruefe_groesse(ansicht, umriss, flaeche)
                if not belegt:
                    self.benennen(plan, ansicht)
                if voll is not None:
                    self.weiterer_plan(ansicht, voll, plan)
                belegt.append(umriss)
                self.platziert += 1
                self.gesetzt.append((plan, fenster, umriss))
                self.messung_merken(ansicht, fenster)
            except Exception as fehler:
                self.hinweise.append(u"{}: {}".format(ansicht.Name, fehler))

    # ---- Auf einen vorhandenen Plan

    def auf_vorhandenen_plan(self, ansichten):
        """Nie einen neuen Plan erstellen: ohne freien Platz kommt die Ansicht
        rechts neben den Plankopf und wird gemeldet."""
        plan = self.zielplan
        kopf = rv.plankopf_rechteck(self.doc, plan)
        flaeche = self.flaeche(plan)
        belegt = rv.belegte_rechtecke(self.doc, plan)
        abstand = rv.fuss(self.wahl["abstand_cm"])
        ablage = (kopf[2] + rv.fuss(5), kopf[1] - 100.0, kopf[2] + 100.0, kopf[3])
        abgelegt = []
        for ansicht in ansichten:
            try:
                self.vorlage_anwenden(ansicht)
                if not Viewport.CanAddViewToSheet(self.doc, plan.Id, ansicht.Id):
                    self.schon_platziert(ansicht)
                    continue
                fenster, umriss = rv.platzieren(self.doc, plan, ansicht, flaeche,
                                                self.typ, belegt, abstand,
                                                ueberlappen=self.wahl["ueberlappen"])
                if fenster is None:
                    fenster, umriss = rv.platzieren(self.doc, plan, ansicht, ablage,
                                                    self.typ, abgelegt, abstand,
                                                    erzwingen=True)
                    abgelegt.append(umriss)
                    self.kein_platz(ansicht, umriss, flaeche, plan)
                else:
                    belegt.append(umriss)
                self.platziert += 1
                self.gesetzt.append((plan, fenster, umriss))
                self.messung_merken(ansicht, fenster)
            except Exception as fehler:
                self.hinweise.append(u"{}: {}".format(ansicht.Name, fehler))

    def aufraeumen(self):
        """Neue Pläne ohne Ansicht wieder löschen (z.B. Ansicht schon platziert)."""
        behalten = []
        for plan in self.neue_plaene:
            if plan.GetAllViewports().Count:
                behalten.append(plan)
            else:
                self.doc.Delete(plan.Id)
        self.neue_plaene = behalten


def vorschau_zeigen(doc, erstellung, wahl):
    """Vorschau je Plan mit neu gesetzten Ansichten.

    Liefert [(Ansichtsfenster, alter Umriss, neuer Umriss)] oder None, wenn
    abgebrochen wurde.
    """
    plaene, je_plan = [], {}
    for index, (plan, fenster, umriss) in enumerate(erstellung.gesetzt):
        schluessel = rv.id_wert(plan.Id)
        if schluessel not in je_plan:
            je_plan[schluessel] = []
            plaene.append(plan)
        je_plan[schluessel].append(index)

    bewegungen = []
    for plan in plaene:
        indizes = je_plan[rv.id_wert(plan.Id)]
        neue_ids = [erstellung.gesetzt[i][1].Id for i in indizes]
        fest = [vorschau.Eintrag(None, name, r, beweglich=False)
                for name, r in rv.belegte_eintraege(doc, plan, neue_ids)]
        beweglich = []
        for i in indizes:
            _, fenster, umriss = erstellung.gesetzt[i]
            beweglich.append(vorschau.Eintrag(
                i, rv.vorschau_name(doc.GetElement(fenster.ViewId)), umriss))
        kopf = rv.plankopf_rechteck(doc, plan)
        flaeche = lg.zeichenflaeche(kopf, rv.fuss(wahl["rand_cm"]),
                                    rv.fuss(wahl["rechts_cm"]))
        titel = t(u"Vorschau - {}", u"Preview - {}", u"Vista previa - {}").format(
            rv.plan_text(plan))
        ergebnis = vorschau.zeigen(titel, kopf, flaeche, fest + beweglich,
                                   rv.fuss(wahl["abstand_cm"]),
                                   rv.plankopf_linien(doc, plan))
        if ergebnis is None:
            return None
        for i, neu in ergebnis.items():
            _, fenster, umriss = erstellung.gesetzt[i]
            bewegungen.append((fenster, umriss, neu))
    return bewegungen


def main():
    doc, uidoc = revit.doc, revit.uidoc
    views = gewaehlte_ansichten(doc, uidoc)
    if not views:
        forms.alert(t(u"Keine platzierbaren Ansichten ausgewählt.",
                      u"No placeable views selected.",
                      u"No hay vistas colocables seleccionadas."), title=TITEL)
        return

    koepfe = rv.plankoepfe(doc)
    if not koepfe:
        forms.alert(t(u"Keine Planfamilie (Plankopf) im Projekt gefunden.",
                      u"No title block family found in the project.",
                      u"No se encontró ninguna familia de cajetín en el proyecto."),
                    title=TITEL)
        return
    vorlagen = ansichtsvorlagen(doc)
    fenstertypen = rv.ansichtsfenster_typen(doc)
    plaene = {rv.plan_text(p): p for p in rv.plaene(doc)}
    plan_texte = [rv.plan_text(p) for p in rv.plaene(doc)]

    cfg = script.get_config()
    dlg = ViewToSheetDialog(len(views), plan_texte, vorlagen, koepfe,
                            fenstertypen, cfg)
    dlg.ShowDialog()
    wahl = dlg.ergebnis
    if not wahl:
        return
    einstellungen_speichern(cfg, wahl)

    erstellung = Erstellung(doc, wahl, vorlagen.get(wahl["vorlage"]),
                            koepfe[wahl["planfamilie"]],
                            fenstertypen.get(wahl["fenstertyp"]),
                            plaene.get(wahl["ziel"]))
    # Eine Gruppe: vorläufig platzieren, in der Vorschau verschieben, dann
    # übernehmen - "Abbrechen" in der Vorschau nimmt alles zurück
    gruppe = TransactionGroup(doc, TITEL)
    gruppe.Start()
    transaktion = Transaction(doc, TITEL)
    try:
        transaktion.Start()
        if wahl["zusammen"]:
            erstellung.zusammen(views)
        else:
            erstellung.einzeln(views)
        erstellung.aufraeumen()
        transaktion.Commit()

        if wahl["zusammen"] and erstellung.gesetzt:
            bewegungen = vorschau_zeigen(doc, erstellung, wahl)
            if bewegungen is None:
                gruppe.RollBack()
                return
            transaktion = Transaction(doc, TITEL)
            transaktion.Start()
            for fenster, alt, neu in bewegungen:
                if alt != neu:
                    rv.verschiebe_rechteck_auf(fenster, alt, neu)
            transaktion.Commit()
        gruppe.Assimilate()
    except Exception as fehler:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        if gruppe.HasStarted() and not gruppe.HasEnded():
            gruppe.RollBack()
        forms.alert(t(u"Erstellen fehlgeschlagen, nichts wurde geändert.",
                      u"Creation failed, nothing was changed.",
                      u"Error al crear, no se cambió nada."),
                    sub_msg=u"{}".format(fehler), title=TITEL)
        return

    neue = erstellung.neue_plaene
    meldung = t(u"{} Ansicht(en) platziert, {} Plan/Pläne erstellt.",
                u"{} view(s) placed, {} sheet(s) created.",
                u"{} vista(s) colocadas, {} plano(s) creados.").format(
        erstellung.platziert, len(neue))
    if neue:
        meldung += u"\n" + u", ".join(p.SheetNumber for p in neue[:30])
    hinweise = list(erstellung.hinweise)
    if erstellung.fehlende_platzhalter:
        hinweise.insert(0, t(u"Unbekannte Platzhalter (leer eingesetzt): {}",
                             u"Unknown placeholders (left empty): {}",
                             u"Marcadores desconocidos (vacíos): {}").format(
            u", ".join(sorted(erstellung.fehlende_platzhalter))))
    if hinweise:
        meldung += t(u"\n\nHinweise:\n", u"\n\nNotes:\n", u"\n\nNotas:\n") \
            + u"\n".join(hinweise[:20])
    if wahl["zusammen"] and erstellung.masse:
        meldung += t(u"\n\nMaße:\n", u"\n\nSizes:\n", u"\n\nMedidas:\n") \
            + u"\n".join(erstellung.masse[:20])
    forms.alert(meldung, title=TITEL)


main()
