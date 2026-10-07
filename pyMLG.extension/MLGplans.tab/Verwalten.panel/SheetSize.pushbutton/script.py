# -*- coding: utf-8 -*-
# Ändert die Grösse vieler Pläne auf einmal: neuer Plankopf-Typ und/oder
# Format-Parameter des Plankopfs (Ja/Nein, Zahl, Länge - z.B. A0/A1/A2 in
# einer Familie). Danach wahlweise
#   - Inhalt mitnehmen: alles auf dem Plan behält den Abstand zur gewählten
#     Ecke (oder Mitte) des Plankopfs
#   - wie Vorbild-Plan ausrichten: Ansichtsfenster und Bauteillisten liegen
#     wie auf dem Vorbild (wie "Ansichten ausrichten")
# "Vom Vorbild übernehmen" holt Plankopf-Typ und Format eines Plans, den
# man schon von Hand umgestellt hat. Die Technik steckt in
# lib/mlg_plaene/plangroesse.py.
# (Kommentar statt Docstring: pyRevit liest unter IronPython den
#  Docstring als Tooltip und scheitert dabei an Umlauten.)

__title__ = "SheetSize"
__author__ = "Manuel"

import io
import os

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import SubTransaction, Transaction, ViewSheet
from System.Windows import GridLength, GridUnitType, TextWrapping, \
    Thickness, VerticalAlignment, Visibility
from System.Windows.Controls import CheckBox, ColumnDefinition, ComboBox, \
    Grid, TextBlock, TextBox
from pyrevit import forms, revit, script

from mlg_plaene import ausrichten as ar
from mlg_plaene import logik as lg
from mlg_plaene import plangroesse as pg
from mlg_plaene import revit as rv
from mlg_sprache import t, uebersetze_xaml

TITEL = t(u"Plangröße ändern", u"Change Sheet Size", u"Cambiar tamaño de plano")
UNVERAENDERT = t(u"<unverändert>", u"<unchanged>", u"<sin cambios>")
JA, NEIN = t(u"Ja", u"Yes", u"Sí"), t(u"Nein", u"No", u"No")

ANKER_TEXTE = [
    (lg.OBEN_LINKS, t(u"Ecke oben links", u"top left corner", u"esquina superior izquierda")),
    (lg.OBEN_RECHTS, t(u"Ecke oben rechts", u"top right corner",
                       u"esquina superior derecha")),
    (lg.UNTEN_LINKS, t(u"Ecke unten links", u"bottom left corner",
                       u"esquina inferior izquierda")),
    (lg.UNTEN_RECHTS, t(u"Ecke unten rechts (Schriftfeld)", u"bottom right corner (title)",
                        u"esquina inferior derecha (rótulo)")),
    (lg.ANKER_MITTE, t(u"Mitte des Plans", u"centre of the sheet", u"centro del plano")),
]

XAML_TEXTE = {
    "titel": (u"Plangröße ändern", u"Change Sheet Size", u"Cambiar tamaño de plano"),
    "plaene": (u"Pläne", u"Sheets", u"Planos"),
    "suche": (u"Nummer oder Name filtern", u"Filter number or name",
              u"Filtrar número o nombre"),
    "alle": (u"Alle", u"All", u"Todos"),
    "keine": (u"Keine", u"None", u"Ninguno"),
    "vorbild": (u"Vorbild-Plan", u"Reference sheet", u"Plano de referencia"),
    "uebernehmen": (u"Format übernehmen", u"Take format", u"Tomar formato"),
    "uebernehmen_tip": (u"Plankopf-Typ und Format-Parameter vom Vorbild-Plan holen",
                        u"Take title block type and format parameters from the "
                        u"reference sheet",
                        u"Tomar tipo de cajetín y parámetros de formato del plano "
                        u"de referencia"),
    "format": (u"Neues Format", u"New format", u"Nuevo formato"),
    "plankopf": (u"Plankopf-Typ", u"Title block type", u"Tipo de cajetín"),
    "parameter": (u"Format-Parameter des Plankopfs - nur angehakte werden gesetzt",
                  u"Format parameters of the title block - only ticked ones are set",
                  u"Parámetros de formato del cajetín: sólo se asignan los marcados"),
    "ausrichten": (u"Danach ausrichten", u"Then align", u"Después alinear"),
    "mitnehmen": (u"Planinhalt mitnehmen - gleicher Abstand zu",
                  u"Move sheet content along - same distance to",
                  u"Mover el contenido del plano - misma distancia a"),
    "wie_vorbild": (u"Ansichten und Bauteillisten wie auf dem Vorbild-Plan",
                    u"Views and schedules like on the reference sheet",
                    u"Vistas y tablas como en el plano de referencia"),
    "ok": (u"Ändern", u"Change", u"Cambiar"),
    "abbrechen": (u"Abbrechen", u"Cancel", u"Cancelar"),
}


def beschreibungen(doc, plaene):
    """{Plan-Id: "Familie: Typ · 84 x 59 cm"}. Gemessen wird je Format nur
    einmal - Pläne mit gleichem Typ und gleichen Format-Parametern sind
    gleich gross."""
    groessen, texte = {}, {}
    for plan in plaene:
        koepfe = pg.plankoepfe_auf(doc, plan)
        if not koepfe:
            texte[rv.id_wert(plan.Id)] = t(u"kein Plankopf", u"no title block",
                                           u"sin cajetín")
            continue
        schluessel = pg.format_schluessel(doc, koepfe[0])
        if schluessel not in groessen:
            groessen[schluessel] = rv.masse_text(rv.plankopf_rechteck(doc, plan))
        texte[rv.id_wert(plan.Id)] = u"{} · {}".format(pg.typ_name(doc, koepfe[0]),
                                                      groessen[schluessel])
    return texte


class PlangroesseDialog(forms.WPFWindow):
    def __init__(self, doc, plaene, vorgewaehlt, vorbild, koepfe, cfg):
        pfad = os.path.join(os.path.dirname(__file__), "ui.xaml")
        with io.open(pfad, encoding="utf-8") as datei:
            xaml = uebersetze_xaml(datei.read(), XAML_TEXTE)
        forms.WPFWindow.__init__(self, xaml, literal_string=True)
        self.doc = doc
        self.plaene = plaene
        self.koepfe = koepfe
        self.ergebnis = None
        self._zeilen = []          # [(Name, Art, Haken, Eingabe)]
        self._familie = None       # Familie, aus der die Zeilen stammen
        self._fuellen = False

        texte = beschreibungen(doc, plaene)
        self._eintraege = []       # [(Plan, Haken, Suchtext)]
        for plan in plaene:
            haken = CheckBox()
            haken.Content = u"{}    ({})".format(rv.plan_text(plan),
                                                 texte[rv.id_wert(plan.Id)])
            haken.IsChecked = rv.id_wert(plan.Id) in vorgewaehlt
            haken.Checked += self.aktualisieren
            haken.Unchecked += self.aktualisieren
            self.lb_plaene.Items.Add(haken)
            self._eintraege.append((plan, haken, rv.plan_text(plan).lower()))

        self.cb_vorbild.ItemsSource = [rv.plan_text(p) for p in plaene]
        self.cb_vorbild.SelectedIndex = plaene.index(vorbild)
        self.cb_typ.ItemsSource = [UNVERAENDERT] + sorted(koepfe)
        self.cb_typ.SelectedIndex = 0
        self._zeilen_aus(self._vorbild_kopf(), angehakt=False)

        self.cb_anker.ItemsSource = [text for _, text in ANKER_TEXTE]
        anker = cfg.get_option("anker", lg.OBEN_LINKS)
        schluessel = [a for a, _ in ANKER_TEXTE]
        self.cb_anker.SelectedIndex = schluessel.index(anker) if anker in schluessel else 0
        self.cb_mitnehmen.IsChecked = cfg.get_option("mitnehmen", "ja") == "ja"
        self.cb_art.ItemsSource = [text for text, _ in ar.art_texte()]
        self.cb_art.SelectedIndex = 1 if cfg.get_option("art", ar.GEBAEUDE) == ar.FENSTER else 0
        self.cb_wie_vorbild.IsChecked = cfg.get_option("wie_vorbild", "nein") == "ja"

        self.tb_suche.TextChanged += self.suchen
        self.btn_alle.Click += lambda s, a: self._alle(True)
        self.btn_keine.Click += lambda s, a: self._alle(False)
        self.btn_vom_vorbild.Click += self.vom_vorbild
        self.cb_typ.SelectionChanged += self.typ_gewechselt
        self.cb_vorbild.SelectionChanged += self.aktualisieren
        for haken in (self.cb_mitnehmen, self.cb_wie_vorbild):
            haken.Checked += self.aktualisieren
            haken.Unchecked += self.aktualisieren
        self.btn_ok.Click += self.ok_click
        self.btn_cancel.Click += self.cancel_click
        self.aktualisieren(None, None)

    # ------------------------------------------------------------ Pläne

    def _gewaehlt(self):
        return [plan for plan, haken, _ in self._eintraege if haken.IsChecked]

    def _vorbild(self):
        return self.plaene[self.cb_vorbild.SelectedIndex]

    def _vorbild_kopf(self):
        koepfe = pg.plankoepfe_auf(self.doc, self._vorbild())
        return koepfe[0] if koepfe else None

    def suchen(self, sender, args):
        suche = self.tb_suche.Text.strip().lower()
        for _, haken, text in self._eintraege:
            haken.Visibility = Visibility.Visible if suche in text else Visibility.Collapsed

    def _alle(self, wert):
        """Alle sichtbaren (nach dem Filter) an- oder abwählen."""
        for _, haken, _ in self._eintraege:
            if haken.Visibility == Visibility.Visible:
                haken.IsChecked = wert

    # ------------------------------------------------------------ Format

    def _zeilen_aus(self, kopf, angehakt):
        """Baut die Parameterzeilen aus einem Plankopf-Exemplar."""
        self._fuellen = True
        self.sp_parameter.Children.Clear()
        self._zeilen = []
        self._familie = None
        if kopf is None:
            hinweis = TextBlock()
            hinweis.Text = t(u"Kein Plankopf dieser Familie im Projekt platziert - "
                             u"es wird nur der Typ getauscht.",
                             u"No title block of this family placed in the project - "
                             u"only the type is changed.",
                             u"No hay ningún cajetín de esta familia colocado: sólo se "
                             u"cambia el tipo.")
            hinweis.TextWrapping = TextWrapping.Wrap
            self.sp_parameter.Children.Add(hinweis)
            self._fuellen = False
            return
        self._familie = self.doc.GetElement(kopf.GetTypeId()).FamilyName
        for name, art, wert in pg.format_parameter(kopf):
            self._zeile(name, art, wert, angehakt)
        if not self._zeilen:
            hinweis = TextBlock()
            hinweis.Text = t(u"Keine Format-Parameter in dieser Familie.",
                             u"No format parameters in this family.",
                             u"No hay parámetros de formato en esta familia.")
            self.sp_parameter.Children.Add(hinweis)
        self._fuellen = False

    def _zeile(self, name, art, wert, angehakt):
        zeile = Grid()
        zeile.Margin = Thickness(0, 1, 0, 1)
        for breite in (GridLength.Auto, GridLength(1, GridUnitType.Star), GridLength(130)):
            spalte = ColumnDefinition()
            spalte.Width = breite
            zeile.ColumnDefinitions.Add(spalte)
        haken = CheckBox()
        haken.IsChecked = angehakt
        haken.VerticalAlignment = VerticalAlignment.Center
        haken.Checked += self.aktualisieren
        haken.Unchecked += self.aktualisieren
        text = TextBlock()
        text.Text = name
        text.Margin = Thickness(4, 0, 6, 0)
        text.VerticalAlignment = VerticalAlignment.Center
        if art == pg.JA_NEIN:
            eingabe = ComboBox()
            eingabe.ItemsSource = [JA, NEIN]
            eingabe.SelectedIndex = 0 if wert == u"1" else 1
            eingabe.SelectionChanged += lambda s, a: self._geaendert(haken)
        else:
            eingabe = TextBox()
            eingabe.Text = wert
            eingabe.TextChanged += lambda s, a: self._geaendert(haken)
        for spalte, element in enumerate((haken, text, eingabe)):
            Grid.SetColumn(element, spalte)
            zeile.Children.Add(element)
        self.sp_parameter.Children.Add(zeile)
        self._zeilen.append((name, art, haken, eingabe))

    def _geaendert(self, haken):
        """Wer einen Wert ändert, will ihn auch setzen."""
        if not self._fuellen:
            haken.IsChecked = True

    def _werte(self):
        werte = []
        for name, art, haken, eingabe in self._zeilen:
            if not haken.IsChecked:
                continue
            if art == pg.JA_NEIN:
                text = u"1" if eingabe.SelectedIndex == 0 else u"0"
            else:
                text = eingabe.Text.strip()
            werte.append((name, art, text))
        return werte

    def vom_vorbild(self, sender, args):
        kopf = self._vorbild_kopf()
        if kopf is None:
            forms.alert(t(u"Der Vorbild-Plan hat keinen Plankopf.",
                          u"The reference sheet has no title block.",
                          u"El plano de referencia no tiene cajetín."), title=TITEL)
            return
        self._fuellen = True
        self.cb_typ.SelectedItem = pg.typ_name(self.doc, kopf)
        self._fuellen = False
        self._zeilen_aus(kopf, angehakt=True)
        self.aktualisieren(None, None)

    def typ_gewechselt(self, sender, args):
        if self._fuellen:
            return
        typ = self.koepfe.get(self.cb_typ.SelectedItem)
        if typ is None:
            kopf = self._vorbild_kopf()
        elif typ.FamilyName == self._familie:
            self.aktualisieren(None, None)
            return
        else:
            kopf = pg.beispiel_kopf(self.doc, typ)
        self._zeilen_aus(kopf, angehakt=False)
        self.aktualisieren(None, None)

    # ------------------------------------------------------------ Ergebnis

    def _aenderung(self):
        anker = ANKER_TEXTE[self.cb_anker.SelectedIndex][0] \
            if self.cb_mitnehmen.IsChecked else None
        return pg.Aenderung(self.koepfe.get(self.cb_typ.SelectedItem), self._werte(), anker)

    def aktualisieren(self, sender, args):
        gewaehlt = self._gewaehlt()
        self.lbl_anzahl.Text = t(u"{} von {} Plänen gewählt", u"{} of {} sheets selected",
                                 u"{} de {} planos seleccionados").format(
            len(gewaehlt), len(self.plaene))
        self.cb_anker.IsEnabled = bool(self.cb_mitnehmen.IsChecked)
        self.cb_art.IsEnabled = bool(self.cb_wie_vorbild.IsChecked)

        aenderung = self._aenderung()
        teile = []
        if aenderung.typ is not None:
            teile.append(t(u"Plankopf → {}", u"title block → {}",
                           u"cajetín → {}").format(self.cb_typ.SelectedItem))
        if aenderung.werte:
            teile.append(u", ".join(u"{} = {}".format(
                n, (JA if w == u"1" else NEIN) if a == pg.JA_NEIN else w)
                for n, a, w in aenderung.werte))
        if self.cb_wie_vorbild.IsChecked:
            teile.append(t(u"ausrichten wie {}", u"align like {}",
                           u"alinear como {}").format(self._vorbild().SheetNumber))

        if not gewaehlt:
            status = t(u"Bitte Pläne wählen.", u"Please select sheets.",
                       u"Seleccione planos.")
        elif not teile:
            status = t(u"Nichts zu ändern: Plankopf-Typ wählen, Parameter anhaken "
                       u"oder wie Vorbild ausrichten.",
                       u"Nothing to change: choose a title block type, tick parameters "
                       u"or align like the reference.",
                       u"Nada que cambiar: elija un tipo de cajetín, marque parámetros "
                       u"o alinee como la referencia.")
        else:
            status = None
        self.btn_ok.IsEnabled = status is None
        self.lbl_status.Text = status or u"{}: {}".format(
            t(u"{} Plan/Pläne", u"{} sheet(s)", u"{} plano(s)").format(len(gewaehlt)),
            u" · ".join(teile))

    def ok_click(self, sender, args):
        aenderung = self._aenderung()
        leer = [n for n, a, w in aenderung.werte if a != pg.JA_NEIN and not w]
        if leer:
            forms.alert(t(u"Bitte einen Wert eingeben für: {}", u"Please enter a value for: {}",
                          u"Indique un valor para: {}").format(u", ".join(leer)),
                        title=TITEL)
            return
        self.ergebnis = {
            "plaene": self._gewaehlt(),
            "aenderung": aenderung,
            "vorbild": self._vorbild() if self.cb_wie_vorbild.IsChecked else None,
            "art": dict(ar.art_texte()).get(self.cb_art.SelectedItem, ar.GEBAEUDE),
            "anker": ANKER_TEXTE[self.cb_anker.SelectedIndex][0],
            "mitnehmen": bool(self.cb_mitnehmen.IsChecked),
        }
        self.Close()

    def cancel_click(self, sender, args):
        self.Close()


def einstellungen_speichern(cfg, wahl):
    cfg.anker = wahl["anker"]
    cfg.mitnehmen = "ja" if wahl["mitnehmen"] else "nein"
    cfg.wie_vorbild = "ja" if wahl["vorbild"] is not None else "nein"
    cfg.art = wahl["art"]
    script.save_config()


def ausfuehren(doc, wahl):
    """Ändert und richtet aus. Liefert (geändert, ausgerichtet, Hinweise)."""
    aenderung, vorbild = wahl["aenderung"], wahl["vorbild"]
    geaendert, ausgerichtet, hinweise = 0, 0, []

    def schritt(plan, arbeit):
        # je Plan eine Untertransaktion: ein sperriger Plan nimmt nur sich
        # selbst zurück, nicht alle anderen
        unter = SubTransaction(doc)
        unter.Start()
        try:
            ergebnis, texte = arbeit()
            unter.Commit()
        except Exception as fehler:
            unter.RollBack()
            ergebnis, texte = None, [t(u"fehlgeschlagen, unverändert: {}",
                                       u"failed, unchanged: {}",
                                       u"error, sin cambios: {}").format(fehler)]
        hinweise.extend(u"{}: {}".format(plan.SheetNumber, h) for h in texte)
        return ergebnis

    if not aenderung.leer():
        for plan in wahl["plaene"]:
            if schritt(plan, lambda: pg.format_aendern(doc, plan, aenderung)):
                geaendert += 1
    if vorbild is not None:
        doc.Regenerate()
        for plan in wahl["plaene"]:
            if rv.id_wert(plan.Id) == rv.id_wert(vorbild.Id):
                continue
            if schritt(plan, lambda: pg.wie_vorbild(doc, vorbild, plan, wahl["art"])):
                ausgerichtet += 1
    return geaendert, ausgerichtet, hinweise


def main():
    doc, uidoc = revit.doc, revit.uidoc
    plaene = rv.plaene(doc)
    if not plaene:
        forms.alert(t(u"Im Projekt gibt es keine Pläne.", u"There are no sheets in the project.",
                      u"No hay planos en el proyecto."), title=TITEL)
        return
    koepfe = rv.plankoepfe(doc)
    vorgewaehlt = set(rv.id_wert(p.Id) for p in rv.gewaehlte_plaene(doc, uidoc))
    nach_id = dict((rv.id_wert(p.Id), p) for p in plaene)
    if isinstance(doc.ActiveView, ViewSheet) and rv.id_wert(doc.ActiveView.Id) in nach_id:
        vorbild = nach_id[rv.id_wert(doc.ActiveView.Id)]
    else:
        vorbild = next((p for p in plaene if rv.id_wert(p.Id) in vorgewaehlt), plaene[0])

    cfg = script.get_config()
    dlg = PlangroesseDialog(doc, plaene, vorgewaehlt, vorbild, koepfe, cfg)
    dlg.ShowDialog()
    wahl = dlg.ergebnis
    if not wahl:
        return
    einstellungen_speichern(cfg, wahl)

    transaktion = Transaction(doc, TITEL)
    transaktion.Start()
    try:
        geaendert, ausgerichtet, hinweise = ausfuehren(doc, wahl)
        transaktion.Commit()
    except Exception as fehler:
        if transaktion.HasStarted() and not transaktion.HasEnded():
            transaktion.RollBack()
        forms.alert(t(u"Ändern fehlgeschlagen, nichts wurde geändert.",
                      u"Change failed, nothing was changed.",
                      u"Error al cambiar, no se cambió nada."),
                    sub_msg=u"{}".format(fehler), title=TITEL)
        return

    meldung = t(u"{} Plan/Pläne im Format geändert.", u"{} sheet(s) changed in format.",
                u"{} plano(s) cambiados de formato.").format(geaendert)
    if wahl["vorbild"] is not None:
        meldung += u"\n" + t(u"{} Plan/Pläne wie '{}' ausgerichtet.",
                             u"{} sheet(s) aligned like '{}'.",
                             u"{} plano(s) alineados como '{}'.").format(
            ausgerichtet, wahl["vorbild"].SheetNumber)
    if hinweise:
        forms.alert(meldung, sub_msg=u"\n".join(hinweise[:30]), title=TITEL)
    else:
        forms.alert(meldung, title=TITEL)


main()
