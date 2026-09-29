# -*- coding: utf-8 -*-
"""Fenster von LinkKoerper (WPF) und der Ablauf eines Laufs.

    ┌ Link als Allgemeines Modell ───────────────────────┐
    │ Verknüpfung  [TGA.rvt  (TGA.rvt : Position 1)   v] │
    │ (o) Ganze Kategorien                               │
    │     [Suche...........]                             │
    │     [x] Rohre (1204)                               │
    │     [ ] Luftkanäle (388)                           │
    │ ( ) Elemente einzeln in der Verknüpfung wählen     │
    │ ( ) Aktuelle Auswahl: 3 Elemente                   │
    │ Vorhandene Körper: unverändert bleiben ...         │
    │                          [Erstellen] [Abbrechen]   │
    └────────────────────────────────────────────────────┘

Das Fenster ist modal - für die Einzelwahl schliesst es und Revit lässt
danach in der Verknüpfung picken.
"""

import io
import json
import os
import time
import traceback

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")
clr.AddReference("System")

from System.Collections.Generic import List  # noqa: E402
from System.Windows import Thickness, Visibility  # noqa: E402
from System.Windows.Controls import CheckBox, ComboBoxItem  # noqa: E402

from Autodesk.Revit.DB import ElementId  # noqa: E402

from filter_manager import dialoge as dlg  # noqa: E402
from link_koerper import logik as lg  # noqa: E402
from link_koerper import revit as rv  # noqa: E402
from mlg_sprache import t, uebersetze_xaml  # noqa: E402

TITEL = t(u"Link als Allgemeines Modell", u"Link to generic model",
          u"Vínculo a modelo genérico")

KATEGORIEN = "kategorien"
WAEHLEN = "waehlen"
AUSWAHL = "auswahl"

FEHLERPROTOKOLL = os.path.join(dlg.protokollordner(),
                               "LinkKoerper_Fehler.log")
EINSTELLUNGEN = os.path.join(dlg.protokollordner(), "LinkKoerper.json")

XAML_TEXTE = {
    "titel": (u"Link als Allgemeines Modell (pyMLG)",
              u"Link to generic model (pyMLG)",
              u"Vínculo a modelo genérico (pyMLG)"),
    "link": (u"Verknüpfung", u"Link", u"Vínculo"),
    "kategorien": (u"Alle Elemente der markierten Kategorien",
                   u"All elements of the checked categories",
                   u"Todos los elementos de las categorías marcadas"),
    "suche": (u"Suche", u"Search", u"Buscar"),
    "waehlen": (u"Elemente einzeln in der Verknüpfung wählen",
                u"Pick single elements inside the link",
                u"Elegir elementos sueltos del vínculo"),
    "hinweis": (u"Jedes Element wird eine eigene Familie \"Allgemeines "
                u"Modell\" (Volumenkörper) - damit lassen sich Elemente des "
                u"Hauptmodells verschneiden. Schon vorhandene Körper bleiben "
                u"unverändert; hat sich die Lage eines Elements geändert, "
                u"wird nur sein Körper nachgeführt.",
                u"Each element becomes its own \"Generic Model\" family "
                u"(solid) - ready to cut elements of the host model. "
                u"Existing bodies stay untouched; only bodies whose element "
                u"has moved are updated.",
                u"Cada elemento se convierte en su propia familia \"Modelo "
                u"genérico\" (sólido), lista para cortar elementos del "
                u"modelo principal. Los cuerpos existentes no se tocan; "
                u"solo se actualizan aquellos cuyo elemento se ha movido."),
    "erstellen": (u"Erstellen", u"Create", u"Crear"),
    "abbrechen": (u"Abbrechen", u"Cancel", u"Cancelar"),
    "alle": (u"Alle", u"All", u"Todas"),
    "keine": (u"Keine", u"None", u"Ninguna"),
}

XAML = u"""<Window %s
        Title="{{titel}}" Width="520" Height="600"
        MinWidth="420" MinHeight="420"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        FontFamily="Segoe UI" FontSize="12"
        ResizeMode="CanResizeWithGrip">
  <DockPanel Margin="10">
    <DockPanel DockPanel.Dock="Top" Margin="0,0,0,8">
      <TextBlock Text="{{link}}" VerticalAlignment="Center" Width="80"/>
      <ComboBox x:Name="link"/>
    </DockPanel>
    <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal"
                HorizontalAlignment="Right" Margin="0,10,0,0">
      <Button x:Name="erstellen" Content="{{erstellen}}" FontWeight="Bold"
              Padding="12,6" Margin="0,0,6,0" IsDefault="True"/>
      <Button x:Name="abbrechen" Content="{{abbrechen}}" Padding="12,6"
              IsCancel="True"/>
    </StackPanel>
    <TextBlock DockPanel.Dock="Bottom" Text="{{hinweis}}" Foreground="#6E6E6E"
               TextWrapping="Wrap" Margin="0,8,0,0"/>
    <RadioButton x:Name="auswahl" DockPanel.Dock="Bottom" GroupName="modus"
                 Margin="0,6,0,0" Visibility="Collapsed"/>
    <RadioButton x:Name="waehlen" DockPanel.Dock="Bottom" GroupName="modus"
                 Content="{{waehlen}}" Margin="0,6,0,0"/>
    <RadioButton x:Name="kategorien" DockPanel.Dock="Top" GroupName="modus"
                 Content="{{kategorien}}" IsChecked="True" Margin="0,0,0,4"/>
    <DockPanel x:Name="kategorie_block" Margin="20,0,0,0">
      <DockPanel DockPanel.Dock="Top" Margin="0,0,0,4">
        <Button x:Name="keine" DockPanel.Dock="Right" Content="{{keine}}"
                Padding="8,2" Margin="4,0,0,0"/>
        <Button x:Name="alle" DockPanel.Dock="Right" Content="{{alle}}"
                Padding="8,2" Margin="4,0,0,0"/>
        <TextBlock Text="{{suche}}" VerticalAlignment="Center"
                   Margin="0,0,6,0"/>
        <TextBox x:Name="suche"/>
      </DockPanel>
      <TextBlock x:Name="zaehler" DockPanel.Dock="Bottom" Foreground="#6E6E6E"
                 Margin="0,4,0,0"/>
      <ListBox x:Name="liste"
               ScrollViewer.VerticalScrollBarVisibility="Auto"/>
    </DockPanel>
  </DockPanel>
</Window>""" % dlg.XMLNS


def schreibe_fehlerprotokoll(spur):
    try:
        with io.open(FEHLERPROTOKOLL, "a", encoding="utf-8") as datei:
            datei.write(spur + u"\n" + u"-" * 70 + u"\n")
        return FEHLERPROTOKOLL
    except Exception:
        return None


def sicher(besitzer_liefern, funktion):
    """Ereignishandler mit Fehlerfang - eine Ausnahme im Handler würde
    ShowDialog() sonst wortlos beenden."""
    def handler(sender, args):
        try:
            funktion(sender, args)
        except Exception as fehler:
            pfad = schreibe_fehlerprotokoll(traceback.format_exc())
            try:
                dlg.meldung(besitzer_liefern(), u"%s%s" % (
                    dlg.fehlertext(fehler),
                    t(u"\n\nTechnische Details: %s",
                      u"\n\nTechnical details: %s",
                      u"\n\nDetalles técnicos: %s") % pfad if pfad else u""),
                    titel=TITEL, warnung=True)
            except Exception:
                pass
    return handler


class LinkKoerperFenster(object):

    def __init__(self, uiapp, links, anzahl_auswahl):
        self.links = links
        self.gruppen = {}          # Link-ID -> elemente_nach_kategorie
        self.boxen = []            # (Suchtext, CheckBox, Kategorie-ID)
        self.auftrag = None

        f = self.fenster = dlg.lade_xaml(uebersetze_xaml(XAML, XAML_TEXTE))
        dlg.setze_besitzer(f, handle=uiapp.MainWindowHandle)
        for name in ("link", "kategorien", "waehlen", "auswahl",
                     "kategorie_block", "suche", "liste", "zaehler"):
            setattr(self, name, f.FindName(name))

        if anzahl_auswahl:
            self.auswahl.Content = t(
                u"Aktuelle Auswahl: %d Element(e) aus Verknüpfungen",
                u"Current selection: %d element(s) from links",
                u"Selección actual: %d elemento(s) de vínculos") % \
                anzahl_auswahl
            self.auswahl.Visibility = Visibility.Visible
            self.auswahl.IsChecked = True
        if not links:
            self.kategorien.IsEnabled = False
            self.link.IsEnabled = False

        for text, _instanz in links:
            item = ComboBoxItem()
            item.Content = text
            self.link.Items.Add(item)
        self._verdrahte()
        if links:
            self.link.SelectedIndex = 0
        self._modus_geaendert()

    # --- Aufbau -----------------------------------------------------------
    def _verdrahte(self):
        f = self.fenster

        def s(funktion):
            return sicher(lambda: f, funktion)

        self.link.SelectionChanged += s(lambda _s, _a: self._lade_liste())
        for knopf in (self.kategorien, self.waehlen, self.auswahl):
            knopf.Checked += s(lambda _s, _a: self._modus_geaendert())
        self.suche.TextChanged += s(lambda _s, _a: self._filtere())
        f.FindName("alle").Click += s(lambda _s, _a: self._markiere(True))
        f.FindName("keine").Click += s(lambda _s, _a: self._markiere(False))
        f.FindName("erstellen").Click += s(lambda _s, _a: self._erstellen())
        f.FindName("abbrechen").Click += s(lambda _s, _a: f.Close())

    def _instanz(self):
        index = self.link.SelectedIndex
        if index < 0 or index >= len(self.links):
            return None
        return self.links[index][1]

    def _lade_liste(self):
        instanz = self._instanz()
        self.liste.Items.Clear()
        self.boxen = []
        if instanz is None:
            self._zaehle()
            return
        schluessel = rv.id_wert(instanz.Id)
        if schluessel not in self.gruppen:
            self.gruppen[schluessel] = rv.elemente_nach_kategorie(
                instanz.GetLinkDocument())
        gruppen = self.gruppen[schluessel]
        for kategorie_id, (name, elemente) in sorted(
                gruppen.items(), key=lambda paar: paar[1][0].lower()):
            box = CheckBox()
            box.Content = u"%s (%d)" % (name, len(elemente))
            box.Margin = Thickness(2.0, 2.0, 2.0, 2.0)
            box.Click += sicher(lambda: self.fenster,
                                lambda _s, _a: self._zaehle())
            self.liste.Items.Add(box)
            self.boxen.append((name.lower(), box, kategorie_id))
        self._filtere()

    def _filtere(self):
        woerter = (self.suche.Text or u"").lower().split()
        for text, box, _id in self.boxen:
            passt = all(wort in text for wort in woerter)
            box.Visibility = Visibility.Visible if passt \
                else Visibility.Collapsed
        self._zaehle()

    def _markiere(self, zustand):
        for _text, box, _id in self.boxen:
            if box.Visibility == Visibility.Visible:
                box.IsChecked = zustand
        self._zaehle()

    def _markierte(self):
        return [kategorie_id for _text, box, kategorie_id in self.boxen
                if box.IsChecked]

    def _zaehle(self):
        instanz = self._instanz()
        gruppen = self.gruppen.get(rv.id_wert(instanz.Id), {}) \
            if instanz is not None else {}
        markiert = self._markierte()
        elemente = sum(len(gruppen[k][1]) for k in markiert if k in gruppen)
        self.zaehler.Text = t(
            u"%d Kategorie(n) markiert, %d Element(e)",
            u"%d category(ies) checked, %d element(s)",
            u"%d categoría(s) marcadas, %d elemento(s)") % (len(markiert),
                                                          elemente)

    def _modus_geaendert(self):
        self.kategorie_block.IsEnabled = bool(self.kategorien.IsChecked)

    # --- Aktionen ---------------------------------------------------------
    def _erstellen(self):
        if self.auswahl.IsChecked:
            self.auftrag = (AUSWAHL, None, None, None)
        elif self.waehlen.IsChecked:
            self.auftrag = (WAEHLEN, None, None, None)
        else:
            markiert = self._markierte()
            if not markiert:
                dlg.meldung(self.fenster, t(
                    u"Bitte mindestens eine Kategorie markieren.",
                    u"Please check at least one category.",
                    u"Marque al menos una categoría."), titel=TITEL)
                return
            instanz = self._instanz()
            self.auftrag = (KATEGORIEN, instanz, markiert,
                            self.gruppen[rv.id_wert(instanz.Id)])
        self.fenster.Close()

    def zeige(self):
        self.fenster.ShowDialog()
        return self.auftrag


# ---------------------------------------------------------------- Einstellungen

def _lies_einstellungen():
    try:
        with io.open(EINSTELLUNGEN, "r", encoding="utf-8") as datei:
            return json.load(datei)
    except Exception:
        return {}


def _schreibe_einstellungen(werte):
    try:
        with io.open(EINSTELLUNGEN, "w", encoding="utf-8") as datei:
            datei.write(json.dumps(werte, ensure_ascii=False, indent=2))
    except Exception:
        pass


def _frage_vorlage(startordner):
    """Familienvorlage (.rft) per Dateidialog wählen."""
    from Microsoft.Win32 import OpenFileDialog
    dialog = OpenFileDialog()
    dialog.Title = t(u"Familienvorlage \"Allgemeines Modell\" wählen",
                     u"Select the \"Generic Model\" family template",
                     u"Seleccione la plantilla de familia \"Modelo genérico\"")
    dialog.Filter = t(u"Familienvorlagen (*.rft)|*.rft",
                      u"Family templates (*.rft)|*.rft",
                      u"Plantillas de familia (*.rft)|*.rft")
    if startordner and os.path.isdir(startordner):
        dialog.InitialDirectory = startordner
    if dialog.ShowDialog() == True:  # noqa: E712 - Nullable<bool> aus .NET
        return dialog.FileName
    return None


def _vorlage(app):
    einstellungen = _lies_einstellungen()
    pfad = rv.finde_vorlage(app, einstellungen.get("vorlage"))
    if pfad is None:
        startordner = u""
        try:
            startordner = app.FamilyTemplatePath
        except Exception:
            pass
        pfad = _frage_vorlage(startordner)
    if pfad:
        einstellungen["vorlage"] = pfad
        _schreibe_einstellungen(einstellungen)
    return pfad


# ---------------------------------------------------------------- Ablauf

class _Fortschritt(object):
    """Fortschritt im pyRevit-Ausgabefenster - fehlt es, läuft alles ohne."""

    def __init__(self):
        self.balken = None
        try:
            from pyrevit import script
            from schedule_sync import ui
            self.balken = ui.Fortschritt(script.get_output())
        except Exception:
            self.balken = None

    def __enter__(self):
        if self.balken is not None:
            self.balken.__enter__()
        return self

    def aktualisiere(self, wert, max_wert=None):
        if self.balken is not None:
            self.balken.aktualisiere(wert, max_wert)

    def __exit__(self, *ausnahme):
        if self.balken is not None:
            self.balken.__exit__(*ausnahme)
        return False


def _meldung(text, hauptzeile, warnung=False):
    from schedule_sync import ui
    ui.meldung(text, titel=TITEL, hauptzeile=hauptzeile, warnung=warnung)


def starte(uiapp, uidoc):
    from schedule_sync import ui

    doc = uidoc.Document
    aus_auswahl = rv.aus_referenzen(doc, rv.referenzen_der_auswahl(uidoc))
    links = rv.verknuepfungen(doc)
    if not links and not aus_auswahl:
        _meldung(t(u"Das Modell hat keine geladene Revit-Verknüpfung.",
                   u"The model has no loaded Revit link.",
                   u"El modelo no tiene ningún vínculo de Revit cargado."),
                 t(u"Keine Verknüpfung", u"No link", u"Ningún vínculo"),
                 warnung=True)
        return

    auftrag = LinkKoerperFenster(uiapp, links, len(aus_auswahl)).zeige()
    if auftrag is None:
        return
    modus, instanz, kategorie_ids, gruppen = auftrag
    if modus == AUSWAHL:
        quellen = aus_auswahl
    elif modus == WAEHLEN:
        referenzen = rv.waehle(uidoc)
        if referenzen is None:
            return
        quellen = rv.aus_referenzen(doc, referenzen)
    else:
        quellen = rv.aus_kategorien(instanz, gruppen, kategorie_ids)
    if not quellen:
        _meldung(t(u"Keine Elemente aus einer geladenen Verknüpfung gewählt.",
                   u"No elements from a loaded link selected.",
                   u"No se han elegido elementos de un vínculo cargado."),
                 t(u"Nichts zu tun", u"Nothing to do", u"Nada que hacer"))
        return

    beginn = time.time()
    ergebnis = lg.Ergebnis()
    bestand = rv.vorhandene(doc)
    plan = rv.plane(quellen, bestand, ergebnis)
    if modus == KATEGORIEN:
        ergebnis.verwaist = rv.zaehle_verwaiste(bestand, instanz)

    zu_bauen = sum(1 for _q, aktion, _a in plan
                   if aktion in (lg.NEU, lg.ERSETZEN))
    if zu_bauen > lg.VIELE and not ui.frage(
            t(u"Es entstehen %d neue Familien - das kann einige Minuten "
              u"dauern. Fortfahren?",
              u"%d new families will be created - this may take a few "
              u"minutes. Continue?",
              u"Se crearán %d familias nuevas; puede tardar unos minutos. "
              u"¿Continuar?") % zu_bauen,
            titel=TITEL, hauptzeile=t(u"Viele Elemente", u"Many elements",
                                      u"Muchos elementos")):
        return

    vorlage = None
    if zu_bauen:
        vorlage = _vorlage(uiapp.Application)
        if not vorlage:
            return

    if zu_bauen:
        with _Fortschritt() as fortschritt:
            rv.ausfuehren(doc, uiapp.Application, plan, vorlage, ergebnis,
                          fortschritt)
    elif plan:
        # Nur verschieben bzw. Ebenen angleichen - ohne Fortschrittsbalken
        rv.ausfuehren(doc, uiapp.Application, plan, vorlage, ergebnis)

    if ergebnis.neue_ids:
        ids = List[ElementId]()
        for element_id in ergebnis.neue_ids:
            ids.Add(element_id)
        try:
            uidoc.Selection.SetElementIds(ids)
        except Exception:
            pass

    text = ergebnis.text()
    if ergebnis.fehler:
        pfad = rv.schreibe_protokoll(ergebnis, time.time() - beginn)
        text += u"\n\n" + u"\n".join(ergebnis.fehler[:5])
        if pfad:
            text += t(u"\n\nProtokoll: %s", u"\n\nLog: %s",
                      u"\n\nRegistro: %s") % pfad
    _meldung(text, t(u"Fertig", u"Done", u"Terminado"),
             warnung=bool(ergebnis.fehler))
