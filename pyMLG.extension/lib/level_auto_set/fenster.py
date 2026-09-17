# -*- coding: utf-8 -*-
"""Hauptfenster von LevelAutoSet (WPF), angelehnt an "LevelAutoSet".

    ┌ Ebene ───────────────── Höhe ┐ Alle Keine
    │ □ Dach - OKFF          29,627 │ Unten (Basis)   Oben
    │ □ DG - OKFF            26,847 │ ○ näher         ○ näher
    │ ...                           │ ○ darüber       ● darüber
    └───────────────────────────────┘ ○ darunter      ○ darunter
                                      ● ignorieren    ○ ignorieren
                                      Optionen / Ebenen filtern
                                      2 / 0  Elemente / Ebenen
                                      [Ebenen setzen] [Abbrechen]

Keine Ebene markiert = alle angezeigten Ebenen kommen in Frage.
"""

import io
import json
import os
import traceback

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import GridLength, GridUnitType, HorizontalAlignment, \
    Thickness, VerticalAlignment, Visibility  # noqa: E402
from System.Windows.Controls import CheckBox, ColumnDefinition, Control, \
    Grid, ListBoxItem, TextBlock  # noqa: E402
from System.Windows.Media import Color, SolidColorBrush  # noqa: E402

from filter_manager import dialoge as dlg  # noqa: E402
from level_auto_set import logik as lg  # noqa: E402
from level_auto_set import revit as rv  # noqa: E402
from workset_creator.logik import filterfunktion  # noqa: E402

TITEL = u"LevelAutoSet"


def _pinsel(r, g, b):
    # Eingefroren, sonst gehört der Pinsel dem Thread, der das Modul geladen
    # hat (siehe filter_manager/fenster.py)
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


ROT = _pinsel(192, 57, 43)

FEHLERPROTOKOLL = os.path.join(dlg.protokollordner(),
                               "LevelAutoSet_Fehler.log")
EINSTELLUNGEN = os.path.join(dlg.protokollordner(), "LevelAutoSet.json")

MODI = ((lg.NAEHER, "naeher"), (lg.DARUEBER, "darueber"),
        (lg.DARUNTER, "darunter"), (lg.IGNORIEREN, "ignorieren"))

OPTIONEN = ("opt_oben_unverbunden", "opt_ignorierte_auswaehlen",
            "opt_bericht", "opt_profil", "opt_stuetzen", "opt_fehler",
            "filter_name", "filter_regex", "nur_geschoss", "nur_tragwerk")

XAML = u"""
<Window %s Title="LevelAutoSet (pyMLG)" Width="860" Height="820"
        MinWidth="700" MinHeight="600" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12"
        ResizeMode="CanResizeWithGrip">
  <Window.Resources>
    <Style TargetType="GroupBox">
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Margin" Value="0,2"/>
    </Style>
    <Style TargetType="RadioButton">
      <Setter Property="Margin" Value="0,2"/>
    </Style>
  </Window.Resources>
  <Grid Margin="10">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="10"/>
      <ColumnDefinition Width="250"/>
    </Grid.ColumnDefinitions>

    <!-- Ebenen -->
    <DockPanel>
      <Border DockPanel.Dock="Top" BorderBrush="#ABADB3"
              BorderThickness="1,1,1,0" Background="#F3F3F3">
        <Grid Margin="6,4,24,4">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="90"/>
          </Grid.ColumnDefinitions>
          <TextBlock Text="Ebene" FontWeight="SemiBold"/>
          <TextBlock Grid.Column="1" Text="Höhe" FontWeight="SemiBold"
                     HorizontalAlignment="Right"/>
        </Grid>
      </Border>
      <ListBox x:Name="liste" HorizontalContentAlignment="Stretch"
               ScrollViewer.VerticalScrollBarVisibility="Visible"/>
    </DockPanel>

    <!-- Optionen -->
    <DockPanel Grid.Column="2" LastChildFill="False">
      <StackPanel DockPanel.Dock="Top">
        <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
          <Button x:Name="alle" Content="Alle markieren" Padding="8,3"
                  Margin="0,0,6,0"/>
          <Button x:Name="keine" Content="Keine" Padding="8,3"/>
        </StackPanel>

        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="6"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <GroupBox Header="Unten (Basis)">
            <StackPanel>
              <RadioButton x:Name="basis_naeher" GroupName="basis"
                           Content="Näher"/>
              <RadioButton x:Name="basis_darueber" GroupName="basis"
                           Content="Darüber"/>
              <RadioButton x:Name="basis_darunter" GroupName="basis"
                           Content="Darunter"/>
              <RadioButton x:Name="basis_ignorieren" GroupName="basis"
                           Content="Ignorieren"/>
            </StackPanel>
          </GroupBox>
          <GroupBox Header="Oben" Grid.Column="2"
                    ToolTip="Gilt auch für Geschossdecken - ihre Ebene bezieht sich auf die Oberkante">
            <StackPanel>
              <RadioButton x:Name="oben_naeher" GroupName="oben"
                           Content="Näher"/>
              <RadioButton x:Name="oben_darueber" GroupName="oben"
                           Content="Darüber"/>
              <RadioButton x:Name="oben_darunter" GroupName="oben"
                           Content="Darunter"/>
              <RadioButton x:Name="oben_ignorieren" GroupName="oben"
                           Content="Ignorieren"/>
            </StackPanel>
          </GroupBox>
        </Grid>

        <GroupBox Header="Optionen">
          <StackPanel>
            <CheckBox x:Name="opt_oben_unverbunden"
                      ToolTip="Wände mit 'Nicht verbunden' oben bleiben oben frei">
              <TextBlock Text="Oben ignorieren, wenn nicht verbunden"
                         TextWrapping="Wrap"/>
            </CheckBox>
            <CheckBox x:Name="opt_ignorierte_auswaehlen"
                      ToolTip="Danach sind nur die übersprungenen, unveränderten und fehlerhaften Elemente ausgewählt">
              <TextBlock Text="Ignorierte danach auswählen"
                         TextWrapping="Wrap"/>
            </CheckBox>
            <CheckBox x:Name="opt_bericht" Content="Bericht am Ende anzeigen"/>
            <CheckBox x:Name="opt_profil"
                      Content="Wände mit bearbeitetem Profil ignorieren"/>
            <CheckBox x:Name="opt_stuetzen"
                      Content="Angehängte Stützen ignorieren"/>
            <CheckBox x:Name="opt_fehler"
                      ToolTip="Aus: der erste Fehler bricht alles ab (nichts wird geändert)">
              <TextBlock Text="Fehler ignorieren (fehlerhafte Elemente auslassen)"
                         TextWrapping="Wrap"/>
            </CheckBox>
          </StackPanel>
        </GroupBox>

        <GroupBox Header="Ebenen filtern">
          <StackPanel>
            <StackPanel Orientation="Horizontal">
              <CheckBox x:Name="filter_name" Content="Nach Name"
                        Margin="0,2,14,2"/>
              <CheckBox x:Name="filter_regex" Content="Regex"/>
            </StackPanel>
            <TextBox x:Name="filter_text" Padding="3" Margin="0,2,0,6"/>
            <CheckBox x:Name="nur_geschoss"
                      Content="Nur Gebäudegeschosse"/>
            <CheckBox x:Name="nur_tragwerk" Content="Nur Tragwerksebenen"/>
          </StackPanel>
        </GroupBox>
      </StackPanel>

      <StackPanel DockPanel.Dock="Bottom">
        <TextBlock x:Name="anzahl" FontSize="22" FontWeight="SemiBold"/>
        <TextBlock x:Name="anzahl_text" Foreground="#555" TextWrapping="Wrap"
                   Margin="0,0,0,10"/>
        <Button x:Name="setzen" Content="Ebenen setzen" FontWeight="Bold"
                Padding="8,9" Margin="0,0,0,6"/>
        <Button x:Name="abbrechen" Content="Abbrechen" Padding="8,5"
                IsCancel="True"/>
      </StackPanel>
    </DockPanel>
  </Grid>
</Window>""" % dlg.XMLNS


# ---------------------------------------------------------------------------
# Hilfen
# ---------------------------------------------------------------------------

def schreibe_fehlerprotokoll(spur):
    try:
        with io.open(FEHLERPROTOKOLL, "a", encoding="utf-8") as datei:
            datei.write(spur + u"\n" + u"-" * 70 + u"\n")
        return FEHLERPROTOKOLL
    except Exception:
        return None


def meldung(besitzer, text, warnung=False):
    dlg.meldung(besitzer, text, titel=TITEL, warnung=warnung)


def sicher(besitzer_liefern, funktion):
    """Ereignishandler, der Ausnahmen abfängt und verständlich anzeigt -
    eine Ausnahme im Handler würde sonst ShowDialog() beenden."""
    def handler(sender, args):
        try:
            funktion(sender, args)
        except Exception as fehler:
            pfad = schreibe_fehlerprotokoll(traceback.format_exc())
            try:
                meldung(besitzer_liefern(), u"%s%s" % (
                    dlg.fehlertext(fehler),
                    u"\n\nTechnische Details: %s" % pfad if pfad else u""),
                    warnung=True)
            except Exception:
                pass
    return handler


def lade_einstellungen():
    try:
        with io.open(EINSTELLUNGEN, "r", encoding="utf-8") as datei:
            return json.load(datei)
    except Exception:
        return {}


def speichere_einstellungen(werte):
    try:
        with io.open(EINSTELLUNGEN, "w", encoding="utf-8") as datei:
            datei.write(json.dumps(werte, ensure_ascii=False, indent=2))
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Fenster
# ---------------------------------------------------------------------------

class Zeile(object):
    def __init__(self, info):
        self.info = info
        self.item = None
        self.box = None


class LevelAutoSetFenster(object):

    def __init__(self, uiapp, uidoc, elemente):
        self.uiapp = uiapp
        self.uidoc = uidoc
        self.doc = uidoc.Document
        self.elemente = elemente
        self.zeilen = [Zeile(e) for e in rv.ebenen(self.doc)]

        f = self.fenster = dlg.lade_xaml(XAML)
        try:
            dlg.setze_besitzer(f, handle=uiapp.MainWindowHandle)
        except Exception:
            pass
        for name in ("liste", "filter_text", "anzahl", "anzahl_text") \
                + OPTIONEN:
            setattr(self, name, f.FindName(name))

        einst = lade_einstellungen()
        for name in OPTIONEN:
            getattr(self, name).IsChecked = bool(
                einst.get(name, name == "opt_bericht"))
        self.filter_text.Text = einst.get("filter_text", u"")
        for seite, vorgabe in (("basis", lg.IGNORIEREN),
                               ("oben", lg.DARUEBER)):
            gewaehlt = einst.get(seite, vorgabe)
            for modus, endung in MODI:
                f.FindName("%s_%s" % (seite, endung)).IsChecked = \
                    modus == gewaehlt
        markiert = set(einst.get("markiert", []))

        self._baue_liste(markiert)
        self._verdrahte()
        self.filtere()

    # --- Aufbau -----------------------------------------------------------
    def _baue_liste(self, markiert):
        for zeile in self.zeilen:
            raster = Grid()
            for breite in (None, 90.0):
                spalte = ColumnDefinition()
                spalte.Width = (GridLength(breite) if breite
                                else GridLength(1.0, GridUnitType.Star))
                raster.ColumnDefinitions.Add(spalte)
            box = CheckBox()
            box.Content = zeile.info.name
            box.IsChecked = zeile.info.name in markiert
            box.VerticalAlignment = VerticalAlignment.Center
            box.Click += sicher(lambda: self.fenster,
                                lambda _s, _a: self.aktualisiere_anzahl())
            raster.Children.Add(box)
            hoehe = TextBlock()
            hoehe.Text = zeile.info.anzeige
            hoehe.HorizontalAlignment = HorizontalAlignment.Right
            hoehe.VerticalAlignment = VerticalAlignment.Center
            Grid.SetColumn(hoehe, 1)
            raster.Children.Add(hoehe)
            item = ListBoxItem()
            item.Content = raster
            item.Padding = Thickness(4.0, 1.0, 4.0, 1.0)
            zeile.item, zeile.box = item, box
            self.liste.Items.Add(item)

    def _verdrahte(self):
        f = self.fenster

        def s(funktion):
            return sicher(lambda: f, funktion)

        f.FindName("alle").Click += s(lambda _s, _a: self.markiere(True))
        f.FindName("keine").Click += s(lambda _s, _a: self.markiere(False))
        f.FindName("setzen").Click += s(lambda _s, _a: self.setze())
        for name in ("filter_name", "filter_regex", "nur_geschoss",
                     "nur_tragwerk"):
            box = getattr(self, name)
            box.Checked += s(lambda _s, _a: self.filtere())
            box.Unchecked += s(lambda _s, _a: self.filtere())

        def bei_text(_s, _a):
            if not self.filter_name.IsChecked:
                self.filter_name.IsChecked = True   # löst filtere() aus
            else:
                self.filtere()
        self.filter_text.TextChanged += s(bei_text)
        f.Closing += s(lambda _s, _a: self._speichere())

    def _speichere(self):
        werte = dict((n, bool(getattr(self, n).IsChecked)) for n in OPTIONEN)
        werte["filter_text"] = self.filter_text.Text or u""
        werte["basis"] = self.modus("basis")
        werte["oben"] = self.modus("oben")
        werte["markiert"] = [z.info.name for z in self.zeilen
                             if z.box.IsChecked]
        speichere_einstellungen(werte)

    # --- Zustand ----------------------------------------------------------
    def modus(self, seite):
        for modus, endung in MODI:
            if self.fenster.FindName("%s_%s" % (seite, endung)).IsChecked:
                return modus
        return lg.IGNORIEREN

    def sichtbare(self):
        return [z for z in self.zeilen
                if z.item.Visibility == Visibility.Visible]

    def kandidaten(self):
        """Markierte angezeigte Ebenen - ohne Markierung alle angezeigten."""
        sichtbar = self.sichtbare()
        markiert = [z for z in sichtbar if z.box.IsChecked]
        return [z.info for z in (markiert or sichtbar)]

    def filtere(self):
        self.filter_text.ClearValue(Control.ForegroundProperty)
        self.filter_text.ToolTip = None
        passt = None
        if self.filter_name.IsChecked:
            try:
                passt = filterfunktion(self.filter_text.Text,
                                       bool(self.filter_regex.IsChecked))
            except Exception as fehler:
                self.filter_text.Foreground = ROT
                self.filter_text.ToolTip = u"Ungültiger Ausdruck: %s" % fehler
        for zeile in self.zeilen:
            zeigen = ((passt is None or passt(zeile.info.name))
                      and (zeile.info.geschoss
                           or not self.nur_geschoss.IsChecked)
                      and (zeile.info.tragwerk
                           or not self.nur_tragwerk.IsChecked))
            zeile.item.Visibility = (Visibility.Visible if zeigen
                                     else Visibility.Collapsed)
        self.aktualisiere_anzahl()

    def markiere(self, zustand):
        for zeile in self.sichtbare():
            zeile.box.IsChecked = zustand
        self.aktualisiere_anzahl()

    def aktualisiere_anzahl(self):
        sichtbar = self.sichtbare()
        markiert = sum(1 for z in sichtbar if z.box.IsChecked)
        self.anzahl.Text = u"%d / %d" % (len(self.elemente), markiert)
        self.anzahl_text.Text = (
            u"Elemente / Ebenen markiert" + (
                u"\n(keine Ebene markiert: alle %d angezeigten Ebenen "
                u"kommen in Frage)" % len(sichtbar) if not markiert else u""))

    # --- Ausführen --------------------------------------------------------
    def setze(self):
        if not self.elemente:
            meldung(self.fenster, u"Es sind keine Elemente ausgewählt. Bitte "
                    u"Elemente in Revit auswählen und das Werkzeug neu "
                    u"starten.")
            return
        basis, oben = self.modus("basis"), self.modus("oben")
        if basis == lg.IGNORIEREN and oben == lg.IGNORIEREN:
            meldung(self.fenster, u"Unten und oben stehen auf 'Ignorieren' - "
                    u"es gibt nichts zu tun.")
            return
        ebenen = self.kandidaten()
        if not ebenen:
            meldung(self.fenster, u"Der Filter blendet alle Ebenen aus.")
            return
        try:
            ergebnis = rv.setze_ebenen(
                self.doc, self.elemente, ebenen, basis, oben,
                oben_unverbunden_ignorieren=bool(
                    self.opt_oben_unverbunden.IsChecked),
                profil_ignorieren=bool(self.opt_profil.IsChecked),
                angehaengte_stuetzen_ignorieren=bool(
                    self.opt_stuetzen.IsChecked),
                fehler_ignorieren=bool(self.opt_fehler.IsChecked))
        except RuntimeError as fehler:
            meldung(self.fenster, u"Abgebrochen, es wurde nichts geändert:"
                    u"\n\n%s\n\nMit 'Fehler ignorieren' werden solche "
                    u"Elemente ausgelassen." % fehler, warnung=True)
            return

        if self.opt_ignorierte_auswaehlen.IsChecked:
            rv.waehle_aus(self.uidoc, ergebnis.ignorierte())
        if self.opt_bericht.IsChecked or ergebnis.verschoben:
            meldung(self.fenster, self._bericht(ergebnis),
                    warnung=bool(ergebnis.fehler or ergebnis.verschoben))
        self.fenster.Close()

    @staticmethod
    def _bericht(ergebnis, maximal=12):
        def liste(titel, eintraege):
            if not eintraege:
                return []
            zeilen = [u"", u"%s (%d):" % (titel, len(eintraege))]
            zeilen += [u"  " + t for t in eintraege[:maximal]]
            if len(eintraege) > maximal:
                zeilen.append(u"  ... und %d weitere"
                              % (len(eintraege) - maximal))
            return zeilen

        zeilen = [u"Geändert: %d Elemente" % len(ergebnis.geaendert),
                  u"  Basis auf andere Ebene: %d" % ergebnis.basis,
                  u"  Oberkante auf andere Ebene: %d" % ergebnis.oben,
                  u"Bereits auf der passenden Ebene: %d"
                  % len(ergebnis.unveraendert)]
        zeilen += liste(u"Übersprungen", [u"%s - %s" % (rv.beschreibung(e), g)
                                         for e, g in ergebnis.uebersprungen])
        zeilen += liste(u"Fehler", [u"%s - %s" % (rv.beschreibung(e), t)
                                   for e, t in ergebnis.fehler])
        zeilen += liste(u"ACHTUNG, Höhe hat sich verändert - bitte prüfen",
                        [rv.beschreibung(e) for e in ergebnis.verschoben])
        return u"\n".join(zeilen)

    def zeige(self):
        self.fenster.ShowDialog()


def starte(uiapp, uidoc):
    elemente = [uidoc.Document.GetElement(i)
                for i in uidoc.Selection.GetElementIds()]
    elemente = [e for e in elemente if e is not None]
    LevelAutoSetFenster(uiapp, uidoc, elemente).zeige()
