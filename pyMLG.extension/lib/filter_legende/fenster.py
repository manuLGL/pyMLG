# -*- coding: utf-8 -*-
"""Dialog von FilterLegend (WPF): Ansichten wählen, Filter mit
Farbvorschau an-/abhaken, Beschriftung je Filter anpassen.

    ┌ 1. Ansichten ───────────┬ 2. Filter ───────────────────────────────────┐
    │ Suchen [        ]       │ ☑ ■ AR_Wand_tragend   [Tragende Wand      ]  │
    │ ☑ EG (Grundriss, 5)     │ ☑ ▨ AR_Wand_Bestand   [Bestand            ]  │
    │ ☐ OG (Grundriss, 5)     │ ☐ – AR_Abbruch        ausgeblendet           │
    │ [Alle] [Keine]          │ [Alle] [Keine]   Reihenfolge [wie Ansicht ▾] │
    ├ 3. Einstellungen ───────┴──────────────────────────────────────────────┤
    │ Name [Legende Filter]  Massstab 1:[100]  Feld [10]×[5] mm  Abstand [2] │
    │ Spalten [1] Breite [70] mm  Farbe aus [Automatisch ▾] Texttyp [▾] ☑ ☐  │
    └─────────────────────────────── 7 Filter [Legende erstellen] [Abbrechen]

Eigene Beschriftungen werden je Filtername gespeichert und beim nächsten
Mal wieder vorgeschlagen.
"""

import io
import json
import os
import traceback

from System import Byte
from System.Windows import GridLength, GridUnitType, HorizontalAlignment, \
    TextTrimming, Thickness, VerticalAlignment, Visibility
from System.Windows.Controls import (
    Border,
    CheckBox,
    ColumnDefinition,
    ComboBoxItem,
    Grid,
    TextBlock,
    TextBox,
)
from System.Windows.Input import Cursors, Mouse
from System.Windows.Media import Color as FarbeWpf, SolidColorBrush

from dxf_legende import logik as dxf_lg
from filter_legende import logik as lg
from filter_legende import revit as rv
from filter_manager import dialoge as dlg
from mlg_sprache import t, tt, uebersetze_xaml

TITEL = t(u"Filterlegende", u"Filter legend", u"Leyenda de filtros")

FEHLERPROTOKOLL = os.path.join(dlg.protokollordner(), "FilterLegend_Fehler.log")
EINSTELLUNGEN = os.path.join(dlg.protokollordner(), "FilterLegend.json")

XAML_TEXTE = {
    "titel": (u"Filterlegende (pyMLG)", u"Filter legend (pyMLG)",
              u"Leyenda de filtros (pyMLG)"),
    "ansichten": (u"1. Ansichten und Vorlagen", u"1. Views and templates",
                  u"1. Vistas y plantillas"),
    "filter": (u"2. Filter  (Häkchen, Farbe, Text in der Legende)",
               u"2. Filters  (tick, colour, text in the legend)",
               u"2. Filtros  (marca, color, texto en la leyenda)"),
    "suchen": (u"Suchen:", u"Search:", u"Buscar:"),
    "alle": (u"Alle", u"All", u"Todos"),
    "keine": (u"Keine", u"None", u"Ninguno"),
    "reihenfolge": (u"Reihenfolge:", u"Order:", u"Orden:"),
    "einstellungen": (u"3. Einstellungen", u"3. Settings", u"3. Ajustes"),
    "name": (u"Name:", u"Name:", u"Nombre:"),
    "massstab": (u"Massstab 1:", u"Scale 1:", u"Escala 1:"),
    "feld": (u"Farbfeld (mm):", u"Swatch (mm):", u"Muestra (mm):"),
    "feld_tip": (u"Breite × Höhe des Farbfelds auf dem Plan.",
                 u"Width × height of the swatch on the sheet.",
                 u"Ancho × alto de la muestra en el plano."),
    "abstand": (u"Abstand (mm):", u"Spacing (mm):", u"Separación (mm):"),
    "abstand_tip": (u"Lücke zwischen zwei Farbfeldern untereinander, auf dem Plan.",
                    u"Gap between two swatches stacked vertically, on the sheet.",
                    u"Hueco entre dos muestras apiladas, en el plano."),
    "spalten": (u"Spalten:", u"Columns:", u"Columnas:"),
    "spaltenbreite": (u"Spaltenbreite (mm):", u"Column width (mm):",
                      u"Ancho de columna (mm):"),
    "spaltenbreite_tip": (u"Von Farbfeld zu Farbfeld nebeneinander, inklusive Text. "
                          u"Nur bei mehr als einer Spalte.",
                          u"From swatch to swatch side by side, text included. "
                          u"Only with more than one column.",
                          u"De muestra a muestra en horizontal, texto incluido. "
                          u"Solo con más de una columna."),
    "quelle": (u"Farbe aus:", u"Colour from:", u"Color de:"),
    "quelle_tip": (u"Automatisch: Oberflächenmuster, sonst Schnittmuster, sonst die "
                   u"Linienfarbe als Vollton. Oder fest Oberfläche bzw. Schnitt (z.B. "
                   u"für Grundrisse, in denen Wände geschnitten sind).",
                   u"Automatic: surface pattern, else cut pattern, else the line "
                   u"colour as solid fill. Or always surface or cut (e.g. for plans "
                   u"where walls are cut).",
                   u"Automático: patrón de superficie, si no de corte, si no el color "
                   u"de línea como relleno sólido. O siempre superficie o corte (p.ej. "
                   u"en plantas con muros cortados)."),
    "texttyp": (u"Texttyp:", u"Text type:", u"Tipo de texto:"),
    "rahmen": (u"Rahmen um Farbfeld", u"Outline around swatch",
               u"Contorno de la muestra"),
    "ersetzen": (u"Gleichnamige Legende ersetzen", u"Replace legend with the same name",
                 u"Reemplazar leyenda del mismo nombre"),
    "ersetzen_tip": (u"Gibt es schon eine Legende mit dem Namen, wird ihr Inhalt "
                     u"gelöscht und neu aufgebaut. Auf Plänen bleibt sie an ihrem "
                     u"Platz. Ohne Haken entsteht \"Name (2)\".",
                     u"If a legend with this name exists, its content is deleted "
                     u"and rebuilt. It stays in place on sheets. Unticked, "
                     u"\"Name (2)\" is created.",
                     u"Si ya existe una leyenda con ese nombre, se borra su "
                     u"contenido y se vuelve a crear. En los planos se queda en "
                     u"su sitio. Sin marcar se crea \"Nombre (2)\"."),
    "erstellen": (u"Legende erstellen", u"Create legend", u"Crear leyenda"),
    "abbrechen": (u"Abbrechen", u"Cancel", u"Cancelar"),
}

XAML = u"""
<Window %s Title="{{titel}}" Width="1100" Height="720"
        MinWidth="860" MinHeight="480" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12"
        ResizeMode="CanResizeWithGrip">
  <Window.Resources>
    <Style TargetType="TextBox">
      <Setter Property="Padding" Value="3"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
    </Style>
    <Style TargetType="GroupBox">
      <Setter Property="Padding" Value="6"/>
    </Style>
  </Window.Resources>
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="1.6*"/>
      </Grid.ColumnDefinitions>

      <GroupBox Header="{{ansichten}}" Margin="0,0,8,0">
        <DockPanel>
          <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
            <TextBlock Text="{{suchen}}" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBox x:Name="suche"/>
          </DockPanel>
          <DockPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
            <Button x:Name="ansichten_alle" Content="{{alle}}" Padding="10,2"
                    Margin="0,0,6,0"/>
            <Button x:Name="ansichten_keine" Content="{{keine}}" Padding="10,2"/>
            <TextBlock x:Name="ansichten_zahl" Foreground="#666" Margin="10,0,0,0"
                       VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
          </DockPanel>
          <ListBox x:Name="ansichtliste"/>
        </DockPanel>
      </GroupBox>

      <GroupBox Header="{{filter}}" Grid.Column="1">
        <DockPanel>
          <DockPanel DockPanel.Dock="Bottom" Margin="0,6,0,0" LastChildFill="False">
            <Button x:Name="filter_alle" Content="{{alle}}" Padding="10,2"
                    Margin="0,0,6,0"/>
            <Button x:Name="filter_keine" Content="{{keine}}" Padding="10,2"/>
            <ComboBox x:Name="reihenfolge" DockPanel.Dock="Right" Width="150"/>
            <TextBlock Text="{{reihenfolge}}" DockPanel.Dock="Right"
                       VerticalAlignment="Center" Margin="0,0,6,0"/>
          </DockPanel>
          <TextBlock x:Name="filter_leer" DockPanel.Dock="Top" Foreground="#666"
                     TextWrapping="Wrap" Margin="2,0,2,6"/>
          <ScrollViewer VerticalScrollBarVisibility="Auto"
                        HorizontalScrollBarVisibility="Disabled">
            <StackPanel x:Name="filterliste"/>
          </ScrollViewer>
        </DockPanel>
      </GroupBox>
    </Grid>

    <GroupBox Header="{{einstellungen}}" Grid.Row="1" Margin="0,8,0,0">
      <StackPanel>
        <StackPanel Orientation="Horizontal">
          <TextBlock Text="{{name}}" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <TextBox x:Name="name" Width="240"/>
          <TextBlock Text="{{massstab}}" VerticalAlignment="Center" Margin="14,0,6,0"/>
          <TextBox x:Name="massstab" Width="55"/>
          <TextBlock Text="{{feld}}" VerticalAlignment="Center" Margin="14,0,6,0"
                     ToolTip="{{feld_tip}}"/>
          <TextBox x:Name="feld_breite" Width="45" ToolTip="{{feld_tip}}"/>
          <TextBlock Text="×" VerticalAlignment="Center" Margin="4,0,4,0"/>
          <TextBox x:Name="feld_hoehe" Width="45" ToolTip="{{feld_tip}}"/>
          <TextBlock Text="{{abstand}}" VerticalAlignment="Center" Margin="14,0,6,0"
                     ToolTip="{{abstand_tip}}"/>
          <TextBox x:Name="abstand" Width="45" ToolTip="{{abstand_tip}}"/>
          <TextBlock Text="{{spalten}}" VerticalAlignment="Center" Margin="14,0,6,0"/>
          <TextBox x:Name="spalten" Width="35"/>
          <TextBlock Text="{{spaltenbreite}}" VerticalAlignment="Center"
                     Margin="14,0,6,0" ToolTip="{{spaltenbreite_tip}}"/>
          <TextBox x:Name="spaltenbreite" Width="45" ToolTip="{{spaltenbreite_tip}}"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
          <TextBlock Text="{{quelle}}" VerticalAlignment="Center" Margin="0,0,6,0"
                     ToolTip="{{quelle_tip}}"/>
          <ComboBox x:Name="quelle" Width="170" ToolTip="{{quelle_tip}}"/>
          <TextBlock Text="{{texttyp}}" VerticalAlignment="Center" Margin="14,0,6,0"/>
          <ComboBox x:Name="texttyp" Width="240"/>
          <CheckBox x:Name="rahmen" Content="{{rahmen}}" VerticalAlignment="Center"
                    Margin="18,0,0,0"/>
          <CheckBox x:Name="ersetzen" Content="{{ersetzen}}" ToolTip="{{ersetzen_tip}}"
                    VerticalAlignment="Center" Margin="18,0,0,0"/>
        </StackPanel>
        <TextBlock x:Name="namehinweis" Foreground="#A05A00" TextWrapping="Wrap"
                   Margin="0,6,0,0"/>
      </StackPanel>
    </GroupBox>

    <DockPanel Grid.Row="2" Margin="0,10,0,0">
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
        <Button x:Name="ok" Content="{{erstellen}}" Padding="14,3"
                Margin="0,0,6,0" IsDefault="True"/>
        <Button x:Name="abbrechen" Content="{{abbrechen}}" Width="90"
                IsCancel="True"/>
      </StackPanel>
      <TextBlock x:Name="zusammenfassung" VerticalAlignment="Center"
                 TextWrapping="Wrap"/>
    </DockPanel>
  </Grid>
</Window>""" % dlg.XMLNS


class EingabeFehler(Exception):
    """Fachlicher Fehler - nur als Text anzeigen."""


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
    """Ereignishandler mit Fehlerfang - eine Ausnahme im Handler würde
    ShowDialog() sonst wortlos beenden."""
    def handler(sender, args):
        try:
            funktion(sender, args)
        except EingabeFehler as fehler:
            meldung(besitzer_liefern(), u"%s" % fehler, warnung=True)
        except Exception as fehler:
            pfad = schreibe_fehlerprotokoll(traceback.format_exc())
            try:
                meldung(besitzer_liefern(), u"%s%s" % (
                    dlg.fehlertext(fehler),
                    t(u"\n\nTechnische Details: %s", u"\n\nTechnical details: %s",
                      u"\n\nDetalles técnicos: %s") % pfad if pfad else u""),
                    warnung=True)
            except Exception:
                pass
    return handler


def lies_einstellungen():
    try:
        with io.open(EINSTELLUNGEN, encoding="utf-8") as datei:
            return json.loads(datei.read())
    except Exception:
        return {}


def schreibe_einstellungen(werte):
    try:
        with io.open(EINSTELLUNGEN, "w", encoding="utf-8") as datei:
            datei.write(json.dumps(werte, ensure_ascii=False, indent=2))
    except Exception:
        pass


def _pinsel(rgb):
    pinsel = SolidColorBrush(FarbeWpf.FromRgb(Byte(int(rgb[0])), Byte(int(rgb[1])),
                                              Byte(int(rgb[2]))))
    pinsel.Freeze()
    return pinsel


def _box(text, markiert=False):
    box = CheckBox()
    box.Content = text
    box.IsChecked = markiert
    box.Margin = Thickness(2.0, 2.0, 2.0, 2.0)
    return box


class _Zeile(object):
    """Eine Filterzeile: Häkchen, Farbfeld, Filtername, Beschriftung."""

    def __init__(self, eintrag, markiert, text):
        self.eintrag = eintrag
        self.raster = Grid()
        self.raster.Margin = Thickness(0, 2, 0, 2)
        for breite in (24.0, 40.0, None, None):
            spalte = ColumnDefinition()
            spalte.Width = GridLength(breite) if breite is not None \
                else GridLength(1.0, GridUnitType.Star)
            self.raster.ColumnDefinitions.Add(spalte)
        self.box = CheckBox()
        self.box.IsChecked = markiert
        self.box.VerticalAlignment = VerticalAlignment.Center
        self.feld = Border()
        self.feld.Width = 32.0
        self.feld.Height = 16.0
        self.feld.BorderThickness = Thickness(1)
        self.feld.BorderBrush = _pinsel((90, 90, 90))
        self.feld.HorizontalAlignment = HorizontalAlignment.Left
        self.feld.VerticalAlignment = VerticalAlignment.Center
        self.name = TextBlock()
        self.name.Text = eintrag.name
        self.name.VerticalAlignment = VerticalAlignment.Center
        self.name.TextTrimming = TextTrimming.CharacterEllipsis
        self.name.Margin = Thickness(4, 0, 8, 0)
        self.text = TextBox()
        self.text.Text = text
        for index, element in enumerate((self.box, self.feld, self.name, self.text)):
            Grid.SetColumn(element, index)
            self.raster.Children.Add(element)

    def zeige_farbe(self, fuellung, muster_name):
        inhalt = TextBlock()
        inhalt.HorizontalAlignment = HorizontalAlignment.Center
        inhalt.VerticalAlignment = VerticalAlignment.Center
        inhalt.FontSize = 10.0
        hinweis = []
        if fuellung is None:
            self.feld.Background = _pinsel((255, 255, 255))
            inhalt.Text = u"–"
            hinweis.append(t(u"keine Farbe überschrieben", u"no colour overridden",
                             u"sin color modificado"))
        else:
            if fuellung.vg_muster == lg.VOLL:
                self.feld.Background = _pinsel(fuellung.vg_farbe)
            else:
                hinter = fuellung.hg_farbe if fuellung.hg_muster else (255, 255, 255)
                self.feld.Background = _pinsel(hinter)
                if fuellung.vg_muster:
                    inhalt.Text = u"▨▨"
                    inhalt.Foreground = _pinsel(fuellung.vg_farbe)
            if fuellung.vg_muster:
                hinweis.append(u"%s %s" % (muster_name(fuellung.vg_muster),
                                           lg.hexfarbe(fuellung.vg_farbe)))
            if fuellung.hg_muster:
                hinweis.append(t(u"Hintergrund: %s %s", u"Background: %s %s",
                                 u"Fondo: %s %s") % (muster_name(fuellung.hg_muster),
                                                     lg.hexfarbe(fuellung.hg_farbe)))
        self.feld.Child = inhalt
        zustand = []
        if not self.eintrag.sichtbar:
            zustand.append(t(u"in der Ansicht ausgeblendet", u"hidden in the view",
                             u"oculto en la vista"))
        if not self.eintrag.aktiv:
            zustand.append(t(u"Filter deaktiviert", u"filter disabled",
                             u"filtro desactivado"))
        if self.eintrag.abweichend:
            zustand.append(t(u"anders dargestellt in: %s - es gilt \"%s\"",
                             u"shown differently in: %s - \"%s\" applies",
                             u"se muestra distinto en: %s - se usa \"%s\"")
                           % (u", ".join(self.eintrag.abweichend), self.eintrag.ansicht))
        self.feld.ToolTip = u"\n".join(hinweis) or None
        self.name.ToolTip = u"\n".join([self.eintrag.name] + hinweis + zustand)
        self.name.Foreground = _pinsel((150, 90, 0) if zustand else (0, 0, 0))
        self.name.Text = self.eintrag.name + (u"  ⚠" if zustand else u"")


class LegendenFenster(object):

    def __init__(self, uiapp, doc, aktive_ansicht):
        self.doc = doc
        self.einstellungen = lies_einstellungen()
        self.fenster = dlg.lade_xaml(uebersetze_xaml(XAML, XAML_TEXTE))
        dlg.setze_besitzer(self.fenster, handle=uiapp.MainWindowHandle)
        self.c = self.fenster.FindName
        self.ergebnis = None
        self.muster = rv.muster(doc)
        self._funde = {}          # Ansicht-Id (Zahl) -> [Eintrag]
        self._eintraege = []
        self._zeilen = []
        self._zeilen_quelle = lg.QUELLE_AUTO     # Farbquelle der angezeigten Zeilen
        self._markiert = {}       # Filter-Id (Zahl) -> bool, in dieser Sitzung geklickt
        self._texte = dict(self.einstellungen.get(u"beschriftungen") or {})
        self._ansicht_boxen = []  # (Suchtext, CheckBox, Ansicht)
        self._fuelle_ansichten(aktive_ansicht)
        self._fuelle_einstellungen()
        self._verdrahte()
        self._aktualisiere()

    def _h(self, funktion):
        return sicher(lambda: self.fenster, funktion)

    def zeige(self):
        self.fenster.ShowDialog()
        return self.ergebnis

    # -- Aufbau ------------------------------------------------------------------

    def _fuelle_ansichten(self, aktive_ansicht):
        liste = self.c("ansichtliste")
        aktiv_id = aktive_ansicht.Id if aktive_ansicht is not None else None
        for ansicht, anzahl, art in rv.quellen(self.doc):
            text = u"%s  (%s, %d Filter)" % (rv.name_von(ansicht), tt(art), anzahl)
            box = _box(text, ansicht.Id == aktiv_id)
            box.Click += self._h(lambda s, a: self._aktualisiere())
            liste.Items.Add(box)
            self._ansicht_boxen.append((text.lower(), box, ansicht))

    def _auswahl(self, feld, eintraege, gespeichert, vorgabe):
        liste = self.c(feld)
        for _wert, texte in eintraege:
            liste.Items.Add(tt(texte))
        werte = [w for w, _ in eintraege]
        liste.SelectedIndex = werte.index(gespeichert) if gespeichert in werte \
            else werte.index(vorgabe)

    def _fuelle_einstellungen(self):
        c = self.c
        e = self.einstellungen
        c("name").Text = e.get(u"name") or t(u"Legende Filter", u"Filter legend",
                                             u"Leyenda de filtros")
        c("massstab").Text = u"%d" % e.get(u"massstab", 100)
        for feld, vorgabe in ((u"feld_breite", 10.0), (u"feld_hoehe", 5.0),
                              (u"abstand", 2.0), (u"spaltenbreite", 70.0)):
            c(feld).Text = dxf_lg.zahl_text(e.get(feld, vorgabe))
        c("spalten").Text = u"%d" % e.get(u"spalten", 1)
        self._auswahl("quelle", lg.QUELLEN, e.get(u"quelle"), lg.QUELLE_AUTO)
        self._auswahl("reihenfolge", lg.REIHENFOLGEN, e.get(u"reihenfolge"),
                      lg.REIHENFOLGE_ANSICHT)

        texttyp = c("texttyp")
        self._texttypen, vorgabe = rv.texttypen(self.doc)
        gewuenscht = e.get(u"texttyp")
        start = 0
        for index, (name, typ_id) in enumerate(self._texttypen):
            eintrag = ComboBoxItem()
            eintrag.Content = name
            texttyp.Items.Add(eintrag)
            if name == gewuenscht or (not gewuenscht and typ_id == vorgabe):
                start = index
        if self._texttypen:
            texttyp.SelectedIndex = start
        c("rahmen").IsChecked = bool(e.get(u"rahmen", True))
        c("ersetzen").IsChecked = bool(e.get(u"ersetzen", False))
        self._namehinweis()

    def _verdrahte(self):
        c = self.c
        c("suche").TextChanged += self._h(self._suche)
        c("ansichten_alle").Click += self._h(lambda s, a: self._markiere_ansichten(True))
        c("ansichten_keine").Click += self._h(lambda s, a: self._markiere_ansichten(False))
        c("filter_alle").Click += self._h(lambda s, a: self._markiere_filter(True))
        c("filter_keine").Click += self._h(lambda s, a: self._markiere_filter(False))
        c("reihenfolge").SelectionChanged += self._h(lambda s, a: self._zeige_filter())
        c("quelle").SelectionChanged += self._h(lambda s, a: self._zeige_filter())
        c("name").TextChanged += self._h(lambda s, a: self._namehinweis())
        c("ersetzen").Click += self._h(lambda s, a: self._namehinweis())
        c("ok").Click += self._h(self._ok)
        self.fenster.Loaded += self._h(lambda s, a: c("suche").Focus())

    # -- Ansichten ---------------------------------------------------------------

    def _suche(self, sender, args):
        woerter = (self.c("suche").Text or u"").lower().split()
        for text, box, _ansicht in self._ansicht_boxen:
            passt = all(w in text for w in woerter)
            box.Visibility = Visibility.Visible if passt else Visibility.Collapsed

    def _markiere_ansichten(self, zustand):
        for _text, box, _ansicht in self._ansicht_boxen:
            if box.Visibility == Visibility.Visible:
                box.IsChecked = zustand
        self._aktualisiere()

    def _gewaehlte_ansichten(self):
        return [a for _t, box, a in self._ansicht_boxen if box.IsChecked]

    def _aktualisiere(self):
        self._merke()
        gewaehlt = self._gewaehlte_ansichten()
        neu = [a for a in gewaehlt if rv.id_wert(a.Id) not in self._funde]
        if neu:
            Mouse.OverrideCursor = Cursors.Wait
            try:
                for ansicht in neu:
                    self._funde[rv.id_wert(ansicht.Id)] = rv.lies_filter(self.doc, ansicht)
            finally:
                Mouse.OverrideCursor = None
        self._eintraege = lg.zusammenfuehren(
            [self._funde[rv.id_wert(a.Id)] for a in gewaehlt])
        self.c("ansichten_zahl").Text = t(u"%d markiert", u"%d checked", u"%d marcadas") \
            % len(gewaehlt)
        leer = self.c("filter_leer")
        if not gewaehlt:
            leer.Text = t(u"Links mindestens eine Ansicht oder Vorlage markieren.",
                          u"Check at least one view or template on the left.",
                          u"Marque al menos una vista o plantilla a la izquierda.")
        else:
            leer.Text = u""
        leer.Visibility = Visibility.Visible if leer.Text else Visibility.Collapsed
        self._zeige_filter()

    # -- Filter ------------------------------------------------------------------

    def _quelle(self):
        index = self.c("quelle").SelectedIndex
        return lg.QUELLEN[index][0] if 0 <= index < len(lg.QUELLEN) else lg.QUELLE_AUTO

    def _reihenfolge(self):
        index = self.c("reihenfolge").SelectedIndex
        return lg.REIHENFOLGEN[index][0] if 0 <= index < len(lg.REIHENFOLGEN) \
            else lg.REIHENFOLGE_ANSICHT

    def _merke(self):
        """Häkchen und Texte der angezeigten Zeilen übernehmen."""
        for zeile in self._zeilen:
            schluessel = zeile.eintrag.schluessel
            markiert = bool(zeile.box.IsChecked)
            if markiert != lg.vorgewaehlt(zeile.eintrag, self._zeilen_quelle) \
                    or schluessel in self._markiert:
                self._markiert[schluessel] = markiert
            text = (zeile.text.Text or u"").strip()
            if text and text != zeile.eintrag.name:
                self._texte[zeile.eintrag.name] = text
            else:
                self._texte.pop(zeile.eintrag.name, None)

    def _zeige_filter(self):
        self._merke()
        quelle = self._quelle()
        liste = self.c("filterliste")
        liste.Children.Clear()
        self._zeilen = []
        self._zeilen_quelle = quelle
        for eintrag in lg.sortiere(self._eintraege, self._reihenfolge()):
            markiert = self._markiert.get(eintrag.schluessel,
                                          lg.vorgewaehlt(eintrag, quelle))
            zeile = _Zeile(eintrag, markiert, self._texte.get(eintrag.name, eintrag.name))
            zeile.zeige_farbe(lg.darstellung(eintrag.grafik, quelle), self.muster.name)
            zeile.box.Click += self._h(lambda s, a: self._zaehle())
            liste.Children.Add(zeile.raster)
            self._zeilen.append(zeile)
        if self._gewaehlte_ansichten() and not self._zeilen:
            leer = self.c("filter_leer")
            leer.Text = t(u"Die markierten Ansichten haben keine Filter.",
                          u"The checked views have no filters.",
                          u"Las vistas marcadas no tienen filtros.")
            leer.Visibility = Visibility.Visible
        self._zaehle()

    def _markiere_filter(self, zustand):
        for zeile in self._zeilen:
            zeile.box.IsChecked = zustand
            self._markiert[zeile.eintrag.schluessel] = zustand
        self._zaehle()

    def _gewaehlte_posten(self):
        quelle = self._quelle()
        posten = []
        for zeile in self._zeilen:
            if zeile.box.IsChecked:
                posten.append((zeile.eintrag,
                               lg.beschriftung(zeile.eintrag, zeile.text.Text),
                               lg.darstellung(zeile.eintrag.grafik, quelle) is not None))
        return posten

    def _zaehle(self):
        posten = self._gewaehlte_posten()
        mit_farbe = sum(1 for _e, _t, farbe in posten if farbe)
        text = t(u"%d Filter in der Legende", u"%d filters in the legend",
                 u"%d filtros en la leyenda") % mit_farbe if mit_farbe else u""
        if len(posten) > mit_farbe:
            text += t(u"  (%d ohne Farbe werden übersprungen)",
                      u"  (%d without colour are skipped)",
                      u"  (%d sin color se omiten)") % (len(posten) - mit_farbe)
        self.c("zusammenfassung").Text = text
        self.c("ok").IsEnabled = mit_farbe > 0

    # -- Einstellungen -----------------------------------------------------------

    def _namehinweis(self):
        name = (self.c("name").Text or u"").strip()
        hinweis = u""
        if name and rv.vorhandene_ansicht(self.doc, name) is not None:
            if self.c("ersetzen").IsChecked:
                hinweis = t(u"\"%s\" wird geleert und neu aufgebaut.",
                            u"\"%s\" will be cleared and rebuilt.",
                            u"\"%s\" se vaciará y se volverá a crear.") % name
            else:
                hinweis = t(u"\"%s\" gibt es schon - es entsteht \"%s (2)\". Zum "
                            u"Aktualisieren \"Gleichnamige Legende ersetzen\" anhaken.",
                            u"\"%s\" already exists - \"%s (2)\" is created. To update "
                            u"it tick \"Replace legend with the same name\".",
                            u"\"%s\" ya existe - se crea \"%s (2)\". Para actualizarla "
                            u"marque \"Reemplazar leyenda del mismo nombre\".") % (name, name)
        elif rv.zielart(self.doc) == u"zeichnung":
            hinweis = t(u"Im Projekt gibt es keine Legende zum Duplizieren - es entsteht "
                        u"eine Zeichnungsansicht (kann nur auf einen Plan).",
                        u"The project has no legend to duplicate - a drafting view is "
                        u"created (can only go on one sheet).",
                        u"El proyecto no tiene leyenda para duplicar - se crea una vista "
                        u"de diseño (solo puede ir en un plano).")
        self.c("namehinweis").Text = hinweis
        self.c("namehinweis").Visibility = Visibility.Visible if hinweis \
            else Visibility.Collapsed

    def _zahl(self, feld, kleinst, groesst, text, ganz=False):
        wert = lg.pruefe(self.c(feld).Text, kleinst, groesst, ganz)
        if wert is None:
            raise EingabeFehler(text)
        return wert

    # -- Abschluss ---------------------------------------------------------------

    def _ok(self, sender, args):
        name = (self.c("name").Text or u"").strip()
        if not name:
            raise EingabeFehler(t(u"Bitte einen Namen für die Legende eingeben.",
                                  u"Please enter a name for the legend.",
                                  u"Introduzca un nombre para la leyenda."))
        massstab = self._zahl("massstab", 1, 100000, t(
            u"Der Massstab muss eine ganze Zahl sein, z.B. 100.",
            u"The scale must be a whole number, e.g. 100.",
            u"La escala debe ser un número entero, p.ej. 100."), ganz=True)
        feld_text = t(u"Das Farbfeld muss zwischen 0.5 und 200 mm gross sein.",
                      u"The swatch must be between 0.5 and 200 mm.",
                      u"La muestra debe medir entre 0.5 y 200 mm.")
        feld_breite = self._zahl("feld_breite", 0.5, 200, feld_text)
        feld_hoehe = self._zahl("feld_hoehe", 0.5, 200, feld_text)
        abstand = self._zahl("abstand", 0, 200, t(
            u"Der Abstand muss zwischen 0 und 200 mm liegen.",
            u"The spacing must be between 0 and 200 mm.",
            u"La separación debe estar entre 0 y 200 mm."))
        spalten = self._zahl("spalten", 1, 20, t(
            u"Spalten: ganze Zahl von 1 bis 20.", u"Columns: whole number from 1 to 20.",
            u"Columnas: número entero de 1 a 20."), ganz=True)
        spaltenbreite = self._zahl("spaltenbreite", 1, 1000, t(
            u"Die Spaltenbreite muss zwischen 1 und 1000 mm liegen.",
            u"The column width must be between 1 and 1000 mm.",
            u"El ancho de columna debe estar entre 1 y 1000 mm."))
        posten = [(e, text) for e, text, farbe in self._gewaehlte_posten()]
        if not posten:
            raise EingabeFehler(t(u"Bitte mindestens einen Filter markieren.",
                                  u"Please check at least one filter.",
                                  u"Marque al menos un filtro."))
        texttyp_name, texttyp_id = (None, None)
        index = self.c("texttyp").SelectedIndex
        if 0 <= index < len(self._texttypen):
            texttyp_name, texttyp_id = self._texttypen[index]
        self._merke()
        werte = {u"name": name, u"massstab": massstab, u"feld_breite": feld_breite,
                 u"feld_hoehe": feld_hoehe, u"abstand": abstand, u"spalten": spalten,
                 u"spaltenbreite": spaltenbreite, u"quelle": self._quelle(),
                 u"rahmen": bool(self.c("rahmen").IsChecked),
                 u"ersetzen": bool(self.c("ersetzen").IsChecked)}
        self.ergebnis = dict(werte, posten=posten, texttyp=texttyp_id)
        werte.update({u"texttyp": texttyp_name, u"reihenfolge": self._reihenfolge(),
                      u"beschriftungen": self._texte})
        schreibe_einstellungen(werte)
        self.fenster.Close()


def starte(uiapp, doc, aktive_ansicht):
    """Dialog zeigen. Rückgabe Auswahl (dict) oder None."""
    return LegendenFenster(uiapp, doc, aktive_ansicht).zeige()
