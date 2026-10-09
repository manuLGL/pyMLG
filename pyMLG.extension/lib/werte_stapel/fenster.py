# -*- coding: utf-8 -*-
"""Fenster von StackValues (WPF).

    Quelle: 37 Elemente - 12 Tragwerksstützen, 20 Wände, 5 Unterzüge
    ┌ Parameter ───────────┐ ┌ Zielgeschosse ──── Alle Keine ┐
    │ [Suchen           ]  │ │ ☑ DG              +9,40    12 │
    │ ☑ Kennzeichen        │ │ ☑ 1.OG            +3,20    37 │
    │ ☐ Kommentare         │ │ ☑ EG (Quelle)      0,00     – │
    └──────────────────────┘ └────────────────────────────────┘
    Lagetoleranz (cm) [5]  ☐ Vorhandene Werte überschreiben
    ☐ Wände/Unterzüge: gleiche Achse genügt
    49 Gegenstücke auf 2 Geschossen       [Übertragen] [Abbrechen]

Beide Listen erlauben Mehrfachauswahl (Klick, Strg, Shift): ein Haken in
einer markierten Zeile - oder die Leertaste - gilt für alle markierten.

Die Zahl je Geschoss ist die Anzahl der Elemente dort, die an derselben
Stelle wie eine Quelle liegen - sie folgt der Toleranz.
Die Einstellungen stehen in %LOCALAPPDATA%\\pyMLG\\StackValues.json.
"""

import io
import json
import os

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import GridLength, GridUnitType, HorizontalAlignment, \
    TextWrapping, Thickness, VerticalAlignment, Visibility  # noqa: E402
from System.Windows.Controls import CheckBox, ColumnDefinition, Control, \
    Grid, ListBoxItem, TextBlock  # noqa: E402
from System.Windows.Input import Key  # noqa: E402
from System.Windows.Media import Color, SolidColorBrush  # noqa: E402

from filter_manager import dialoge as dlg  # noqa: E402
from level_auto_set.revit import ebenen as alle_ebenen  # noqa: E402
from mlg_sprache import t, uebersetze_xaml  # noqa: E402
from werte_stapel import logik as lg  # noqa: E402
from werte_stapel import revit as rv  # noqa: E402

TITEL = t(u"Werte nach oben/unten übertragen", u"Copy Values Up/Down",
          u"Copiar valores arriba/abajo")

EINSTELLUNGEN = os.path.join(dlg.protokollordner(), "StackValues.json")


def _pinsel(r, g, b):
    # Eingefroren, sonst gehört der Pinsel dem Thread, der das Modul geladen
    # hat (siehe filter_manager/fenster.py)
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


ROT = _pinsel(192, 57, 43)
GRAU = _pinsel(140, 140, 140)

XAML_TEXTE = {
    "titel": (u"Werte nach oben/unten übertragen", u"Copy Values Up/Down",
              u"Copiar valores arriba/abajo"),
    "parameter": (u"Parameter übertragen", u"Parameters to copy",
                  u"Parámetros a copiar"),
    "suchen": (u"Suchen", u"Search", u"Buscar"),
    "geschosse": (u"Zielgeschosse", u"Target levels",
                  u"Niveles de destino"),
    "alle": (u"Alle", u"All", u"Todos"),
    "keine": (u"Keine", u"None", u"Ninguno"),
    "ebene": (u"Ebene", u"Level", u"Nivel"),
    "treffer": (u"Gegenstücke", u"Matches", u"Coincidencias"),
    "toleranz": (u"Lagetoleranz (cm)", u"Position tolerance (cm)",
                 u"Tolerancia de posición (cm)"),
    "toleranz_tip": (u"Stützen: Abstand der Einfügepunkte. Wände und "
                     u"Unterzüge: Abstand von Anfang, Ende und Mitte der "
                     u"Achse - die Richtung ist egal. Höhe (Z) zählt nie.",
                     u"Columns: distance of the insertion points. Walls and "
                     u"beams: distance of start, end and middle of the axis "
                     u"- direction does not matter. Height (Z) never counts.",
                     u"Pilares: distancia de los puntos de inserción. Muros "
                     u"y vigas: distancia del inicio, fin y centro del eje; "
                     u"la dirección no importa. La altura (Z) nunca cuenta."),
    "ueberschreiben": (u"Vorhandene Werte überschreiben (sonst nur leere "
                       u"Felder füllen)",
                       u"Overwrite existing values (otherwise only fill "
                       u"empty fields)",
                       u"Sobrescribir valores existentes (si no, solo "
                       u"rellenar campos vacíos)"),
    "achse": (u"Wände/Unterzüge: gleiche Achse genügt (Länge darf "
              u"abweichen, mind. halb überlappend)",
              u"Walls/beams: same axis is enough (length may differ, at "
              u"least half overlapping)",
              u"Muros/vigas: basta el mismo eje (la longitud puede variar, "
              u"solapando al menos la mitad)"),
    "achse_tip": (u"Aus: Anfang und Ende müssen innerhalb der Toleranz "
                  u"liegen. An: eine Wand darüber zählt auch, wenn sie "
                  u"kürzer oder länger ist, solange sie auf derselben Achse "
                  u"liegt und sich mit mindestens der Hälfte der kürzeren "
                  u"Wand überlappt.",
                  u"Off: start and end must lie within the tolerance. On: a "
                  u"wall above also counts if it is shorter or longer, as "
                  u"long as it lies on the same axis and overlaps at least "
                  u"half of the shorter wall.",
                  u"Desactivado: inicio y fin deben estar dentro de la "
                  u"tolerancia. Activado: un muro superior también cuenta "
                  u"si es más corto o más largo, siempre que esté en el "
                  u"mismo eje y solape al menos la mitad del muro más "
                  u"corto."),
    "uebertragen": (u"Übertragen", u"Copy", u"Copiar"),
    "abbrechen": (u"Abbrechen", u"Cancel", u"Cancelar"),
}

XAML = u"""
<Window %s Title="{{titel}} (pyMLG)" Width="760" Height="580"
        MinWidth="600" MinHeight="420" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12"
        ResizeMode="CanResizeWithGrip">
  <Window.Resources>
    <Style TargetType="GroupBox">
      <Setter Property="Padding" Value="6,4"/>
    </Style>
  </Window.Resources>
  <DockPanel Margin="10">
    <TextBlock x:Name="quelle" DockPanel.Dock="Top" TextWrapping="Wrap"
               FontWeight="SemiBold" Margin="0,0,0,8"/>

    <StackPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
      <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
        <TextBlock Text="{{toleranz}}" VerticalAlignment="Center"
                   Margin="0,0,8,0" ToolTip="{{toleranz_tip}}"/>
        <TextBox x:Name="toleranz" Width="60" Padding="3"
                 ToolTip="{{toleranz_tip}}"/>
        <CheckBox x:Name="ueberschreiben" VerticalAlignment="Center"
                  Margin="20,0,0,0" Content="{{ueberschreiben}}"/>
      </StackPanel>
      <CheckBox x:Name="achse" Content="{{achse}}" ToolTip="{{achse_tip}}"
                Margin="0,0,0,8"/>
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <TextBlock x:Name="zusammenfassung" VerticalAlignment="Center"
                   TextWrapping="Wrap" Margin="0,0,10,0"/>
        <StackPanel Grid.Column="1" Orientation="Horizontal">
          <Button x:Name="uebertragen" Content="{{uebertragen}}"
                  FontWeight="Bold" Padding="18,5" Margin="0,0,6,0"
                  IsDefault="True"/>
          <Button Content="{{abbrechen}}" Padding="12,5" IsCancel="True"/>
        </StackPanel>
      </Grid>
    </StackPanel>

    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="10"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <GroupBox Header="{{parameter}}">
        <DockPanel>
          <Grid DockPanel.Dock="Top" Margin="0,2,0,6">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock Text="{{suchen}}" VerticalAlignment="Center"
                       Margin="0,0,8,0"/>
            <TextBox x:Name="suche" Grid.Column="1" Padding="3"/>
          </Grid>
          <ListBox x:Name="parameter" HorizontalContentAlignment="Stretch"
                   SelectionMode="Extended"
                   ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                   ScrollViewer.VerticalScrollBarVisibility="Visible"/>
        </DockPanel>
      </GroupBox>

      <GroupBox Header="{{geschosse}}" Grid.Column="2">
        <DockPanel>
          <StackPanel DockPanel.Dock="Top" Orientation="Horizontal"
                      Margin="0,2,0,6">
            <Button x:Name="alle" Content="{{alle}}" Padding="8,3"
                    Margin="0,0,6,0"/>
            <Button x:Name="keine" Content="{{keine}}" Padding="8,3"/>
          </StackPanel>
          <Border DockPanel.Dock="Top" BorderBrush="#ABADB3"
                  BorderThickness="1,1,1,0" Background="#F3F3F3">
            <Grid Margin="6,3,24,3">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="80"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="{{ebene}}" FontWeight="SemiBold"/>
              <TextBlock Grid.Column="1" Text="{{treffer}}"
                         FontWeight="SemiBold" HorizontalAlignment="Right"/>
            </Grid>
          </Border>
          <ListBox x:Name="ebenen" HorizontalContentAlignment="Stretch"
                   SelectionMode="Extended"
                   ScrollViewer.VerticalScrollBarVisibility="Visible"/>
        </DockPanel>
      </GroupBox>
    </Grid>
  </DockPanel>
</Window>""" % dlg.XMLNS


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


class Einstellungen(object):
    def __init__(self, schluessel, ebenen, toleranz, ueberschreiben,
                 achse, zuordnung):
        self.schluessel = schluessel        # [Parameterschlüssel]
        self.ebenen = ebenen                # set der Ebenen-Id-Werte
        self.toleranz = toleranz            # Fuss
        self.ueberschreiben = ueberschreiben
        self.achse = achse                  # gleiche Achse genügt
        self.zuordnung = zuordnung          # logik.zuordnen(...)


class _Zeile(object):
    """Listenzeile: Kästchen, Name und weitere Spalten rechts. Ein Klick
    auf den Namen markiert die Zeile, nur das Kästchen setzt den Haken."""

    def __init__(self, name, rechts=(), umbrechen=False):
        raster = Grid()
        for breite in [20.0, None] + [b for b, _ in rechts]:
            spalte = ColumnDefinition()
            spalte.Width = (GridLength(breite) if breite
                            else GridLength(1.0, GridUnitType.Star))
            raster.ColumnDefinitions.Add(spalte)
        self.box = CheckBox()
        self.box.VerticalAlignment = VerticalAlignment.Center
        raster.Children.Add(self.box)
        self.name = TextBlock()
        self.name.Text = name
        self.name.VerticalAlignment = VerticalAlignment.Center
        if umbrechen:
            self.name.TextWrapping = TextWrapping.Wrap
        Grid.SetColumn(self.name, 1)
        raster.Children.Add(self.name)
        self.rechts = []
        for i, (_, text) in enumerate(rechts):
            block = TextBlock()
            block.Text = text
            block.HorizontalAlignment = HorizontalAlignment.Right
            block.VerticalAlignment = VerticalAlignment.Center
            Grid.SetColumn(block, 2 + i)
            raster.Children.Add(block)
            self.rechts.append(block)
        self.item = ListBoxItem()
        self.item.Content = raster
        self.item.Padding = Thickness(4.0, 1.0, 4.0, 1.0)


def _markierte(zeilen):
    return [z for z in zeilen if z.item.IsSelected
            and z.item.Visibility == Visibility.Visible]


def _haken_fuer_markierte(zeilen, zeile):
    """Haken in einer markierten Zeile: alle markierten übernehmen ihn."""
    if not zeile.item.IsSelected:
        return
    for z in _markierte(zeilen):
        z.box.IsChecked = zeile.box.IsChecked


def _leertaste(zeilen):
    """Markierte umschalten: alle an - ausser sie sind es schon."""
    markiert = _markierte(zeilen)
    if not markiert:
        return False
    zustand = not all(z.box.IsChecked for z in markiert)
    for z in markiert:
        z.box.IsChecked = zustand
    return True


class _Ebene(object):
    def __init__(self, kennung, name, anzeige, quelle):
        self.kennung = kennung
        self.name = name
        self.anzeige = anzeige
        self.quelle = quelle
        self.item = self.box = self.zahl = None


def quell_text(bestand):
    je = bestand.anzahl_je_kategorie()
    teile = [u"%d %s" % (je[k], rv.kategorie_name(k, je[k]))
             for k, _, _ in rv.KATEGORIEN if k in je]
    gesamt = len(bestand.quellen)
    return (t(u"Quelle: %d Elemente", u"Source: %d elements",
              u"Origen: %d elementos") % gesamt
            + (u" – " + u", ".join(teile) if teile else u""))


class Fenster(object):

    def __init__(self, uiapp, bestand, optionen):
        self.bestand = bestand
        self.optionen = optionen
        self.ergebnis = None
        self.zuordnung = {}
        f = self.fenster = dlg.lade_xaml(uebersetze_xaml(XAML, XAML_TEXTE))
        try:
            dlg.setze_besitzer(f, handle=uiapp.MainWindowHandle)
        except Exception:
            pass
        for name in ("quelle", "toleranz", "ueberschreiben", "achse",
                     "zusammenfassung", "suche", "parameter", "ebenen"):
            setattr(self, name, f.FindName(name))
        self.quelle.Text = quell_text(bestand)

        einst = lade_einstellungen()
        self.toleranz.Text = einst.get("toleranz", u"5")
        self.ueberschreiben.IsChecked = bool(einst.get("ueberschreiben",
                                                       False))
        self.achse.IsChecked = bool(einst.get("achse", False))
        self._baue_parameter(set(einst.get("parameter",
                                           [u"bip:ALL_MODEL_MARK"])))
        self._baue_ebenen(set(einst.get("abgewaehlt", [])))

        def s(funktion):
            return dlg.sicher(lambda: f, funktion)

        f.FindName("alle").Click += s(lambda _s, _a: self.markiere(True))
        f.FindName("keine").Click += s(lambda _s, _a: self.markiere(False))
        f.FindName("uebertragen").Click += s(lambda _s, _a: self.starte())
        self.toleranz.TextChanged += s(lambda _s, _a: self.zaehle())
        self.achse.Click += s(lambda _s, _a: self.zaehle())
        self.suche.TextChanged += s(lambda _s, _a: self.filtere())
        for zeilen, liste in ((self.zeilen, self.ebenen),
                              (self.param_zeilen, self.parameter)):
            for zeile in zeilen:
                zeile.box.Click += s(self._bei_haken(zeilen, zeile))
            liste.PreviewKeyDown += s(self._bei_taste(zeilen))
        self.zaehle()

    def _bei_haken(self, zeilen, zeile):
        def handler(_s, _a):
            _haken_fuer_markierte(zeilen, zeile)
            self.aktualisiere()
        return handler

    def _bei_taste(self, zeilen):
        def handler(_s, args):
            if args.Key == Key.Space and _leertaste(zeilen):
                args.Handled = True
                self.aktualisiere()
        return handler

    # --- Aufbau -----------------------------------------------------------
    def _baue_parameter(self, gemerkt):
        self.param_zeilen = []
        for option in self.optionen:
            zeile = _Zeile(option.anzeige, umbrechen=True)
            zeile.box.IsChecked = option.schluessel in gemerkt
            self.param_zeilen.append(zeile)
            self.parameter.Items.Add(zeile.item)

    def _baue_ebenen(self, abgewaehlt):
        quell_ebenen = set(q.ebene for q in self.bestand.quellen)
        self.zeilen = [_Ebene(e.id, e.name, e.anzeige, e.id in quell_ebenen)
                       for e in alle_ebenen(self.bestand.doc)]
        if any(z.ebene == rv.OHNE_EBENE for z in self.bestand.ziele):
            self.zeilen.append(_Ebene(rv.OHNE_EBENE,
                                      t(u"(ohne Ebene)", u"(no level)",
                                        u"(sin nivel)"), u"", False))
        for zeile in self.zeilen:
            liste = _Zeile(zeile.name + (t(u"  (Quelle)", u"  (source)",
                                           u"  (origen)")
                                         if zeile.quelle else u""),
                           rechts=((80.0, zeile.anzeige), (50.0, u"")))
            liste.box.IsChecked = zeile.name not in abgewaehlt
            liste.rechts[0].Foreground = GRAU
            zeile.item, zeile.box, zeile.zahl = (liste.item, liste.box,
                                                 liste.rechts[1])
            self.ebenen.Items.Add(zeile.item)

    # --- Zustand ----------------------------------------------------------
    def _toleranz(self):
        """Toleranz in Fuss oder None."""
        try:
            cm = lg.zahl(self.toleranz.Text)
        except ValueError:
            return None
        if cm <= 0 or cm > 100:
            return None
        return cm / 100.0 / rv.FUSS

    def gewaehlte_parameter(self):
        return [o.schluessel for o, z in zip(self.optionen, self.param_zeilen)
                if z.box.IsChecked]

    def gewaehlte_ebenen(self):
        return set(z.kennung for z in self.zeilen if z.box.IsChecked)

    def filtere(self):
        suche = (self.suche.Text or u"").strip().lower()
        for option, zeile in zip(self.optionen, self.param_zeilen):
            zeile.item.Visibility = (
                Visibility.Visible
                if not suche or suche in option.anzeige.lower()
                else Visibility.Collapsed)

    def markiere(self, zustand):
        for zeile in self.zeilen:
            zeile.box.IsChecked = zustand
        self.aktualisiere()

    def zaehle(self):
        """Zuordnung mit der aktuellen Toleranz neu bestimmen."""
        self.toleranz.ClearValue(Control.ForegroundProperty)
        toleranz = self._toleranz()
        if toleranz is None:
            self.toleranz.Foreground = ROT
            self.zuordnung = None
        else:
            self.zuordnung = lg.zuordnen(self.bestand.quellen,
                                         self.bestand.ziele, toleranz,
                                         bool(self.achse.IsChecked))
            je_ebene = lg.gegenstuecke_je_ebene(self.zuordnung,
                                                self.bestand.ziele)
            for zeile in self.zeilen:
                anzahl = je_ebene.get(zeile.kennung, 0)
                zeile.zahl.Text = u"%d" % anzahl if anzahl else u"–"
        self.aktualisiere()

    def aktualisiere(self):
        self.zusammenfassung.ClearValue(TextBlock.ForegroundProperty)
        if self.zuordnung is None:
            self.zusammenfassung.Foreground = ROT
            self.zusammenfassung.Text = t(
                u"Lagetoleranz: Zahl zwischen 0 und 100 cm.",
                u"Position tolerance: number between 0 and 100 cm.",
                u"Tolerancia de posición: número entre 0 y 100 cm.")
            return
        ebenen = self.gewaehlte_ebenen()
        ziele = [zi for zi in self.zuordnung
                 if self.bestand.ziele[zi].ebene in ebenen]
        anzahl_ebenen = len(set(self.bestand.ziele[zi].ebene for zi in ziele))
        text = t(u"%d Gegenstücke auf %d Geschossen",
                 u"%d matches on %d levels",
                 u"%d coincidencias en %d niveles") % (len(ziele),
                                                       anzahl_ebenen)
        anzahl_parameter = len(self.gewaehlte_parameter())
        text += u", " + t(u"%d Parameter", u"%d parameters",
                          u"%d parámetros") % anzahl_parameter
        self.zusammenfassung.Text = text
        if not ziele or not anzahl_parameter:
            self.zusammenfassung.Foreground = ROT

    # --- Ende -------------------------------------------------------------
    def starte(self):
        toleranz = self._toleranz()
        schluessel = self.gewaehlte_parameter()
        ebenen = self.gewaehlte_ebenen()
        fehler = None
        if toleranz is None:
            fehler = t(u"Bitte eine Lagetoleranz zwischen 0 und 100 cm "
                       u"eingeben.",
                       u"Please enter a position tolerance between 0 and "
                       u"100 cm.",
                       u"Introduzca una tolerancia de posición entre 0 y "
                       u"100 cm.")
        elif not schluessel:
            fehler = t(u"Bitte mindestens einen Parameter anhaken.",
                       u"Please check at least one parameter.",
                       u"Marque al menos un parámetro.")
        elif not ebenen:
            fehler = t(u"Bitte mindestens ein Zielgeschoss anhaken.",
                       u"Please check at least one target level.",
                       u"Marque al menos un nivel de destino.")
        if fehler:
            dlg.meldung(self.fenster, fehler, titel=TITEL, warnung=True)
            return
        speichere_einstellungen({
            "toleranz": self.toleranz.Text,
            "ueberschreiben": bool(self.ueberschreiben.IsChecked),
            "achse": bool(self.achse.IsChecked),
            "parameter": schluessel,
            "abgewaehlt": [z.name for z in self.zeilen
                           if not z.box.IsChecked],
        })
        self.ergebnis = Einstellungen(schluessel, ebenen, toleranz,
                                      bool(self.ueberschreiben.IsChecked),
                                      bool(self.achse.IsChecked),
                                      self.zuordnung)
        self.fenster.Close()

    def zeige(self):
        self.fenster.ShowDialog()
        return self.ergebnis


def frage_einstellungen(uiapp, bestand, optionen):
    """Einstellungen oder None (abgebrochen)."""
    return Fenster(uiapp, bestand, optionen).zeige()
