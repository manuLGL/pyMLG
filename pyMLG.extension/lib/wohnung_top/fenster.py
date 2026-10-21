# -*- coding: utf-8 -*-
"""Einstellungsfenster von ApartmentNumbering (WPF).

    Eingabe      Bauteil [TU      ]   Stiege [LBS 45   ]
    TOP          Erste Nummer [1]     1, 2, 3, 4 …
    Geschosse                                  Alle  Keine
                 ☑ OG12   +36,00   4 Wohnungen
                 ☑ OG11   +33,00   4 Wohnungen
                 ...
                                       [Weiter] [Abbrechen]

Das Fenster sammelt nur die Einstellungen; die Linie wird danach in Revit
gewählt (ein modales Fenster und PickObject vertragen sich nicht).
Die Einstellungen stehen in %LOCALAPPDATA%\\pyMLG\\ApartmentNumbering.json.
"""

import io
import json
import os

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import GridLength, HorizontalAlignment, \
    Thickness  # noqa: E402
from System.Windows.Controls import CheckBox, ColumnDefinition, Control,     Grid, TextBlock  # noqa: E402
from System.Windows.Media import Color, SolidColorBrush  # noqa: E402

from filter_manager import dialoge as dlg  # noqa: E402
from mlg_sprache import t, uebersetze_xaml  # noqa: E402
from stuetzen_nummer import logik as nummern  # noqa: E402

TITEL = t(u"Wohnungen nummerieren", u"Number Apartments",
          u"Numerar viviendas")

EINSTELLUNGEN = os.path.join(dlg.protokollordner(),
                             "ApartmentNumbering.json")


def _pinsel(r, g, b):
    # Eingefroren, sonst gehört der Pinsel dem Thread, der das Modul geladen
    # hat (siehe filter_manager/fenster.py)
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


ROT = _pinsel(192, 57, 43)
GRAU = _pinsel(110, 110, 110)

XAML_TEXTE = {
    "titel": (u"Wohnungen nummerieren", u"Number Apartments",
              u"Numerar viviendas"),
    "eingabe": (u"Für alle Wohnungen dieses Laufs",
                u"For all apartments of this run",
                u"Para todas las viviendas de esta ejecución"),
    "bauteil": (u"BAUTEIL", u"BAUTEIL", u"BAUTEIL"),
    "stiege": (u"STIEGE", u"STIEGE", u"STIEGE"),
    "leer": (u"Leer = Parameter bleibt unverändert.",
             u"Empty = parameter stays unchanged.",
             u"Vacío = el parámetro no cambia."),
    "top": (u"TOP", u"TOP", u"TOP"),
    "erste": (u"Erste Nummer", u"First number", u"Primer número"),
    "zaehlt": (u"Gezählt wird von unten nach oben, in jedem Geschoss "
               u"entlang der Linie. Führende Nullen bleiben (01, 02 …). "
               u"Für ein einzelnes Geschoss: nur dieses wählen und die "
               u"passende erste Nummer eintragen.",
               u"Counting goes from bottom to top, on each level along the "
               u"line. Leading zeros are kept (01, 02 …). For a single "
               u"level: choose only that level and enter the matching first "
               u"number.",
               u"Se cuenta de abajo arriba y en cada nivel a lo largo de la "
               u"línea. Se conservan los ceros a la izquierda (01, 02 …). "
               u"Para un solo nivel: elija solo ese nivel e introduzca el "
               u"primer número correspondiente."),
    "geschosse": (u"Geschosse", u"Levels", u"Niveles"),
    "alle": (u"Alle", u"All", u"Todos"),
    "keine": (u"Keine", u"None", u"Ninguno"),
    "geschoss_text": (u"Nur Geschosse mit WHG_TYP-Gruppen. Die Linie wird "
                      u"einmal gezeichnet und gilt für alle gewählten "
                      u"Geschosse.",
                      u"Only levels with WHG_TYP groups. The line is drawn "
                      u"once and applies to all chosen levels.",
                      u"Solo niveles con grupos WHG_TYP. La línea se dibuja "
                      u"una vez y vale para todos los niveles elegidos."),
    "weiter": (u"Linie wählen", u"Pick line", u"Seleccionar línea"),
    "abbrechen": (u"Abbrechen", u"Cancel", u"Cancelar"),
}

XAML = u"""
<Window %s Title="{{titel}} (pyMLG)" Width="440" SizeToContent="Height"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12">
  <Window.Resources>
    <Style TargetType="GroupBox">
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Margin" Value="0,2"/>
    </Style>
  </Window.Resources>
  <StackPanel Margin="10">
    <GroupBox Header="{{eingabe}}">
      <StackPanel>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <TextBlock Text="{{bauteil}}" VerticalAlignment="Center"
                     Margin="0,0,8,0"/>
          <TextBox x:Name="bauteil" Grid.Column="1" Padding="3"/>
          <TextBlock Text="{{stiege}}" Grid.Column="2"
                     VerticalAlignment="Center" Margin="12,0,8,0"/>
          <TextBox x:Name="stiege" Grid.Column="3" Padding="3"/>
        </Grid>
        <TextBlock Text="{{leer}}" Foreground="#555" TextWrapping="Wrap"
                   Margin="0,4,0,0"/>
      </StackPanel>
    </GroupBox>

    <GroupBox Header="{{top}}">
      <StackPanel>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="80"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <TextBlock Text="{{erste}}" VerticalAlignment="Center"
                     Margin="0,0,8,0"/>
          <TextBox x:Name="start" Grid.Column="1" Padding="3"/>
          <TextBlock x:Name="vorschau" Grid.Column="2" FontWeight="SemiBold"
                     VerticalAlignment="Center" Margin="12,0,0,0"
                     TextWrapping="Wrap"/>
        </Grid>
        <TextBlock Text="{{zaehlt}}" Foreground="#555" TextWrapping="Wrap"
                   Margin="0,4,0,0"/>
      </StackPanel>
    </GroupBox>

    <GroupBox Header="{{geschosse}}">
      <StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right"
                    Margin="0,0,0,4">
          <Button x:Name="alle" Content="{{alle}}" Padding="10,1"
                  Margin="0,0,4,0"/>
          <Button x:Name="keine" Content="{{keine}}" Padding="10,1"/>
        </StackPanel>
        <ScrollViewer Height="220" VerticalScrollBarVisibility="Auto"
                      BorderBrush="#ccc" BorderThickness="1">
          <StackPanel x:Name="geschosse" Margin="4,2"/>
        </ScrollViewer>
        <TextBlock Text="{{geschoss_text}}" Foreground="#555"
                   TextWrapping="Wrap" Margin="0,4,0,0"/>
      </StackPanel>
    </GroupBox>

    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="weiter" Content="{{weiter}}" FontWeight="Bold"
              Padding="18,5" Margin="0,0,6,0" IsDefault="True"/>
      <Button Content="{{abbrechen}}" Padding="12,5" IsCancel="True"/>
    </StackPanel>
  </StackPanel>
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
    def __init__(self, bauteil, stiege, start, ebenen):
        self.bauteil = bauteil
        self.stiege = stiege
        self.start = start
        self.ebenen = ebenen       # [revit.Ebene], von oben nach unten


def _hoehe_text(ebene):
    """+36,00 - in Metern, wie im Ribbon üblich."""
    meter = ebene.hoehe * 0.3048
    return (u"%+.2f" % meter).replace(u".", u",")


def _zeile(ebene):
    """Geschoss | Höhe | Anzahl - in festen Spalten, damit die Liste
    sauber untereinander steht."""
    zeile = Grid()
    for breite in (110, 70, 110):
        spalte = ColumnDefinition()
        spalte.Width = GridLength(breite)
        zeile.ColumnDefinitions.Add(spalte)
    texte = (ebene.name, _hoehe_text(ebene),
             t(u"%d Wohnungen", u"%d apartments", u"%d viviendas")
             % ebene.anzahl)
    for spalte, text in enumerate(texte):
        feld = TextBlock()
        feld.Text = text
        if spalte == 1:
            feld.HorizontalAlignment = HorizontalAlignment.Right
            feld.Foreground = GRAU
        elif spalte == 2:
            feld.Margin = Thickness(16, 0, 0, 0)
        Grid.SetColumn(feld, spalte)
        zeile.Children.Add(feld)
    return zeile


class Fenster(object):

    def __init__(self, uiapp, ebenen):
        self.ergebnis = None
        self.ebenen = ebenen
        f = self.fenster = dlg.lade_xaml(uebersetze_xaml(XAML, XAML_TEXTE))
        try:
            dlg.setze_besitzer(f, handle=uiapp.MainWindowHandle)
        except Exception:
            pass
        for name in ("bauteil", "stiege", "start", "vorschau", "geschosse"):
            setattr(self, name, f.FindName(name))

        einst = lade_einstellungen()
        self.bauteil.Text = einst.get("bauteil", u"")
        self.stiege.Text = einst.get("stiege", u"")
        self.start.Text = einst.get("start", u"1")

        def s(funktion):
            return dlg.sicher(lambda: f, funktion)

        gemerkt = einst.get("ebenen")
        self.haken = []            # [(Ebene, CheckBox)]
        for ebene in ebenen:
            haken = CheckBox()
            haken.Content = _zeile(ebene)
            haken.IsChecked = gemerkt is None or ebene.name in gemerkt
            self.geschosse.Children.Add(haken)
            self.haken.append((ebene, haken))

        f.FindName("alle").Click += s(lambda _s, _a: self._alle(True))
        f.FindName("keine").Click += s(lambda _s, _a: self._alle(False))
        self.start.TextChanged += s(lambda _s, _a: self.pruefe())
        f.FindName("weiter").Click += s(lambda _s, _a: self.weiter())
        self.pruefe()

    def _alle(self, an):
        for _, haken in self.haken:
            haken.IsChecked = an

    def pruefe(self):
        """Vorschau der Nummern -> Startnummer oder None."""
        self.vorschau.ClearValue(Control.ForegroundProperty)
        start = (self.start.Text or u"").strip()
        try:
            if nummern.pruefe_start(start):
                raise ValueError
            self.vorschau.Text = nummern.vorschau(start, 1)
        except ValueError:
            self.vorschau.Foreground = ROT
            self.vorschau.Text = t(u"Bitte eine Zahl eingeben (z. B. 1).",
                                   u"Please enter a number (e.g. 1).",
                                   u"Introduzca un número (p. ej. 1).")
            return None
        return start

    def weiter(self):
        start = self.pruefe()
        if start is None:
            return
        gewaehlt = [e for e, haken in self.haken if haken.IsChecked]
        if not gewaehlt:
            dlg.meldung(self.fenster, t(
                u"Bitte mindestens ein Geschoss wählen.",
                u"Please choose at least one level.",
                u"Elija al menos un nivel."), titel=TITEL, warnung=True)
            return
        bauteil = (self.bauteil.Text or u"").strip()
        stiege = (self.stiege.Text or u"").strip()
        self.ergebnis = Einstellungen(bauteil, stiege, start, gewaehlt)
        speichere_einstellungen({
            "bauteil": bauteil, "stiege": stiege, "start": start,
            "ebenen": [e.name for e in gewaehlt],
        })
        self.fenster.Close()

    def zeige(self):
        self.fenster.ShowDialog()
        return self.ergebnis


def frage_einstellungen(uiapp, ebenen):
    """Einstellungen oder None (abgebrochen)."""
    return Fenster(uiapp, ebenen).zeige()
