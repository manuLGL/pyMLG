# -*- coding: utf-8 -*-
"""Einstellungsfenster von ColumnNumbering (WPF).

    Vorgehen    ● Anklicken   ○ Pfad-Linie
    Nummer      Erste [S-01]  Schritt [1]   S-01, S-02, S-03 …
    Parameter   [Kennzeichen          v]
    Stützen     ☑ Tragwerk ☑ Architektur  ☑ darüber/darunter gleich
    Pfad        Abstand zur Linie [0,30] m
                                       [Starten] [Abbrechen]

Das Fenster sammelt nur die Einstellungen; das Anklicken passiert danach
in Revit (ein modales Fenster und PickObject vertragen sich nicht).
Die Einstellungen stehen in %LOCALAPPDATA%\\pyMLG\\ColumnNumbering.json.
"""

import io
import json
import os

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows.Controls import Control  # noqa: E402
from System.Windows.Media import Color, SolidColorBrush  # noqa: E402

from filter_manager import dialoge as dlg  # noqa: E402
from mlg_sprache import t, uebersetze_xaml  # noqa: E402
from stuetzen_nummer import logik as lg  # noqa: E402
from stuetzen_nummer import revit as rv  # noqa: E402

TITEL = t(u"Stützen nummerieren", u"Number Columns", u"Numerar pilares")

KLICK = "klick"
PFAD = "pfad"

EINSTELLUNGEN = os.path.join(dlg.protokollordner(), "ColumnNumbering.json")


def _pinsel(r, g, b):
    # Eingefroren, sonst gehört der Pinsel dem Thread, der das Modul geladen
    # hat (siehe filter_manager/fenster.py)
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


ROT = _pinsel(192, 57, 43)

XAML_TEXTE = {
    "titel": (u"Stützen nummerieren", u"Number Columns",
              u"Numerar pilares"),
    "vorgehen": (u"Vorgehen", u"Method", u"Método"),
    "klick": (u"Anklicken", u"Pick", u"Seleccionar"),
    "klick_text": (u"Stützen nacheinander in der Reihenfolge des Statik-Plans "
                   u"anklicken. ESC beendet.",
                   u"Pick the columns one after another in the order of the "
                   u"structural plan. ESC ends.",
                   u"Seleccione los pilares uno tras otro en el orden del "
                   u"plano de estructura. ESC termina."),
    "pfad": (u"Pfad-Linie", u"Path line", u"Línea de recorrido"),
    "pfad_text": (u"Modell- oder Detaillinien durch die Stützen zeichnen und "
                  u"anklicken (getrennte Reihen in der gewünschten "
                  u"Reihenfolge). Nummeriert wird in Zeichenrichtung.",
                  u"Draw model or detail lines through the columns and pick "
                  u"them (separate rows in the desired order). Numbering "
                  u"follows the drawing direction.",
                  u"Dibuje líneas de modelo o de detalle a través de los "
                  u"pilares y selecciónelas (filas separadas en el orden "
                  u"deseado). Se numera en el sentido del dibujo."),
    "nummer": (u"Nummer", u"Number", u"Número"),
    "erste": (u"Erste Nummer", u"First number", u"Primer número"),
    "schritt": (u"Schritt", u"Step", u"Incremento"),
    "zaehlt": (u"Die letzte Zahl wird hochgezählt, führende Nullen bleiben.",
               u"The last number is counted up, leading zeros are kept.",
               u"Se incrementa el último número, se conservan los ceros."),
    "parameter": (u"Schreiben in Parameter", u"Write to parameter",
                  u"Escribir en el parámetro"),
    "stuetzen": (u"Stützen", u"Columns", u"Pilares"),
    "tragwerk": (u"Tragwerksstützen", u"Structural columns",
                 u"Pilares estructurales"),
    "architektur": (u"Architektonische Stützen", u"Architectural columns",
                    u"Pilares arquitectónicos"),
    "stapel": (u"Gleiche Nummer für Stützen darüber und darunter "
               u"(gleiche Lage auf anderen Ebenen)",
               u"Same number for columns above and below "
               u"(same position on other levels)",
               u"Mismo número para los pilares de encima y debajo "
               u"(misma posición en otros niveles)"),
    "abstand": (u"Pfad: max. Abstand Stütze - Linie (m)",
                u"Path: max. distance column - line (m)",
                u"Recorrido: distancia máx. pilar - línea (m)"),
    "starten": (u"Starten", u"Start", u"Iniciar"),
    "abbrechen": (u"Abbrechen", u"Cancel", u"Cancelar"),
}

XAML = u"""
<Window %s Title="{{titel}} (pyMLG)" Width="460" SizeToContent="Height"
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
    <GroupBox Header="{{vorgehen}}">
      <StackPanel>
        <RadioButton x:Name="modus_klick" GroupName="modus"
                     FontWeight="SemiBold" Content="{{klick}}"/>
        <TextBlock Text="{{klick_text}}" TextWrapping="Wrap"
                   Foreground="#555" Margin="18,0,0,6"/>
        <RadioButton x:Name="modus_pfad" GroupName="modus"
                     FontWeight="SemiBold" Content="{{pfad}}"/>
        <TextBlock Text="{{pfad_text}}" TextWrapping="Wrap"
                   Foreground="#555" Margin="18,0,0,2"/>
      </StackPanel>
    </GroupBox>

    <GroupBox Header="{{nummer}}">
      <StackPanel>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="60"/>
          </Grid.ColumnDefinitions>
          <TextBlock Text="{{erste}}" VerticalAlignment="Center"
                     Margin="0,0,8,0"/>
          <TextBox x:Name="start" Grid.Column="1" Padding="3"/>
          <TextBlock Text="{{schritt}}" Grid.Column="2"
                     VerticalAlignment="Center" Margin="12,0,8,0"/>
          <TextBox x:Name="schritt" Grid.Column="3" Padding="3"/>
        </Grid>
        <TextBlock x:Name="vorschau" FontWeight="SemiBold"
                   Margin="0,6,0,0" TextWrapping="Wrap"/>
        <TextBlock Text="{{zaehlt}}" Foreground="#555" TextWrapping="Wrap"/>
      </StackPanel>
    </GroupBox>

    <GroupBox Header="{{parameter}}">
      <ComboBox x:Name="parameter"/>
    </GroupBox>

    <GroupBox Header="{{stuetzen}}">
      <StackPanel>
        <CheckBox x:Name="kat_tragwerk" Content="{{tragwerk}}"/>
        <CheckBox x:Name="kat_architektur" Content="{{architektur}}"/>
        <CheckBox x:Name="stapeln" Margin="0,6,0,2">
          <TextBlock Text="{{stapel}}" TextWrapping="Wrap"/>
        </CheckBox>
      </StackPanel>
    </GroupBox>

    <Grid Margin="0,0,0,10">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="60"/>
      </Grid.ColumnDefinitions>
      <TextBlock Text="{{abstand}}" VerticalAlignment="Center"/>
      <TextBox x:Name="abstand" Grid.Column="1" Padding="3"/>
    </Grid>

    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
      <Button x:Name="starten" Content="{{starten}}" FontWeight="Bold"
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


def merke_naechste(nummer):
    """Nach einem Lauf: nächstes Mal mit der folgenden Nummer beginnen."""
    werte = lade_einstellungen()
    werte["start"] = nummer
    speichere_einstellungen(werte)


class Einstellungen(object):
    def __init__(self, modus, start, schritt, schluessel, kategorien,
                 stapeln, abstand_m):
        self.modus = modus
        self.start = start
        self.schritt = schritt
        self.schluessel = schluessel
        self.kategorien = kategorien
        self.stapeln = stapeln
        self.abstand = abstand_m / rv.FUSS


class Fenster(object):

    def __init__(self, uiapp, doc):
        self.doc = doc
        self.ergebnis = None
        f = self.fenster = dlg.lade_xaml(uebersetze_xaml(XAML, XAML_TEXTE))
        try:
            dlg.setze_besitzer(f, handle=uiapp.MainWindowHandle)
        except Exception:
            pass
        for name in ("modus_klick", "modus_pfad", "start", "schritt",
                     "vorschau", "parameter", "kat_tragwerk",
                     "kat_architektur", "stapeln", "abstand"):
            setattr(self, name, f.FindName(name))

        einst = lade_einstellungen()
        pfad = einst.get("modus") == PFAD
        self.modus_pfad.IsChecked = pfad
        self.modus_klick.IsChecked = not pfad
        self.start.Text = einst.get("start", u"S-01")
        self.schritt.Text = u"%s" % einst.get("schritt", 1)
        self.kat_tragwerk.IsChecked = einst.get("tragwerk", True)
        self.kat_architektur.IsChecked = einst.get("architektur", True)
        self.stapeln.IsChecked = einst.get("stapeln", True)
        self.abstand.Text = einst.get("abstand", u"0,30")

        self.optionen = (rv.parameter_optionen(
            doc, [rv.TRAGWERK, rv.ARCHITEKTUR]) or rv.standard_optionen())
        for _, name in self.optionen:
            self.parameter.Items.Add(name)
        schluessel = [s for s, _ in self.optionen]
        gemerkt = einst.get("parameter")
        self.parameter.SelectedIndex = (schluessel.index(gemerkt)
                                        if gemerkt in schluessel else 0)

        def s(funktion):
            return dlg.sicher(lambda: f, funktion)

        self.start.TextChanged += s(lambda _s, _a: self.pruefe())
        self.schritt.TextChanged += s(lambda _s, _a: self.pruefe())
        self.modus_klick.Checked += s(lambda _s, _a: self._modus_geaendert())
        self.modus_pfad.Checked += s(lambda _s, _a: self._modus_geaendert())
        f.FindName("starten").Click += s(lambda _s, _a: self.starte())
        self._modus_geaendert()
        self.pruefe()

    def _modus_geaendert(self):
        self.abstand.IsEnabled = bool(self.modus_pfad.IsChecked)

    def _fehler(self, text):
        self.vorschau.Foreground = ROT
        self.vorschau.Text = text
        return None

    def pruefe(self):
        """Vorschau aktualisieren -> (Start, Schritt) oder None."""
        self.vorschau.ClearValue(Control.ForegroundProperty)
        start = (self.start.Text or u"").strip()
        if lg.pruefe_start(start):
            return self._fehler(t(
                u"Die erste Nummer braucht eine Zahl (z. B. S-01).",
                u"The first number needs a digit (e.g. S-01).",
                u"El primer número necesita una cifra (p. ej. S-01)."))
        try:
            schritt = int((self.schritt.Text or u"").strip())
            text = lg.vorschau(start, schritt)
        except ValueError:
            return self._fehler(t(
                u"Schritt: ganze Zahl ungleich 0, die Nummern dürfen nicht "
                u"negativ werden.",
                u"Step: whole number other than 0, numbers must not become "
                u"negative.",
                u"Incremento: número entero distinto de 0; los números no "
                u"pueden ser negativos."))
        self.vorschau.Text = text
        return start, schritt

    def starte(self):
        geprueft = self.pruefe()
        if geprueft is None:
            return
        start, schritt = geprueft
        kategorien = []
        if self.kat_tragwerk.IsChecked:
            kategorien.append(rv.TRAGWERK)
        if self.kat_architektur.IsChecked:
            kategorien.append(rv.ARCHITEKTUR)
        if not kategorien:
            dlg.meldung(self.fenster, t(
                u"Bitte mindestens eine Stützen-Kategorie wählen.",
                u"Please choose at least one column category.",
                u"Elija al menos una categoría de pilares."),
                titel=TITEL, warnung=True)
            return
        modus = PFAD if self.modus_pfad.IsChecked else KLICK
        try:
            abstand = lg.zahl(self.abstand.Text)
            if abstand <= 0:
                raise ValueError
        except ValueError:
            if modus == PFAD:
                dlg.meldung(self.fenster, t(
                    u"Abstand zur Linie: bitte eine Zahl größer 0 in Metern.",
                    u"Distance to line: please enter a number above 0 in "
                    u"meters.",
                    u"Distancia a la línea: introduzca un número mayor que "
                    u"0 en metros."), titel=TITEL, warnung=True)
                return
            abstand = 0.3
        schluessel = self.optionen[max(0, self.parameter.SelectedIndex)][0]
        self.ergebnis = Einstellungen(modus, start, schritt, schluessel,
                                      kategorien,
                                      bool(self.stapeln.IsChecked), abstand)
        speichere_einstellungen({
            "modus": modus, "start": start, "schritt": schritt,
            "parameter": schluessel,
            "tragwerk": rv.TRAGWERK in kategorien,
            "architektur": rv.ARCHITEKTUR in kategorien,
            "stapeln": bool(self.stapeln.IsChecked),
            "abstand": self.abstand.Text,
        })
        self.fenster.Close()

    def zeige(self):
        self.fenster.ShowDialog()
        return self.ergebnis


def frage_einstellungen(uiapp, doc):
    """Einstellungen oder None (abgebrochen)."""
    return Fenster(uiapp, doc).zeige()
