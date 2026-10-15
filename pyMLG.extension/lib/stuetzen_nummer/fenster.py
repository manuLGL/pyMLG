# -*- coding: utf-8 -*-
"""Einstellungsfenster von ColumnNumbering (WPF).

    Vorgehen    ● Anklicken   ○ Pfad-Linie
    Nummer      Erste [S-01]  Schritt [1]   S-01, S-02, S-03 …
    Kategorien  [Suche      ]
                ☑ Tragwerksstützen  ☐ Türen  ☐ Wände ...
                ☑ darüber/darunter gleich
    Parameter   [Kennzeichen          v]
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

from System.Windows import Visibility  # noqa: E402
from System.Windows.Controls import CheckBox, Control  # noqa: E402
from System.Windows.Media import Color, SolidColorBrush  # noqa: E402

from filter_manager import dialoge as dlg  # noqa: E402
from mlg_sprache import t, uebersetze_xaml  # noqa: E402
from stuetzen_nummer import logik as lg  # noqa: E402
from stuetzen_nummer import revit as rv  # noqa: E402

TITEL = t(u"Elemente nummerieren", u"Number Elements", u"Numerar elementos")

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
    "titel": (u"Elemente nummerieren", u"Number Elements",
              u"Numerar elementos"),
    "vorgehen": (u"Vorgehen", u"Method", u"Método"),
    "klick": (u"Anklicken", u"Pick", u"Seleccionar"),
    "klick_text": (u"Elemente nacheinander in der Reihenfolge des Plans "
                   u"anklicken. ESC beendet.",
                   u"Pick the elements one after another in the order of "
                   u"the plan. ESC ends.",
                   u"Seleccione los elementos uno tras otro en el orden del "
                   u"plano. ESC termina."),
    "pfad": (u"Pfad-Linie", u"Path line", u"Línea de recorrido"),
    "pfad_text": (u"Modell- oder Detaillinien durch die Elemente zeichnen "
                  u"und anklicken (getrennte Reihen in der gewünschten "
                  u"Reihenfolge). Nummeriert wird in Zeichenrichtung.",
                  u"Draw model or detail lines through the elements and "
                  u"pick them (separate rows in the desired order). "
                  u"Numbering follows the drawing direction.",
                  u"Dibuje líneas de modelo o de detalle a través de los "
                  u"elementos y selecciónelas (filas separadas en el orden "
                  u"deseado). Se numera en el sentido del dibujo."),
    "nummer": (u"Nummer", u"Number", u"Número"),
    "erste": (u"Erste Nummer", u"First number", u"Primer número"),
    "schritt": (u"Schritt", u"Step", u"Incremento"),
    "zaehlt": (u"Die letzte Zahl wird hochgezählt, führende Nullen bleiben.",
               u"The last number is counted up, leading zeros are kept.",
               u"Se incrementa el último número, se conservan los ceros."),
    "parameter": (u"Schreiben in Parameter", u"Write to parameter",
                  u"Escribir en el parámetro"),
    "stuetzen": (u"Kategorien", u"Categories", u"Categorías"),
    "suche": (u"Suchen:", u"Search:", u"Buscar:"),
    "kat_text": (u"Nur Kategorien mit Elementen im Projekt.",
                 u"Only categories with elements in the project.",
                 u"Solo categorías con elementos en el proyecto."),
    "stapel": (u"Gleiche Nummer für Elemente derselben Kategorie darüber "
               u"und darunter (gleiche Lage auf anderen Ebenen; Wände, "
               u"Träger u. ä. bekommen immer eine eigene Nummer)",
               u"Same number for elements of the same category above and "
               u"below (same position on other levels; walls, beams etc. "
               u"always get their own number)",
               u"Mismo número para los elementos de la misma categoría "
               u"encima y debajo (misma posición en otros niveles; muros, "
               u"vigas, etc. siempre reciben su propio número)"),
    "abstand": (u"Pfad: max. Abstand Element - Linie (m)",
                u"Path: max. distance element - line (m)",
                u"Recorrido: distancia máx. elemento - línea (m)"),
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

    <GroupBox Header="{{stuetzen}}">
      <StackPanel>
        <Grid Margin="0,0,0,4">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <TextBlock Text="{{suche}}" VerticalAlignment="Center"
                     Margin="0,0,8,0"/>
          <TextBox x:Name="suche" Grid.Column="1" Padding="3"/>
        </Grid>
        <ScrollViewer Height="170" VerticalScrollBarVisibility="Auto"
                      BorderBrush="#ccc" BorderThickness="1">
          <StackPanel x:Name="kategorien" Margin="4,2"/>
        </ScrollViewer>
        <TextBlock Text="{{kat_text}}" Foreground="#555" TextWrapping="Wrap"/>
        <CheckBox x:Name="stapeln" Margin="0,6,0,2">
          <TextBlock Text="{{stapel}}" TextWrapping="Wrap"/>
        </CheckBox>
      </StackPanel>
    </GroupBox>

    <GroupBox Header="{{parameter}}">
      <ComboBox x:Name="parameter"/>
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


def _gemerkte_kategorien(einst):
    """Id-Werte der zuletzt gewählten Kategorien. Ältere Einstellungen
    kannten nur die Haken Tragwerk / Architektur / Wände."""
    if "kategorien" in einst:
        return [int(k) for k in einst["kategorien"]]
    alt = [(u"tragwerk", rv.TRAGWERK, True),
           (u"architektur", rv.ARCHITEKTUR, True),
           (u"wand", rv.WAND, False)]
    return [wert_ for name, wert_, standard in alt
            if einst.get(name, standard)]


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
                     "vorschau", "parameter", "suche", "kategorien",
                     "stapeln", "abstand"):
            setattr(self, name, f.FindName(name))

        einst = lade_einstellungen()
        pfad = einst.get("modus") == PFAD
        self.modus_pfad.IsChecked = pfad
        self.modus_klick.IsChecked = not pfad
        self.start.Text = einst.get("start", u"S-01")
        self.schritt.Text = u"%s" % einst.get("schritt", 1)
        self.stapeln.IsChecked = einst.get("stapeln", True)
        self.abstand.Text = einst.get("abstand", u"0,30")

        def s(funktion):
            return dlg.sicher(lambda: f, funktion)

        gewaehlt = set(_gemerkte_kategorien(einst))
        self.haken = []            # [(Id-Wert, CheckBox)]
        for wert_, name in rv.kategorie_optionen(doc):
            haken = CheckBox()
            haken.Content = name
            haken.IsChecked = wert_ in gewaehlt
            haken.Checked += s(lambda _s, _a: self._kategorien_geaendert())
            haken.Unchecked += s(lambda _s, _a: self._kategorien_geaendert())
            self.kategorien.Children.Add(haken)
            self.haken.append((wert_, haken))

        self.optionen = []
        self._parameter_fuellen(einst.get("parameter"))

        self.suche.TextChanged += s(lambda _s, _a: self._filtern())
        self.start.TextChanged += s(lambda _s, _a: self.pruefe())
        self.schritt.TextChanged += s(lambda _s, _a: self.pruefe())
        self.modus_klick.Checked += s(lambda _s, _a: self._modus_geaendert())
        self.modus_pfad.Checked += s(lambda _s, _a: self._modus_geaendert())
        f.FindName("starten").Click += s(lambda _s, _a: self.starte())
        self._modus_geaendert()
        self.pruefe()

    def gewaehlte_kategorien(self):
        return [wert_ for wert_, haken in self.haken if haken.IsChecked]

    def _filtern(self):
        suche = (self.suche.Text or u"").strip().lower()
        for _, haken in self.haken:
            sichtbar = (not suche or suche in (u"%s" % haken.Content).lower()
                        or haken.IsChecked)
            haken.Visibility = (Visibility.Visible if sichtbar
                                else Visibility.Collapsed)

    def _kategorien_geaendert(self):
        alt = (self.optionen[self.parameter.SelectedIndex][0]
               if 0 <= self.parameter.SelectedIndex < len(self.optionen)
               else None)
        self._parameter_fuellen(alt)

    def _parameter_fuellen(self, gemerkt):
        """Parameterliste passend zu den gewählten Kategorien."""
        self.optionen = (rv.parameter_optionen(
            self.doc, self.gewaehlte_kategorien())
            or rv.standard_optionen())
        self.parameter.Items.Clear()
        for _, name in self.optionen:
            self.parameter.Items.Add(name)
        schluessel = [s for s, _ in self.optionen]
        self.parameter.SelectedIndex = (schluessel.index(gemerkt)
                                        if gemerkt in schluessel else 0)

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
        kategorien = self.gewaehlte_kategorien()
        if not kategorien:
            dlg.meldung(self.fenster, t(
                u"Bitte mindestens eine Kategorie wählen.",
                u"Please choose at least one category.",
                u"Elija al menos una categoría."),
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
            "kategorien": kategorien,
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
