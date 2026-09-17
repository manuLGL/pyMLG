# -*- coding: utf-8 -*-
"""Hauptfenster von JoinMultiple (WPF), angelehnt an "JoinMultiple".

    ┌ Kategorie ─────────── Elemente ─ Priorität ┐ Alle  Keine
    │ ☑ Wände                    12     [300]    │ Bereich: ● Auswahl ○ Ansicht ○ Modell
    │ ☑ Geschossdecken            4     [200]    │ Verbinden: ☑ nur bei Schnitt ...
    │                                            │ Priorität: □ aus Parameter [____]
    │                                            │ 16 Elemente in der Auswahl
    └────────────────────────────────────────────┘ [Verbindung lösen] [Verbinden]

Höhere Priorität schneidet niedrigere. Prioritäten und Haken werden je
Kategorie in %LOCALAPPDATA%\\pyMLG\\JoinMultiple.json gemerkt.

Das Fenster ist modal: Die Revit-Auswahl wird beim Öffnen übernommen.
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
    Thickness, VerticalAlignment  # noqa: E402
from System.Windows.Controls import CheckBox, ColumnDefinition, Control, \
    Grid, ListBoxItem, TextBlock, TextBox  # noqa: E402
from System.Windows.Media import Color, SolidColorBrush  # noqa: E402

from filter_manager import dialoge as dlg  # noqa: E402
from join_multiple import logik as lg  # noqa: E402
from join_multiple import revit as rv  # noqa: E402

TITEL = u"JoinMultiple"


def _pinsel(r, g, b):
    # Eingefroren, sonst gehört der Pinsel dem Thread, der das Modul geladen
    # hat (siehe filter_manager/fenster.py)
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


ROT = _pinsel(192, 57, 43)

FEHLERPROTOKOLL = os.path.join(dlg.protokollordner(),
                               "JoinMultiple_Fehler.log")
EINSTELLUNGEN = os.path.join(dlg.protokollordner(), "JoinMultiple.json")

BEREICH_TEXT = {
    rv.AUSWAHL: u"Elemente in der Auswahl",
    rv.ANSICHT: u"Elemente in der aktuellen Ansicht",
    rv.MODELL: u"Elemente im ganzen Modell",
}

BREITE_ANZAHL = 80.0
BREITE_PRIO = 90.0

XAML = u"""
<Window %s Title="JoinMultiple (pyMLG)" Width="980" Height="640"
        MinWidth="760" MinHeight="520" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12"
        ResizeMode="CanResizeWithGrip">
  <Window.Resources>
    <Style TargetType="GroupBox">
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Margin" Value="0,3"/>
    </Style>
    <Style TargetType="RadioButton">
      <Setter Property="Margin" Value="0,3"/>
    </Style>
  </Window.Resources>
  <Grid Margin="10">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="10"/>
      <ColumnDefinition Width="270"/>
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Kategorien -->
    <DockPanel>
      <TextBlock DockPanel.Dock="Top" Foreground="#555" Margin="0,0,0,6"
                 TextWrapping="Wrap"
                 Text="Höhere Priorität schneidet niedrigere. Nur markierte Kategorien werden bearbeitet."/>
      <Border DockPanel.Dock="Top" BorderBrush="#ABADB3"
              BorderThickness="1,1,1,0" Background="#F3F3F3">
        <Grid Margin="6,4,24,4">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="80"/>
            <ColumnDefinition Width="90"/>
          </Grid.ColumnDefinitions>
          <TextBlock Text="Kategorie" FontWeight="SemiBold"/>
          <TextBlock Grid.Column="1" Text="Elemente" FontWeight="SemiBold"
                     HorizontalAlignment="Right" Margin="0,0,12,0"/>
          <TextBlock Grid.Column="2" Text="Priorität" FontWeight="SemiBold"/>
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

        <GroupBox Header="Bereich">
          <StackPanel>
            <RadioButton x:Name="bereich_auswahl"
                         Content="Ausgewählte Elemente"/>
            <RadioButton x:Name="bereich_ansicht"
                         Content="Elemente in aktueller Ansicht"/>
            <RadioButton x:Name="bereich_modell" Content="Alle Modellelemente"/>
            <CheckBox x:Name="alle_kategorien"
                      ToolTip="Sonst nur Kategorien, die sich mit 'Geometrie verbinden' zuverlässig bearbeiten lassen (Wände, Decken, Dächer, Stützen, Tragwerk, Fundamente, Allgemeines Modell ...)">
              <TextBlock Text="Alle Modellkategorien anzeigen"
                         TextWrapping="Wrap"/>
            </CheckBox>
          </StackPanel>
        </GroupBox>

        <GroupBox Header="Verbinden">
          <StackPanel>
            <CheckBox x:Name="nur_bei_schnitt"
                      ToolTip="Nur Elemente verbinden, deren Volumen sich überschneiden. Aus: auch Elemente, die sich nur berühren oder deren Umrisse sich überlagern.">
              <TextBlock Text="Nur wenn sich Elemente schneiden"
                         TextWrapping="Wrap"/>
            </CheckBox>
            <CheckBox x:Name="gleiche_kategorie">
              <TextBlock Text="Auch innerhalb derselben Kategorie"
                         TextWrapping="Wrap"/>
            </CheckBox>
            <CheckBox x:Name="bestehende_anpassen"
                      ToolTip="Bereits verbundene Elemente: Schnittreihenfolge nach Priorität umkehren, falls nötig">
              <TextBlock Text="Bereits verbundene nach Priorität anpassen"
                         TextWrapping="Wrap"/>
            </CheckBox>
          </StackPanel>
        </GroupBox>

        <GroupBox Header="Priorität">
          <StackPanel>
            <CheckBox x:Name="prio_parameter"
                      ToolTip="Zahlenwert eines Exemplar- oder Typparameters. Elemente ohne Wert nutzen die Priorität ihrer Kategorie.">
              <TextBlock Text="Priorität aus Parameter:"
                         TextWrapping="Wrap"/>
            </CheckBox>
            <TextBox x:Name="parametername" Padding="3" Margin="18,2,0,6"/>
            <Button x:Name="zuruecksetzen" Padding="8,4"
                    Content="Prioritäten zurücksetzen"
                    ToolTip="Alle Kategorien auf die Vorgabewerte setzen (Stützen 500 ... Wände 300, Decken 200 ...)"/>
          </StackPanel>
        </GroupBox>

        <CheckBox x:Name="auswaehlen" Margin="2,0,0,0">
          <TextBlock Text="Bearbeitete Elemente danach auswählen"
                     TextWrapping="Wrap"/>
        </CheckBox>
      </StackPanel>

      <StackPanel DockPanel.Dock="Bottom">
        <TextBlock x:Name="anzahl" FontSize="22" FontWeight="SemiBold"/>
        <TextBlock x:Name="anzahl_text" Foreground="#555" Margin="0,0,0,10"
                   TextWrapping="Wrap"/>
        <Button x:Name="loesen" Content="Verbindung lösen" Padding="8,6"
                Margin="0,0,0,6"/>
        <Button x:Name="verbinden" Content="Elemente verbinden"
                FontWeight="Bold" Padding="8,9"/>
      </StackPanel>
    </DockPanel>

    <!-- Fußzeile -->
    <DockPanel Grid.Row="1" Grid.ColumnSpan="3" Margin="0,10,0,0">
      <Button x:Name="schliessen" Content="Schließen" Width="100"
              DockPanel.Dock="Right" IsCancel="True"/>
      <TextBlock x:Name="status" VerticalAlignment="Center" Foreground="#555"
                 TextTrimming="CharacterEllipsis"/>
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


def _spalte(breite=None):
    spalte = ColumnDefinition()
    spalte.Width = (GridLength(breite) if breite
                    else GridLength(1.0, GridUnitType.Star))
    return spalte


# ---------------------------------------------------------------------------
# Fenster
# ---------------------------------------------------------------------------

class Kategorie(object):
    def __init__(self, schluessel, name):
        self.schluessel = schluessel      # Id-Wert als Text (JSON-Schlüssel)
        self.name = name
        self.elemente = []
        self.box = None
        self.prio_feld = None


class JoinMultipleFenster(object):

    def __init__(self, uiapp, uidoc):
        self.uiapp = uiapp
        self.uidoc = uidoc
        self.doc = uidoc.Document
        self.kategorien = []           # [Kategorie], sortiert nach Name
        self.vorgaben = dict((str(k), lg.standard_prioritaet(n))
                             for k, (_b, n)
                             in rv.standard_kategorien().items())

        f = self.fenster = dlg.lade_xaml(XAML)
        try:
            dlg.setze_besitzer(f, handle=uiapp.MainWindowHandle)
        except Exception:
            pass
        for name in ("liste", "bereich_auswahl", "bereich_ansicht",
                     "bereich_modell", "alle_kategorien", "nur_bei_schnitt",
                     "gleiche_kategorie", "bestehende_anpassen",
                     "prio_parameter", "parametername", "auswaehlen",
                     "anzahl", "anzahl_text", "status"):
            setattr(self, name, f.FindName(name))

        einst = self.einst = lade_einstellungen()
        self.prioritaeten = dict(einst.get("prioritaeten", {}))
        self.abgewaehlt = set(einst.get("abgewaehlt", []))
        self.alle_kategorien.IsChecked = einst.get("alle_kategorien", False)
        self.nur_bei_schnitt.IsChecked = einst.get("nur_bei_schnitt", True)
        self.gleiche_kategorie.IsChecked = einst.get("gleiche_kategorie",
                                                     True)
        self.bestehende_anpassen.IsChecked = einst.get(
            "bestehende_anpassen", True)
        self.prio_parameter.IsChecked = einst.get("prio_parameter", False)
        self.parametername.Text = einst.get("parametername", u"")
        self.auswaehlen.IsChecked = einst.get("auswaehlen", False)

        # Mit Auswahl startet der Bereich "Auswahl", sonst die Ansicht
        hat_auswahl = self.uidoc.Selection.GetElementIds().Count > 0
        self.bereich_auswahl.IsEnabled = hat_auswahl
        if hat_auswahl:
            self.bereich_auswahl.IsChecked = True
        elif einst.get("bereich") == rv.MODELL:
            self.bereich_modell.IsChecked = True
        else:
            self.bereich_ansicht.IsChecked = True
        if not hat_auswahl:
            self.bereich_auswahl.Content = u"Ausgewählte Elemente (keine)"

        self._verdrahte()
        self.lade()

    # --- Verdrahtung ------------------------------------------------------
    def _verdrahte(self):
        f = self.fenster

        def s(funktion):
            return sicher(lambda: f, funktion)

        def klick(name, funktion):
            f.FindName(name).Click += s(lambda _s, _a: funktion())

        klick("alle", lambda: self.markiere(True))
        klick("keine", lambda: self.markiere(False))
        klick("zuruecksetzen", self.setze_zurueck)
        klick("verbinden", self.verbinde)
        klick("loesen", self.loese)
        klick("schliessen", f.Close)
        for box in (self.bereich_auswahl, self.bereich_ansicht,
                    self.bereich_modell):
            box.Checked += s(lambda _s, _a: self.lade())
        self.alle_kategorien.Checked += s(lambda _s, _a: self.lade())
        self.alle_kategorien.Unchecked += s(lambda _s, _a: self.lade())
        f.Closing += s(lambda _s, _a: self._speichere())

    def _speichere(self):
        self._uebernimm_felder()
        speichere_einstellungen({
            "prioritaeten": self.prioritaeten,
            "abgewaehlt": sorted(self.abgewaehlt),
            "bereich": self.bereich(),
            "alle_kategorien": bool(self.alle_kategorien.IsChecked),
            "nur_bei_schnitt": bool(self.nur_bei_schnitt.IsChecked),
            "gleiche_kategorie": bool(self.gleiche_kategorie.IsChecked),
            "bestehende_anpassen": bool(self.bestehende_anpassen.IsChecked),
            "prio_parameter": bool(self.prio_parameter.IsChecked),
            "parametername": self.parametername.Text or u"",
            "auswaehlen": bool(self.auswaehlen.IsChecked),
        })

    # --- Liste ------------------------------------------------------------
    def bereich(self):
        if self.bereich_auswahl.IsChecked:
            return rv.AUSWAHL
        if self.bereich_modell.IsChecked:
            return rv.MODELL
        return rv.ANSICHT

    def prioritaet_text(self, schluessel):
        if schluessel in self.prioritaeten:
            return u"%s" % self.prioritaeten[schluessel]
        return u"%s" % self.vorgaben.get(schluessel, 0)

    def lade(self):
        """Elemente des Bereichs sammeln und nach Kategorie auflisten."""
        self._uebernimm_felder()
        elemente = rv.sammle(self.doc, self.uidoc, self.bereich(),
                             bool(self.alle_kategorien.IsChecked))
        nach_schluessel = {}
        for element in elemente:
            schluessel = str(rv.kategorie_schluessel(element))
            kategorie = nach_schluessel.get(schluessel)
            if kategorie is None:
                kategorie = nach_schluessel[schluessel] = Kategorie(
                    schluessel, element.Category.Name)
            kategorie.elemente.append(element)
        self.kategorien = sorted(nach_schluessel.values(),
                                 key=lambda k: k.name.casefold())
        self.zeichne()

    def zeichne(self):
        self.liste.Items.Clear()
        for kategorie in self.kategorien:
            zeile = Grid()
            for breite in (None, BREITE_ANZAHL, BREITE_PRIO):
                zeile.ColumnDefinitions.Add(_spalte(breite))

            box = CheckBox()
            box.Content = kategorie.name
            box.IsChecked = kategorie.schluessel not in self.abgewaehlt
            box.VerticalAlignment = VerticalAlignment.Center
            box.Click += sicher(lambda: self.fenster,
                                self._haken_handler(kategorie))
            zeile.Children.Add(box)

            anzahl = TextBlock()
            anzahl.Text = u"%d" % len(kategorie.elemente)
            anzahl.HorizontalAlignment = HorizontalAlignment.Right
            anzahl.VerticalAlignment = VerticalAlignment.Center
            anzahl.Margin = Thickness(0.0, 0.0, 12.0, 0.0)
            Grid.SetColumn(anzahl, 1)
            zeile.Children.Add(anzahl)

            feld = TextBox()
            feld.Text = self.prioritaet_text(kategorie.schluessel)
            feld.Width = 70.0
            feld.Padding = Thickness(2.0, 1.0, 2.0, 1.0)
            feld.HorizontalAlignment = HorizontalAlignment.Left
            feld.TextChanged += sicher(lambda: self.fenster,
                                       self._prio_handler(kategorie))
            Grid.SetColumn(feld, 2)
            zeile.Children.Add(feld)

            item = ListBoxItem()
            item.Content = zeile
            item.Padding = Thickness(4.0, 2.0, 4.0, 2.0)
            kategorie.box, kategorie.prio_feld = box, feld
            self.liste.Items.Add(item)
        self.aktualisiere_anzahl()

    def _haken_handler(self, kategorie):
        def handler(sender, _args):
            if sender.IsChecked:
                self.abgewaehlt.discard(kategorie.schluessel)
            else:
                self.abgewaehlt.add(kategorie.schluessel)
            self.aktualisiere_anzahl()
        return handler

    def _prio_handler(self, kategorie):
        def handler(sender, _args):
            try:
                lg.lies_prioritaet(sender.Text)
            except ValueError as fehler:
                sender.Foreground = ROT
                sender.ToolTip = u"%s" % fehler
                return
            sender.ClearValue(Control.ForegroundProperty)
            sender.ToolTip = None
        return handler

    def _uebernimm_felder(self):
        """Gültige Prioritäten aus den Eingabefeldern merken."""
        for kategorie in self.kategorien:
            if kategorie.prio_feld is None:
                continue
            try:
                wert = lg.lies_prioritaet(kategorie.prio_feld.Text)
            except ValueError:
                continue
            if wert == self.vorgaben.get(kategorie.schluessel, 0):
                self.prioritaeten.pop(kategorie.schluessel, None)
            else:
                self.prioritaeten[kategorie.schluessel] = wert

    def markierte(self):
        return [k for k in self.kategorien
                if k.schluessel not in self.abgewaehlt]

    def aktualisiere_anzahl(self):
        elemente = sum(len(k.elemente) for k in self.markierte())
        self.anzahl.Text = u"%d" % elemente
        self.anzahl_text.Text = BEREICH_TEXT[self.bereich()] + (
            u"" if len(self.markierte()) == len(self.kategorien)
            else u" (markierte Kategorien)")

    def markiere(self, zustand):
        for kategorie in self.kategorien:
            kategorie.box.IsChecked = zustand
            if zustand:
                self.abgewaehlt.discard(kategorie.schluessel)
            else:
                self.abgewaehlt.add(kategorie.schluessel)
        self.aktualisiere_anzahl()

    def setze_zurueck(self):
        if not dlg.frage(self.fenster, u"Die Prioritäten aller Kategorien "
                         u"auf die Vorgabewerte zurücksetzen?", titel=TITEL):
            return
        self.prioritaeten = {}
        for kategorie in self.kategorien:
            kategorie.prio_feld.Text = u"%s" % self.vorgaben.get(
                kategorie.schluessel, 0)
        self.status.Text = u"Prioritäten zurückgesetzt."

    # --- Ausführen --------------------------------------------------------
    def _vorbereiten(self, aktion):
        """Markierte Elemente und Prioritätsfunktion oder None."""
        kategorien = self.markierte()
        elemente = [e for k in kategorien for e in k.elemente]
        if not elemente:
            meldung(self.fenster, u"Keine Elemente in markierten "
                    u"Kategorien.")
            return None
        prios = {}
        for kategorie in kategorien:
            try:
                prios[kategorie.schluessel] = lg.lies_prioritaet(
                    kategorie.prio_feld.Text)
            except ValueError as fehler:
                meldung(self.fenster, u"%s: %s" % (kategorie.name, fehler),
                        warnung=True)
                return None
        self._uebernimm_felder()

        parameter = None
        if self.prio_parameter.IsChecked:
            parameter = (self.parametername.Text or u"").strip()
            if not parameter:
                meldung(self.fenster, u"Bitte den Namen des Parameters für "
                        u"die Priorität eingeben.", warnung=True)
                return None

        def prioritaet(element):
            if parameter:
                wert = rv.parameter_zahl(self.doc, element, parameter)
                if wert is not None:
                    return wert
            return prios[str(rv.kategorie_schluessel(element))]

        if self.bereich() == rv.MODELL and len(elemente) > 2000:
            if not dlg.frage(self.fenster, u"%d Elemente im ganzen Modell "
                             u"%s? Das kann einige Minuten dauern."
                             % (len(elemente), aktion), titel=TITEL):
                return None
        return elemente, prioritaet

    def verbinde(self):
        vorbereitet = self._vorbereiten(u"verbinden")
        if vorbereitet is None:
            return
        elemente, prioritaet = vorbereitet
        ergebnis = rv.verbinde(
            self.doc, elemente, prioritaet,
            nur_bei_schnitt=bool(self.nur_bei_schnitt.IsChecked),
            gleiche_kategorie=bool(self.gleiche_kategorie.IsChecked),
            bestehende_anpassen=bool(self.bestehende_anpassen.IsChecked))
        n = ergebnis.anzahl
        zeilen = [u"Neu verbunden: %d" % n[lg.NEU],
                  u"Bereits verbunden: %d" % n[lg.BEREITS]]
        if self.nur_bei_schnitt.IsChecked:
            zeilen.append(u"Übersprungen, schneiden sich nicht: %d"
                          % n[lg.UEBERSPRUNGEN])
        zeilen += [u"",
                   u"Schnittreihenfolge nach Priorität:",
                   u"  umgekehrt: %d" % n[lg.UMGEDREHT],
                   u"  stimmte bereits: %d" % n[lg.RICHTIG]]
        if n[lg.GLEICH]:
            zeilen.append(u"  gleiche Priorität, unverändert: %d"
                          % n[lg.GLEICH])
        if n[lg.NICHT_GEPRUEFT]:
            zeilen.append(
                u"  NICHT geprüft: %d bereits verbundene Paare - dafür "
                u"'Bereits verbundene nach Priorität anpassen' einschalten"
                % n[lg.NICHT_GEPRUEFT])
        if n[lg.ABGELEHNT]:
            zeilen.append(u"  von Revit nicht übernommen: %d"
                          % n[lg.ABGELEHNT])
            zeilen.extend(u"    " + t for t in ergebnis.abgelehnt[:10])
        if ergebnis.falsch_nach_speichern:
            zeilen.append(u"  nach dem Speichern FALSCH: %d"
                          % len(ergebnis.falsch_nach_speichern))
            zeilen.extend(u"    " + t
                          for t in ergebnis.falsch_nach_speichern[:10])
        else:
            zeilen.append(u"  nach dem Speichern kontrolliert: alles richtig")
        zeilen += [u"", u"Protokoll: %s" % rv.PROTOKOLL]
        self._abschluss(elemente, ergebnis, zeilen,
                        u"%d verbunden, %d umgekehrt" % (
                            n[lg.NEU], n[lg.UMGEDREHT]))

    def loese(self):
        vorbereitet = self._vorbereiten(u"lösen")
        if vorbereitet is None:
            return
        elemente, _prioritaet = vorbereitet
        ergebnis = rv.loese(
            self.doc, elemente,
            gleiche_kategorie=bool(self.gleiche_kategorie.IsChecked))
        self._abschluss(elemente, ergebnis,
                        [u"Verbindungen gelöst: %d" % ergebnis.geloest],
                        u"%d Verbindungen gelöst" % ergebnis.geloest)

    def _abschluss(self, elemente, ergebnis, zeilen, kurz, maximal=10):
        if self.auswaehlen.IsChecked and ergebnis.bearbeitet:
            rv.waehle_aus(self.uidoc, elemente, ergebnis.bearbeitet)
        if ergebnis.fehler:
            zeilen.append(u"\nFehlgeschlagen: %d" % len(ergebnis.fehler))
            zeilen.extend(u"  " + t for t in ergebnis.fehler[:maximal])
            if len(ergebnis.fehler) > maximal:
                zeilen.append(u"  ... und %d weitere"
                              % (len(ergebnis.fehler) - maximal))
        self.status.Text = kurz + (u", %d fehlgeschlagen"
                                   % len(ergebnis.fehler)
                                   if ergebnis.fehler else u"")
        meldung(self.fenster, u"\n".join(zeilen),
                warnung=bool(ergebnis.fehler
                             or getattr(ergebnis, "abgelehnt", None)))

    def zeige(self):
        self.fenster.ShowDialog()


def starte(uiapp, uidoc):
    JoinMultipleFenster(uiapp, uidoc).zeige()
