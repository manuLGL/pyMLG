# -*- coding: utf-8 -*-
"""Hauptfenster des Workset-Creators (WPF), angelehnt an "WorksetCreator".

    ┌ Vorhandene Worksets ─┬ Neue Worksets ──────────┬ [Worksets erstellen]  │
    │                      │ Alle Keine Umb. Entf.   │ Wenn vorhanden:        │
    │ Liste (Mehrfachwahl) │ ☑ Name     → Hinweis    │  ○ nicht erstellen     │
    │                      │                         │  ● umbenennen          │
    ├ Filter ──────────────┼ Filter ─────────────────┤ ☑ in allen Ansichten   │
    │ □ aktiv □ Regex [__] │ □ aktiv □ Regex [____]  │ ☑ an Liste anhängen    │
    └──────────────────────┴─────────────────────────┘ [offene Projekte ▾]    │
                                                      Aus Projekt / Datei /  │
                                                      Zwischenablage / Namen │

Die Liste "Neue Worksets" wird gesammelt; erst "Worksets erstellen" ändert das
Modell - in einer Transaktion (ein Rückgängig-Schritt). Das Fenster bleibt
danach offen.

Übernahme aus anderen Projekten: aus einem in Revit geöffneten Dokument
(auch Verknüpfungen) oder direkt aus einer .rvt-Datei über
WorksharingUtils.GetUserWorksetInfo - ohne die Datei zu öffnen.
"""

import io
import json
import os
import traceback

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import (  # noqa: E402
    FilteredWorksetCollector,
    ModelPathUtils,
    NamingUtils,
    Transaction,
    Workset,
    WorksetDefaultVisibilitySettings,
    WorksetKind,
    WorksetTable,
    WorksharingUtils,
)
from Microsoft.Win32 import OpenFileDialog  # noqa: E402
from System.Windows import Clipboard, Thickness, \
    VerticalAlignment, Visibility  # noqa: E402
from System.Windows.Controls import CheckBox, ComboBoxItem, Control, Dock, \
    DockPanel, ListBoxItem, TextBlock  # noqa: E402
from System.Windows.Input import Key, Keyboard  # noqa: E402
from System.Windows.Media import Color, SolidColorBrush  # noqa: E402

from filter_manager import dialoge as dlg  # noqa: E402
from workset_creator import logik as lg  # noqa: E402

TITEL = u"Workset-Creator"

# Vorgaben der deutschen Revit-Version beim Aktivieren der Teamarbeit
WS_EBENEN_RASTER = u"Gemeinsam genutzte Ebenen und Raster"
WS_STANDARD = u"Bearbeitungsbereich1"


def _pinsel(r, g, b):
    # Eingefroren, sonst gehört der Pinsel dem Thread, der das Modul geladen
    # hat (siehe filter_manager/fenster.py)
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


GRAU = _pinsel(120, 120, 120)
ORANGE = _pinsel(200, 110, 20)
ROT = _pinsel(192, 57, 43)
GRUEN = _pinsel(40, 130, 70)

FEHLERPROTOKOLL = os.path.join(dlg.protokollordner(),
                               "WorksetCreator_Fehler.log")
EINSTELLUNGEN = os.path.join(dlg.protokollordner(), "WorksetCreator.json")

XAML = u"""
<Window %s Title="Workset-Creator (pyMLG)" Width="1120" Height="660"
        MinWidth="820" MinHeight="440" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12"
        ResizeMode="CanResizeWithGrip">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="8,4"/>
      <Setter Property="Margin" Value="0,0,0,6"/>
    </Style>
    <Style TargetType="GroupBox">
      <Setter Property="Padding" Value="4"/>
    </Style>
  </Window.Resources>
  <Grid Margin="10">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*" MinWidth="240"/>
      <ColumnDefinition Width="8"/>
      <ColumnDefinition Width="*" MinWidth="300"/>
      <ColumnDefinition Width="10"/>
      <ColumnDefinition Width="210"/>
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Vorhandene Worksets -->
    <GroupBox x:Name="kopf_vorhanden" Grid.Column="0"
              Header="Vorhandene Worksets">
      <ListBox x:Name="liste_vorhanden" SelectionMode="Extended"
               ToolTip="Strg+C kopiert die markierten Namen"/>
    </GroupBox>
    <GroupBox Header="Filter vorhandene Worksets" Grid.Column="0"
              Grid.Row="1" Margin="0,6,0,0">
      <DockPanel>
        <CheckBox x:Name="filter_v_aktiv" Content="Aktiv"
                  VerticalAlignment="Center" Margin="0,0,10,0"/>
        <CheckBox x:Name="filter_v_regex" Content="Regex"
                  VerticalAlignment="Center" Margin="0,0,10,0"/>
        <TextBox x:Name="filter_v_text" Padding="3"/>
      </DockPanel>
    </GroupBox>

    <!-- Neue Worksets -->
    <GroupBox x:Name="kopf_neu" Grid.Column="2" Header="Neue Worksets">
      <DockPanel>
        <WrapPanel DockPanel.Dock="Top" Margin="0,0,0,4">
          <Button x:Name="alle" Content="Alle markieren" Margin="0,0,4,0"/>
          <Button x:Name="keine" Content="Keine" Margin="0,0,4,0"/>
          <Button x:Name="umbenennen" Content="Umbenennen..."
                  Margin="0,0,4,0" ToolTip="F2 oder Doppelklick"/>
          <Button x:Name="entfernen" Content="Entfernen" Margin="0,0,4,0"
                  ToolTip="Entf"/>
          <Button x:Name="leeren" Content="Liste leeren" Margin="0,0,4,0"/>
        </WrapPanel>
        <ListBox x:Name="liste_neu" SelectionMode="Extended"
                 HorizontalContentAlignment="Stretch"
                 ToolTip="Leertaste: Haken setzen/entfernen - Entf: entfernen - F2: umbenennen"/>
      </DockPanel>
    </GroupBox>
    <GroupBox Header="Filter neue Worksets" Grid.Column="2" Grid.Row="1"
              Margin="0,6,0,0">
      <DockPanel>
        <CheckBox x:Name="filter_n_aktiv" Content="Aktiv"
                  VerticalAlignment="Center" Margin="0,0,10,0"/>
        <CheckBox x:Name="filter_n_regex" Content="Regex"
                  VerticalAlignment="Center" Margin="0,0,10,0"/>
        <TextBox x:Name="filter_n_text" Padding="3"/>
      </DockPanel>
    </GroupBox>

    <!-- Aktionen -->
    <DockPanel Grid.Column="4" Grid.RowSpan="2" LastChildFill="False">
      <StackPanel DockPanel.Dock="Top">
        <Button x:Name="erstellen" Content="Worksets erstellen"
                FontWeight="Bold" Padding="8,9"/>
        <GroupBox Header="Wenn Workset existiert" Margin="0,2,0,8">
          <StackPanel Margin="2">
            <RadioButton x:Name="modus_ueberspringen" Content="Nicht erstellen"
                         Margin="0,2"/>
            <RadioButton x:Name="modus_umbenennen"
                         Content="Umbenennen und erstellen" Margin="0,2"
                         ToolTip="Hängt eine Nummer an: Name (2), Name (3) ..."/>
          </StackPanel>
        </GroupBox>
        <CheckBox x:Name="sichtbar" Margin="2,0,0,4"
                  ToolTip="Einstellung 'In allen Ansichten sichtbar' der neuen Worksets">
          <TextBlock Text="In allen Ansichten sichtbar" TextWrapping="Wrap"/>
        </CheckBox>
      </StackPanel>

      <StackPanel DockPanel.Dock="Bottom">
        <CheckBox x:Name="anhaengen" Margin="2,0,0,6"
                  ToolTip="Aus: die Liste 'Neue Worksets' wird ersetzt">
          <TextBlock Text="An Liste 'Neue Worksets' anhängen"
                     TextWrapping="Wrap"/>
        </CheckBox>
        <ComboBox x:Name="dokumente" Margin="0,0,0,4"
                  ToolTip="In Revit geöffnete Projekte mit Teamarbeit"/>
        <Button x:Name="aus_projekt" Content="Aus offenem Projekt"/>
        <Button x:Name="aus_datei" Content="Aus Datei (.rvt)..."
                ToolTip="Liest die Worksets, ohne die Datei zu öffnen"/>
        <Button x:Name="aus_zwischenablage" Content="Aus Zwischenablage"
                ToolTip="Ein Name je Zeile bzw. Zelle (z.B. aus Excel)"/>
        <Button x:Name="eingeben" Content="Namen eingeben..."/>
        <Button x:Name="in_zwischenablage"
                Content="Vorhandene in Zwischenablage"
                ToolTip="Markierte vorhandene Worksets, sonst alle angezeigten"/>
      </StackPanel>
    </DockPanel>

    <!-- Fußzeile -->
    <DockPanel Grid.Row="2" Grid.ColumnSpan="5" Margin="0,10,0,0">
      <Button x:Name="schliessen" Content="Schließen" Width="100"
              DockPanel.Dock="Right" IsCancel="True" Margin="0"/>
      <TextBlock x:Name="status" VerticalAlignment="Center" Foreground="#555"
                 TextTrimming="CharacterEllipsis"/>
    </DockPanel>
  </Grid>
</Window>""" % dlg.XMLNS

_XAML_NAMEN = u"""
<Window %s Title="Namen eingeben" Width="460" Height="480"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        FontFamily="Segoe UI" FontSize="12" MinWidth="320" MinHeight="260">
  <DockPanel Margin="12">
    <TextBlock DockPanel.Dock="Top" TextWrapping="Wrap" Margin="0,0,0,8"
               Text="Ein Workset-Name pro Zeile:"/>
    <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal"
                HorizontalAlignment="Right" Margin="0,10,0,0">
      <Button x:Name="ok" Content="Übernehmen" Width="100" Margin="0,0,8,0"/>
      <Button Content="Abbrechen" Width="90" IsCancel="True"/>
    </StackPanel>
    <TextBox x:Name="eingabe" AcceptsReturn="True" AcceptsTab="True"
             VerticalScrollBarVisibility="Auto" Padding="3"
             FontFamily="Consolas"/>
  </DockPanel>
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


def workset_namen(doc):
    if not doc.IsWorkshared:
        return []
    return [ws.Name for ws in
            FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)]


def workset_namen_aus_datei(pfad):
    """Worksets einer .rvt-Datei, ohne sie zu öffnen."""
    modellpfad = ModelPathUtils.ConvertUserVisiblePathToModelPath(pfad)
    try:
        vorschau = WorksharingUtils.GetUserWorksetInfo(modellpfad)
    except Exception as fehler:
        raise ValueError(
            u"Die Worksets der Datei konnten nicht gelesen werden.\n\n"
            u"Ist die Datei für Teamarbeit aktiviert?\n\n%s"
            % dlg.fehlertext(fehler))
    return [ws.Name for ws in vorschau]


def revit_gueltig(name):
    if not lg.ist_gueltig(name):
        return False
    try:
        return bool(NamingUtils.IsValidName(name))
    except Exception:
        return True


def sortiert(namen):
    return sorted(namen, key=lambda n: n.casefold())


# ---------------------------------------------------------------------------
# Fenster
# ---------------------------------------------------------------------------

class Eintrag(object):
    """Ein Name der Liste 'Neue Worksets'."""

    def __init__(self, name, markiert=True):
        self.name = name
        self.markiert = markiert
        self.item = None
        self.box = None


class WorksetCreatorFenster(object):

    def __init__(self, uiapp, doc):
        self.uiapp = uiapp
        self.doc = doc
        self.neu = []                  # [Eintrag]
        self.vorhanden = []            # [str]
        self.vorhanden_items = []      # [(ListBoxItem, Name)]
        self.fremde_dokumente = []     # parallel zur ComboBox

        f = self.fenster = dlg.lade_xaml(XAML)
        try:
            dlg.setze_besitzer(f, handle=uiapp.MainWindowHandle)
        except Exception:
            pass
        for name in ("kopf_vorhanden", "liste_vorhanden", "filter_v_aktiv",
                     "filter_v_regex", "filter_v_text", "kopf_neu",
                     "liste_neu", "filter_n_aktiv", "filter_n_regex",
                     "filter_n_text", "modus_ueberspringen",
                     "modus_umbenennen", "sichtbar", "anhaengen",
                     "dokumente", "status"):
            setattr(self, name, f.FindName(name))

        einst = lade_einstellungen()
        self.letzter_ordner = einst.get("ordner", u"")
        umbenennen = einst.get("umbenennen", True)
        self.modus_umbenennen.IsChecked = umbenennen
        self.modus_ueberspringen.IsChecked = not umbenennen
        self.sichtbar.IsChecked = einst.get("sichtbar", True)
        self.anhaengen.IsChecked = einst.get("anhaengen", True)

        self._verdrahte()
        self.lade_vorhandene()
        self.fuelle_dokumente()
        self.zeichne_neu()

    # --- Verdrahtung ------------------------------------------------------
    def _verdrahte(self):
        f = self.fenster

        def s(funktion):
            return sicher(lambda: f, funktion)

        def klick(name, funktion):
            f.FindName(name).Click += s(lambda _s, _a: funktion())

        klick("erstellen", self.erstelle)
        klick("alle", lambda: self.markiere(True))
        klick("keine", lambda: self.markiere(False))
        klick("umbenennen", self.benenne_um)
        klick("entfernen", self.entferne_markierte_zeilen)
        klick("leeren", self.leere)
        klick("aus_projekt", self.aus_projekt)
        klick("aus_datei", self.aus_datei)
        klick("aus_zwischenablage", self.aus_zwischenablage)
        klick("eingeben", self.eingeben)
        klick("in_zwischenablage", self.in_zwischenablage)
        klick("schliessen", f.Close)

        for box in (self.modus_ueberspringen, self.modus_umbenennen):
            box.Checked += s(lambda _s, _a: self.zeichne_neu())

        def filter_geaendert(aktiv, aufruf):
            def handler(_s, _a):
                if not aktiv.IsChecked:
                    aktiv.IsChecked = True     # löst aufruf erneut aus
                else:
                    aufruf()
            return handler

        self.filter_v_text.TextChanged += s(filter_geaendert(
            self.filter_v_aktiv, self.filtere_vorhandene))
        self.filter_n_text.TextChanged += s(filter_geaendert(
            self.filter_n_aktiv, self.filtere_neu))
        for box in (self.filter_v_aktiv, self.filter_v_regex):
            box.Checked += s(lambda _s, _a: self.filtere_vorhandene())
            box.Unchecked += s(lambda _s, _a: self.filtere_vorhandene())
        for box in (self.filter_n_aktiv, self.filter_n_regex):
            box.Checked += s(lambda _s, _a: self.filtere_neu())
            box.Unchecked += s(lambda _s, _a: self.filtere_neu())

        self.liste_neu.PreviewKeyDown += s(self._tasten_neu)
        self.liste_neu.MouseDoubleClick += s(
            lambda _s, _a: self.benenne_um(still=True))
        self.liste_vorhanden.PreviewKeyDown += s(self._tasten_vorhanden)
        self.dokumente.DropDownOpened += s(
            lambda _s, _a: self.fuelle_dokumente())
        f.Closing += s(lambda _s, _a: self._speichere())

    def _tasten_neu(self, _sender, args):
        if args.Key == Key.Delete:
            self.entferne_markierte_zeilen()
        elif args.Key == Key.F2:
            self.benenne_um()
        elif args.Key == Key.Space:
            gewaehlt = self._gewaehlte_neu()
            if gewaehlt:
                zustand = not all(e.markiert for e in gewaehlt)
                for e in gewaehlt:
                    e.markiert = zustand
                    e.box.IsChecked = zustand
                self.aktualisiere_status()
        else:
            return
        args.Handled = True

    def _tasten_vorhanden(self, _sender, args):
        strg = (Keyboard.IsKeyDown(Key.LeftCtrl)
                or Keyboard.IsKeyDown(Key.RightCtrl))
        if args.Key == Key.C and strg:
            self.in_zwischenablage()
            args.Handled = True

    def _speichere(self):
        speichere_einstellungen({
            "umbenennen": bool(self.modus_umbenennen.IsChecked),
            "sichtbar": bool(self.sichtbar.IsChecked),
            "anhaengen": bool(self.anhaengen.IsChecked),
            "ordner": self.letzter_ordner or u"",
        })

    # --- Vorhandene -------------------------------------------------------
    def lade_vorhandene(self):
        self.vorhanden = sortiert(workset_namen(self.doc))
        self.liste_vorhanden.Items.Clear()
        self.vorhanden_items = []
        for name in self.vorhanden:
            item = ListBoxItem()
            item.Content = name
            self.liste_vorhanden.Items.Add(item)
            self.vorhanden_items.append((item, name))
        self.filtere_vorhandene()

    def _filter(self, aktiv, regex, textfeld):
        """Prüffunktion oder None (kein Filter). Ungültiger Regex -> rot."""
        textfeld.ClearValue(Control.ForegroundProperty)
        textfeld.ToolTip = None
        if not aktiv.IsChecked:
            return None
        try:
            return lg.filterfunktion(textfeld.Text, bool(regex.IsChecked))
        except Exception as fehler:
            textfeld.Foreground = ROT
            textfeld.ToolTip = u"Ungültiger regulärer Ausdruck: %s" % fehler
            return None

    def filtere_vorhandene(self):
        passt = self._filter(self.filter_v_aktiv, self.filter_v_regex,
                             self.filter_v_text)
        sichtbar = 0
        for item, name in self.vorhanden_items:
            zeigen = passt is None or passt(name)
            item.Visibility = (Visibility.Visible if zeigen
                               else Visibility.Collapsed)
            sichtbar += zeigen
        if sichtbar == len(self.vorhanden):
            anzahl = u"%d" % sichtbar
        else:
            anzahl = u"%d von %d" % (sichtbar, len(self.vorhanden))
        kopf = u"Vorhandene Worksets (%s)" % anzahl
        if not self.doc.IsWorkshared:
            kopf = u"Vorhandene Worksets - Teamarbeit nicht aktiviert"
        self.kopf_vorhanden.Header = kopf

    def in_zwischenablage(self):
        angezeigt = [(item, n) for item, n in self.vorhanden_items
                     if item.Visibility == Visibility.Visible]
        # Markierte, die der Filter gerade ausblendet, bleiben außen vor
        namen = ([n for item, n in angezeigt if item.IsSelected]
                 or [n for _item, n in angezeigt])
        if not namen:
            meldung(self.fenster, u"Keine vorhandenen Worksets angezeigt.")
            return
        Clipboard.SetText(u"\r\n".join(namen))
        self.status.Text = u"%d Namen in die Zwischenablage kopiert." % len(
            namen)

    # --- Neue Worksets ----------------------------------------------------
    def zeichne_neu(self):
        """Liste neu aufbauen - mit Hinweis, was beim Erstellen passiert."""
        self.liste_neu.Items.Clear()
        plan = lg.plane([e.name for e in self.neu], self.vorhanden,
                        bool(self.modus_umbenennen.IsChecked), revit_gueltig)
        for eintrag, (_name, ziel, status) in zip(self.neu, plan):
            item = ListBoxItem()
            zeile = DockPanel()
            box = CheckBox()
            box.IsChecked = eintrag.markiert
            box.VerticalAlignment = VerticalAlignment.Center
            box.Margin = Thickness(0.0, 1.0, 6.0, 1.0)
            box.Click += sicher(lambda: self.fenster,
                                self._haken_handler(eintrag))
            DockPanel.SetDock(box, Dock.Left)
            zeile.Children.Add(box)

            hinweis = TextBlock()
            hinweis.Margin = Thickness(10.0, 0.0, 0.0, 0.0)
            hinweis.VerticalAlignment = VerticalAlignment.Center
            if status == lg.UMBENANNT:
                hinweis.Text, hinweis.Foreground = u"→ %s" % ziel, ORANGE
            elif status == lg.UEBERSPRUNGEN:
                hinweis.Text, hinweis.Foreground = u"existiert", GRAU
            elif status == lg.UNGUELTIG:
                hinweis.Text, hinweis.Foreground = u"ungültige Zeichen", ROT
            DockPanel.SetDock(hinweis, Dock.Right)
            zeile.Children.Add(hinweis)

            text = TextBlock()
            text.Text = eintrag.name
            text.VerticalAlignment = VerticalAlignment.Center
            if status == lg.UEBERSPRUNGEN:
                text.Foreground = GRAU
            zeile.Children.Add(text)

            item.Content = zeile
            eintrag.item, eintrag.box = item, box
            self.liste_neu.Items.Add(item)
        self.filtere_neu()

    def _haken_handler(self, eintrag):
        def handler(sender, _args):
            eintrag.markiert = bool(sender.IsChecked)
            self.aktualisiere_status()
        return handler

    def filtere_neu(self):
        passt = self._filter(self.filter_n_aktiv, self.filter_n_regex,
                             self.filter_n_text)
        for e in self.neu:
            zeigen = passt is None or passt(e.name)
            e.item.Visibility = (Visibility.Visible if zeigen
                                 else Visibility.Collapsed)
        self.aktualisiere_status()

    def _sichtbare_neu(self):
        return [e for e in self.neu
                if e.item is None or e.item.Visibility == Visibility.Visible]

    def _gewaehlte_neu(self):
        return [e for e in self.neu if e.item is not None
                and e.item.IsSelected
                and e.item.Visibility == Visibility.Visible]

    def aktualisiere_status(self):
        sichtbar = len(self._sichtbare_neu())
        anzahl = (u"%d" % len(self.neu) if sichtbar == len(self.neu)
                  else u"%d von %d" % (sichtbar, len(self.neu)))
        self.kopf_neu.Header = u"Neue Worksets (%s)" % anzahl
        markiert = [e.name for e in self.neu if e.markiert]
        plan = lg.plane(markiert, self.vorhanden,
                        bool(self.modus_umbenennen.IsChecked), revit_gueltig)
        zaehl = dict((s, 0) for s in (lg.NEU, lg.UMBENANNT,
                                      lg.UEBERSPRUNGEN, lg.UNGUELTIG))
        for _n, _z, status in plan:
            zaehl[status] += 1
        teile = [u"%d markiert" % len(markiert)]
        if markiert:
            teile.append(u"werden erstellt: %d" % (
                zaehl[lg.NEU] + zaehl[lg.UMBENANNT]))
            if zaehl[lg.UMBENANNT]:
                teile.append(u"davon umbenannt: %d" % zaehl[lg.UMBENANNT])
            if zaehl[lg.UEBERSPRUNGEN]:
                teile.append(u"übersprungen: %d" % zaehl[lg.UEBERSPRUNGEN])
            if zaehl[lg.UNGUELTIG]:
                teile.append(u"ungültig: %d" % zaehl[lg.UNGUELTIG])
        self.status.Text = u" - ".join(teile)

    def markiere(self, zustand):
        """Wirkt nur auf die angezeigten Einträge (Filter)."""
        for e in self._sichtbare_neu():
            e.markiert = zustand
            e.box.IsChecked = zustand
        self.aktualisiere_status()

    def uebernimm(self, namen, quelle):
        if not namen:
            self.status.Text = u"%s: keine Namen gefunden." % quelle
            return
        anhaengen = bool(self.anhaengen.IsChecked)
        bisher = [e.name for e in self.neu]
        liste, hinzu = lg.zusammenfuehren(bisher, namen, anhaengen)
        alte = dict((lg.schluessel(e.name), e) for e in self.neu)
        self.neu = [alte.get(lg.schluessel(n)) or Eintrag(n) for n in liste]
        self.zeichne_neu()
        if anhaengen:
            text = u"%s: %d Namen gelesen, %d neu in der Liste." % (
                quelle, len(namen), hinzu)
        else:
            text = u"%s: Liste mit %d Namen ersetzt." % (quelle, len(liste))
        self.status.Text = text

    def entferne_markierte_zeilen(self):
        gewaehlt = self._gewaehlte_neu()
        if not gewaehlt:
            return
        self.neu = [e for e in self.neu if e not in gewaehlt]
        self.zeichne_neu()

    def leere(self):
        if not self.neu:
            return
        if not dlg.frage(self.fenster, u"Liste 'Neue Worksets' mit %d "
                         u"Einträgen leeren?" % len(self.neu), titel=TITEL):
            return
        self.neu = []
        self.zeichne_neu()

    def benenne_um(self, still=False):
        gewaehlt = self._gewaehlte_neu()
        if len(gewaehlt) != 1:
            if still:
                return
            meldung(self.fenster, u"Bitte genau einen Eintrag in "
                    u"'Neue Worksets' auswählen.")
            return
        eintrag = gewaehlt[0]
        andere = set(lg.schluessel(e.name) for e in self.neu
                     if e is not eintrag)

        def pruefen(text):
            name = lg.bereinige(text)
            if not name:
                raise ValueError(u"Der Name darf nicht leer sein.")
            if not revit_gueltig(name):
                raise ValueError(u"Nicht erlaubt: %s"
                                 % u" ".join(lg.VERBOTENE_ZEICHEN))
            if lg.schluessel(name) in andere:
                raise ValueError(u"Der Name steht schon in der Liste.")
            return name

        name = dlg.frage_text(self.fenster, u"Umbenennen",
                              u"Neuer Name des Worksets:", eintrag.name,
                              pruefen)
        if name:
            eintrag.name = name
            self.zeichne_neu()
            eintrag.item.IsSelected = True

    # --- Quellen ----------------------------------------------------------
    def fuelle_dokumente(self):
        vorher = self.dokumente.SelectedIndex
        vorher_titel = (self.fremde_dokumente[vorher].Title
                        if 0 <= vorher < len(self.fremde_dokumente) else None)
        self.dokumente.Items.Clear()
        self.fremde_dokumente = []
        for d in self.uiapp.Application.Documents:
            try:
                if (d.Equals(self.doc) or d.IsFamilyDocument
                        or not d.IsWorkshared):
                    continue
            except Exception:
                continue
            item = ComboBoxItem()
            item.Content = d.Title + (u"  (Verknüpfung)" if d.IsLinked
                                      else u"")
            self.dokumente.Items.Add(item)
            self.fremde_dokumente.append(d)
        if not self.fremde_dokumente:
            item = ComboBoxItem()
            item.Content = u"(keine weiteren Projekte geöffnet)"
            item.IsEnabled = False
            self.dokumente.Items.Add(item)
            return
        titel = [d.Title for d in self.fremde_dokumente]
        self.dokumente.SelectedIndex = (titel.index(vorher_titel)
                                        if vorher_titel in titel else 0)

    def aus_projekt(self):
        index = self.dokumente.SelectedIndex
        if not 0 <= index < len(self.fremde_dokumente):
            meldung(self.fenster, u"Es ist kein weiteres Projekt mit "
                    u"Teamarbeit geöffnet.\n\nWorksets lassen sich auch "
                    u"über 'Aus Datei (.rvt)...' übernehmen, ohne die Datei "
                    u"zu öffnen.")
            return
        quelle = self.fremde_dokumente[index]
        self.uebernimm(workset_namen(quelle), quelle.Title)

    def aus_datei(self):
        dialog = OpenFileDialog()
        dialog.Title = u"Projekt wählen, dessen Worksets übernommen werden"
        dialog.Filter = u"Revit-Projekte (*.rvt)|*.rvt|Alle Dateien (*.*)|*.*"
        dialog.CheckFileExists = True
        if self.letzter_ordner and os.path.isdir(self.letzter_ordner):
            dialog.InitialDirectory = self.letzter_ordner
        if dialog.ShowDialog(self.fenster) != True:  # noqa: E712 - Nullable
            return
        pfad = dialog.FileName
        self.letzter_ordner = os.path.dirname(pfad)
        try:
            namen = workset_namen_aus_datei(pfad)
        except ValueError as fehler:
            meldung(self.fenster, u"%s" % fehler, warnung=True)
            return
        self.uebernimm(namen, os.path.basename(pfad))

    def aus_zwischenablage(self):
        text = Clipboard.GetText() if Clipboard.ContainsText() else u""
        if not text.strip():
            meldung(self.fenster, u"Die Zwischenablage enthält keinen Text.")
            return
        self.uebernimm(lg.namen_aus_text(text), u"Zwischenablage")

    def eingeben(self):
        f = dlg.lade_xaml(_XAML_NAMEN)
        dlg.setze_besitzer(f, besitzer=self.fenster)
        eingabe = f.FindName("eingabe")
        ergebnis = {"text": None}

        def bei_ok(_s, _a):
            ergebnis["text"] = eingabe.Text
            f.DialogResult = True

        f.FindName("ok").Click += sicher(lambda: f, bei_ok)
        f.Loaded += lambda _s, _a: eingabe.Focus()
        f.ShowDialog()
        if ergebnis["text"] is not None:
            self.uebernimm(lg.namen_aus_text(ergebnis["text"]), u"Eingabe")

    # --- Erstellen --------------------------------------------------------
    def _teamarbeit_sicherstellen(self):
        if self.doc.IsWorkshared:
            return True
        if not dlg.frage(
                self.fenster,
                u"Für dieses Projekt ist die Teamarbeit nicht aktiviert.\n\n"
                u"Teamarbeit jetzt aktivieren? Revit legt dabei die Worksets "
                u"'%s' und '%s' an. Danach muss das Projekt als "
                u"Zentralmodell gespeichert werden.\n\nDas lässt sich nicht "
                u"rückgängig machen." % (WS_EBENEN_RASTER, WS_STANDARD),
                titel=TITEL, warnung=True):
            return False
        self.doc.EnableWorksharing(WS_EBENEN_RASTER, WS_STANDARD)
        self.lade_vorhandene()
        return True

    def erstelle(self):
        markiert = [e for e in self.neu if e.markiert]
        if not markiert:
            meldung(self.fenster, u"In 'Neue Worksets' ist nichts markiert.")
            return
        if not self._teamarbeit_sicherstellen():
            return

        umbenennen = bool(self.modus_umbenennen.IsChecked)
        plan = lg.plane([e.name for e in markiert],
                        workset_namen(self.doc), umbenennen, revit_gueltig)
        erstellt, umbenannt, uebersprungen, fehler = [], [], [], []
        for name, ziel, status in plan:
            if status == lg.UEBERSPRUNGEN:
                uebersprungen.append(name)
            elif status == lg.UNGUELTIG:
                fehler.append(u"%s: ungültige Zeichen" % name)

        aufgaben = [(n, z, s) for n, z, s in plan if z]
        if aufgaben:
            sichtbar = bool(self.sichtbar.IsChecked)
            t = Transaction(self.doc, u"pyMLG Worksets erstellen")
            t.Start()
            try:
                einstellungen = WorksetDefaultVisibilitySettings \
                    .GetWorksetDefaultVisibilitySettings(self.doc)
                for name, ziel, status in aufgaben:
                    if not WorksetTable.IsWorksetNameUnique(self.doc, ziel):
                        fehler.append(u"%s: Name bereits vergeben" % ziel)
                        continue
                    try:
                        ws = Workset.Create(self.doc, ziel)
                        if not sichtbar:
                            einstellungen.SetWorksetVisibility(ws.Id, False)
                    except Exception as ausnahme:
                        fehler.append(u"%s: %s" % (
                            ziel, dlg.fehlertext(ausnahme)))
                        continue
                    erstellt.append(name)
                    if status == lg.UMBENANNT:
                        umbenannt.append(u"%s → %s" % (name, ziel))
                t.Commit()
            except Exception:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
                raise

        fertig = set(lg.schluessel(n) for n in erstellt)
        self.neu = [e for e in self.neu
                    if lg.schluessel(e.name) not in fertig]
        self.lade_vorhandene()
        self.zeichne_neu()
        self.status.Text = u"%d Worksets erstellt." % len(erstellt)
        meldung(self.fenster, self._bericht(erstellt, umbenannt,
                                            uebersprungen, fehler),
                warnung=bool(fehler))

    @staticmethod
    def _bericht(erstellt, umbenannt, uebersprungen, fehler, maximal=15):
        def liste(titel, namen):
            if not namen:
                return u""
            zeilen = [u"  " + n for n in namen[:maximal]]
            if len(namen) > maximal:
                zeilen.append(u"  ... und %d weitere" % (len(namen) - maximal))
            return u"\n\n%s (%d):\n%s" % (titel, len(namen),
                                          u"\n".join(zeilen))

        return (u"%d Worksets erstellt." % len(erstellt)
                + liste(u"Umbenannt", umbenannt)
                + liste(u"Übersprungen, existieren bereits", uebersprungen)
                + liste(u"Nicht erstellt", fehler))

    def zeige(self):
        self.fenster.ShowDialog()


def starte(uiapp, doc):
    WorksetCreatorFenster(uiapp, doc).zeige()
