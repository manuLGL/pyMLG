# -*- coding: utf-8 -*-
"""Hauptfenster von FilterMore (WPF), angelehnt an "FilterMore".

    ┌ Elemente ──────────────────────────────┐ Aufklappen Zuklappen ...
    │ ☑ Alle                              2  │ Bereich: ● Auswahl ○ Ansicht
    │   ☑ Geschossdecken                  1  │ Auswahl erweitern: □ gleiche …
    │     ☑ Geschossdecke                 1  │ Anzeigefilter: [____▾]
    │       ☑ DA_Dachaufbau_25cm          1  │   Nur behalten  Entfernen
    │         ☑ DA_Dachaufbau_25cm [123]  1  │ 2 Elemente markiert
    └────────────────────────────────────────┘ [Auswahl übernehmen]

Die Markierung ist eine Menge von Element-Ids (self.markiert). Kontrollkästchen
zeigen sie nur an: Klick auf einen Knoten markiert bzw. entfernt alle
Elemente darunter, danach werden alle sichtbaren Kästchen neu gesetzt
(teilweise markiert = Zwischenzustand).

Der Baum wird erst beim Aufklappen befüllt - bei "Alle Modellelemente"
entstehen sonst zehntausende Steuerelemente auf einmal.
"""

import io
import json
import os
import traceback

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import FontWeights, Thickness, \
    VerticalAlignment  # noqa: E402
from System.Windows.Controls import CheckBox, ComboBoxItem, Orientation, \
    StackPanel, TextBlock, TreeViewItem  # noqa: E402
from System.Windows.Media import Color, SolidColorBrush  # noqa: E402

from filter_manager import dialoge as dlg  # noqa: E402
from filter_more import baum as bm  # noqa: E402
from filter_more import revit as rv  # noqa: E402

TITEL = u"FilterMore"


def _pinsel(r, g, b):
    # Eingefroren, sonst gehört der Pinsel dem Thread, der das Modul geladen
    # hat (siehe filter_manager/fenster.py)
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


GRAU = _pinsel(120, 120, 120)
BLAU = _pinsel(40, 90, 160)

FEHLERPROTOKOLL = os.path.join(dlg.protokollordner(), "FilterMore_Fehler.log")
EINSTELLUNGEN = os.path.join(dlg.protokollordner(), "FilterMore.json")

BEREICH_TEXT = {
    rv.AUSWAHL: u"in der aktuellen Auswahl",
    rv.ANSICHT: u"in der aktuellen Ansicht",
    rv.MODELL: u"im ganzen Modell",
}

# Name des Kontrollkästchens -> Erweiterungs-Schlüssel
ERWEITERUNGEN = (
    ("was_kategorie", rv.GLEICHE_KATEGORIE),
    ("was_familie", rv.GLEICHE_FAMILIE),
    ("was_typ", rv.GLEICHER_TYP),
    ("was_workset", rv.GLEICHES_WORKSET),
    ("was_host", rv.HOST),
    ("was_gehostete", rv.GEHOSTETE),
    ("was_verschachtelte", rv.VERSCHACHTELTE),
    ("was_verbundene", rv.VERBUNDENE),
    ("was_uebergeordnete", rv.UEBERGEORDNETE),
    ("was_abhaengige", rv.ABHAENGIGE),
)

OPTIONEN = ("nur_3d", "wo_ansicht", "wie_neu", "ohne_gruppe",
            "ohne_baugruppe") + tuple(n for n, _k in ERWEITERUNGEN)

XAML = u"""
<Window %s Title="FilterMore (pyMLG)" Width="1120" Height="860"
        MinWidth="820" MinHeight="600" WindowStartupLocation="CenterOwner"
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
    <Style x:Key="klein" TargetType="Button">
      <Setter Property="Padding" Value="6,2"/>
      <Setter Property="Margin" Value="0,0,4,4"/>
    </Style>
  </Window.Resources>
  <Grid Margin="10">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="10"/>
      <ColumnDefinition Width="260"/>
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Baum -->
    <DockPanel>
      <WrapPanel DockPanel.Dock="Top">
        <Button x:Name="aufklappen" Content="Aufklappen" Style="{StaticResource klein}"
                ToolTip="Bis zur Typ-Ebene aufklappen"/>
        <Button x:Name="zuklappen" Content="Zuklappen" Style="{StaticResource klein}"/>
        <Button x:Name="alle" Content="Alle markieren" Style="{StaticResource klein}"/>
        <Button x:Name="keine" Content="Keine" Style="{StaticResource klein}"/>
        <Button x:Name="umkehren" Content="Umkehren" Style="{StaticResource klein}"
                ToolTip="Markierung umkehren"/>
      </WrapPanel>
      <TreeView x:Name="baum" Padding="2,4"/>
    </DockPanel>

    <!-- Optionen -->
    <ScrollViewer Grid.Column="2" VerticalScrollBarVisibility="Auto">
      <StackPanel>
        <GroupBox Header="Elemente">
          <StackPanel>
            <RadioButton x:Name="bereich_auswahl" Content="Aktuelle Auswahl"/>
            <RadioButton x:Name="bereich_ansicht"
                         Content="Elemente in aktueller Ansicht"/>
            <RadioButton x:Name="bereich_modell" Content="Alle Modellelemente"/>
            <CheckBox x:Name="nur_3d" Content="Nur 3D-modellierte Elemente"
                      ToolTip="Ohne Beschriftungen und ansichtsspezifische Elemente"/>
          </StackPanel>
        </GroupBox>

        <GroupBox Header="Markierte erweitern">
          <StackPanel>
            <TextBlock Text="Was" FontWeight="SemiBold" Margin="0,0,0,2"/>
            <CheckBox x:Name="was_kategorie" Content="Gleiche Kategorie"/>
            <CheckBox x:Name="was_familie" Content="Gleiche Familie"/>
            <CheckBox x:Name="was_typ" Content="Gleicher Typ"/>
            <CheckBox x:Name="was_workset" Content="Gleiches Workset"/>
            <CheckBox x:Name="was_host" Content="Host der Elemente"
                      ToolTip="z.B. die Wand einer Tür"/>
            <CheckBox x:Name="was_gehostete" Content="Gehostete Elemente"
                      ToolTip="z.B. Türen und Fenster einer Wand"/>
            <CheckBox x:Name="was_verschachtelte"
                      Content="Verschachtelte Elemente"
                      ToolTip="Gemeinsam genutzte Unterkomponenten von Familien"/>
            <CheckBox x:Name="was_verbundene" Content="Verbundene Elemente"
                      ToolTip="Über 'Geometrie verbinden'"/>
            <CheckBox x:Name="was_uebergeordnete"
                      Content="Übergeordnete Komponente"/>
            <CheckBox x:Name="was_abhaengige" Content="Abhängige Elemente"/>

            <TextBlock Text="Wo" FontWeight="SemiBold" Margin="0,8,0,2"/>
            <RadioButton x:Name="wo_modell" Content="Ganzes Modell"/>
            <RadioButton x:Name="wo_ansicht" Content="Aktuelle Ansicht"/>

            <TextBlock Text="Wie" FontWeight="SemiBold" Margin="0,8,0,2"/>
            <RadioButton x:Name="wie_hinzu" Content="Zur Liste hinzufügen"/>
            <RadioButton x:Name="wie_neu" Content="Neue Liste erstellen"
                         ToolTip="Nur die gefundenen Elemente anzeigen"/>

            <TextBlock Text="Nicht übernehmen, wenn" FontWeight="SemiBold"
                       Margin="0,8,0,2"/>
            <CheckBox x:Name="ohne_gruppe" Content="Teil einer Gruppe"/>
            <CheckBox x:Name="ohne_baugruppe" Content="Teil einer Baugruppe"/>

            <Button x:Name="erweitern" Content="Auswahl erweitern"
                    Padding="8,5" Margin="0,8,0,2"/>
          </StackPanel>
        </GroupBox>

        <GroupBox Header="Anzeigefilter auf markierte anwenden">
          <StackPanel>
            <ComboBox x:Name="filterliste" Margin="0,0,0,6"
                      IsTextSearchEnabled="True"/>
            <UniformGrid Columns="2">
              <Button x:Name="nur_behalten" Content="Nur behalten"
                      Padding="6,4" Margin="0,0,3,0"
                      ToolTip="Markierung nur bei Elementen lassen, die der Filter erfasst"/>
              <Button x:Name="entfernen" Content="Entfernen" Padding="6,4"
                      Margin="3,0,0,0"
                      ToolTip="Markierung bei Elementen entfernen, die der Filter erfasst"/>
            </UniformGrid>
          </StackPanel>
        </GroupBox>

        <TextBlock x:Name="anzahl" FontSize="22" FontWeight="SemiBold"
                   Margin="0,6,0,0"/>
        <TextBlock x:Name="anzahl_text" Foreground="#555" TextWrapping="Wrap"
                   Margin="0,0,0,8"/>
        <Button x:Name="uebernehmen" Content="Auswahl übernehmen"
                FontWeight="Bold" Padding="8,9"
                ToolTip="Markierte Elemente in Revit auswählen und schließen"/>
      </StackPanel>
    </ScrollViewer>

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


# ---------------------------------------------------------------------------
# Fenster
# ---------------------------------------------------------------------------

class FilterMoreFenster(object):

    def __init__(self, uiapp, uidoc):
        self.uiapp = uiapp
        self.uidoc = uidoc
        self.doc = uidoc.Document
        self.elemente = {}          # Id -> Element
        self.markiert = set()       # Ids
        self.wurzel = None
        self.ui_knoten = []         # Knoten mit Steuerelementen
        self.eigene_liste = False   # nach "Auswahl erweitern"
        self.filter = []            # [(Name, ParameterFilterElement)]

        f = self.fenster = dlg.lade_xaml(XAML)
        try:
            dlg.setze_besitzer(f, handle=uiapp.MainWindowHandle)
        except Exception:
            pass
        for name in ("baum", "bereich_auswahl", "bereich_ansicht",
                     "bereich_modell", "wo_modell", "wie_hinzu",
                     "filterliste", "anzahl", "anzahl_text",
                     "status") + OPTIONEN:
            setattr(self, name, f.FindName(name))

        einst = lade_einstellungen()
        for name in OPTIONEN:
            getattr(self, name).IsChecked = bool(einst.get(name, False))
        self.wo_modell.IsChecked = not self.wo_ansicht.IsChecked
        self.wie_hinzu.IsChecked = not self.wie_neu.IsChecked

        hat_auswahl = self.uidoc.Selection.GetElementIds().Count > 0
        if hat_auswahl:
            self.bereich_auswahl.IsChecked = True
        else:
            self.bereich_auswahl.IsEnabled = False
            self.bereich_auswahl.Content = u"Aktuelle Auswahl (keine)"
            if einst.get("bereich") == rv.MODELL:
                self.bereich_modell.IsChecked = True
            else:
                self.bereich_ansicht.IsChecked = True

        self._fuelle_filter()
        self._verdrahte()
        self.lade()

    # --- Verdrahtung ------------------------------------------------------
    def _verdrahte(self):
        f = self.fenster

        def s(funktion):
            return sicher(lambda: f, funktion)

        def klick(name, funktion):
            f.FindName(name).Click += s(lambda _s, _a: funktion())

        klick("aufklappen", lambda: self.klappe(True))
        klick("zuklappen", lambda: self.klappe(False))
        klick("alle", lambda: self.setze_markierung(set(self.elemente)))
        klick("keine", lambda: self.setze_markierung(set()))
        klick("umkehren", lambda: self.setze_markierung(
            set(self.elemente) - self.markiert))
        klick("erweitern", self.erweitere)
        klick("nur_behalten", lambda: self.wende_filter_an(True))
        klick("entfernen", lambda: self.wende_filter_an(False))
        klick("uebernehmen", self.uebernehme)
        klick("schliessen", f.Close)
        for box in (self.bereich_auswahl, self.bereich_ansicht,
                    self.bereich_modell):
            box.Checked += s(lambda _s, _a: self.lade())
        self.nur_3d.Checked += s(lambda _s, _a: self.lade())
        self.nur_3d.Unchecked += s(lambda _s, _a: self.lade())
        f.Closing += s(lambda _s, _a: self._speichere())

    def _speichere(self):
        werte = dict((n, bool(getattr(self, n).IsChecked)) for n in OPTIONEN)
        werte["bereich"] = self.bereich()
        speichere_einstellungen(werte)

    def bereich(self):
        if self.bereich_auswahl.IsChecked:
            return rv.AUSWAHL
        if self.bereich_modell.IsChecked:
            return rv.MODELL
        return rv.ANSICHT

    # --- Laden und Baum ---------------------------------------------------
    def lade(self):
        elemente = rv.sammle(self.doc, self.uidoc, self.bereich(),
                             bool(self.nur_3d.IsChecked))
        self.elemente = dict((rv.id_wert(e.Id), e) for e in elemente)
        self.markiert = set(self.elemente)
        self.eigene_liste = False
        self.baue_baum()
        self.status.Text = u"%d Elemente geladen." % len(self.elemente)

    def baue_baum(self):
        self.wurzel = rv.baue_baum(self.doc, self.elemente.values())
        self.baum.Items.Clear()
        self.ui_knoten = []
        item = self._item(self.wurzel)
        self.baum.Items.Add(item)
        item.IsExpanded = True
        self.aktualisiere()

    def _item(self, knoten):
        item = TreeViewItem()
        kopf = StackPanel()
        kopf.Orientation = Orientation.Horizontal

        box = CheckBox()
        box.VerticalAlignment = VerticalAlignment.Center
        box.Margin = Thickness(0.0, 1.0, 5.0, 1.0)
        box.Click += sicher(lambda: self.fenster, self._klick(knoten))
        kopf.Children.Add(box)

        name = TextBlock()
        name.Text = knoten.name
        name.VerticalAlignment = VerticalAlignment.Center
        if knoten.ebene == bm.ELEMENT:
            name.Foreground = GRAU
        elif knoten.ebene in (bm.WURZEL, bm.KATEGORIE):
            name.FontWeight = FontWeights.SemiBold
        kopf.Children.Add(name)

        zaehler = TextBlock()
        zaehler.Margin = Thickness(10.0, 0.0, 0.0, 0.0)
        zaehler.Foreground = BLAU
        zaehler.VerticalAlignment = VerticalAlignment.Center
        if knoten.ebene != bm.ELEMENT:
            kopf.Children.Add(zaehler)

        item.Header = kopf
        knoten.item, knoten.box, knoten.zaehler = item, box, zaehler
        knoten.geladen = False
        if knoten.kinder:
            item.Items.Add(u"…")     # Platzhalter, damit der Pfeil erscheint
            item.Expanded += sicher(lambda: self.fenster,
                                    self._aufklappen(knoten))
        self.ui_knoten.append(knoten)
        self._zeige_zustand(knoten)
        return item

    def _aufklappen(self, knoten):
        def handler(_sender, args):
            # Das Ereignis steigt von Unterknoten auf - nur das eigene zählt
            if not args.OriginalSource.Equals(knoten.item) or knoten.geladen:
                return
            knoten.geladen = True
            knoten.item.Items.Clear()
            for kind in knoten.kinder:
                knoten.item.Items.Add(self._item(kind))
        return handler

    def _klick(self, knoten):
        def handler(sender, _args):
            if sender.IsChecked:
                self.markiert.update(knoten.ids)
            else:
                self.markiert.difference_update(knoten.ids)
            self.aktualisiere()
        return handler

    def _zeige_zustand(self, knoten):
        zustand, text = bm.anzeige(knoten, self.markiert)
        knoten.box.IsChecked = zustand
        knoten.zaehler.Text = text

    def aktualisiere(self):
        for knoten in self.ui_knoten:
            self._zeige_zustand(knoten)
        self.anzahl.Text = u"%d" % len(self.markiert)
        herkunft = (u"in der erweiterten Liste" if self.eigene_liste
                    else BEREICH_TEXT[self.bereich()])
        self.anzahl_text.Text = u"von %d Elementen markiert - %s" % (
            len(self.elemente), herkunft)

    def setze_markierung(self, ids):
        self.markiert = set(ids)
        self.aktualisiere()

    def klappe(self, auf):
        """Aufklappen bis zur Typ-Ebene bzw. alles unter 'Alle' zuklappen."""
        if self.wurzel is None:
            return

        def auf_bis_typ(knoten):
            if knoten.item is None or knoten.ebene >= bm.TYP:
                return
            knoten.item.IsExpanded = True     # lädt die Kinder
            for kind in knoten.kinder:
                auf_bis_typ(kind)

        if auf:
            auf_bis_typ(self.wurzel)
        else:
            for knoten in self.ui_knoten:
                if knoten.ebene != bm.WURZEL and knoten.item is not None:
                    knoten.item.IsExpanded = False

    # --- Erweitern --------------------------------------------------------
    def erweitere(self):
        was = set(k for n, k in ERWEITERUNGEN if getattr(self, n).IsChecked)
        if not was:
            meldung(self.fenster, u"Bitte unter 'Was' mindestens eine "
                    u"Erweiterung ankreuzen.")
            return
        quellen = [self.elemente[i] for i in self.markiert]
        if not quellen:
            meldung(self.fenster, u"Es ist kein Element markiert.")
            return
        gefunden = rv.erweitere(
            self.doc, self.uidoc, quellen, was,
            nur_ansicht=bool(self.wo_ansicht.IsChecked),
            nur_3d=bool(self.nur_3d.IsChecked),
            ohne_gruppe=bool(self.ohne_gruppe.IsChecked),
            ohne_baugruppe=bool(self.ohne_baugruppe.IsChecked))
        neue = dict((rv.id_wert(e.Id), e) for e in gefunden)
        if not neue:
            # Eine leere "neue Liste" würde alles verwerfen
            self.status.Text = u"Keine weiteren Elemente gefunden."
            meldung(self.fenster, u"Keine weiteren Elemente gefunden.")
            return
        if self.wie_neu.IsChecked:
            self.elemente = neue
            self.markiert = set(neue)
            text = u"Neue Liste mit %d gefundenen Elementen." % len(neue)
        else:
            hinzu = set(neue) - set(self.elemente)
            self.elemente.update(neue)
            self.markiert.update(neue)
            text = u"%d Elemente gefunden, %d neu in der Liste." % (
                len(neue), len(hinzu))
        self.eigene_liste = True
        self.baue_baum()
        self.status.Text = text

    # --- Anzeigefilter ----------------------------------------------------
    def _fuelle_filter(self):
        self.filter = rv.anzeigefilter(self.doc)
        self.filterliste.Items.Clear()
        for name, _element in self.filter:
            item = ComboBoxItem()
            item.Content = name
            self.filterliste.Items.Add(item)
        if self.filter:
            self.filterliste.SelectedIndex = 0

    def wende_filter_an(self, behalten):
        index = self.filterliste.SelectedIndex
        if not 0 <= index < len(self.filter):
            meldung(self.fenster, u"Bitte einen Anzeigefilter wählen.")
            return
        name, filter_element = self.filter[index]
        markierte = [self.elemente[i] for i in self.markiert]
        treffer = rv.filter_treffer(filter_element, markierte)
        vorher = len(self.markiert)
        if behalten:
            self.markiert &= treffer
        else:
            self.markiert -= treffer
        self.aktualisiere()
        self.status.Text = u"Filter '%s': %d von %d markierten erfasst, " \
                           u"noch %d markiert." % (name, len(treffer), vorher,
                                                   len(self.markiert))

    # --- Übernehmen -------------------------------------------------------
    def uebernehme(self):
        if not self.markiert:
            meldung(self.fenster, u"Es ist kein Element markiert.")
            return
        rv.waehle_aus(self.uidoc, [self.elemente[i] for i in self.markiert])
        self.fenster.Close()

    def zeige(self):
        self.fenster.ShowDialog()


def starte(uiapp, uidoc):
    FilterMoreFenster(uiapp, uidoc).zeige()
