# -*- coding: utf-8 -*-
"""Hauptfenster von TransferSingle (WPF), angelehnt an "TransferSingle".

    Von: [Projekt ▾]                          ☑ Verknüpfungen als Quelle
                                              Transformation ○ ○ ○
    ┌ Was ───────────────────────────────┐ Aufklappen Zuklappen Alle Keine
    │ ☐ Alle                        2808 │ [Suche______] [Weitersuchen]
    │   ☐ Wände – Typen               12 │ Markierte bearbeiten (Quelle)
    │     ☐ Basiswand: AW 36.5           │  Löschen, Suchen/Ersetzen, ...
    └────────────────────────────────────┘ Markiert: 0
    ┌ Nach ──────────────────────────────┐ Bei Duplikaten ○ ○ ○
    │ ☐ Projekt 2                        │ ☐ Pläne mit Ansichten ...
    └────────────────────────────────────┘ [Übertragen] [Protokoll]

Der Baum wird erst beim Aufklappen befüllt (wie FilterMore). Die Markierung
ist eine Menge von Element-Ids der aktuellen Quelle.
"""

import io
import json
import os
import traceback

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Clipboard, FontWeights, Thickness, \
    VerticalAlignment  # noqa: E402
from System.Windows.Controls import CheckBox, ComboBoxItem, ListBoxItem, \
    Orientation, StackPanel, TextBlock, TreeViewItem  # noqa: E402
from System.Windows.Media import Color, SolidColorBrush  # noqa: E402

from filter_manager import dialoge as dlg  # noqa: E402
from filter_more import baum as bm  # noqa: E402
from mlg_sprache import t, uebersetze_xaml  # noqa: E402
from transfer_single import logik as lg  # noqa: E402
from transfer_single import revit as rv  # noqa: E402

TITEL = u"TransferSingle"


def _pinsel(r, g, b):
    # Eingefroren, sonst gehört der Pinsel dem Thread, der das Modul geladen
    # hat (siehe filter_manager/fenster.py)
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


GRAU = _pinsel(120, 120, 120)
BLAU = _pinsel(40, 90, 160)
ROT = _pinsel(192, 57, 43)

ORDNER = dlg.protokollordner()
FEHLERPROTOKOLL = os.path.join(ORDNER, "TransferSingle_Fehler.log")
EINSTELLUNGEN = os.path.join(ORDNER, "TransferSingle.json")
PROTOKOLL = os.path.join(ORDNER, "TransferSingle_Protokoll.txt")

OPTIONEN = ("mit_links", "mit_modell", "trans_keine", "trans_link",
            "trans_gemeinsam", "dup_ziel", "dup_abbrechen", "dup_fragen",
            "dialoge", "plaene_ansichten", "ansichtselemente")
VORGABEN = {"trans_keine": True, "dup_ziel": True, "dialoge": True,
            "plaene_ansichten": True, "ansichtselemente": True}

XAML_TEXTE = {
    "titel": (u"TransferSingle (pyMLG)", u"TransferSingle (pyMLG)",
              u"TransferSingle (pyMLG)"),
    "von": (u"Von:", u"From:", u"Desde:"),
    "links": (u"Verknüpfungen als Quelle", u"Include links as source",
              u"Incluir vínculos como origen"),
    "trans": (u"Transformation (Modellelemente):",
              u"Transform model elements by:",
              u"Transformar elementos de modelo por:"),
    "keine": (u"Keine", u"None", u"Ninguna"),
    "link": (u"Verknüpfung", u"Link", u"Vínculo"),
    "gemeinsam": (u"Gemeinsame Koordinaten", u"Shared coordinates",
                  u"Coordenadas compartidas"),
    "was": (u"Was:", u"What:", u"Qué:"),
    "modell": (u"Modellelemente anzeigen", u"Show model elements",
               u"Mostrar elementos de modelo"),
    "modell_tipp": (u"Wände, Decken, Familien ... - bei großen Modellen langsam",
                    u"Walls, floors, families ... - slow in large models",
                    u"Muros, suelos, familias ... - lento en modelos grandes"),
    "auf": (u"Aufklappen", u"Expand", u"Expandir"),
    "zu": (u"Zuklappen", u"Collapse", u"Contraer"),
    "alle": (u"Alle", u"All", u"Todo"),
    "keine_m": (u"Keine", u"None", u"Ninguno"),
    "weiter": (u"Weitersuchen", u"Search next", u"Buscar siguiente"),
    "verwalten": (u"Markierte bearbeiten (in der Quelle)",
                  u"Manage checked (in the source)",
                  u"Gestionar marcados (en el origen)"),
    "loeschen": (u"Löschen", u"Delete", u"Eliminar"),
    "ersetzen": (u"Suchen und Ersetzen", u"Find and replace",
                 u"Buscar y reemplazar"),
    "praefix": (u"Präfix hinzufügen", u"Add prefix", u"Añadir prefijo"),
    "suffix": (u"Suffix hinzufügen", u"Add suffix", u"Añadir sufijo"),
    "gross": (u"GROSSBUCHSTABEN", u"UPPER CASE", u"MAYÚSCULAS"),
    "klein": (u"kleinbuchstaben", u"lower case", u"minúsculas"),
    "erster": (u"Erster Buchstabe Groß", u"Proper Case",
               u"Primera Letra Mayúscula"),
    "zahlen": (u"Zahlen ersetzen", u"Find and replace numbers",
               u"Reemplazar números"),
    "markiert": (u"Markierte Elemente:", u"Elements checked:",
                 u"Elementos marcados:"),
    "ids": (u"IDs kopieren", u"Copy ids", u"Copiar ids"),
    "nach": (u"Nach:", u"To:", u"A:"),
    "duplikate": (u"Bei doppelten Typnamen", u"On duplicate type names",
                  u"Con nombres de tipo duplicados"),
    "dup_ziel": (u"Zieltypen verwenden", u"Use destination types",
                 u"Usar tipos del destino"),
    "dup_abbrechen": (u"Abbrechen", u"Abort", u"Cancelar"),
    "dup_fragen": (u"Nachfragen", u"Ask user", u"Preguntar"),
    "dialoge": (u"Revit-Warnungen automatisch bestätigen",
                u"Accept Revit warnings automatically",
                u"Aceptar avisos de Revit automáticamente"),
    "plaene": (u"Pläne mit Ansichten übertragen",
               u"Transfer sheets with views", u"Transferir planos con vistas"),
    "plaene_tipp": (u"Zeichenansichten, Legenden und Bauteillisten werden mit übertragen und platziert. Modellansichten lassen sich nicht übertragen.",
                    u"Drafting views, legends and schedules are transferred and placed as well. Model views cannot be transferred.",
                    u"Las vistas de diseño, leyendas y tablas se transfieren y colocan. Las vistas de modelo no se pueden transferir."),
    "elemente": (u"Ansichtselemente mit übertragen",
                 u"Transfer view elements", u"Transferir elementos de vista"),
    "uebertragen": (u"Übertragen", u"Transfer", u"Transferir"),
    "protokoll": (u"Protokoll", u"View log", u"Ver registro"),
    "schliessen": (u"Schließen", u"Close", u"Cerrar"),
}

XAML = u"""
<Window %s Title="{{titel}}" Width="1180" Height="900"
        MinWidth="900" MinHeight="650" WindowStartupLocation="CenterOwner"
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
      <Setter Property="Margin" Value="0,2,10,2"/>
    </Style>
    <Style x:Key="aktion" TargetType="Button">
      <Setter Property="Padding" Value="6,3"/>
      <Setter Property="Margin" Value="0,0,0,4"/>
    </Style>
  </Window.Resources>
  <Grid Margin="10">
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="*"/>
      <ColumnDefinition Width="10"/>
      <ColumnDefinition Width="270"/>
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="3*"/>
      <RowDefinition Height="10"/>
      <RowDefinition Height="1*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- Quelle -->
    <StackPanel Grid.Row="0" Grid.Column="0" Margin="0,0,0,8">
      <TextBlock Text="{{von}}" Margin="0,0,0,3"/>
      <ComboBox x:Name="quellen"/>
    </StackPanel>
    <StackPanel Grid.Row="0" Grid.Column="2" Margin="0,0,0,8">
      <CheckBox x:Name="mit_links" Content="{{links}}"/>
      <TextBlock Text="{{trans}}" Margin="0,4,0,0"/>
      <WrapPanel>
        <RadioButton x:Name="trans_keine" GroupName="trans" Content="{{keine}}"/>
        <RadioButton x:Name="trans_link" GroupName="trans" Content="{{link}}"/>
        <RadioButton x:Name="trans_gemeinsam" GroupName="trans"
                     Content="{{gemeinsam}}"/>
      </WrapPanel>
    </StackPanel>

    <!-- Was -->
    <DockPanel Grid.Row="1" Grid.Column="0">
      <DockPanel DockPanel.Dock="Top" Margin="0,0,0,3">
        <CheckBox x:Name="mit_modell" Content="{{modell}}"
                  ToolTip="{{modell_tipp}}" DockPanel.Dock="Right"/>
        <TextBlock Text="{{was}}" VerticalAlignment="Center"/>
      </DockPanel>
      <TreeView x:Name="baum" Padding="2,4"/>
    </DockPanel>

    <DockPanel Grid.Row="1" Grid.Column="2" LastChildFill="False">
      <StackPanel DockPanel.Dock="Top">
        <WrapPanel Margin="0,0,0,6">
          <Button x:Name="aufklappen" Content="{{auf}}" Padding="6,2" Margin="0,0,4,4"/>
          <Button x:Name="zuklappen" Content="{{zu}}" Padding="6,2" Margin="0,0,4,4"/>
          <Button x:Name="alle" Content="{{alle}}" Padding="6,2" Margin="0,0,4,4"/>
          <Button x:Name="keine" Content="{{keine_m}}" Padding="6,2" Margin="0,0,4,4"/>
        </WrapPanel>
        <TextBox x:Name="suche" Padding="3" Margin="0,0,0,4"/>
        <Button x:Name="weitersuchen" Content="{{weiter}}" Style="{StaticResource aktion}"
                Margin="0,0,0,10"/>
        <GroupBox Header="{{verwalten}}">
          <StackPanel>
            <Button x:Name="loeschen" Content="{{loeschen}}" Style="{StaticResource aktion}"/>
            <Button x:Name="ersetzen" Content="{{ersetzen}}" Style="{StaticResource aktion}"/>
            <Button x:Name="praefix" Content="{{praefix}}" Style="{StaticResource aktion}"/>
            <Button x:Name="suffix" Content="{{suffix}}" Style="{StaticResource aktion}"/>
            <Button x:Name="gross" Content="{{gross}}" Style="{StaticResource aktion}"/>
            <Button x:Name="klein" Content="{{klein}}" Style="{StaticResource aktion}"/>
            <Button x:Name="erster" Content="{{erster}}" Style="{StaticResource aktion}"/>
            <Button x:Name="zahlen" Content="{{zahlen}}" Style="{StaticResource aktion}"/>
          </StackPanel>
        </GroupBox>
      </StackPanel>
      <StackPanel DockPanel.Dock="Bottom">
        <TextBlock Text="{{markiert}}" Foreground="#555"/>
        <DockPanel>
          <Button x:Name="ids" Content="{{ids}}" Padding="6,2"
                  DockPanel.Dock="Right" VerticalAlignment="Center"/>
          <TextBlock x:Name="anzahl" FontSize="22" FontWeight="SemiBold"/>
        </DockPanel>
      </StackPanel>
    </DockPanel>

    <!-- Nach -->
    <DockPanel Grid.Row="3" Grid.Column="0">
      <TextBlock DockPanel.Dock="Top" Text="{{nach}}" Margin="0,0,0,3"/>
      <ListBox x:Name="ziele"/>
    </DockPanel>
    <StackPanel Grid.Row="3" Grid.Column="2">
      <GroupBox Header="{{duplikate}}">
        <StackPanel>
          <RadioButton x:Name="dup_ziel" GroupName="dup" Content="{{dup_ziel}}"/>
          <RadioButton x:Name="dup_abbrechen" GroupName="dup" Content="{{dup_abbrechen}}"/>
          <RadioButton x:Name="dup_fragen" GroupName="dup" Content="{{dup_fragen}}"/>
          <CheckBox x:Name="dialoge" Margin="0,6,0,0">
            <TextBlock Text="{{dialoge}}" TextWrapping="Wrap"/>
          </CheckBox>
        </StackPanel>
      </GroupBox>
      <CheckBox x:Name="plaene_ansichten" ToolTip="{{plaene_tipp}}">
        <TextBlock Text="{{plaene}}" TextWrapping="Wrap"/>
      </CheckBox>
      <CheckBox x:Name="ansichtselemente" Content="{{elemente}}"/>
    </StackPanel>

    <!-- Fußzeile -->
    <DockPanel Grid.Row="4" Grid.ColumnSpan="3" Margin="0,10,0,0">
      <Button x:Name="schliessen" Content="{{schliessen}}" Width="100"
              DockPanel.Dock="Right" IsCancel="True"/>
      <Button x:Name="protokoll" Content="{{protokoll}}" Width="100"
              DockPanel.Dock="Right" Margin="0,0,8,0"/>
      <Button x:Name="uebertragen" Content="{{uebertragen}}" Width="160"
              DockPanel.Dock="Right" Margin="0,0,8,0" FontWeight="Bold"/>
      <TextBlock x:Name="status" VerticalAlignment="Center" Foreground="#555"
                 TextTrimming="CharacterEllipsis"/>
    </DockPanel>
  </Grid>
</Window>""" % dlg.XMLNS

ZWEI_FELDER_TEXTE = {
    "ok": (u"OK", u"OK", u"Aceptar"),
    "abbrechen": (u"Abbrechen", u"Cancel", u"Cancelar"),
}

XAML_ZWEI_FELDER = u"""
<Window %s Title="{titel}" Width="420" SizeToContent="Height"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12">
  <StackPanel Margin="14">
    <TextBlock x:Name="text1" Margin="0,0,0,3"/>
    <TextBox x:Name="feld1" Padding="3" Margin="0,0,0,8"/>
    <TextBlock x:Name="text2" Margin="0,0,0,3"/>
    <TextBox x:Name="feld2" Padding="3"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right"
                Margin="0,14,0,0">
      <Button x:Name="ok" Content="{{ok}}" Width="90" Margin="0,0,8,0"
              IsDefault="True"/>
      <Button Content="{{abbrechen}}" Width="90" IsCancel="True"/>
    </StackPanel>
  </StackPanel>
</Window>""" % dlg.XMLNS

PROTOKOLL_TEXTE = {
    "schliessen": (u"Schließen", u"Close", u"Cerrar"),
}

XAML_PROTOKOLL = u"""
<Window %s Title="{titel}" Width="900" Height="600"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        FontFamily="Segoe UI" FontSize="12">
  <DockPanel Margin="10">
    <DockPanel DockPanel.Dock="Bottom" Margin="0,8,0,0">
      <Button Content="{{schliessen}}" Width="100" IsCancel="True"
              DockPanel.Dock="Right"/>
      <TextBlock x:Name="pfad" Foreground="#555" VerticalAlignment="Center"/>
    </DockPanel>
    <TextBox x:Name="text" IsReadOnly="True" FontFamily="Consolas"
             VerticalScrollBarVisibility="Auto"
             HorizontalScrollBarVisibility="Auto"/>
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
                    t(u"\n\nTechnische Details: %s", u"\n\nTechnical details: %s",
                      u"\n\nDetalles técnicos: %s") % pfad if pfad else u""),
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

class TransferSingleFenster(object):

    def __init__(self, uiapp):
        self.uiapp = uiapp
        self.app = uiapp.Application
        self.quellen = []
        self.quelle = None
        self.ziel_boxen = []        # [(CheckBox, Document)]
        self.wurzel = None
        self.ui_knoten = []
        self.markiert = set()
        self.letzter_treffer = None
        self.protokoll_text = u""

        f = self.fenster = dlg.lade_xaml(uebersetze_xaml(XAML, XAML_TEXTE))
        try:
            dlg.setze_besitzer(f, handle=uiapp.MainWindowHandle)
        except Exception:
            pass
        for name in ("quellen", "baum", "suche", "anzahl", "ziele",
                     "status") + OPTIONEN:
            setattr(self, name, f.FindName(name))

        einst = lade_einstellungen()
        for name in OPTIONEN:
            getattr(self, name).IsChecked = bool(
                einst.get(name, VORGABEN.get(name, False)))
        self._fuelle_quellen(einst.get("quelle"))     # lädt auch den Baum
        self._verdrahte()

    # --- Verdrahtung ------------------------------------------------------
    def _verdrahte(self):
        f = self.fenster

        def s(funktion):
            return sicher(lambda: f, funktion)

        def klick(name, funktion):
            f.FindName(name).Click += s(lambda _s, _a: funktion())

        klick("aufklappen", lambda: self.klappe(True))
        klick("zuklappen", lambda: self.klappe(False))
        klick("alle", lambda: self.setze_markierung(set(
            self.wurzel.ids if self.wurzel else [])))
        klick("keine", lambda: self.setze_markierung(set()))
        klick("weitersuchen", self.weitersuchen)
        klick("loeschen", self.loesche)
        klick("ersetzen", lambda: self.benenne_um(lg.ERSETZEN))
        klick("praefix", lambda: self.benenne_um(lg.PRAEFIX))
        klick("suffix", lambda: self.benenne_um(lg.SUFFIX))
        klick("gross", lambda: self.benenne_um(lg.GROSS))
        klick("klein", lambda: self.benenne_um(lg.KLEIN))
        klick("erster", lambda: self.benenne_um(lg.ERSTER_GROSS))
        klick("zahlen", lambda: self.benenne_um(lg.ZAHLEN))
        klick("ids", self.kopiere_ids)
        klick("uebertragen", self.uebertrage)
        klick("protokoll", self.zeige_protokoll)
        klick("schliessen", f.Close)

        self.quellen.SelectionChanged += s(lambda _s, _a: self.lade())
        self.mit_links.Checked += s(lambda _s, _a: self._fuelle_quellen())
        self.mit_links.Unchecked += s(lambda _s, _a: self._fuelle_quellen())
        self.mit_modell.Checked += s(lambda _s, _a: self.lade())
        self.mit_modell.Unchecked += s(lambda _s, _a: self.lade())

        def bei_taste(_s, args):
            from System.Windows.Input import Key
            if args.Key == Key.Enter:
                self.weitersuchen()
        self.suche.KeyDown += s(bei_taste)
        f.Closing += s(lambda _s, _a: self._speichere())

    def _speichere(self):
        werte = dict((n, bool(getattr(self, n).IsChecked)) for n in OPTIONEN)
        werte["quelle"] = self.quelle.titel if self.quelle else u""
        speichere_einstellungen(werte)

    # --- Quellen und Ziele ------------------------------------------------
    def _fuelle_quellen(self, vorauswahl=None):
        vorher = vorauswahl or (self.quelle.titel if self.quelle else None)
        aktiv = self.uiapp.ActiveUIDocument.Document \
            if self.uiapp.ActiveUIDocument else None
        self.quellen_liste = rv.quellen(self.app,
                                        bool(self.mit_links.IsChecked))
        self._still = True
        try:
            self.quellen.Items.Clear()
            index = 0
            for i, quelle in enumerate(self.quellen_liste):
                item = ComboBoxItem()
                item.Content = quelle.titel
                self.quellen.Items.Add(item)
                if quelle.titel == vorher:
                    index = i
                elif vorher is None and aktiv is not None \
                        and quelle.doc.Equals(aktiv):
                    index = i
            self.quellen.SelectedIndex = index if self.quellen_liste else -1
        finally:
            self._still = False
        self.lade()

    def _fuelle_ziele(self):
        gewaehlt = set(d.Title for b, d in self.ziel_boxen if b.IsChecked)
        self.ziele.Items.Clear()
        self.ziel_boxen = []
        if self.quelle is None:
            return
        for doc in rv.ziele(self.app, self.quelle):
            box = CheckBox()
            box.Content = doc.Title
            box.IsChecked = doc.Title in gewaehlt
            box.Margin = Thickness(2.0, 2.0, 2.0, 2.0)
            item = ListBoxItem()
            item.Content = box
            self.ziele.Items.Add(item)
            self.ziel_boxen.append((box, doc))
        if not self.ziel_boxen:
            item = ListBoxItem()
            item.Content = t(u"(kein weiteres Projekt geöffnet)",
                             u"(no other project open)",
                             u"(no hay otro proyecto abierto)")
            item.Foreground = GRAU
            item.IsEnabled = False
            self.ziele.Items.Add(item)

    # --- Baum -------------------------------------------------------------
    def lade(self):
        if getattr(self, "_still", False):
            return
        index = self.quellen.SelectedIndex
        self.quelle = (self.quellen_liste[index]
                       if 0 <= index < len(self.quellen_liste) else None)
        self._fuelle_ziele()
        for name in ("loeschen", "ersetzen", "praefix", "suffix", "gross",
                     "klein", "erster", "zahlen"):
            # Verknüpfte Dokumente sind schreibgeschützt
            self.fenster.FindName(name).IsEnabled = bool(
                self.quelle and not self.quelle.ist_verknuepfung)
        if self.quelle is None:
            self.wurzel = None
            self.baum.Items.Clear()
            self.aktualisiere()
            return
        daten = rv.datensaetze(self.quelle.doc,
                               bool(self.mit_modell.IsChecked))
        vorhanden = set(w for _g, _b, w in daten)
        self.markiert &= vorhanden
        self.wurzel = lg.baue(daten, t(u"Alle", u"All", u"Todo"))
        self.baum.Items.Clear()
        self.ui_knoten = []
        item = self._item(self.wurzel)
        self.baum.Items.Add(item)
        item.IsExpanded = True
        self.letzter_treffer = None
        self.aktualisiere()
        self.status.Text = t(u"%d Elemente in '%s'", u"%d elements in '%s'",
                             u"%d elementos en '%s'") % (
            len(vorhanden), self.quelle.titel)

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
        if knoten.ebene == lg.ELEMENT:
            name.Foreground = GRAU
        else:
            name.FontWeight = FontWeights.SemiBold
        kopf.Children.Add(name)
        zaehler = TextBlock()
        zaehler.Margin = Thickness(10.0, 0.0, 0.0, 0.0)
        zaehler.Foreground = BLAU
        zaehler.VerticalAlignment = VerticalAlignment.Center
        if knoten.ebene != lg.ELEMENT:
            kopf.Children.Add(zaehler)
        item.Header = kopf
        knoten.item, knoten.box, knoten.zaehler = item, box, zaehler
        knoten.geladen = False
        if knoten.kinder:
            item.Items.Add(u"…")
            item.Expanded += sicher(lambda: self.fenster,
                                    self._aufklappen(knoten))
        self.ui_knoten.append(knoten)
        self._zeige_zustand(knoten)
        return item

    def _aufklappen(self, knoten):
        def handler(_sender, args):
            if not args.OriginalSource.Equals(knoten.item):
                return
            self._lade_kinder(knoten)
        return handler

    def _lade_kinder(self, knoten):
        if knoten.geladen:
            return
        knoten.geladen = True
        knoten.item.Items.Clear()
        for kind in knoten.kinder:
            knoten.item.Items.Add(self._item(kind))

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

    def setze_markierung(self, ids):
        self.markiert = set(ids)
        self.aktualisiere()

    MAX_AUFKLAPPEN = 3000

    def klappe(self, auf):
        """Aufklappen: alle Gruppen öffnen - bei sehr vielen Elementen nur
        die Gruppenliste, sonst entstehen zehntausende Einträge auf einmal."""
        if self.wurzel is None:
            return
        self.wurzel.item.IsExpanded = True
        viele = len(self.wurzel.ids) > self.MAX_AUFKLAPPEN
        for gruppe in self.wurzel.kinder:
            if gruppe.item is not None:
                gruppe.item.IsExpanded = bool(auf and not viele)
        if auf and viele:
            self.status.Text = t(
                u"Zu viele Elemente zum Aufklappen - Gruppen einzeln öffnen.",
                u"Too many elements to expand - open groups individually.",
                u"Demasiados elementos para expandir: abra los grupos por "
                u"separado.")

    def weitersuchen(self):
        if self.wurzel is None:
            return
        treffer = lg.naechster_treffer(self.wurzel, self.suche.Text,
                                       self.letzter_treffer)
        if treffer is None:
            self.status.Text = t(u"Nichts gefunden.", u"Nothing found.",
                                 u"No se encontró nada.")
            return
        gruppe, blatt = treffer
        self._lade_kinder(self.wurzel)
        self.wurzel.item.IsExpanded = True
        gruppe.item.IsExpanded = True
        self._lade_kinder(gruppe)
        blatt.item.IsSelected = True
        blatt.item.BringIntoView()
        self.letzter_treffer = blatt.element_id
        self.status.Text = u"%s  /  %s" % (gruppe.name, blatt.name)

    def kopiere_ids(self):
        if not self.markiert:
            return
        Clipboard.SetText(u", ".join(u"%d" % w for w in sorted(self.markiert)))
        self.status.Text = t(u"%d IDs in die Zwischenablage kopiert.",
                             u"%d ids copied to the clipboard.",
                             u"%d ids copiados al portapapeles.") % len(
            self.markiert)

    # --- Protokoll ----------------------------------------------------------
    def _merke_protokoll(self, ueberschrift, protokoll):
        zeilen = [ueberschrift, u"=" * len(ueberschrift)] + protokoll.zeilen
        self.protokoll_text = u"\n".join(zeilen)
        try:
            with io.open(PROTOKOLL, "w", encoding="utf-8") as datei:
                datei.write(self.protokoll_text + u"\n")
        except Exception:
            pass

    def zeige_protokoll(self):
        if not self.protokoll_text:
            meldung(self.fenster, t(u"Es gibt noch kein Protokoll.",
                                    u"There is no log yet.",
                                    u"Todavía no hay registro."))
            return
        f = dlg.lade_xaml(uebersetze_xaml(
            XAML_PROTOKOLL.replace(u"{titel}", t(
                u"TransferSingle - Protokoll", u"TransferSingle - Log",
                u"TransferSingle - Registro")), PROTOKOLL_TEXTE))
        dlg.setze_besitzer(f, besitzer=self.fenster)
        f.FindName("text").Text = self.protokoll_text
        f.FindName("pfad").Text = PROTOKOLL
        f.ShowDialog()

    # --- Verwalten ----------------------------------------------------------
    def _markierte_pruefen(self):
        if not self.markiert:
            meldung(self.fenster, t(u"Es ist kein Element markiert.",
                                    u"No element is checked.",
                                    u"No hay ningún elemento marcado."))
            return False
        return True

    def loesche(self):
        if not self._markierte_pruefen():
            return
        if not dlg.frage(self.fenster, t(
                u"%d markierte Elemente in '%s' löschen?\n\nAbhängige "
                u"Elemente werden von Revit mit gelöscht.",
                u"Delete %d checked elements in '%s'?\n\nRevit also deletes "
                u"dependent elements.",
                u"¿Eliminar %d elementos marcados en '%s'?\n\nRevit también "
                u"elimina los elementos dependientes.") % (
                    len(self.markiert), self.quelle.titel),
                titel=TITEL, warnung=True):
            return
        protokoll = rv.loesche(self.quelle.doc, sorted(self.markiert))
        self._abschluss_verwalten(t(u"Löschen", u"Delete", u"Eliminar"),
                                  protokoll)

    def benenne_um(self, aktion):
        if not self._markierte_pruefen():
            return
        a = b = u""
        if aktion in (lg.ERSETZEN, lg.ZAHLEN):
            werte = self._zwei_felder(
                t(u"Suchen und Ersetzen", u"Find and replace",
                  u"Buscar y reemplazar") if aktion == lg.ERSETZEN
                else t(u"Zahlen ersetzen", u"Replace numbers",
                       u"Reemplazar números"),
                t(u"Suchen:", u"Find:", u"Buscar:"),
                t(u"Ersetzen durch:", u"Replace with:", u"Reemplazar por:"))
            if werte is None:
                return
            a, b = werte
        elif aktion in (lg.PRAEFIX, lg.SUFFIX):
            a = dlg.frage_text(
                self.fenster,
                t(u"Präfix", u"Prefix", u"Prefijo") if aktion == lg.PRAEFIX
                else t(u"Suffix", u"Suffix", u"Sufijo"),
                t(u"Text, der an den Namen angefügt wird:",
                  u"Text to add to the name:",
                  u"Texto que se añade al nombre:"))
            if not a:
                return
        try:
            protokoll = rv.benenne_um(self.quelle.doc, sorted(self.markiert),
                                      aktion, a, b)
        except ValueError:
            meldung(self.fenster, t(
                u"Bitte ganze Zahlen eingeben.", u"Please enter whole numbers.",
                u"Introduzca números enteros."), warnung=True)
            return
        self._abschluss_verwalten(t(u"Umbenennen", u"Rename", u"Renombrar"),
                                  protokoll)

    def _zwei_felder(self, titel, text1, text2):
        f = dlg.lade_xaml(uebersetze_xaml(
            XAML_ZWEI_FELDER.replace(u"{titel}", titel), ZWEI_FELDER_TEXTE))
        dlg.setze_besitzer(f, besitzer=self.fenster)
        f.FindName("text1").Text = text1
        f.FindName("text2").Text = text2
        ergebnis = {}

        def bei_ok(_s, _a):
            ergebnis["werte"] = (f.FindName("feld1").Text,
                                 f.FindName("feld2").Text)
            f.DialogResult = True
        f.FindName("ok").Click += sicher(lambda: f, bei_ok)
        f.Loaded += lambda _s, _a: f.FindName("feld1").Focus()
        f.ShowDialog()
        return ergebnis.get("werte")

    def _abschluss_verwalten(self, aktion, protokoll):
        self._merke_protokoll(u"%s - %s" % (aktion, self.quelle.titel),
                              protokoll)
        self.lade()
        self.status.Text = t(u"%s: %d erfolgreich, %d fehlgeschlagen.",
                             u"%s: %d succeeded, %d failed.",
                             u"%s: %d correctos, %d con error.") % (
            aktion, protokoll.ok, protokoll.fehler)
        if protokoll.fehler:
            meldung(self.fenster, self.status.Text + t(
                u"\n\nDetails im Protokoll.", u"\n\nDetails in the log.",
                u"\n\nDetalles en el registro."), warnung=True)

    # --- Übertragen ---------------------------------------------------------
    def _optionen(self):
        transformation = (rv.VERKNUEPFUNG if self.trans_link.IsChecked
                          else rv.GEMEINSAM if self.trans_gemeinsam.IsChecked
                          else rv.KEINE)
        duplikate = (rv.DUP_ABBRECHEN if self.dup_abbrechen.IsChecked
                     else rv.DUP_FRAGEN if self.dup_fragen.IsChecked
                     else rv.DUP_ZIEL)
        return rv.Optionen(
            transformation=transformation, duplikate=duplikate,
            dialoge_bestaetigen=bool(self.dialoge.IsChecked),
            plaene_mit_ansichten=bool(self.plaene_ansichten.IsChecked),
            ansichtselemente=bool(self.ansichtselemente.IsChecked))

    def uebertrage(self):
        if not self._markierte_pruefen():
            return
        ziele = [d for b, d in self.ziel_boxen if b.IsChecked]
        if not ziele:
            meldung(self.fenster, t(u"Bitte unter 'Nach' mindestens ein "
                                    u"Zielprojekt ankreuzen.",
                                    u"Please tick at least one target project "
                                    u"under 'To'.",
                                    u"Marque al menos un proyecto de destino "
                                    u"en 'A'."))
            return
        protokoll = rv.uebertrage(self.quelle, ziele, sorted(self.markiert),
                                  self._optionen())
        self._merke_protokoll(t(u"Übertragung", u"Transfer",
                                u"Transferencia"), protokoll)
        text = t(u"%d erfolgreich, %d fehlgeschlagen.",
                 u"%d succeeded, %d failed.",
                 u"%d correctos, %d con error.") % (protokoll.ok,
                                                     protokoll.fehler)
        self.status.Text = text
        meldung(self.fenster, text + (t(u"\n\nDetails im Protokoll.",
                                        u"\n\nDetails in the log.",
                                        u"\n\nDetalles en el registro.")
                                      if protokoll.fehler else u""),
                warnung=bool(protokoll.fehler))

    def zeige(self):
        self.fenster.ShowDialog()


def starte(uiapp):
    TransferSingleFenster(uiapp).zeige()
