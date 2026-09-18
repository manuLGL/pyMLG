# -*- coding: utf-8 -*-
"""Hauptfenster des Filter-Managers (WPF), aufgebaut wie der native Dialog
"Sichtbarkeit/Grafiken > Filter > Filter bearbeiten" - plus Suchfeld.

    ┌ Filter ─────────────┬ Kategorien ─────────┬ Filterregeln ──────────────┐
    │ Suchen: [______]    │ Liste filtern: [__] │ [UND ▾] Regel+  Satz+      │
    │ Suchen in: [Name ▾] │ □ nicht aktivierte  │  [Parameter▾][Operator▾]   │
    │ ┌─────────────────┐ │ ┌─────────────────┐ │  [Wert________] [–]        │
    │ │ Filterliste     │ │ │ ☑ Kategorien    │ │  ┌ [ODER ▾] ...            │
    │ └─────────────────┘ │ └─────────────────┘ │                            │
    │ Neu Dupl. Umb. Lösch│ Alle   Keine        │                            │
    └─────────────────────┴─────────────────────┴────────────────────────────┘
    Aktive Ansicht: …  [Auf aktive Ansicht] [Auf Ansichten…] [Entfernen]
                                                  [OK] [Abbrechen] [Anwenden]

Transaktionen: Das ganze Fenster läuft in einer TransactionGroup. Jede Aktion
ist eine eigene, benannte Transaktion darin. OK -> Assimilate (ein
Rückgängig-Schritt "pyMLG Filter-Manager"), Abbrechen -> RollBack. Das
entspricht dem Verhalten des nativen Dialogs.
"""

from Autodesk.Revit.DB import (
    BuiltInParameter,
    ElementId,
    FilteredElementCollector,
    ParameterFilterElement,
    StorageType,
    Transaction,
    TransactionGroup,
    TransactionStatus,
)
from System.Windows import (
    FontWeights,
    GridLength,
    GridUnitType,
    Thickness,
    Visibility,
)
from System.Windows.Controls import (
    Border,
    Button,
    ColumnDefinition,
    ComboBox,
    ComboBoxItem,
    Grid,
    ListBoxItem,
    StackPanel,
    TextBlock,
    TextBox,
    CheckBox,
    VirtualizingPanel,
    WrapPanel,
)
from System import Action
from System.Windows.Media import Color, FontFamily, SolidColorBrush
from System.Windows.Threading import DispatcherPriority

from filter_manager import darstellung as ds
from filter_manager import dialoge as dlg
from filter_manager import modell as mo
from filter_manager import regelbaum as rb
from mlg_sprache import t, uebersetze_xaml
from filter_manager.aufloesung import RevitAufloeser, id_liste, id_wert, \
    ist_ja_nein

FACHLICH = (mo.FilterFehler,)

# ViewType -> Bezeichnung wie im Projektbrowser der jeweiligen Sprache
ANSICHTSTYPEN = {
    "FloorPlan": t(u"Grundriss", u"Floor Plan", u"Plano de planta"),
    "CeilingPlan": t(u"Deckenplan", u"Ceiling Plan", u"Plano de techo"),
    "EngineeringPlan": t(u"Tragwerksplan", u"Structural Plan",
                         u"Plano estructural"),
    "AreaPlan": t(u"Flächenplan", u"Area Plan", u"Plano de área"),
    "Section": t(u"Schnitt", u"Section", u"Sección"),
    "Elevation": t(u"Ansicht", u"Elevation", u"Alzado"),
    "Detail": t(u"Detail", u"Detail View", u"Vista de detalle"),
    "ThreeD": t(u"3D-Ansicht", u"3D View", u"Vista 3D"),
    "DraftingView": t(u"Zeichenansicht", u"Drafting View",
                      u"Vista de diseño"),
    "Legend": t(u"Legende", u"Legend", u"Leyenda"),
    "Walkthrough": t(u"Walkthrough", u"Walkthrough", u"Recorrido"),
    "Rendering": t(u"Rendering", u"Rendering", u"Renderización"),
    "Schedule": t(u"Bauteilliste", u"Schedule", u"Tabla de planificación"),
}

SUCHE_NAME = 0
SUCHE_REGELN = 1


def _pinsel(r, g, b):
    # Eingefroren, sonst gehört der Pinsel dem Thread, der das Modul geladen
    # hat, und WPF verweigert ihn in Fenstern anderer Threads
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


GRUEN = _pinsel(112, 160, 60)
ROT = _pinsel(192, 57, 43)
GRAU = _pinsel(110, 110, 110)
HINTERGRUND_SATZ = _pinsel(246, 250, 240)

XAML_TEXTE = {
    "t0": (u"Filter-Manager (pyMLG)",
           u"Filter Manager (pyMLG)",
           u"Gestor de filtros (pyMLG)"),
    "t1": (u"Filter",
           u"Filters",
           u"Filtros"),
    "t2": (u"Suchen:",
           u"Search:",
           u"Buscar:"),
    "t3": (u"Mehrere Wörter: alle müssen vorkommen",
           u"Several words: all must occur",
           u"Varias palabras: deben aparecer todas"),
    "t4": (u"Filtername",
           u"Filter name",
           u"Nombre de filtro"),
    "t5": (u"Name, Parameter und Regelwerte",
           u"Name, parameters and rule values",
           u"Nombre, parámetros y valores de reglas"),
    "t6": (u"Neu…",
           u"New…",
           u"Nuevo…"),
    "t7": (u"Duplizieren",
           u"Duplicate",
           u"Duplicar"),
    "t8": (u"Umbenennen…",
           u"Rename…",
           u"Renombrar…"),
    "t9": (u"Löschen",
           u"Delete",
           u"Eliminar"),
    "t10": (u"Kategorien",
           u"Categories",
           u"Categorías"),
    "t11": (u"Wählen Sie eine oder mehrere Kategorien für den Filter aus. Für die Regeln stehen die Parameter zur Verfügung, die alle gewählten Kategorien gemeinsam haben.",
           u"Select one or more categories for the filter. The rules can use the parameters that all selected categories have in common.",
           u"Seleccione una o varias categorías para el filtro. Las reglas pueden usar los parámetros comunes a todas las categorías seleccionadas."),
    "t12": (u"Liste filtern:",
           u"Filter list:",
           u"Filtrar lista:"),
    "t13": (u"Nicht aktivierte Kategorien ausblenden",
           u"Hide un-checked categories",
           u"Ocultar categorías no marcadas"),
    "t14": (u"Alle markieren",
           u"Check all",
           u"Marcar todo"),
    "t15": (u"Keine markieren",
           u"Uncheck all",
           u"Desmarcar todo"),
    "t16": (u"Filterregeln",
           u"Filter rules",
           u"Reglas de filtro"),
    "t17": (u"Abbrechen",
           u"Cancel",
           u"Cancelar"),
    "t18": (u"Anwenden",
           u"Apply",
           u"Aplicar"),
    "t19": (u"Auf aktive Ansicht anwenden…",
           u"Apply to active view…",
           u"Aplicar a la vista activa…"),
    "t20": (u"Auf Ansichten anwenden…",
           u"Apply to views…",
           u"Aplicar a vistas…"),
    "t21": (u"Von aktiver Ansicht entfernen",
           u"Remove from active view",
           u"Quitar de la vista activa"),
}

XAML = u"""
<Window %s Title="{{t0}}" Width="1240" Height="700"
        MinWidth="900" MinHeight="460" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12"
        ResizeMode="CanResizeWithGrip">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="10,3"/>
      <Setter Property="Margin" Value="0,0,6,0"/>
    </Style>
    <Style TargetType="GroupBox">
      <Setter Property="Padding" Value="6"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
    </Style>
  </Window.Resources>
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="1*" MinWidth="240"/>
      <ColumnDefinition Width="1*" MinWidth="220"/>
      <ColumnDefinition Width="1.9*" MinWidth="380"/>
    </Grid.ColumnDefinitions>

    <!-- Filter -->
    <GroupBox Header="{{t1}}" Grid.Column="0">
      <DockPanel>
        <Grid DockPanel.Dock="Top" Margin="0,0,0,4">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <TextBlock Text="{{t2}}" VerticalAlignment="Center"
                     Margin="0,0,6,0"/>
          <TextBox x:Name="suche" Grid.Column="1" Padding="3"
                   ToolTip="{{t3}}"/>
          <TextBlock Grid.Row="1" Text="in:" VerticalAlignment="Center"
                     Margin="0,4,6,0"/>
          <ComboBox x:Name="suchmodus" Grid.Row="1" Grid.Column="1"
                    Margin="0,4,0,0" SelectedIndex="0">
            <ComboBoxItem Content="{{t4}}"/>
            <ComboBoxItem Content="{{t5}}"/>
          </ComboBox>
        </Grid>
        <WrapPanel DockPanel.Dock="Bottom" Margin="0,2,0,0">
          <Button x:Name="neu" Content="{{t6}}" Margin="0,4,6,0"/>
          <Button x:Name="duplizieren" Content="{{t7}}" Margin="0,4,6,0"/>
          <Button x:Name="umbenennen" Content="{{t8}}" Margin="0,4,6,0"/>
          <Button x:Name="loeschen" Content="{{t9}}" Margin="0,4,6,0"/>
        </WrapPanel>
        <TextBlock x:Name="filterzaehler" DockPanel.Dock="Bottom"
                   Foreground="#666" Margin="0,4,0,0"/>
        <ListBox x:Name="filterliste" SelectionMode="Extended"/>
      </DockPanel>
    </GroupBox>

    <!-- Kategorien -->
    <GroupBox Header="{{t10}}" Grid.Column="1">
      <DockPanel>
        <TextBlock DockPanel.Dock="Top" TextWrapping="Wrap" Margin="0,0,0,6"
                   Text="{{t11}}"/>
        <DockPanel DockPanel.Dock="Top" Margin="0,0,0,4">
          <TextBlock Text="{{t12}}" VerticalAlignment="Center"
                     Margin="0,0,6,0"/>
          <TextBox x:Name="katsuche" Padding="3"/>
        </DockPanel>
        <CheckBox x:Name="nuraktive" DockPanel.Dock="Top" Margin="0,2,0,6"
                  Content="{{t13}}"/>
        <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal"
                    Margin="0,6,0,0">
          <Button x:Name="katalle" Content="{{t14}}"/>
          <Button x:Name="katkeine" Content="{{t15}}"/>
        </StackPanel>
        <TextBlock x:Name="kathinweis" DockPanel.Dock="Bottom"
                   TextWrapping="Wrap" Foreground="#C0392B" Margin="0,4,0,0"
                   Visibility="Collapsed"/>
        <Border BorderBrush="#ABADB3" BorderThickness="1">
          <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel x:Name="katliste" Margin="4,2"/>
          </ScrollViewer>
        </Border>
      </DockPanel>
    </GroupBox>

    <!-- Filterregeln -->
    <GroupBox Header="{{t16}}" Grid.Column="2" Margin="0">
      <DockPanel>
        <TextBlock x:Name="regelhinweis" DockPanel.Dock="Top"
                   TextWrapping="Wrap" Margin="0,0,0,6" Foreground="#666"/>
        <TextBlock x:Name="formel" DockPanel.Dock="Bottom" TextWrapping="Wrap"
                   Margin="0,6,0,0" Foreground="#666"/>
        <ScrollViewer x:Name="regelscroll" VerticalScrollBarVisibility="Auto"
                      HorizontalScrollBarVisibility="Disabled">
          <StackPanel x:Name="regeln"/>
        </ScrollViewer>
      </DockPanel>
    </GroupBox>

    <!-- Fusszeile -->
    <DockPanel Grid.Row="1" Grid.ColumnSpan="3" Margin="0,10,0,0">
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
        <Button x:Name="ok" Content="OK" Width="90"/>
        <Button x:Name="abbrechen" Content="{{t17}}" Width="90"
                IsCancel="True"/>
        <Button x:Name="anwenden" Content="{{t18}}" Width="90"
                Margin="0"/>
      </StackPanel>
      <StackPanel Orientation="Horizontal">
        <Button x:Name="ansicht_an" Content="{{t19}}"/>
        <Button x:Name="ansichten_an" Content="{{t20}}"/>
        <Button x:Name="ansicht_weg" Content="{{t21}}"/>
        <TextBlock x:Name="ansichtstatus" VerticalAlignment="Center"
                   Foreground="#666" Margin="8,0" TextTrimming="CharacterEllipsis"/>
      </StackPanel>
    </DockPanel>
  </Grid>
</Window>""" % dlg.XMLNS


class FilterManagerFenster(object):

    def __init__(self, uiapp, doc):
        self.uiapp = uiapp
        self.doc = doc
        self.aufloeser = RevitAufloeser(doc)
        self.fenster = dlg.lade_xaml(uebersetze_xaml(XAML, XAML_TEXTE))
        dlg.setze_besitzer(self.fenster, handle=uiapp.MainWindowHandle)
        self.c = self.fenster.FindName

        self.eintraege = {}          # Id-Wert -> Eintrag (dict)
        self.ausgewaehlt = []        # Id-Werte der markierten Filter
        self.kategorien = set()      # Kategorien des aktuellen Filters
        self.satz = None             # Regelmodell (None = nicht bearbeitbar)
        self.nicht_bearbeitbar = u""
        self.geaendert = False       # ungespeicherte Änderungen im Editor
        self.sitzung_geaendert = False
        self.uebernehmen = False
        self._still = False          # Ereignisse bei Programmänderung ignorieren
        self._leser = []             # (Regel, Funktion) - Eingaben einlesen
        self._parameter_cache = {}
        self._neubau_geplant = False

        self._kategorie_boxen = []
        self._verdrahte()

    # ------------------------------------------------------------------
    # Grundgerüst
    # ------------------------------------------------------------------

    def _h(self, funktion):
        return dlg.sicher(lambda: self.fenster, funktion, FACHLICH)

    def _verdrahte(self):
        c = self.c
        c("suche").TextChanged += self._h(lambda s, a: self._fuelle_liste())
        c("suchmodus").SelectionChanged += self._h(
            lambda s, a: self._fuelle_liste())
        c("filterliste").SelectionChanged += self._h(self._bei_filterauswahl)
        c("neu").Click += self._h(self._neu)
        c("duplizieren").Click += self._h(self._duplizieren)
        c("umbenennen").Click += self._h(self._umbenennen)
        c("loeschen").Click += self._h(self._loeschen)
        c("katsuche").TextChanged += self._h(
            lambda s, a: self._filtere_kategorien())
        c("nuraktive").Click += self._h(
            lambda s, a: self._filtere_kategorien())
        c("katalle").Click += self._h(lambda s, a: self._markiere_alle(True))
        c("katkeine").Click += self._h(
            lambda s, a: self._markiere_alle(False))
        c("ansicht_an").Click += self._h(self._auf_aktive_ansicht)
        c("ansichten_an").Click += self._h(self._auf_ansichten)
        c("ansicht_weg").Click += self._h(self._von_aktiver_ansicht)
        c("anwenden").Click += self._h(lambda s, a: self._speichern())
        c("ok").Click += self._h(self._ok)
        self.fenster.Closing += self._h(self._beim_schliessen)

    def zeige(self):
        gruppe = TransactionGroup(self.doc, t(u"pyMLG Filter-Manager", u"pyMLG Filter Manager", u"pyMLG Gestor de filtros"))
        gruppe.Start()
        try:
            self._lade_kategorien()
            self._lade_filter()
            self._fuelle_liste()
            self._zeige_auswahl()
            self.fenster.ShowDialog()
        finally:
            if self.uebernehmen and self.sitzung_geaendert:
                gruppe.Assimilate()
            else:
                gruppe.RollBack()

    def _transaktion(self, name, funktion):
        """Führt funktion() in einer eigenen Transaktion aus."""
        transaktion = Transaction(self.doc, name)
        transaktion.Start()
        try:
            ergebnis = funktion()
            if transaktion.Commit() != TransactionStatus.Committed:
                raise mo.FilterFehler(t(u"Revit hat die Änderung \"%s\" nicht "
                                      u"übernommen.", u"Revit did not accept the change \"%s\".", u"Revit no aceptó el cambio \"%s\".") % name)
        except Exception:
            if transaktion.HasStarted() and not transaktion.HasEnded():
                transaktion.RollBack()
            raise
        self.sitzung_geaendert = True
        return ergebnis

    # ------------------------------------------------------------------
    # Daten laden
    # ------------------------------------------------------------------

    def _analysiere(self, filter_element):
        ergebnis = rb.analysiere(filter_element, self.aufloeser)
        ergebnis["formel"] = rb.als_formel(ergebnis["baum"]) \
            if ergebnis["baum"] is not None else u""
        ergebnis["suchtext_name"] = ergebnis["name"].lower()
        ergebnis["suchtext_regeln"] = (
            u"%s — Parameter: %s | %s | %s" % (
                ergebnis["name"], u", ".join(ergebnis["parameter"]),
                ergebnis["formel"], u", ".join(ergebnis["kategorien"]))
        ).lower()
        return ergebnis

    def _lade_filter(self):
        self.eintraege = {}
        for element in FilteredElementCollector(self.doc).OfClass(
                ParameterFilterElement):
            try:
                eintrag = self._analysiere(element)
            except Exception as fehler:
                eintrag = {"element": element, "id": id_wert(element.Id),
                           "name": element.Name, "kategorien": [],
                           "baum": None, "parameter": [],
                           "fehler": mo.fehlertext(fehler), "formel": u"",
                           "suchtext_name": element.Name.lower(),
                           "suchtext_regeln": element.Name.lower()}
            self.eintraege[eintrag["id"]] = eintrag

    def _lade_kategorien(self):
        panel = self.c("katliste")
        panel.Children.Clear()
        self._kategorie_boxen = []
        for name, wert in mo.filterbare_kategorien(self.doc, self.aufloeser):
            box = CheckBox()
            box.Content = name
            box.Tag = wert
            box.Margin = Thickness(0.0, 1.0, 0.0, 1.0)
            box.Click += self._h(self._bei_kategorie)
            panel.Children.Add(box)
            self._kategorie_boxen.append((name.lower(), box, wert))

    # ------------------------------------------------------------------
    # Filterliste
    # ------------------------------------------------------------------

    def _fuelle_liste(self):
        woerter = self.c("suche").Text.lower().split()
        modus = self.c("suchmodus").SelectedIndex
        liste = self.c("filterliste")
        schluessel = ("suchtext_regeln" if modus == SUCHE_REGELN
                      else "suchtext_name")
        passend = [e for e in self.eintraege.values()
                   if all(w in e[schluessel] for w in woerter)]
        passend.sort(key=lambda e: e["name"].lower())

        self._still = True
        try:
            liste.Items.Clear()
            for eintrag in passend:
                item = ListBoxItem()
                item.Content = eintrag["name"]
                item.Tag = eintrag["id"]
                item.ToolTip = self._tooltip(eintrag)
                if eintrag.get("fehler"):
                    item.Foreground = ROT
                liste.Items.Add(item)
                if eintrag["id"] in self.ausgewaehlt:
                    item.IsSelected = True
        finally:
            self._still = False
        self.c("filterzaehler").Text = t(u"%d von %d Filtern", u"%d of %d filters", u"%d de %d filtros") % (
            len(passend), len(self.eintraege))

    def _tooltip(self, eintrag):
        if eintrag.get("fehler"):
            return t(u"Regeln nicht lesbar: %s", u"Rules not readable: %s", u"Reglas no legibles: %s") % eintrag["fehler"]
        return t(u"Kategorien: %s\nRegeln: %s", u"Categories: %s\nRules: %s", u"Categorías: %s\nReglas: %s") % (
            u", ".join(eintrag["kategorien"]) or u"-",
            eintrag["formel"] or u"-")

    def _markierte_ids(self):
        return [item.Tag for item in self.c("filterliste").SelectedItems]

    def _bei_filterauswahl(self, sender, args):
        if self._still:
            return
        neu = self._markierte_ids()
        if not neu or sorted(neu) == sorted(self.ausgewaehlt):
            # Leere Markierung entsteht auch beim Neuaufbau der Liste
            # (Suche) - der bearbeitete Filter bleibt dann erhalten.
            return
        if self.geaendert:
            antwort = dlg.frage(
                self.fenster, t(u"Die Änderungen an \"%s\" wurden noch nicht "
                u"angewendet. Jetzt übernehmen?", u"The changes to \"%s\" have not been applied yet. Apply now?", u"Los cambios en \"%s\" aún no se han aplicado. ¿Aplicarlos ahora?") % self._aktuell()["name"],
                abbrechen=True)
            if antwort is None or (antwort and not self._speichern()):
                self._setze_markierung(self.ausgewaehlt)
                return
        self.ausgewaehlt = neu
        self.geaendert = False
        self._zeige_auswahl()

    def _setze_markierung(self, ids):
        self._still = True
        try:
            for item in self.c("filterliste").Items:
                item.IsSelected = item.Tag in ids
        finally:
            self._still = False

    def _aktuell(self):
        if len(self.ausgewaehlt) == 1:
            return self.eintraege.get(self.ausgewaehlt[0])
        return None

    def _markierte_eintraege(self):
        return [self.eintraege[i] for i in self.ausgewaehlt
                if i in self.eintraege]

    # ------------------------------------------------------------------
    # Anzeige des gewählten Filters
    # ------------------------------------------------------------------

    def _zeige_auswahl(self):
        eintrag = self._aktuell()
        anzahl = len(self.ausgewaehlt)
        einzeln = eintrag is not None
        for name in ("umbenennen",):
            self.c(name).IsEnabled = einzeln
        for name in ("duplizieren", "loeschen", "ansicht_an", "ansichten_an",
                     "ansicht_weg"):
            self.c(name).IsEnabled = anzahl > 0
        for name in ("katsuche", "nuraktive", "katalle", "katkeine",
                     "katliste", "anwenden"):
            self.c(name).IsEnabled = einzeln

        if einzeln:
            element = eintrag["element"]
            self.kategorien = set(id_wert(k) for k in element.GetCategories())
            self.aufloeser.setze_kontext([ElementId(k)
                                          for k in self.kategorien])
            try:
                if eintrag.get("fehler") or (
                        eintrag["baum"] is None
                        and element.GetElementFilter() is not None):
                    raise mo.NichtBearbeitbar(t(u"Regeln nicht lesbar: %s", u"Rules not readable: %s", u"Reglas no legibles: %s")
                                              % eintrag.get("fehler"))
                self.satz = mo.satz_aus_baum(eintrag["baum"], self.aufloeser)
                self.nicht_bearbeitbar = u""
            except mo.NichtBearbeitbar as grund:
                self.satz = None
                self.nicht_bearbeitbar = u"%s" % grund
        else:
            self.kategorien = set()
            self.satz = None
            self.nicht_bearbeitbar = u""
        self.geaendert = False

        self._still = True
        try:
            for _name, box, wert in self._kategorie_boxen:
                box.IsChecked = wert in self.kategorien
        finally:
            self._still = False
        self._filtere_kategorien()
        self._baue_regeln()
        self._zeige_ansichtstatus()

    def _zeige_ansichtstatus(self):
        ansicht = self.doc.ActiveView
        eintrag = self._aktuell()
        text = t(u"Aktive Ansicht: %s", u"Active view: %s", u"Vista activa: %s") % ansicht.Name
        try:
            if not ansicht.AreGraphicsOverridesAllowed():
                text += t(u" (unterstützt keine Filter)", u" (does not support filters)", u" (no admite filtros)")
            elif eintrag is not None:
                fid = eintrag["element"].Id
                if ansicht.IsFilterApplied(fid):
                    text += t(u" – Filter angewendet, %s", u" – filter applied, %s", u" – filtro aplicado, %s") % (
                        t(u"sichtbar", u"visible", u"visible") if ansicht.GetFilterVisibility(fid)
                        else t(u"unsichtbar", u"hidden", u"oculto"))
                else:
                    text += t(u" – Filter nicht angewendet", u" – filter not applied", u" – filtro no aplicado")
        except Exception:
            pass
        self.c("ansichtstatus").Text = text

    # ------------------------------------------------------------------
    # Kategorien
    # ------------------------------------------------------------------

    def _filtere_kategorien(self):
        woerter = self.c("katsuche").Text.lower().split()
        nur_aktive = bool(self.c("nuraktive").IsChecked)
        for name, box, _wert in self._kategorie_boxen:
            sichtbar = all(w in name for w in woerter) and (
                not nur_aktive or bool(box.IsChecked))
            box.Visibility = (Visibility.Visible if sichtbar
                              else Visibility.Collapsed)

    def _bei_kategorie(self, sender, args):
        if self._still:
            return
        if sender.IsChecked:
            self.kategorien.add(sender.Tag)
        else:
            self.kategorien.discard(sender.Tag)
        self._kategorien_geaendert()

    def _markiere_alle(self, zustand):
        for _name, box, wert in self._kategorie_boxen:
            if box.Visibility == Visibility.Visible:
                box.IsChecked = zustand
                if zustand:
                    self.kategorien.add(wert)
                else:
                    self.kategorien.discard(wert)
        self._kategorien_geaendert()

    def _kategorien_geaendert(self):
        self._lies_eingaben()
        self.geaendert = True
        self.aufloeser.setze_kontext([ElementId(k) for k in self.kategorien])
        self._spaeter_neu_aufbauen()

    def _parameterliste(self):
        schluessel = tuple(sorted(self.kategorien))
        if schluessel not in self._parameter_cache:
            self._parameter_cache[schluessel] = mo.filterbare_parameter(
                self.doc, self.kategorien, self.aufloeser)
        return self._parameter_cache[schluessel]

    # ------------------------------------------------------------------
    # Regeleditor
    # ------------------------------------------------------------------

    def _lies_eingaben(self):
        """Texteingaben in das Modell übernehmen."""
        for regel, leser in self._leser:
            alt = (regel.wert, regel.wert_id)
            leser()
            if (regel.wert, regel.wert_id) != alt:
                self.geaendert = True

    def _baue_regeln(self):
        panel = self.c("regeln")
        scroll = self.c("regelscroll")
        position = scroll.VerticalOffset
        panel.Children.Clear()
        self._leser = []
        hinweis = self.c("regelhinweis")
        formel = self.c("formel")
        eintrag = self._aktuell()

        if eintrag is None:
            hinweis.Text = (t(u"%d Filter markiert. Kategorien und Regeln lassen "
                            u"sich nur für einen einzelnen Filter bearbeiten.", u"%d filters selected. Categories and rules can only be edited for a single filter.", u"%d filtros seleccionados. Las categorías y reglas solo se pueden editar para un único filtro.")
                            % len(self.ausgewaehlt) if self.ausgewaehlt
                            else t(u"Links einen Filter auswählen.", u"Select a filter on the left.", u"Seleccione un filtro a la izquierda."))
            formel.Text = u""
            self._zeige_kategoriehinweis([])
            return

        if self.satz is None:
            hinweis.Text = (
                t(u"Dieser Filter enthält Regeln, die hier nicht bearbeitet "
                u"werden können (%s). Er wird nur angezeigt; Kategorien lassen "
                u"sich ändern. Regeln bitte im nativen Dialog bearbeiten.", u"This filter contains rules that cannot be edited here (%s). It is only displayed; categories can be changed. Please edit the rules in the native dialog.", u"Este filtro contiene reglas que no se pueden editar aquí (%s). Solo se muestra; las categorías se pueden cambiar. Edite las reglas en el cuadro de diálogo nativo.")
                % self.nicht_bearbeitbar)
            zeilen = (rb.als_zeilen(eintrag["baum"])
                      if eintrag["baum"] is not None else [])
            for zeile in zeilen:
                text = TextBlock()
                text.Text = zeile
                text.FontFamily = FontFamily("Consolas")
                panel.Children.Add(text)
            formel.Text = u""
            self._zeige_kategoriehinweis(
                [r.parametername for r in rb.regeln(eintrag["baum"])
                 if r.parameter_id is not None
                 and r.parameter_id not in set(
                     w for _n, w in self._parameterliste())])
            return

        hinweis.Text = (t(u"Keine Kategorien gewählt - Regeln sind erst nach "
                        u"Auswahl von Kategorien möglich.", u"No categories selected - rules are only possible after selecting categories.", u"No hay categorías seleccionadas: las reglas solo son posibles después de elegir categorías.")
                        if not self.kategorien else u"")
        panel.Children.Add(self._satz_ui(self.satz, None, 0))
        try:
            formel.Text = t(u"Regeln: ", u"Rules: ", u"Reglas: ") + self._formel_vorschau(self.satz)
        except Exception:
            formel.Text = u""
        self._zeige_kategoriehinweis(mo.ungueltige_regeln(
            self.doc, self.kategorien, self.satz, self.aufloeser)
            if self.kategorien else [])
        scroll.ScrollToVerticalOffset(position)

    def _zeige_kategoriehinweis(self, ungueltig):
        hinweis = self.c("kathinweis")
        if ungueltig:
            hinweis.Text = (t(u"Achtung: Für die gewählten Kategorien nicht "
                            u"verfügbar: %s. Diese Regeln werden beim "
                            u"Anwenden entfernt.", u"Warning: not available for the selected categories: %s. These rules are removed when applying.", u"Atención: no disponible para las categorías seleccionadas: %s. Estas reglas se eliminarán al aplicar.") % u", ".join(ungueltig))
            hinweis.Visibility = Visibility.Visible
        else:
            hinweis.Visibility = Visibility.Collapsed

    def _formel_vorschau(self, satz):
        teile = []
        for element in satz.elemente:
            if isinstance(element, mo.Satz):
                if element.elemente:
                    teile.append(u"(" + self._formel_vorschau(element) + u")")
            else:
                name = (self.aufloeser.parametername(element.parameter)
                        if element.parameter is not None else u"?")
                anzeige = element.wert
                if element.parameter is not None and self.aufloeser.info(
                        element.parameter).bearbeitungsbereich:
                    try:
                        anzeige = self.aufloeser.workset_name(
                            int(element.wert))
                    except (TypeError, ValueError):
                        pass
                wert = u"" if element.operator in mo.OHNE_WERT \
                    else u" „%s“" % anzeige
                teile.append(u"%s %s%s" % (
                    name, mo.OPERATOR_TEXT[element.operator], wert))
        return (t(u" ODER ", u" OR ", u" O ") if satz.oder else t(u" UND ", u" AND ", u" Y ")).join(teile) or u"-"

    def _knopf(self, text, funktion, tooltip=None, farbe=None):
        knopf = Button()
        knopf.Content = text
        knopf.Margin = Thickness(0.0, 0.0, 6.0, 2.0)
        knopf.Padding = Thickness(8.0, 2.0, 8.0, 2.0)
        if tooltip:
            knopf.ToolTip = tooltip
        if farbe is not None:
            knopf.Foreground = farbe
            knopf.FontWeight = FontWeights.Bold
        knopf.Click += self._h(funktion)
        return knopf

    def _satz_ui(self, satz, eltern, tiefe):
        rahmen = Border()
        rahmen.BorderBrush = GRUEN
        rahmen.BorderThickness = Thickness(3.0 if tiefe == 0 else 2.0)
        rahmen.Background = HINTERGRUND_SATZ
        rahmen.Padding = Thickness(6.0)
        rahmen.Margin = Thickness(0.0 if tiefe == 0 else 18.0, 4.0, 0.0, 4.0)
        inhalt = StackPanel()
        rahmen.Child = inhalt

        kopf = WrapPanel()
        kopf.Margin = Thickness(0.0, 0.0, 0.0, 2.0)
        verknuepfung = ComboBox()
        verknuepfung.Width = 250.0
        verknuepfung.Margin = Thickness(0.0, 0.0, 6.0, 2.0)
        verknuepfung.FontWeight = FontWeights.Bold
        for text in (t(u"UND (Alle Regeln müssen wahr sein)", u"AND (all rules must be true)", u"Y (todas las reglas deben cumplirse)"),
                     t(u"ODER (Eine Regel muss wahr sein)", u"OR (one rule must be true)", u"O (una regla debe cumplirse)")):
            item = ComboBoxItem()
            item.Content = text
            verknuepfung.Items.Add(item)
        verknuepfung.SelectedIndex = 1 if satz.oder else 0

        def bei_verknuepfung(sender, args):
            if self._still:
                return
            satz.oder = sender.SelectedIndex == 1
            self._aendern()

        verknuepfung.SelectionChanged += self._h(bei_verknuepfung)
        kopf.Children.Add(verknuepfung)

        def regel_dazu(sender, args):
            self._lies_eingaben()
            satz.elemente.append(mo.Regel())
            self._aendern()

        def satz_dazu(sender, args):
            self._lies_eingaben()
            satz.elemente.append(mo.Satz(oder=not satz.oder,
                                         elemente=[mo.Regel()]))
            self._aendern()

        def satz_weg(sender, args):
            self._lies_eingaben()
            eltern.elemente.remove(satz)
            self._aendern()

        kopf.Children.Add(self._knopf(t(u"Regel hinzufügen", u"Add rule", u"Añadir regla"), regel_dazu))
        kopf.Children.Add(self._knopf(t(u"Satz hinzufügen", u"Add set", u"Añadir conjunto"), satz_dazu))
        if eltern is not None:
            kopf.Children.Add(self._knopf(t(u"Satz entfernen", u"Remove set", u"Quitar conjunto"), satz_weg,
                                          farbe=ROT))
        inhalt.Children.Add(kopf)

        if not satz.elemente:
            leer = TextBlock()
            leer.Text = t(u"(keine Regeln)", u"(no rules)", u"(sin reglas)")
            leer.Foreground = GRAU
            leer.Margin = Thickness(4.0)
            inhalt.Children.Add(leer)
        for element in list(satz.elemente):
            if isinstance(element, mo.Satz):
                inhalt.Children.Add(self._satz_ui(element, satz, tiefe + 1))
            else:
                inhalt.Children.Add(self._regel_ui(element, satz))
        return rahmen

    def _regel_ui(self, regel, satz):
        zeile = Grid()
        zeile.Margin = Thickness(12.0, 2.0, 0.0, 2.0)
        for breite in (GridLength(1.2, GridUnitType.Star),
                       GridLength(1.0, GridUnitType.Star),
                       GridLength(1.2, GridUnitType.Star),
                       GridLength(1.0, GridUnitType.Auto)):
            spalte = ColumnDefinition()
            spalte.Width = breite
            zeile.ColumnDefinitions.Add(spalte)

        parameter = self._parameterliste()
        verfuegbar = set(w for _n, w in parameter)
        info = (self.aufloeser.info_mit_spec(regel.parameter)
                if regel.parameter is not None else None)

        # Parameter
        param_box = ComboBox()
        param_box.IsTextSearchEnabled = True
        param_box.Margin = Thickness(0.0, 0.0, 4.0, 0.0)
        param_box.ToolTip = (t(u"Aufklappen und Anfangsbuchstaben tippen, um in "
                             u"der Liste zu springen", u"Open the list and type the first letters to jump", u"Abra la lista y escriba las primeras letras para saltar"))
        self._still = True
        try:
            if regel.parameter is not None and regel.parameter not in verfuegbar:
                item = ComboBoxItem()
                item.Content = t(u"%s (nicht verfügbar)", u"%s (not available)", u"%s (no disponible)") % info.name
                item.Tag = regel.parameter
                item.Foreground = ROT
                param_box.Items.Add(item)
                param_box.SelectedItem = item
            for name, wert in parameter:
                item = ComboBoxItem()
                item.Content = name
                item.Tag = wert
                param_box.Items.Add(item)
                if wert == regel.parameter:
                    param_box.SelectedItem = item
        finally:
            self._still = False

        def bei_parameter(sender, args):
            if self._still or sender.IsDropDownOpen \
                    or sender.SelectedItem is None:
                return
            wert = sender.SelectedItem.Tag
            if wert == regel.parameter:
                return
            self._lies_eingaben()
            regel.parameter = wert
            erlaubt = mo.operatoren_fuer(self.aufloeser.info_mit_spec(wert))
            if regel.operator not in erlaubt:
                regel.operator = erlaubt[0]
            regel.wert, regel.wert_id = u"", None
            self._aendern()

        param_box.DropDownClosed += self._h(bei_parameter)
        param_box.LostKeyboardFocus += self._h(bei_parameter)
        Grid.SetColumn(param_box, 0)
        zeile.Children.Add(param_box)

        # Operator
        op_box = ComboBox()
        op_box.Margin = Thickness(0.0, 0.0, 4.0, 0.0)
        self._still = True
        try:
            for schluessel in mo.operatoren_fuer(info):
                item = ComboBoxItem()
                item.Content = mo.OPERATOR_TEXT[schluessel]
                item.Tag = schluessel
                op_box.Items.Add(item)
                if schluessel == regel.operator:
                    op_box.SelectedItem = item
        finally:
            self._still = False
        Grid.SetColumn(op_box, 1)
        zeile.Children.Add(op_box)

        # Wert
        wert_element = self._wert_ui(regel, info)
        wert_element.Margin = Thickness(0.0, 0.0, 4.0, 0.0)
        wert_element.Visibility = (Visibility.Collapsed
                                   if regel.operator in mo.OHNE_WERT
                                   else Visibility.Visible)
        Grid.SetColumn(wert_element, 2)
        zeile.Children.Add(wert_element)

        def bei_operator(sender, args):
            if self._still or sender.SelectedItem is None:
                return
            regel.operator = sender.SelectedItem.Tag
            wert_element.Visibility = (Visibility.Collapsed
                                       if regel.operator in mo.OHNE_WERT
                                       else Visibility.Visible)
            self._lies_eingaben()
            self._aendern(neu_aufbauen=False)

        op_box.SelectionChanged += self._h(bei_operator)

        def weg(sender, args):
            self._lies_eingaben()
            satz.elemente.remove(regel)
            self._aendern()

        entfernen = self._knopf(u"–", weg, t(u"Regel entfernen", u"Remove rule", u"Quitar regla"), ROT)
        entfernen.Margin = Thickness(0.0)
        Grid.SetColumn(entfernen, 3)
        zeile.Children.Add(entfernen)
        return zeile

    def _wert_ui(self, regel, info):
        art = info.speicherart if info is not None else None

        if info is not None and info.bearbeitungsbereich:
            return self._workset_ui(regel)

        if art == StorageType.Integer and ist_ja_nein(info.spec):
            box = ComboBox()
            ja, nein = t(u"Ja", u"Yes", u"Sí"), t(u"Nein", u"No", u"No")
            for text in (ja, nein):
                item = ComboBoxItem()
                item.Content = text
                box.Items.Add(item)
            box.SelectedIndex = 0 if regel.wert in (ja, u"Ja", u"1") else (
                1 if regel.wert in (nein, u"Nein", u"0") else -1)

            def lies():
                if box.SelectedItem is not None:
                    regel.wert = box.SelectedItem.Content
            self._leser.append((regel, lies))
            box.SelectionChanged += self._h(
                lambda s, a: None if self._still else self._lies_eingaben())
            return box

        if art == StorageType.ElementId:
            box = ComboBox()
            box.ToolTip = t(u"Vorhandene Werte werden beim Aufklappen gesucht", u"Existing values are searched when opening the list", u"Los valores existentes se buscan al abrir la lista")
            if regel.wert_id is not None:
                item = ComboBoxItem()
                item.Content = regel.wert or u"<Id %s>" % regel.wert_id
                item.Tag = regel.wert_id
                box.Items.Add(item)
                box.SelectedItem = item
            def laden():
                if regel.parameter is None:
                    return
                vorhanden = set(i.Tag for i in box.Items)
                for text, wert in mo.wertvorschlaege(
                        self.doc, self.kategorien, regel.parameter,
                        self.aufloeser):
                    if wert in vorhanden:
                        continue
                    item = ComboBoxItem()
                    item.Content = text
                    item.Tag = wert
                    box.Items.Add(item)

            def lies():
                if box.SelectedItem is not None:
                    regel.wert_id = box.SelectedItem.Tag
                    regel.wert = box.SelectedItem.Content

            self._vorschlaege_beim_oeffnen(box, laden, vorab=True)
            box.SelectionChanged += self._h(
                lambda s, a: None if self._still else self._lies_eingaben())
            self._leser.append((regel, lies))
            return box

        if art in (StorageType.Double, StorageType.Integer):
            feld = TextBox()
            feld.Text = regel.wert or u""
            feld.Padding = Thickness(3.0, 2.0, 3.0, 2.0)
            feld.ToolTip = (t(u"Wert in Projekteinheiten, z.B. 2750 oder "
                            u"2,75 m", u"Value in project units, e.g. 2750 or 2.75 m", u"Valor en unidades del proyecto, p. ej. 2750 o 2,75 m") if art == StorageType.Double
                            else t(u"Ganzzahl", u"Integer", u"Número entero"))

            def lies():
                regel.wert = feld.Text
            self._leser.append((regel, lies))
            return feld

        # Text (oder unbekannt): Eingabe mit Vorschlägen vorhandener Werte
        box = ComboBox()
        box.IsEditable = True
        box.Text = regel.wert or u""
        box.ToolTip = t(u"Wert eintippen oder vorhandenen Wert aufklappen", u"Type a value or open the list of existing values", u"Escriba un valor o abra la lista de valores existentes")
        def laden_text():
            if regel.parameter is None:
                return
            text = box.Text
            for vorschlag, _wert in mo.wertvorschlaege(
                    self.doc, self.kategorien, regel.parameter,
                    self.aufloeser):
                box.Items.Add(vorschlag)
            box.Text = text

        def lies_text():
            regel.wert = box.Text or u""
        # Eingabefeld: nicht schon beim Hineinklicken laden (man will tippen)
        self._vorschlaege_beim_oeffnen(box, laden_text, vorab=False)
        self._leser.append((regel, lies_text))
        return box

    def _vorschlaege_beim_oeffnen(self, box, laden, vorab):
        """Werteliste einer ComboBox erst bei Bedarf befüllen - aber bevor sie
        sich öffnet.

        Kommen die Einträge erst im DropDownOpened hinzu, zeigt WPF eine
        leere Liste in passender Höhe: das Aufklappfeld ist dann schon
        aufgebaut und zeichnet die neuen Einträge nicht. Deshalb:
          vorab=True  laden schon beim Mausklick bzw. Tastendruck auf das Feld
          sonst       beim Öffnen laden, schließen und sofort neu öffnen
        """
        VirtualizingPanel.SetIsVirtualizing(box, False)
        geladen = {"ja": False}

        def einmal_laden():
            if geladen["ja"]:
                return False
            geladen["ja"] = True
            laden()
            return True

        def neu_oeffnen():
            box.IsDropDownOpen = True

        def beim_oeffnen(sender, args):
            if not einmal_laden():
                return
            box.IsDropDownOpen = False
            self.fenster.Dispatcher.InvokeAsync(Action(neu_oeffnen),
                                                DispatcherPriority.Background)

        if vorab:
            box.PreviewMouseLeftButtonDown += self._h(
                lambda s, a: einmal_laden())
            box.PreviewKeyDown += self._h(lambda s, a: einmal_laden())
        box.DropDownOpened += self._h(beim_oeffnen)

    def _workset_ui(self, regel):
        """Bearbeitungsbereich: Revit speichert die WorksetId als Ganzzahl,
        gewählt wird wie im nativen Dialog der Name."""
        box = ComboBox()
        try:
            aktuell = int((regel.wert or u"").strip())
        except ValueError:
            aktuell = None
        eintraege = self.aufloeser.benutzer_worksets()
        if aktuell is not None and aktuell not in set(w for _n, w
                                                      in eintraege):
            item = ComboBoxItem()
            item.Content = self.aufloeser.workset_name(aktuell)
            item.Tag = aktuell
            item.Foreground = ROT
            box.Items.Add(item)
            box.SelectedItem = item
        for name, wert in eintraege:
            item = ComboBoxItem()
            item.Content = name
            item.Tag = wert
            box.Items.Add(item)
            if wert == aktuell:
                box.SelectedItem = item
        if not eintraege:
            box.ToolTip = t(u"Das Projekt hat keine Teamarbeit (Worksets)", u"The project is not workshared (worksets)", u"El proyecto no tiene trabajo compartido (subproyectos)")

        def lies():
            if box.SelectedItem is not None:
                regel.wert = u"%d" % box.SelectedItem.Tag
        self._leser.append((regel, lies))
        box.SelectionChanged += self._h(
            lambda s, a: None if self._still else self._lies_eingaben())
        return box

    def _spaeter_neu_aufbauen(self):
        """Regelansicht nach dem laufenden Ereignis neu aufbauen - das
        auslösende Steuerelement wird sonst mitten im Ereignis entfernt."""
        if self._neubau_geplant:
            return
        self._neubau_geplant = True

        def ausfuehren():
            self._neubau_geplant = False
            try:
                self._baue_regeln()
            except Exception as fehler:
                dlg.zeige_fehler(self.fenster, fehler, FACHLICH)

        self.fenster.Dispatcher.InvokeAsync(Action(ausfuehren),
                                            DispatcherPriority.Background)

    def _aendern(self, neu_aufbauen=True):
        self.geaendert = True
        if neu_aufbauen:
            self._spaeter_neu_aufbauen()
        else:
            try:
                self.c("formel").Text = t(u"Regeln: ", u"Rules: ", u"Reglas: ") + self._formel_vorschau(
                    self.satz)
            except Exception:
                pass

    # ------------------------------------------------------------------
    # Speichern
    # ------------------------------------------------------------------

    def _speichern(self):
        """Kategorien/Regeln des aktuellen Filters schreiben. True bei Erfolg."""
        eintrag = self._aktuell()
        if eintrag is None or not self.geaendert:
            return True
        self._lies_eingaben()
        element = eintrag["element"]
        kategorien = sorted(self.kategorien)

        if self.satz is None:
            # Nicht bearbeitbare Regeln: nur Kategorien setzen, Revit entfernt
            # dabei selbst die nicht mehr passenden Regeln
            betroffen = self.c("kathinweis").Visibility == Visibility.Visible
            if betroffen and not dlg.frage(
                    self.fenster, self.c("kathinweis").Text + t(u"\n\n"
                    u"Kategorien trotzdem ändern?", u"\n\nChange categories anyway?", u"\n\n¿Cambiar las categorías de todos modos?"), warnung=True):
                return False
            self._transaktion(t(u"Filterkategorien ändern", u"Change filter categories", u"Cambiar categorías de filtro"),
                              lambda: element.SetCategories(
                                  id_liste(kategorien)))
        else:
            ungueltig = mo.ungueltige_regeln(self.doc, kategorien, self.satz,
                                             self.aufloeser)
            if ungueltig:
                if not dlg.frage(
                        self.fenster,
                        t(u"Die Regeln mit folgenden Parametern gelten nicht "
                        u"für alle gewählten Kategorien und werden entfernt:"
                        u"\n\n%s\n\nFortfahren?", u"The rules with the following parameters do not apply to all selected categories and will be removed:\n\n%s\n\nContinue?", u"Las reglas con los siguientes parámetros no se aplican a todas las categorías seleccionadas y se eliminarán:\n\n%s\n\n¿Continuar?") % u"\n".join(ungueltig),
                        warnung=True):
                    return False
                zulaessig = set(w for _n, w in self._parameterliste())
                self._entferne_ungueltige(self.satz, zulaessig)
            if not self.satz.ist_leer() and not kategorien:
                raise mo.FilterFehler(t(u"Ohne Kategorien sind keine Regeln "
                                      u"möglich. Bitte Kategorien wählen.", u"Rules are not possible without categories. Please select categories.", u"Sin categorías no hay reglas posibles. Seleccione categorías."))
            self._transaktion(t(u"Filterregeln ändern", u"Change filter rules", u"Cambiar reglas de filtro"),
                              lambda: mo.speichere_regeln(
                                  self.doc, element, kategorien, self.satz,
                                  self.aufloeser))

        self._aktualisiere_eintrag(element)
        self.geaendert = False
        self._zeige_auswahl()
        return True

    def _entferne_ungueltige(self, satz, zulaessig):
        for element in list(satz.elemente):
            if isinstance(element, mo.Satz):
                self._entferne_ungueltige(element, zulaessig)
            elif element.parameter not in zulaessig:
                satz.elemente.remove(element)

    def _aktualisiere_eintrag(self, element):
        self.eintraege[id_wert(element.Id)] = self._analysiere(element)
        self._fuelle_liste()

    # ------------------------------------------------------------------
    # Filteraktionen
    # ------------------------------------------------------------------

    def _vorher_speichern(self):
        if not self.geaendert:
            return True
        antwort = dlg.frage(self.fenster, t(u"Die Änderungen an \"%s\" wurden "
                            u"noch nicht angewendet. Jetzt übernehmen?", u"The changes to \"%s\" have not been applied yet. Apply now?", u"Los cambios en \"%s\" aún no se han aplicado. ¿Aplicarlos ahora?")
                            % self._aktuell()["name"], abbrechen=True)
        if antwort is None:
            return False
        if antwort:
            return self._speichern()
        self.geaendert = False
        self._zeige_auswahl()
        return True

    def _waehle_filter(self, ids):
        self.ausgewaehlt = list(ids)
        self.c("suche").Text = u""        # neue Filter sicher sichtbar
        self._fuelle_liste()
        self._zeige_auswahl()

    def _neu(self, sender, args):
        if not self._vorher_speichern():
            return
        name = dlg.frage_text(
            self.fenster, t(u"Neuer Filter", u"New filter", u"Filtro nuevo"), t(u"Name des neuen Filters:", u"Name of the new filter:", u"Nombre del filtro nuevo:"),
            mo.freier_name(self.doc, t(u"Filter 1", u"Filter 1", u"Filtro 1")),
            pruefen=lambda n: mo.pruefe_name(self.doc, n))
        if name is None:
            return
        element = self._transaktion(
            t(u"Filter erstellen", u"Create filter", u"Crear filtro"), lambda: mo.neuer_filter(self.doc, name))
        self._aktualisiere_eintrag(element)
        self._waehle_filter([id_wert(element.Id)])

    def _duplizieren(self, sender, args):
        if not self._vorher_speichern():
            return
        markiert = self._markierte_eintraege()
        if len(markiert) == 1:
            name = dlg.frage_text(
                self.fenster, t(u"Filter duplizieren", u"Duplicate filter", u"Duplicar filtro"),
                t(u"Name der Kopie von \"%s\":", u"Name of the copy of \"%s\":", u"Nombre de la copia de \"%s\":") % markiert[0]["name"],
                mo.freier_name(self.doc, markiert[0]["name"] + t(u" - Kopie", u" - Copy", u" - Copia")),
                pruefen=lambda n: mo.pruefe_name(self.doc, n))
            if name is None:
                return
            namen = [name]
        else:
            if not dlg.frage(self.fenster, t(u"%d Filter duplizieren? Die Kopien "
                             u"erhalten den Zusatz \" - Kopie\".", u"Duplicate %d filters? The copies get the suffix \" - Copy\".", u"¿Duplicar %d filtros? Las copias reciben el sufijo \" - Copia\".")
                             % len(markiert)):
                return
            namen = None

        def ausfuehren():
            kopien = []
            for index, eintrag in enumerate(markiert):
                name = namen[0] if namen else mo.freier_name(
                    self.doc, eintrag["name"] + t(u" - Kopie", u" - Copy", u" - Copia"))
                kopien.append(mo.dupliziere(self.doc, eintrag["element"],
                                            name))
            return kopien

        kopien = self._transaktion(t(u"Filter duplizieren", u"Duplicate filter", u"Duplicar filtro"), ausfuehren)
        for kopie in kopien:
            self.eintraege[id_wert(kopie.Id)] = self._analysiere(kopie)
        self._waehle_filter([id_wert(k.Id) for k in kopien])

    def _umbenennen(self, sender, args):
        eintrag = self._aktuell()
        if eintrag is None or not self._vorher_speichern():
            return
        alt = eintrag["name"]
        name = dlg.frage_text(
            self.fenster, t(u"Filter umbenennen", u"Rename filter", u"Renombrar filtro"), t(u"Neuer Name für \"%s\":", u"New name for \"%s\":", u"Nombre nuevo para \"%s\":") % alt,
            alt, pruefen=lambda n: mo.pruefe_name(self.doc, n, alt))
        if name is None or name == alt:
            return
        element = eintrag["element"]

        def ausfuehren():
            element.Name = name
        self._transaktion(t(u"Filter umbenennen", u"Rename filter", u"Renombrar filtro"), ausfuehren)
        self._aktualisiere_eintrag(element)
        self._zeige_auswahl()

    def _loeschen(self, sender, args):
        markiert = self._markierte_eintraege()
        if not markiert:
            return
        namen = u"\n".join(u"• " + e["name"] for e in markiert[:15])
        if len(markiert) > 15:
            namen += t(u"\n… und %d weitere", u"\n… and %d more", u"\n… y %d más") % (len(markiert) - 15)
        if not dlg.frage(self.fenster,
                         t(u"%d Filter löschen?\n\n%s\n\nDie Filter werden auch "
                         u"aus allen Ansichten und Ansichtsvorlagen entfernt.", u"Delete %d filters?\n\n%s\n\nThe filters are also removed from all views and view templates.", u"¿Eliminar %d filtros?\n\n%s\n\nLos filtros también se quitan de todas las vistas y plantillas de vista.")
                         % (len(markiert), namen), warnung=True):
            return
        ids = [e["element"].Id for e in markiert]
        self._transaktion(t(u"Filter löschen", u"Delete filters", u"Eliminar filtros"),
                          lambda: self.doc.Delete(id_liste(ids)))
        for eintrag in markiert:
            self.eintraege.pop(eintrag["id"], None)
        self.geaendert = False
        self._waehle_filter([])

    # ------------------------------------------------------------------
    # Ansichten
    # ------------------------------------------------------------------

    def _steuernde_vorlage(self, ansicht):
        """Vorlage, die die Filter dieser Ansicht steuert, sonst None."""
        try:
            vorlage_id = ansicht.ViewTemplateId
            if vorlage_id is None or id_wert(vorlage_id) == -1:
                return None
            vorlage = self.doc.GetElement(vorlage_id)
            if vorlage is None:
                return None
            filter_param = id_wert(ElementId(
                BuiltInParameter.VIS_GRAPHICS_FILTERS))
            frei = set(id_wert(i) for i in
                       vorlage.GetNonControlledTemplateParameterIds())
            return None if filter_param in frei else vorlage
        except Exception:
            return None

    def _darstellung_waehlen(self, titel, text, ansicht=None):
        """Dialog Sichtbarkeit/Grafiken. Rückgabe ds.Darstellung oder None.

        Ist genau ein Filter markiert und in der Ansicht schon angewendet,
        zeigt der Dialog dessen aktuelle Sichtbarkeit und Überschreibungen.
        """
        vorbelegung, sichtbar = None, True
        eintrag = self._aktuell()
        if ansicht is not None and eintrag is not None:
            fid = eintrag["element"].Id
            try:
                if ansicht.IsFilterApplied(fid):
                    vorbelegung = ansicht.GetFilterOverrides(fid)
                    sichtbar = ansicht.GetFilterVisibility(fid)
            except Exception:
                pass
        return ds.frage_darstellung(self.fenster, self.doc, titel, text,
                                    vorbelegung, sichtbar)

    def _anwenden_auf(self, ansichten, darstellung):
        """Rückgabe: (Anzahl ok, [(Ansicht, Filter, Grund)])"""
        markiert = self._markierte_eintraege()
        fehler = []

        def ausfuehren():
            zaehler = 0
            for ansicht in ansichten:
                vorlage = self._steuernde_vorlage(ansicht)
                if vorlage is not None:
                    fehler.append((ansicht.Name, t(u"alle", u"all", u"todos"),
                                   t(u"Filter werden von der Vorlage \"%s\" "
                                   u"gesteuert", u"Filters are controlled by the template \"%s\"", u"Los filtros los controla la plantilla \"%s\"") % vorlage.Name))
                    continue
                for eintrag in markiert:
                    fid = eintrag["element"].Id
                    try:
                        if not ansicht.IsFilterApplied(fid):
                            ansicht.AddFilter(fid)
                        darstellung.anwenden(ansicht, fid, self.doc)
                        zaehler += 1
                    except Exception as ausnahme:
                        fehler.append((ansicht.Name, eintrag["name"],
                                       mo.fehlertext(ausnahme)))
            return zaehler

        erfolgreich = self._transaktion(t(u"Filter auf Ansichten anwenden", u"Apply filters to views", u"Aplicar filtros a vistas"),
                                        ausfuehren)
        return erfolgreich, fehler

    def _bericht(self, erfolgreich, fehler, was):
        text = t(u"%d Zuordnung(en) %s.", u"%d assignment(s) %s.", u"%d asignación(es) %s.") % (erfolgreich, was)
        if fehler:
            text += t(u"\n\nNicht möglich (%d):\n", u"\n\nNot possible (%d):\n", u"\n\nNo es posible (%d):\n") % len(fehler)
            text += u"\n".join(u"• %s / %s: %s" % f for f in fehler[:20])
            if len(fehler) > 20:
                text += t(u"\n… und %d weitere", u"\n… and %d more", u"\n… y %d más") % (len(fehler) - 20)
        dlg.meldung(self.fenster, text, warnung=bool(fehler))

    def _auf_aktive_ansicht(self, sender, args):
        if not self._vorher_speichern():
            return
        ansicht = self.doc.ActiveView
        if not ansicht.AreGraphicsOverridesAllowed():
            raise mo.FilterFehler(t(u"Die aktive Ansicht \"%s\" unterstützt "
                                  u"keine Anzeigefilter.", u"The active view \"%s\" does not support view filters.", u"La vista activa \"%s\" no admite filtros de vista.") % ansicht.Name)
        vorlage = self._steuernde_vorlage(ansicht)
        if vorlage is not None:
            if not dlg.frage(self.fenster, t(u"Die Filter der aktiven Ansicht "
                             u"werden von der Ansichtsvorlage \"%s\" gesteuert."
                             u"\n\nStattdessen auf diese Vorlage anwenden? "
                             u"(Wirkt auf alle Ansichten mit dieser Vorlage.)", u"The filters of the active view are controlled by the view template \"%s\".\n\nApply to this template instead? (Affects all views using this template.)", u"Los filtros de la vista activa los controla la plantilla de vista \"%s\".\n\n¿Aplicar a esta plantilla en su lugar? (Afecta a todas las vistas con esta plantilla.)")
                             % vorlage.Name):
                return
            ansicht = vorlage
        darstellung = self._darstellung_waehlen(
            t(u"Auf aktive Ansicht anwenden", u"Apply to active view", u"Aplicar a la vista activa"),
            t(u"%d Filter auf \"%s\" anwenden - Sichtbarkeit und Grafiken wie in Sichtbarkeit/Grafiken:", u"Apply %d filters to \"%s\" - visibility and graphics as in Visibility/Graphics:", u"Aplicar %d filtros a \"%s\" - visibilidad y gráficos como en Visibilidad/Gráficos:")
            % (len(self.ausgewaehlt), ansicht.Name), ansicht)
        if darstellung is None:
            return
        erfolgreich, fehler = self._anwenden_auf([ansicht], darstellung)
        self._bericht(erfolgreich, fehler, t(u"angewendet", u"applied", u"aplicadas"))
        self._zeige_ansichtstatus()

    def _auf_ansichten(self, sender, args):
        if not self._vorher_speichern():
            return
        eintraege = []
        for ansicht in mo.ansichten_mit_filtern(self.doc):
            art = t(u"Vorlage", u"Template", u"Plantilla") if ansicht.IsTemplate else ANSICHTSTYPEN.get(
                u"%s" % ansicht.ViewType, u"%s" % ansicht.ViewType)
            vorlage = None if ansicht.IsTemplate \
                else self._steuernde_vorlage(ansicht)
            zusatz = (t(u"   [gesteuert von Vorlage \"%s\"]", u"   [controlled by template \"%s\"]", u"   [controlada por la plantilla \"%s\"]") % vorlage.Name
                      if vorlage is not None else u"")
            eintraege.append((u"%s: %s%s" % (art, ansicht.Name, zusatz),
                              id_wert(ansicht.Id)))
        eintraege.sort(key=lambda e: e[0].lower())
        ergebnis = dlg.waehle(
            self.fenster, t(u"Auf Ansichten anwenden", u"Apply to views", u"Aplicar a vistas"),
            t(u"Ansichten und Ansichtsvorlagen wählen, auf die %d Filter "
            u"angewendet werden sollen. Ansichten, deren Filter von einer "
            u"Vorlage gesteuert werden, werden übersprungen - dann die Vorlage "
            u"wählen.", u"Choose the views and view templates to apply %d filters to. Views whose filters are controlled by a template are skipped - choose the template instead.", u"Elija las vistas y plantillas de vista a las que aplicar %d filtros. Se omiten las vistas cuyos filtros controla una plantilla: elija entonces la plantilla.") % len(self.ausgewaehlt),
            eintraege, mehrfach=True)
        if ergebnis is None:
            return
        ansichten = [self.doc.GetElement(ElementId(i)) for i in ergebnis]
        darstellung = self._darstellung_waehlen(
            t(u"Auf Ansichten anwenden", u"Apply to views", u"Aplicar a vistas"),
            t(u"%d Filter auf %d Ansicht(en) anwenden - Sichtbarkeit und Grafiken:", u"Apply %d filters to %d view(s) - visibility and graphics:", u"Aplicar %d filtros a %d vista(s) - visibilidad y gráficos:")
            % (len(self.ausgewaehlt), len(ansichten)),
            ansichten[0] if len(ansichten) == 1 else None)
        if darstellung is None:
            return
        erfolgreich, fehler = self._anwenden_auf(ansichten, darstellung)
        self._bericht(erfolgreich, fehler, t(u"angewendet", u"applied", u"aplicadas"))
        self._zeige_ansichtstatus()

    def _von_aktiver_ansicht(self, sender, args):
        ansicht = self.doc.ActiveView
        if not ansicht.AreGraphicsOverridesAllowed():
            raise mo.FilterFehler(t(u"Die aktive Ansicht \"%s\" unterstützt "
                                  u"keine Anzeigefilter.", u"The active view \"%s\" does not support view filters.", u"La vista activa \"%s\" no admite filtros de vista.") % ansicht.Name)
        vorlage = self._steuernde_vorlage(ansicht)
        if vorlage is not None:
            if not dlg.frage(self.fenster, t(u"Die Filter der aktiven Ansicht "
                             u"werden von der Ansichtsvorlage \"%s\" gesteuert."
                             u"\n\nStattdessen aus dieser Vorlage entfernen?", u"The filters of the active view are controlled by the view template \"%s\".\n\nRemove from this template instead?", u"Los filtros de la vista activa los controla la plantilla de vista \"%s\".\n\n¿Quitar de esta plantilla en su lugar?")
                             % vorlage.Name):
                return
            ansicht = vorlage
        markiert = self._markierte_eintraege()
        angewendet = [e for e in markiert
                      if ansicht.IsFilterApplied(e["element"].Id)]
        if not angewendet:
            dlg.meldung(self.fenster, t(u"Keiner der markierten Filter ist in "
                        u"\"%s\" angewendet.", u"None of the selected filters is applied in \"%s\".", u"Ninguno de los filtros seleccionados está aplicado en \"%s\".") % ansicht.Name)
            return

        def ausfuehren():
            for eintrag in angewendet:
                ansicht.RemoveFilter(eintrag["element"].Id)
        self._transaktion(t(u"Filter aus Ansicht entfernen", u"Remove filters from view", u"Quitar filtros de la vista"), ausfuehren)
        self._bericht(len(angewendet), [], t(u"aus \"%s\" entfernt", u"removed from \"%s\"", u"quitadas de \"%s\"")
                      % ansicht.Name)
        self._zeige_ansichtstatus()

    # ------------------------------------------------------------------
    # Schliessen
    # ------------------------------------------------------------------

    def _ok(self, sender, args):
        if self.geaendert and not self._speichern():
            return
        self.uebernehmen = True
        self.fenster.Close()

    def _beim_schliessen(self, sender, args):
        if self.uebernehmen or not (self.sitzung_geaendert or self.geaendert):
            return
        if not dlg.frage(self.fenster, t(u"Alle Änderungen dieser Sitzung "
                         u"verwerfen?", u"Discard all changes of this session?", u"¿Descartar todos los cambios de esta sesión?"), warnung=True):
            args.Cancel = True


def starte(uiapp, doc):
    FilterManagerFenster(uiapp, doc).zeige()
