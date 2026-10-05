# -*- coding: utf-8 -*-
"""Hauptfenster von LinkedViews (WPF) - die verknüpfte Ansicht einer
RVT-Verknüpfung suchen und setzen, ohne sich durch Sichtbarkeit/Grafiken >
Revit-Verknüpfungen > Anzeigeeinstellungen zu klicken.

    ┌ Verknüpfungen ──────┬ Ansichten in den Verknüpfungen ─────────────────┐
    │ Suchen: [______]    │ Suchen: [__________]  Typ: [Passend ▾] □ benutzt│
    │ ┌─────────────────┐ │ ┌─────────────────────────────────────────────┐ │
    │ │ AR.rvt          │ │ │ ● Grundriss: UG02            AR.rvt         │ │
    │ │ TW.rvt          │ │ │   Ebene UG02 · in 3 Ansichten               │ │
    │ └─────────────────┘ │ └─────────────────────────────────────────────┘ │
    ├ Aktive Ansicht ─────┴─────────────────────────────────────────────────┤
    │ M100 - UG02 (gesteuert von Vorlage …)                                │
    │ AR.rvt: Nach verknüpfter Ansicht "Grundriss: UG02"                   │
    └──────────────────────────────────────────────────────────────────────┘
    [Auf aktive Ansicht] [Auf Ansichten…] [Zurücksetzen]     [OK] [Abbrechen]

Steuert die Ansichtsvorlage der Zielansicht die RVT-Verknüpfungen (Haken in
der Vorlage), wird die Vorlage geändert - nach Rückfrage.

Transaktionen wie im Filter-Manager: alles läuft in einer TransactionGroup,
OK -> Assimilate (ein Rückgängig-Schritt), Abbrechen -> RollBack.
"""

import io
import os
import traceback

from Autodesk.Revit.DB import (
    Transaction,
    TransactionGroup,
    TransactionStatus,
)
from System.Windows import (
    FontWeights,
    GridLength,
    GridUnitType,
    TextTrimming,
    Thickness,
)
from System.Windows.Controls import (
    ColumnDefinition,
    Grid,
    ListBoxItem,
    Orientation,
    StackPanel,
    TextBlock,
)
from System.Windows.Media import Color, SolidColorBrush

from filter_manager import dialoge as dlg
from filter_manager.aufloesung import id_wert
from link_ansichten import logik as lg
from link_ansichten import revit as rv
from mlg_sprache import t, tt, uebersetze_xaml

TITEL = t(u"Verknüpfte Ansichten", u"Linked Views", u"Vistas vinculadas")

FEHLERPROTOKOLL = os.path.join(dlg.protokollordner(),
                               "LinkedViews_Fehler.log")


class LinkFehler(Exception):
    """Fachlicher Fehler - nur als Text anzeigen."""


def _pinsel(r, g, b):
    # Eingefroren, sonst gehört der Pinsel dem Thread, der das Modul geladen
    # hat (siehe filter_manager/fenster.py)
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


GRUEN = _pinsel(33, 115, 70)
GRAU = _pinsel(110, 110, 110)
ROT = _pinsel(192, 57, 43)

TYP_PASSEND = 0
TYP_ALLE = 1

XAML_TEXTE = {
    "titel": (u"Verknüpfte Ansichten (pyMLG)",
              u"Linked Views (pyMLG)",
              u"Vistas vinculadas (pyMLG)"),
    "links": (u"Verknüpfungen",
              u"Links",
              u"Vínculos"),
    "suchen": (u"Suchen:",
               u"Search:",
               u"Buscar:"),
    "suche_tip": (u"Mehrere Wörter: alle müssen vorkommen",
                  u"Several words: all must occur",
                  u"Varias palabras: deben aparecer todas"),
    "links_hinweis": (u"Ohne Markierung werden die Ansichten aller "
                      u"Verknüpfungen gezeigt.",
                      u"Without a selection the views of all links are "
                      u"shown.",
                      u"Sin selección se muestran las vistas de todos los "
                      u"vínculos."),
    "ansichten": (u"Ansichten in den Verknüpfungen",
                  u"Views in the links",
                  u"Vistas de los vínculos"),
    "ansicht_tip": (u"Name, Ansichtstyp, Ebene, Ansichtsvorlage oder "
                    u"Verknüpfung - mehrere Wörter: alle müssen vorkommen",
                    u"Name, view type, level, view template or link - "
                    u"several words: all must occur",
                    u"Nombre, tipo de vista, nivel, plantilla o vínculo - "
                    u"varias palabras: deben aparecer todas"),
    "typ": (u"Typ:",
            u"Type:",
            u"Tipo:"),
    "passend": (u"Passend zur aktiven Ansicht",
                u"Matching the active view",
                u"Según la vista activa"),
    "alle": (u"Alle Ansichtstypen",
             u"All view types",
             u"Todos los tipos de vista"),
    "benutzt": (u"Nur bereits verwendete",
                u"Only already used",
                u"Solo las ya usadas"),
    "benutzt_tip": (u"Nur verknüpfte Ansichten, die im Projekt schon in "
                    u"einer Ansicht oder Vorlage eingestellt sind",
                    u"Only linked views that are already set in a view or "
                    u"template of the project",
                    u"Solo vistas vinculadas ya ajustadas en una vista o "
                    u"plantilla del proyecto"),
    "liste_tip": (u"Doppelklick: auf die aktive Ansicht anwenden. "
                  u"Mehrfachauswahl mit Strg - eine Ansicht je Verknüpfung.",
                  u"Double-click: apply to the active view. Multi-select "
                  u"with Ctrl - one view per link.",
                  u"Doble clic: aplicar a la vista activa. Selección "
                  u"múltiple con Ctrl - una vista por vínculo."),
    "aktiv": (u"Aktive Ansicht",
              u"Active view",
              u"Vista activa"),
    "an_aktiv": (u"Auf aktive Ansicht anwenden",
                 u"Apply to active view",
                 u"Aplicar a la vista activa"),
    "an_viele": (u"Auf Ansichten anwenden…",
                 u"Apply to views…",
                 u"Aplicar a vistas…"),
    "weg": (u"Aktive Ansicht: Nach Basisbauteilansicht",
            u"Active view: By host view",
            u"Vista activa: Por vista anfitriona"),
    "weg_tip": (u"Setzt die markierten Verknüpfungen (oder die der "
                u"markierten Ansichten) in der aktiven Ansicht zurück. "
                u"Bei einem Exemplar gilt danach wieder die Einstellung "
                u"des Typs.",
                u"Resets the selected links (or those of the selected "
                u"views) in the active view. For an instance, the type "
                u"setting applies again afterwards.",
                u"Restablece los vínculos marcados (o los de las vistas "
                u"marcadas) en la vista activa. En un ejemplar vuelve a "
                u"regir el ajuste del tipo."),
    "abbrechen": (u"Abbrechen",
                  u"Cancel",
                  u"Cancelar"),
}

XAML = u"""
<Window %s Title="{{titel}}" Width="1100" Height="720"
        MinWidth="780" MinHeight="460" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12"
        ResizeMode="CanResizeWithGrip">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Padding" Value="10,3"/>
      <Setter Property="Margin" Value="0,0,6,0"/>
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
    <Grid.ColumnDefinitions>
      <ColumnDefinition Width="1*" MinWidth="240"/>
      <ColumnDefinition Width="2*" MinWidth="420"/>
    </Grid.ColumnDefinitions>

    <!-- Verknüpfungen -->
    <GroupBox Header="{{links}}" Grid.Column="0" Margin="0,0,8,0">
      <DockPanel>
        <DockPanel DockPanel.Dock="Top" Margin="0,0,0,4">
          <TextBlock Text="{{suchen}}" VerticalAlignment="Center"
                     Margin="0,0,6,0"/>
          <TextBox x:Name="linksuche" Padding="3" ToolTip="{{suche_tip}}"/>
        </DockPanel>
        <TextBlock DockPanel.Dock="Bottom" Foreground="#666"
                   TextWrapping="Wrap" Margin="0,4,0,0"
                   Text="{{links_hinweis}}"/>
        <TextBlock x:Name="linkzaehler" DockPanel.Dock="Bottom"
                   Foreground="#666" Margin="0,4,0,0"/>
        <ListBox x:Name="linkliste" SelectionMode="Extended"/>
      </DockPanel>
    </GroupBox>

    <!-- Ansichten der Verknüpfungen -->
    <GroupBox Header="{{ansichten}}" Grid.Column="1">
      <DockPanel>
        <Grid DockPanel.Dock="Top" Margin="0,0,0,4">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <TextBlock Text="{{suchen}}" VerticalAlignment="Center"
                     Margin="0,0,6,0"/>
          <TextBox x:Name="suche" Grid.Column="1" Grid.ColumnSpan="3"
                   Padding="3" ToolTip="{{ansicht_tip}}"/>
          <TextBlock Grid.Row="1" Text="{{typ}}" VerticalAlignment="Center"
                     Margin="0,4,6,0"/>
          <ComboBox x:Name="typfilter" Grid.Row="1" Grid.Column="1"
                    Margin="0,4,0,0"/>
          <CheckBox x:Name="nurbenutzt" Grid.Row="1" Grid.Column="2"
                    Margin="10,4,0,0" VerticalAlignment="Center"
                    Content="{{benutzt}}" ToolTip="{{benutzt_tip}}"/>
        </Grid>
        <TextBlock x:Name="ansichtzaehler" DockPanel.Dock="Bottom"
                   Foreground="#666" Margin="0,4,0,0"/>
        <ListBox x:Name="ansichtliste" SelectionMode="Extended"
                 HorizontalContentAlignment="Stretch"
                 ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                 VirtualizingPanel.IsVirtualizing="True"
                 ToolTip="{{liste_tip}}"/>
      </DockPanel>
    </GroupBox>

    <!-- Aktive Ansicht -->
    <GroupBox Header="{{aktiv}}" Grid.Row="1" Grid.ColumnSpan="2"
              Margin="0,8,0,0">
      <ScrollViewer MaxHeight="110" VerticalScrollBarVisibility="Auto">
        <StackPanel>
          <TextBlock x:Name="aktivkopf" TextWrapping="Wrap"
                     FontWeight="SemiBold"/>
          <TextBlock x:Name="aktivtext" TextWrapping="Wrap"
                     Margin="0,4,0,0"/>
        </StackPanel>
      </ScrollViewer>
    </GroupBox>

    <!-- Fusszeile -->
    <DockPanel Grid.Row="2" Grid.ColumnSpan="2" Margin="0,10,0,0">
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
        <Button x:Name="ok" Content="OK" Width="90"/>
        <Button x:Name="abbrechen" Content="{{abbrechen}}" Width="90"
                Margin="0" IsCancel="True"/>
      </StackPanel>
      <WrapPanel>
        <Button x:Name="an_aktiv" Content="{{an_aktiv}}"/>
        <Button x:Name="an_viele" Content="{{an_viele}}"/>
        <Button x:Name="weg" Content="{{weg}}" ToolTip="{{weg_tip}}"/>
      </WrapPanel>
    </DockPanel>
  </Grid>
</Window>""" % dlg.XMLNS


def schreibe_fehlerprotokoll(spur):
    try:
        with io.open(FEHLERPROTOKOLL, "a", encoding="utf-8") as datei:
            datei.write(spur + u"\n" + u"-" * 70 + u"\n")
        return FEHLERPROTOKOLL
    except Exception:
        return None


def meldung(besitzer, text, warnung=False):
    dlg.meldung(besitzer, text, titel=TITEL, warnung=warnung)


def frage(besitzer, text, warnung=False):
    return dlg.frage(besitzer, text, titel=TITEL, warnung=warnung)


def sicher(besitzer_liefern, funktion):
    """Ereignishandler mit Fehlerfang - eine Ausnahme im Handler würde
    ShowDialog() sonst wortlos beenden."""
    def handler(sender, args):
        try:
            funktion(sender, args)
        except LinkFehler as fehler:
            meldung(besitzer_liefern(), u"%s" % fehler, warnung=True)
        except Exception as fehler:
            pfad = schreibe_fehlerprotokoll(traceback.format_exc())
            try:
                meldung(besitzer_liefern(), u"%s%s" % (
                    dlg.fehlertext(fehler),
                    t(u"\n\nTechnische Details: %s",
                      u"\n\nTechnical details: %s",
                      u"\n\nDetalles técnicos: %s") % pfad if pfad else u""),
                    warnung=True)
            except Exception:
                pass
    return handler


class LinkAnsichtenFenster(object):

    def __init__(self, uiapp, doc, verknuepfungen=None):
        self.uiapp = uiapp
        self.doc = doc
        self.fenster = dlg.lade_xaml(uebersetze_xaml(XAML, XAML_TEXTE))
        dlg.setze_besitzer(self.fenster, handle=uiapp.MainWindowHandle)
        self.c = self.fenster.FindName

        self.typen = lg.ansichtstypen()
        self.verknuepfungen = (verknuepfungen if verknuepfungen is not None
                               else rv.verknuepfungen(doc))
        self.nach_wert = dict((v.wert, v) for v in self.verknuepfungen)
        self.link_ansicht = {}       # (Verknüpfung, Ansicht) -> LinkAnsicht
        self.nach_tag = {}           # Text-Schlüssel der Listenzeile
        for v in self.verknuepfungen:
            for a in v.ansichten:
                self.link_ansicht[(v.wert, a.wert)] = a
                # Tag als Text: Python-Tupel kämen über pythonnet nicht
                # als Tupel zurück
                self.nach_tag[u"%d:%d" % (v.wert, a.wert)] = a
        self.host = rv.host_ansichten(doc)
        self.verwendung = {}
        self.sitzung_geaendert = False
        self.uebernehmen = False
        self._still = False
        self._verdrahte()

    # ------------------------------------------------------------------
    # Grundgerüst
    # ------------------------------------------------------------------

    def _h(self, funktion):
        return sicher(lambda: self.fenster, funktion)

    def _verdrahte(self):
        c = self.c
        typfilter = c("typfilter")
        namen = lg.gruppen_namen()
        for text in [tt(XAML_TEXTE["passend"]), tt(XAML_TEXTE["alle"])] + \
                [namen[g] for g in lg.GRUPPEN_REIHENFOLGE]:
            typfilter.Items.Add(text)
        typfilter.SelectedIndex = TYP_PASSEND

        c("linksuche").TextChanged += self._h(
            lambda s, a: self._fuelle_links())
        c("linkliste").SelectionChanged += self._h(self._bei_linkauswahl)
        c("suche").TextChanged += self._h(lambda s, a: self._fuelle_ansichten())
        typfilter.SelectionChanged += self._h(
            lambda s, a: self._fuelle_ansichten())
        c("nurbenutzt").Click += self._h(lambda s, a: self._fuelle_ansichten())
        c("ansichtliste").MouseDoubleClick += self._h(self._doppelklick)
        c("an_aktiv").Click += self._h(self._auf_aktive_ansicht)
        c("an_viele").Click += self._h(self._auf_ansichten)
        c("weg").Click += self._h(self._zuruecksetzen)
        c("ok").Click += self._h(self._ok)
        self.fenster.Closing += self._h(self._beim_schliessen)

    def zeige(self):
        gruppe = TransactionGroup(self.doc, t(u"pyMLG Verknüpfte Ansichten",
                                              u"pyMLG Linked Views",
                                              u"pyMLG Vistas vinculadas"))
        gruppe.Start()
        try:
            self._aktualisiere_verwendung()
            self._fuelle_links()
            self._fuelle_ansichten()
            self._zeige_aktive()
            self.fenster.Loaded += lambda s, a: self.c("suche").Focus()
            self.fenster.ShowDialog()
        finally:
            if self.uebernehmen and self.sitzung_geaendert:
                gruppe.Assimilate()
            else:
                gruppe.RollBack()

    def _transaktion(self, name, funktion):
        transaktion = Transaction(self.doc, name)
        transaktion.Start()
        try:
            ergebnis = funktion()
            if transaktion.Commit() != TransactionStatus.Committed:
                raise LinkFehler(t(u"Revit hat die Änderung \"%s\" nicht "
                                   u"übernommen.",
                                   u"Revit did not accept the change \"%s\".",
                                   u"Revit no aceptó el cambio \"%s\".")
                                 % name)
        except Exception:
            if transaktion.HasStarted() and not transaktion.HasEnded():
                transaktion.RollBack()
            raise
        self.sitzung_geaendert = True
        return ergebnis

    def _aktualisiere_verwendung(self):
        self.verwendung = rv.verwendung(self.doc, self.host,
                                        self.verknuepfungen)

    def _aktive_ansicht(self):
        return self.doc.ActiveView

    def _ansichtsname(self, ansicht):
        if ansicht.IsTemplate:
            return t(u"Vorlage \"%s\"", u"template \"%s\"",
                     u"plantilla \"%s\"") % ansicht.Name
        return u"%s: %s" % (self.typen.get(u"%s" % ansicht.ViewType,
                                           u"%s" % ansicht.ViewType),
                            ansicht.Name)

    # ------------------------------------------------------------------
    # Listen
    # ------------------------------------------------------------------

    def _fuelle_links(self):
        woerter = lg.suchwoerter(self.c("linksuche").Text)
        liste = self.c("linkliste")
        markiert = set(self._markierte_link_werte())
        passend = [v for v in self.verknuepfungen
                   if lg.trifft((v.name + u" " + v.typ_name).lower(),
                                woerter)]
        self._still = True
        try:
            liste.Items.Clear()
            for v in passend:
                item = ListBoxItem()
                item.Content = (u"    └ " + v.name) if v.exemplar else v.name
                item.Tag = v.wert
                if not v.geladen:
                    item.Foreground = ROT
                    item.ToolTip = t(u"Nicht geladen - keine Ansichten "
                                     u"lesbar.",
                                     u"Not loaded - views cannot be read.",
                                     u"No cargado - no se pueden leer las "
                                     u"vistas.")
                elif v.exemplar:
                    item.ToolTip = t(u"Einzelnes Exemplar - überschreibt in "
                                     u"der Ansicht die Einstellung des Typs.",
                                     u"Single instance - overrides the type "
                                     u"setting in the view.",
                                     u"Ejemplar individual - sobrescribe el "
                                     u"ajuste del tipo en la vista.")
                else:
                    item.ToolTip = t(u"%d Ansichten", u"%d views",
                                     u"%d vistas") % len(v.ansichten)
                liste.Items.Add(item)
                if v.wert in markiert:
                    item.IsSelected = True
        finally:
            self._still = False
        self.c("linkzaehler").Text = t(u"%d von %d Verknüpfungen",
                                       u"%d of %d links",
                                       u"%d de %d vínculos") % (
            len(passend), len(self.verknuepfungen))

    def _markierte_link_werte(self):
        return [int(item.Tag) for item in self.c("linkliste").SelectedItems]

    def _bei_linkauswahl(self, sender, args):
        if self._still:
            return
        self._fuelle_ansichten()
        self._zeige_aktive()

    def _gezeigte_verknuepfungen(self):
        werte = self._markierte_link_werte()
        if werte:
            return [self.nach_wert[w] for w in werte]
        # ohne Markierung: alle Typen (Exemplare hätten dieselben Ansichten)
        return [v for v in self.verknuepfungen if not v.exemplar]

    def _gruppe_filter(self):
        index = self.c("typfilter").SelectedIndex
        if index == TYP_PASSEND:
            return lg.gruppe(self._aktive_ansicht().ViewType)
        if index == TYP_ALLE or index < 0:
            return None
        return lg.GRUPPEN_REIHENFOLGE[index - 2]

    def _fuelle_ansichten(self):
        woerter = lg.suchwoerter(self.c("suche").Text)
        gruppe = self._gruppe_filter()
        nur_benutzt = bool(self.c("nurbenutzt").IsChecked)
        liste = self.c("ansichtliste")
        markiert = set(self._markierte_tags())
        aktiv = self._aktiv_gesetzt()

        gesamt = 0
        passend = []
        for v in self._gezeigte_verknuepfungen():
            for a in v.ansichten:
                gesamt += 1
                if gruppe is not None and a.gruppe != gruppe:
                    continue
                if nur_benutzt and not self.verwendung.get((v.wert, a.wert)):
                    continue
                if lg.trifft(a.suchtext, woerter):
                    passend.append(a)

        self._still = True
        try:
            liste.Items.Clear()
            for a in passend:
                schluessel = (a.verknuepfung.wert, a.wert)
                item = ListBoxItem()
                item.Content = self._zeile(a, schluessel in aktiv)
                tag = u"%d:%d" % schluessel
                item.Tag = tag
                benutzt = self.verwendung.get(schluessel, [])
                if benutzt:
                    item.ToolTip = t(u"Eingestellt in:\n", u"Set in:\n",
                                     u"Ajustada en:\n") + u"\n".join(
                        sorted(benutzt)[:25]) + (
                        u"\n…" if len(benutzt) > 25 else u"")
                liste.Items.Add(item)
                if tag in markiert:
                    item.IsSelected = True
        finally:
            self._still = False
        self.c("ansichtzaehler").Text = t(
            u"%d von %d Ansichten  ·  ● = in der aktiven Ansicht eingestellt",
            u"%d of %d views  ·  ● = set in the active view",
            u"%d de %d vistas  ·  ● = ajustada en la vista activa") % (
            len(passend), gesamt)

    def _zeile(self, link_ansicht, aktiv):
        raster = Grid()
        for breite in (None, GridLength.Auto):
            spalte = ColumnDefinition()
            spalte.Width = breite if breite is not None \
                else GridLength(1.0, GridUnitType.Star)
            raster.ColumnDefinitions.Add(spalte)

        links = StackPanel()
        kopf = StackPanel()
        kopf.Orientation = Orientation.Horizontal
        punkt = TextBlock()
        punkt.Text = u"● " if aktiv else u"   "
        punkt.Foreground = GRUEN
        kopf.Children.Add(punkt)
        name = TextBlock()
        name.Text = link_ansicht.text
        name.TextTrimming = TextTrimming.CharacterEllipsis
        if aktiv:
            name.FontWeight = FontWeights.SemiBold
            name.Foreground = GRUEN
        kopf.Children.Add(name)
        links.Children.Add(kopf)

        teile = []
        if link_ansicht.ebene:
            teile.append(t(u"Ebene: %s", u"Level: %s", u"Nivel: %s")
                         % link_ansicht.ebene)
        if link_ansicht.vorlage:
            teile.append(t(u"Vorlage: %s", u"Template: %s",
                           u"Plantilla: %s") % link_ansicht.vorlage)
        anzahl = len(self.verwendung.get(
            (link_ansicht.verknuepfung.wert, link_ansicht.wert), []))
        if anzahl:
            teile.append(t(u"in %d Ansicht(en) eingestellt",
                           u"set in %d view(s)",
                           u"ajustada en %d vista(s)") % anzahl)
        if teile:
            unter = TextBlock()
            unter.Text = u"   " + u"  ·  ".join(teile)
            unter.Foreground = GRAU
            unter.FontSize = 11.0
            unter.TextTrimming = TextTrimming.CharacterEllipsis
            links.Children.Add(unter)
        raster.Children.Add(links)

        link = TextBlock()
        link.Text = link_ansicht.verknuepfung.name
        link.Foreground = GRAU
        link.Margin = Thickness(12.0, 0.0, 0.0, 0.0)
        Grid.SetColumn(link, 1)
        raster.Children.Add(link)
        return raster

    def _markierte_tags(self):
        return [u"%s" % item.Tag
                for item in self.c("ansichtliste").SelectedItems]

    def _markierte_link_ansichten(self):
        ergebnis = []
        for tag in self._markierte_tags():
            a = self.nach_tag.get(tag)
            if a is not None:
                ergebnis.append(a)
        return ergebnis

    # ------------------------------------------------------------------
    # Aktive Ansicht
    # ------------------------------------------------------------------

    def _ziel_der_aktiven(self):
        """(Ansicht, Vorlage) - Vorlage, wenn sie die RVT-Verknüpfungen der
        aktiven Ansicht steuert."""
        ansicht = self._aktive_ansicht()
        return ansicht, rv.steuernde_vorlage(self.doc, ansicht)

    def _aktiv_gesetzt(self):
        """Schlüssel (Verknüpfung, Ansicht), die in der aktiven Ansicht
        (bzw. ihrer steuernden Vorlage) gerade eingestellt sind."""
        ansicht, vorlage = self._ziel_der_aktiven()
        quelle = vorlage or ansicht
        if not rv.unterstuetzt(quelle):
            return set()
        ergebnis = set()
        for v in self.verknuepfungen:
            art, wert, _geerbt = rv.zustand(quelle, v)
            if art == rv.LINK and wert is not None:
                ergebnis.add((v.wert, wert))
        return ergebnis

    def _zustandstext(self, quelle, v):
        art, wert, geerbt = rv.zustand(quelle, v)
        if art == rv.LINK:
            a = self.link_ansicht.get((v.wert, wert))
            text = t(u"Nach verknüpfter Ansicht \"%s\"",
                     u"By linked view \"%s\"",
                     u"Por vista vinculada \"%s\"") % (
                a.text if a is not None else u"?")
        elif art == rv.EIGEN:
            text = t(u"Benutzerdefiniert", u"Custom", u"Personalizado")
        else:
            text = t(u"Nach Basisbauteilansicht", u"By host view",
                     u"Por vista anfitriona")
        if geerbt:
            text += t(u" (vom Typ)", u" (from type)", u" (del tipo)")
        return u"%s: %s" % (v.name, text)

    def _zeige_aktive(self):
        ansicht, vorlage = self._ziel_der_aktiven()
        kopf = self._ansichtsname(ansicht)
        if vorlage is not None:
            kopf += t(u"   →  RVT-Verknüpfungen gesteuert von Vorlage \"%s\" "
                      u"(Änderungen gehen an die Vorlage)",
                      u"   →  RVT links controlled by template \"%s\" "
                      u"(changes go to the template)",
                      u"   →  vínculos RVT controlados por la plantilla "
                      u"\"%s\" (los cambios van a la plantilla)") % \
                vorlage.Name
        quelle = vorlage or ansicht
        unterstuetzt = rv.unterstuetzt(quelle)
        self.c("aktivkopf").Text = kopf
        self.c("an_aktiv").IsEnabled = unterstuetzt
        self.c("weg").IsEnabled = unterstuetzt
        if not unterstuetzt:
            self.c("aktivtext").Text = t(
                u"In dieser Ansicht lässt sich keine verknüpfte Ansicht "
                u"einstellen (nur Grundrisse, Schnitte, Ansichten, 3D).",
                u"No linked view can be set in this view (plans, sections, "
                u"elevations and 3D only).",
                u"En esta vista no se puede ajustar una vista vinculada "
                u"(solo plantas, secciones, alzados y 3D).")
            return
        werte = self._markierte_link_werte()
        gezeigt = ([self.nach_wert[w] for w in werte] if werte
                   else self.verknuepfungen)
        self.c("aktivtext").Text = u"\n".join(
            self._zustandstext(quelle, v) for v in gezeigt)

    # ------------------------------------------------------------------
    # Anwenden
    # ------------------------------------------------------------------

    def _auswahl_pruefen(self):
        auswahl = self._markierte_link_ansichten()
        if not auswahl:
            raise LinkFehler(t(u"Bitte rechts mindestens eine verknüpfte "
                               u"Ansicht markieren.",
                               u"Please select at least one linked view on "
                               u"the right.",
                               u"Marque al menos una vista vinculada a la "
                               u"derecha."))
        doppelt = lg.doppelte([(a.verknuepfung.wert, a.verknuepfung.name)
                               for a in auswahl])
        if doppelt:
            raise LinkFehler(t(u"Pro Verknüpfung lässt sich nur eine Ansicht "
                               u"einstellen. Mehrfach markiert: %s",
                               u"Only one view can be set per link. Selected "
                               u"more than once: %s",
                               u"Solo se puede ajustar una vista por "
                               u"vínculo. Marcado varias veces: %s")
                             % u", ".join(doppelt))
        return auswahl

    def _anwenden(self, ziele, auswahl):
        """Rückgabe: (Anzahl ok, [Hinweise], [Fehler])"""
        fehler = []
        hinweise = []

        def ausfuehren():
            zaehler = 0
            for ziel in ziele:
                for a in auswahl:
                    v = a.verknuepfung
                    if not lg.passt(ziel.ViewType, a.typ):
                        fehler.append(t(
                            u"%s / %s: \"%s\" passt nicht zum Ansichtstyp",
                            u"%s / %s: \"%s\" does not match the view type",
                            u"%s / %s: \"%s\" no corresponde al tipo de "
                            u"vista") % (ziel.Name, v.name, a.text))
                        continue
                    try:
                        rv.setze(ziel, v, a)
                        zaehler += 1
                    except Exception as ausnahme:
                        fehler.append(u"%s / %s: %s" % (
                            ziel.Name, v.name, dlg.fehlertext(ausnahme)))
                        continue
                    if not v.exemplar:
                        eigene = rv.exemplare_mit_eigener(
                            ziel, self.verknuepfungen, v)
                        if eigene:
                            hinweise.append(t(
                                u"%s: eigene Einstellung bleibt bei %s",
                                u"%s: own setting stays for %s",
                                u"%s: se mantiene el ajuste propio de %s")
                                % (ziel.Name, u", ".join(eigene)))
            return zaehler

        erfolgreich = self._transaktion(t(u"Verknüpfte Ansicht setzen",
                                          u"Set linked view",
                                          u"Ajustar vista vinculada"),
                                        ausfuehren)
        return erfolgreich, hinweise, fehler

    def _bericht(self, erfolgreich, hinweise, fehler):
        text = t(u"%d Einstellung(en) gesetzt.", u"%d setting(s) applied.",
                 u"%d ajuste(s) aplicado(s).") % erfolgreich
        for titel, zeilen in (
                (t(u"Hinweis", u"Note", u"Nota"), hinweise),
                (t(u"Nicht möglich", u"Not possible", u"No es posible"),
                 fehler)):
            if zeilen:
                text += u"\n\n%s (%d):\n" % (titel, len(zeilen))
                text += u"\n".join(u"• " + z for z in zeilen[:20])
                if len(zeilen) > 20:
                    text += t(u"\n… und %d weitere", u"\n… and %d more",
                              u"\n… y %d más") % (len(zeilen) - 20)
        meldung(self.fenster, text, warnung=bool(fehler))

    def _nach_aenderung(self):
        self._aktualisiere_verwendung()
        self._fuelle_ansichten()
        self._zeige_aktive()

    def _doppelklick(self, sender, args):
        if self.c("ansichtliste").SelectedItem is not None \
                and self.c("an_aktiv").IsEnabled:
            self._auf_aktive_ansicht(sender, args)

    def _auf_aktive_ansicht(self, sender, args):
        auswahl = self._auswahl_pruefen()
        ansicht, vorlage = self._ziel_der_aktiven()
        ziel = ansicht
        if vorlage is not None:
            if not frage(self.fenster, t(
                    u"Die RVT-Verknüpfungen der aktiven Ansicht werden von "
                    u"der Ansichtsvorlage \"%s\" gesteuert.\n\nStattdessen "
                    u"die Vorlage ändern? (Wirkt auf alle Ansichten mit "
                    u"dieser Vorlage.)",
                    u"The RVT links of the active view are controlled by the "
                    u"view template \"%s\".\n\nChange the template instead? "
                    u"(Affects all views using this template.)",
                    u"Los vínculos RVT de la vista activa los controla la "
                    u"plantilla de vista \"%s\".\n\n¿Cambiar la plantilla en "
                    u"su lugar? (Afecta a todas las vistas con esta "
                    u"plantilla.)") % vorlage.Name):
                return
            ziel = vorlage
        if not rv.unterstuetzt(ziel):
            raise LinkFehler(t(u"In \"%s\" lässt sich keine verknüpfte "
                               u"Ansicht einstellen.",
                               u"No linked view can be set in \"%s\".",
                               u"En \"%s\" no se puede ajustar una vista "
                               u"vinculada.") % ziel.Name)
        erfolgreich, hinweise, fehler = self._anwenden([ziel], auswahl)
        self._nach_aenderung()
        if fehler or hinweise:
            self._bericht(erfolgreich, hinweise, fehler)

    def _auf_ansichten(self, sender, args):
        auswahl = self._auswahl_pruefen()
        gruppen = set(a.gruppe for a in auswahl)
        if len(gruppen) > 1:
            raise LinkFehler(t(u"Die markierten Ansichten sind verschiedener "
                               u"Art (z.B. Grundriss und Schnitt). Bitte nur "
                               u"eine Art markieren.",
                               u"The selected views are of different kinds "
                               u"(e.g. plan and section). Please select one "
                               u"kind only.",
                               u"Las vistas marcadas son de distinto tipo "
                               u"(p. ej. planta y sección). Marque solo un "
                               u"tipo."))
        gruppe = gruppen.pop()
        einzeln = auswahl[0] if len(auswahl) == 1 else None

        eintraege = []
        for ansicht in self.host:
            if lg.gruppe(ansicht.ViewType) != gruppe:
                continue
            text = self._ansichtsname(ansicht)
            vorlage = rv.steuernde_vorlage(self.doc, ansicht)
            if vorlage is not None:
                text += t(u"   [über Vorlage \"%s\"]",
                          u"   [via template \"%s\"]",
                          u"   [vía plantilla \"%s\"]") % vorlage.Name
            elif einzeln is not None:
                art, wert, _g = rv.zustand(ansicht, einzeln.verknuepfung)
                if art == rv.LINK and wert == einzeln.wert:
                    text += t(u"   [schon eingestellt]", u"   [already set]",
                              u"   [ya ajustada]")
            eintraege.append((text, id_wert(ansicht.Id)))
        eintraege.sort(key=lambda e: e[0].lower())
        if not eintraege:
            raise LinkFehler(t(u"Im Projekt gibt es keine passenden "
                               u"Ansichten.",
                               u"There are no matching views in the "
                               u"project.",
                               u"No hay vistas adecuadas en el proyecto."))

        ergebnis = dlg.waehle(
            self.fenster, tt(XAML_TEXTE["an_viele"]),
            t(u"Ansichten und Ansichtsvorlagen wählen, in denen %s "
              u"eingestellt werden soll. Steuert die Vorlage einer Ansicht "
              u"die RVT-Verknüpfungen, wird die Vorlage geändert.",
              u"Choose the views and view templates in which %s should be "
              u"set. If a view's template controls the RVT links, the "
              u"template is changed.",
              u"Elija las vistas y plantillas en las que se ajustará %s. Si "
              u"la plantilla de una vista controla los vínculos RVT, se "
              u"cambia la plantilla.") % u", ".join(
                u"\"%s\" (%s)" % (a.text, a.verknuepfung.name)
                for a in auswahl),
            eintraege, mehrfach=True)
        if ergebnis is None:
            return
        gewaehlt = [a for a in self.host if id_wert(a.Id) in set(ergebnis)]
        ziele, umgeleitet = lg.ziele(
            gewaehlt, lambda a: rv.steuernde_vorlage(self.doc, a),
            lambda e: id_wert(e.Id))
        if umgeleitet:
            zeilen = [u"• %s: %s" % (vorlage.Name, lg.kurzliste(
                a.Name for a in ansichten))
                for vorlage, ansichten in umgeleitet.values()]
            if not frage(self.fenster, t(
                    u"Bei diesen Ansichten steuert die Ansichtsvorlage die "
                    u"RVT-Verknüpfungen - geändert wird deshalb die Vorlage "
                    u"(wirkt auf alle Ansichten mit dieser Vorlage):\n\n%s"
                    u"\n\nFortfahren?",
                    u"For these views the view template controls the RVT "
                    u"links - so the template is changed (affects all views "
                    u"using it):\n\n%s\n\nContinue?",
                    u"En estas vistas la plantilla controla los vínculos "
                    u"RVT, así que se cambia la plantilla (afecta a todas "
                    u"las vistas que la usan):\n\n%s\n\n¿Continuar?")
                    % u"\n".join(zeilen)):
                return
        erfolgreich, hinweise, fehler = self._anwenden(ziele, auswahl)
        self._nach_aenderung()
        self._bericht(erfolgreich, hinweise, fehler)

    def _zuruecksetzen(self, sender, args):
        werte = self._markierte_link_werte()
        if werte:
            betroffen = [self.nach_wert[w] for w in werte]
        else:
            betroffen = []
            for a in self._markierte_link_ansichten():
                if a.verknuepfung not in betroffen:
                    betroffen.append(a.verknuepfung)
        if not betroffen:
            raise LinkFehler(t(u"Bitte links eine Verknüpfung oder rechts "
                               u"eine Ansicht markieren.",
                               u"Please select a link on the left or a view "
                               u"on the right.",
                               u"Marque un vínculo a la izquierda o una "
                               u"vista a la derecha."))
        ansicht, vorlage = self._ziel_der_aktiven()
        ziel = vorlage or ansicht
        if not frage(self.fenster, t(
                u"In %s auf \"Nach Basisbauteilansicht\" zurücksetzen:\n%s",
                u"Reset to \"By host view\" in %s:\n%s",
                u"Restablecer a \"Por vista anfitriona\" en %s:\n%s") % (
                self._ansichtsname(ziel),
                u"\n".join(u"• " + v.name for v in betroffen))):
            return

        def ausfuehren():
            for v in betroffen:
                rv.zuruecksetzen(ziel, v)
        self._transaktion(t(u"Verknüpfte Ansicht zurücksetzen",
                            u"Reset linked view",
                            u"Restablecer vista vinculada"), ausfuehren)
        self._nach_aenderung()

    # ------------------------------------------------------------------
    # Schliessen
    # ------------------------------------------------------------------

    def _ok(self, sender, args):
        self.uebernehmen = True
        self.fenster.Close()

    def _beim_schliessen(self, sender, args):
        if self.uebernehmen or not self.sitzung_geaendert:
            return
        if not frage(self.fenster, t(u"Alle Änderungen dieser Sitzung "
                                     u"verwerfen?",
                                     u"Discard all changes of this session?",
                                     u"¿Descartar todos los cambios de esta "
                                     u"sesión?"), warnung=True):
            args.Cancel = True


def starte(uiapp, doc, verknuepfungen=None):
    LinkAnsichtenFenster(uiapp, doc, verknuepfungen).zeige()
