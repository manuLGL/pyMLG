# -*- coding: utf-8 -*-
"""Dialog von ComponentLegend (WPF): Ansichten wählen, daraus gelesene
Kategorien wählen, Reihenfolge prüfen, Massstab und Abstand einstellen.

    ┌ Ansichten ──────────┬ Kategorien ─────────┬ Reihenfolge ───────────┐
    │ Suchen [        ]   │ ☑ Türen (7 Typen)   │ Türen                  │
    │ ☑ EG (Grundriss)    │ ☐ Fenster (4 Typen) │   Tür 1-flg: T1 (12)   │
    │ ☐ OG (Grundriss)    │                     │   Tür 1-flg: T2 (3)    │
    │ [Alle] [Keine]      │ [Alle] [Keine]      │                        │
    ├ Einstellungen ──────┴─────────────────────┴────────────────────────┤
    │ Name [Legende Türen]  Massstab 1:[50]  Abstand [10] mm (Papier)    │
    │ Beschriftung [Familie: Typ ▾]  Texttyp [2.5mm Arial ▾]  ☐ ersetzen │
    └────────────────────────────────── 7 Bauteile [Legende erstellen] [Abbrechen]

Die Kategorien werden beim Markieren einer Ansicht gelesen (je Ansicht
einmal, dann zwischengespeichert).
"""

import io
import json
import os
import traceback

from System.Windows import FontWeights, Thickness, Visibility
from System.Windows.Controls import CheckBox, ComboBoxItem, TextBlock
from System.Windows.Input import Cursors, Mouse

from bauteil_legende import logik as lg
from bauteil_legende import revit as rv
from dxf_legende import logik as dxf_lg
from filter_manager import dialoge as dlg
from mlg_sprache import t, tt, uebersetze_xaml

TITEL = t(u"Legende aus Ansichten", u"Legend from views", u"Leyenda desde vistas")

FEHLERPROTOKOLL = os.path.join(dlg.protokollordner(), "ComponentLegend_Fehler.log")
EINSTELLUNGEN = os.path.join(dlg.protokollordner(), "ComponentLegend.json")

XAML_TEXTE = {
    "titel": (u"Legende aus Ansichten (pyMLG)", u"Legend from views (pyMLG)",
              u"Leyenda desde vistas (pyMLG)"),
    "ansichten": (u"1. Ansichten", u"1. Views", u"1. Vistas"),
    "kategorien": (u"2. Kategorien", u"2. Categories", u"2. Categorías"),
    "reihenfolge": (u"Reihenfolge in der Legende", u"Order in the legend",
                    u"Orden en la leyenda"),
    "suchen": (u"Suchen:", u"Search:", u"Buscar:"),
    "alle": (u"Alle", u"All", u"Todas"),
    "keine": (u"Keine", u"None", u"Ninguna"),
    "einstellungen": (u"3. Einstellungen", u"3. Settings", u"3. Ajustes"),
    "name": (u"Name:", u"Name:", u"Nombre:"),
    "massstab": (u"Massstab 1:", u"Scale 1:", u"Escala 1:"),
    "abstand": (u"Abstand (mm auf dem Plan):", u"Spacing (mm on sheet):",
                u"Separación (mm en plano):"),
    "abstand_tip": (u"Lücke zwischen Unterkante eines Bauteils und Oberkante des "
                    u"nächsten, gemessen auf dem Papier.",
                    u"Gap between the bottom edge of one component and the top "
                    u"edge of the next, measured on paper.",
                    u"Hueco entre el borde inferior de un componente y el borde "
                    u"superior del siguiente, medido en el papel."),
    "beschriftung": (u"Beschriftung:", u"Label:", u"Etiqueta:"),
    "texttyp": (u"Texttyp:", u"Text type:", u"Tipo de texto:"),
    "ersetzen": (u"Gleichnamige Legende ersetzen", u"Replace legend with the same name",
                 u"Reemplazar leyenda del mismo nombre"),
    "ersetzen_tip": (u"Gibt es schon eine Legende mit dem Namen, wird ihr Inhalt "
                     u"gelöscht und neu aufgebaut. Auf Plänen bleibt sie an ihrem "
                     u"Platz. Ohne Haken entsteht \"Name (2)\".",
                     u"If a legend with this name exists, its content is deleted "
                     u"and rebuilt. It stays in place on sheets. Unticked, "
                     u"\"Name (2)\" is created.",
                     u"Si ya existe una leyenda con ese nombre, se borra su "
                     u"contenido y se vuelve a crear. En los planos se queda en "
                     u"su sitio. Sin marcar se crea \"Nombre (2)\"."),
    "erstellen": (u"Legende erstellen", u"Create legend", u"Crear leyenda"),
    "abbrechen": (u"Abbrechen", u"Cancel", u"Cancelar"),
}

BESCHRIFTUNGEN = (
    (lg.BESCHRIFTUNG_KEINE, (u"keine", u"none", u"ninguna")),
    (lg.BESCHRIFTUNG_TYP, (u"Typ", u"Type", u"Tipo")),
    (lg.BESCHRIFTUNG_FAMILIE_TYP, (u"Familie: Typ", u"Family: Type", u"Familia: Tipo")),
)

XAML = u"""
<Window %s Title="{{titel}}" Width="1180" Height="720"
        MinWidth="860" MinHeight="480" WindowStartupLocation="CenterOwner"
        ShowInTaskbar="False" FontFamily="Segoe UI" FontSize="12"
        ResizeMode="CanResizeWithGrip">
  <Window.Resources>
    <Style TargetType="TextBox">
      <Setter Property="Padding" Value="3"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
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

    <Grid>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="1.2*"/>
      </Grid.ColumnDefinitions>

      <GroupBox Header="{{ansichten}}" Margin="0,0,8,0">
        <DockPanel>
          <DockPanel DockPanel.Dock="Top" Margin="0,0,0,6">
            <TextBlock Text="{{suchen}}" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBox x:Name="suche"/>
          </DockPanel>
          <DockPanel DockPanel.Dock="Bottom" Margin="0,6,0,0">
            <Button x:Name="ansichten_alle" Content="{{alle}}" Padding="10,2"
                    Margin="0,0,6,0"/>
            <Button x:Name="ansichten_keine" Content="{{keine}}" Padding="10,2"/>
            <TextBlock x:Name="ansichten_zahl" Foreground="#666" Margin="10,0,0,0"
                       VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
          </DockPanel>
          <ListBox x:Name="ansichtliste"/>
        </DockPanel>
      </GroupBox>

      <GroupBox Header="{{kategorien}}" Grid.Column="1" Margin="0,0,8,0">
        <DockPanel>
          <DockPanel DockPanel.Dock="Bottom" Margin="0,6,0,0" LastChildFill="False">
            <Button x:Name="kategorien_alle" Content="{{alle}}" Padding="10,2"
                    Margin="0,0,6,0"/>
            <Button x:Name="kategorien_keine" Content="{{keine}}" Padding="10,2"/>
          </DockPanel>
          <TextBlock x:Name="kategorien_leer" DockPanel.Dock="Top" Foreground="#666"
                     TextWrapping="Wrap" Margin="2,0,2,6"/>
          <ListBox x:Name="kategorieliste"/>
        </DockPanel>
      </GroupBox>

      <GroupBox Header="{{reihenfolge}}" Grid.Column="2">
        <ListBox x:Name="vorschau" BorderThickness="0"
                 ScrollViewer.HorizontalScrollBarVisibility="Auto"/>
      </GroupBox>
    </Grid>

    <GroupBox Header="{{einstellungen}}" Grid.Row="1" Margin="0,8,0,0">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="70"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="70"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Text="{{name}}" VerticalAlignment="Center" Margin="0,0,6,0"/>
        <TextBox x:Name="name" Grid.Column="1"/>
        <TextBlock Text="{{massstab}}" Grid.Column="2" VerticalAlignment="Center"
                   Margin="14,0,6,0"/>
        <TextBox x:Name="massstab" Grid.Column="3"/>
        <TextBlock Text="{{abstand}}" Grid.Column="4" VerticalAlignment="Center"
                   Margin="14,0,6,0" ToolTip="{{abstand_tip}}"/>
        <TextBox x:Name="abstand" Grid.Column="5" ToolTip="{{abstand_tip}}"/>

        <StackPanel Grid.Row="1" Grid.ColumnSpan="6" Orientation="Horizontal"
                    Margin="0,8,0,0">
          <TextBlock Text="{{beschriftung}}" VerticalAlignment="Center" Margin="0,0,6,0"/>
          <ComboBox x:Name="beschriftung" Width="140"/>
          <TextBlock Text="{{texttyp}}" VerticalAlignment="Center" Margin="14,0,6,0"/>
          <ComboBox x:Name="texttyp" Width="240"/>
          <CheckBox x:Name="ersetzen" Content="{{ersetzen}}" ToolTip="{{ersetzen_tip}}"
                    VerticalAlignment="Center" Margin="24,0,0,0"/>
        </StackPanel>
        <TextBlock x:Name="namehinweis" Grid.Row="2" Grid.ColumnSpan="6"
                   Foreground="#A05A00" TextWrapping="Wrap" Margin="0,6,0,0"/>
      </Grid>
    </GroupBox>

    <DockPanel Grid.Row="2" Margin="0,10,0,0">
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
        <Button x:Name="ok" Content="{{erstellen}}" Padding="14,3"
                Margin="0,0,6,0" IsDefault="True"/>
        <Button x:Name="abbrechen" Content="{{abbrechen}}" Width="90"
                IsCancel="True"/>
      </StackPanel>
      <TextBlock x:Name="zusammenfassung" VerticalAlignment="Center"
                 TextWrapping="Wrap"/>
    </DockPanel>
  </Grid>
</Window>""" % dlg.XMLNS


class EingabeFehler(Exception):
    """Fachlicher Fehler - nur als Text anzeigen."""


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
    """Ereignishandler mit Fehlerfang - eine Ausnahme im Handler würde
    ShowDialog() sonst wortlos beenden."""
    def handler(sender, args):
        try:
            funktion(sender, args)
        except EingabeFehler as fehler:
            meldung(besitzer_liefern(), u"%s" % fehler, warnung=True)
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


def lies_einstellungen():
    try:
        with io.open(EINSTELLUNGEN, encoding="utf-8") as datei:
            return json.loads(datei.read())
    except Exception:
        return {}


def schreibe_einstellungen(werte):
    try:
        with io.open(EINSTELLUNGEN, "w", encoding="utf-8") as datei:
            datei.write(json.dumps(werte, ensure_ascii=False, indent=2))
    except Exception:
        pass


def _box(text, markiert=False):
    box = CheckBox()
    box.Content = text
    box.IsChecked = markiert
    box.Margin = Thickness(2.0, 2.0, 2.0, 2.0)
    return box


class LegendenFenster(object):

    def __init__(self, uiapp, doc, aktive_ansicht):
        self.doc = doc
        self.einstellungen = lies_einstellungen()
        self.fenster = dlg.lade_xaml(uebersetze_xaml(XAML, XAML_TEXTE))
        dlg.setze_besitzer(self.fenster, handle=uiapp.MainWindowHandle)
        self.c = self.fenster.FindName
        self.ergebnis = None
        self._funde = {}          # Ansicht-Id (Zahl) -> Funde
        self._typen = {}
        self._kat_markiert = set(self.einstellungen.get(u"kategorien", []))
        self._kat_boxen = []      # (kat_id, CheckBox)
        self._reihenfolge = []
        self._ansicht_boxen = []  # (Suchtext, CheckBox, Ansicht)
        self._fuelle_ansichten(aktive_ansicht)
        self._fuelle_einstellungen()
        self._verdrahte()
        self._aktualisiere()

    def _h(self, funktion):
        return sicher(lambda: self.fenster, funktion)

    def zeige(self):
        self.fenster.ShowDialog()
        return self.ergebnis

    # -- Aufbau ------------------------------------------------------------------

    def _fuelle_ansichten(self, aktive_ansicht):
        liste = self.c("ansichtliste")
        aktiv_id = aktive_ansicht.Id if rv.ist_lesbar(aktive_ansicht) else None
        for ansicht in rv.ansichten(self.doc):
            text = u"%s  (%s)" % (rv.name_von(ansicht), tt(rv.ANSICHTSARTEN[ansicht.ViewType]))
            box = _box(text, ansicht.Id == aktiv_id)
            box.Click += self._h(lambda s, a: self._aktualisiere())
            liste.Items.Add(box)
            self._ansicht_boxen.append((text.lower(), box, ansicht))

    def _fuelle_einstellungen(self):
        c = self.c
        e = self.einstellungen
        c("name").Text = e.get(u"name") or t(u"Legende Bauteile", u"Component legend",
                                             u"Leyenda de componentes")
        c("massstab").Text = u"%d" % e.get(u"massstab", 50)
        c("abstand").Text = dxf_lg.zahl_text(e.get(u"abstand", 10.0))
        art = c("beschriftung")
        for _wert, texte in BESCHRIFTUNGEN:
            art.Items.Add(tt(texte))
        werte = [w for w, _ in BESCHRIFTUNGEN]
        gespeichert = e.get(u"beschriftung", lg.BESCHRIFTUNG_FAMILIE_TYP)
        art.SelectedIndex = werte.index(gespeichert) if gespeichert in werte else 0

        texttyp = c("texttyp")
        self._texttypen, vorgabe = rv.texttypen(self.doc)
        gewuenscht = e.get(u"texttyp")
        start = 0
        for index, (name, typ_id) in enumerate(self._texttypen):
            eintrag = ComboBoxItem()
            eintrag.Content = name
            texttyp.Items.Add(eintrag)
            if name == gewuenscht or (not gewuenscht and typ_id == vorgabe):
                start = index
        if self._texttypen:
            texttyp.SelectedIndex = start
        c("ersetzen").IsChecked = bool(e.get(u"ersetzen", False))
        self._beschriftung_geaendert()

    def _verdrahte(self):
        c = self.c
        c("suche").TextChanged += self._h(self._suche)
        c("ansichten_alle").Click += self._h(lambda s, a: self._markiere_ansichten(True))
        c("ansichten_keine").Click += self._h(lambda s, a: self._markiere_ansichten(False))
        c("kategorien_alle").Click += self._h(lambda s, a: self._markiere_kategorien(True))
        c("kategorien_keine").Click += self._h(lambda s, a: self._markiere_kategorien(False))
        c("beschriftung").SelectionChanged += self._h(
            lambda s, a: self._beschriftung_geaendert())
        c("name").TextChanged += self._h(lambda s, a: self._namehinweis())
        c("ersetzen").Click += self._h(lambda s, a: self._namehinweis())
        c("ok").Click += self._h(self._ok)
        self.fenster.Loaded += self._h(lambda s, a: c("suche").Focus())

    # -- Ansichten ---------------------------------------------------------------

    def _suche(self, sender, args):
        woerter = (self.c("suche").Text or u"").lower().split()
        for text, box, _ansicht in self._ansicht_boxen:
            passt = all(w in text for w in woerter)
            box.Visibility = Visibility.Visible if passt else Visibility.Collapsed

    def _markiere_ansichten(self, zustand):
        for _text, box, _ansicht in self._ansicht_boxen:
            if box.Visibility == Visibility.Visible:
                box.IsChecked = zustand
        self._aktualisiere()

    def _gewaehlte_ansichten(self):
        return [a for _t, box, a in self._ansicht_boxen if box.IsChecked]

    def _lies_typen(self):
        """Typen aller markierten Ansichten - neue Ansichten werden gelesen."""
        gewaehlt = self._gewaehlte_ansichten()
        neu = [a for a in gewaehlt if rv.id_wert(a.Id) not in self._funde]
        if neu:
            Mouse.OverrideCursor = Cursors.Wait
            try:
                for ansicht in neu:
                    self._funde[rv.id_wert(ansicht.Id)] = rv.lies_ansicht(self.doc, ansicht)
            finally:
                Mouse.OverrideCursor = None
        typen = {}
        for ansicht in gewaehlt:
            lg.sammle(self._funde[rv.id_wert(ansicht.Id)], typen)
        return typen

    # -- Kategorien und Vorschau -------------------------------------------------

    def _merke_kategorien(self):
        for kat_id, box in self._kat_boxen:
            if box.IsChecked:
                self._kat_markiert.add(kat_id)
            else:
                self._kat_markiert.discard(kat_id)

    def _markiere_kategorien(self, zustand):
        for _kat_id, box in self._kat_boxen:
            box.IsChecked = zustand
        self._merke_kategorien()
        self._zeige_vorschau()

    def _aktualisiere(self):
        self._merke_kategorien()
        self._typen = self._lies_typen()
        liste = self.c("kategorieliste")
        liste.Items.Clear()
        self._kat_boxen = []
        for kat_id, name, anzahl in lg.kategorien(self._typen):
            text = t(u"%s  (%d Typen)", u"%s  (%d types)", u"%s  (%d tipos)") % (name, anzahl) \
                if anzahl != 1 else t(u"%s  (1 Typ)", u"%s  (1 type)", u"%s  (1 tipo)") % name
            box = _box(text, kat_id in self._kat_markiert)
            box.Click += self._h(self._kategorie_geklickt)
            liste.Items.Add(box)
            self._kat_boxen.append((kat_id, box))

        gewaehlt = len(self._gewaehlte_ansichten())
        self.c("ansichten_zahl").Text = t(u"%d markiert", u"%d checked", u"%d marcadas") \
            % gewaehlt
        leer = self.c("kategorien_leer")
        if not gewaehlt:
            leer.Text = t(u"Links mindestens eine Ansicht markieren.",
                          u"Check at least one view on the left.",
                          u"Marque al menos una vista a la izquierda.")
        elif not self._kat_boxen:
            leer.Text = t(u"In den markierten Ansichten ist kein passendes Bauteil sichtbar.",
                          u"No suitable component is visible in the checked views.",
                          u"No hay ningún componente adecuado visible en las vistas "
                          u"marcadas.")
        else:
            leer.Text = u""
        leer.Visibility = Visibility.Visible if leer.Text else Visibility.Collapsed
        self._zeige_vorschau()

    def _kategorie_geklickt(self, sender, args):
        self._merke_kategorien()
        self._zeige_vorschau()

    def _zeige_vorschau(self):
        gewaehlt = [k for k, box in self._kat_boxen if box.IsChecked]
        self._reihenfolge = lg.reihenfolge(self._typen, gewaehlt)
        liste = self.c("vorschau")
        liste.Items.Clear()
        art = self._beschriftungsart()
        kategorie = None
        for typ in self._reihenfolge:
            if typ.kategorie != kategorie:
                kategorie = typ.kategorie
                kopf = TextBlock()
                kopf.Text = kategorie
                kopf.FontWeight = FontWeights.SemiBold
                kopf.Margin = Thickness(0, 6 if liste.Items.Count else 0, 0, 2)
                liste.Items.Add(kopf)
            zeile = TextBlock()
            text = lg.beschriftung(typ, lg.BESCHRIFTUNG_FAMILIE_TYP)
            zeile.Text = u"%s   (%d×)" % (text, typ.anzahl)
            zeile.Margin = Thickness(14, 0, 0, 0)
            if art == lg.BESCHRIFTUNG_KEINE:
                zeile.ToolTip = text
            liste.Items.Add(zeile)
        anzahl = len(self._reihenfolge)
        self.c("zusammenfassung").Text = (
            t(u"%d Legendenbauteile untereinander", u"%d legend components stacked",
              u"%d componentes de leyenda apilados") % anzahl if anzahl else u"")
        self.c("ok").IsEnabled = anzahl > 0

    # -- Einstellungen -----------------------------------------------------------

    def _beschriftungsart(self):
        index = self.c("beschriftung").SelectedIndex
        return BESCHRIFTUNGEN[index][0] if 0 <= index < len(BESCHRIFTUNGEN) \
            else lg.BESCHRIFTUNG_KEINE

    def _beschriftung_geaendert(self):
        self.c("texttyp").IsEnabled = self._beschriftungsart() != lg.BESCHRIFTUNG_KEINE

    def _namehinweis(self):
        name = (self.c("name").Text or u"").strip()
        hinweis = u""
        if name and rv.vorhandene_legende(self.doc, name) is not None:
            if self.c("ersetzen").IsChecked:
                hinweis = t(u"\"%s\" wird geleert und neu aufgebaut.",
                            u"\"%s\" will be cleared and rebuilt.",
                            u"\"%s\" se vaciará y se volverá a crear.") % name
            else:
                hinweis = t(u"\"%s\" gibt es schon - es entsteht \"%s (2)\". Zum "
                            u"Aktualisieren \"Gleichnamige Legende ersetzen\" anhaken.",
                            u"\"%s\" already exists - \"%s (2)\" is created. To update "
                            u"it tick \"Replace legend with the same name\".",
                            u"\"%s\" ya existe - se crea \"%s (2)\". Para actualizarla "
                            u"marque \"Reemplazar leyenda del mismo nombre\".") % (name, name)
        self.c("namehinweis").Text = hinweis
        self.c("namehinweis").Visibility = Visibility.Visible if hinweis \
            else Visibility.Collapsed

    # -- Abschluss ---------------------------------------------------------------

    def _ok(self, sender, args):
        name = (self.c("name").Text or u"").strip()
        if not name:
            raise EingabeFehler(t(u"Bitte einen Namen für die Legende eingeben.",
                                  u"Please enter a name for the legend.",
                                  u"Introduzca un nombre para la leyenda."))
        massstab = lg.pruefe_massstab(self.c("massstab").Text)
        if massstab is None:
            raise EingabeFehler(t(u"Der Massstab muss eine ganze Zahl sein, z.B. 50.",
                                  u"The scale must be a whole number, e.g. 50.",
                                  u"La escala debe ser un número entero, p.ej. 50."))
        abstand = lg.pruefe_abstand(self.c("abstand").Text)
        if abstand is None:
            raise EingabeFehler(t(u"Der Abstand muss zwischen 0 und 500 mm liegen.",
                                  u"The spacing must be between 0 and 500 mm.",
                                  u"La separación debe estar entre 0 y 500 mm."))
        if not self._reihenfolge:
            raise EingabeFehler(t(u"Bitte mindestens eine Kategorie markieren.",
                                  u"Please check at least one category.",
                                  u"Marque al menos una categoría."))
        art = self._beschriftungsart()
        texttyp_name, texttyp_id = (None, None)
        index = self.c("texttyp").SelectedIndex
        if 0 <= index < len(self._texttypen):
            texttyp_name, texttyp_id = self._texttypen[index]
        self._merke_kategorien()
        ersetzen = bool(self.c("ersetzen").IsChecked)
        self.ergebnis = {u"typen": list(self._reihenfolge), u"name": name,
                         u"massstab": massstab, u"abstand": abstand,
                         u"beschriftung": art, u"texttyp": texttyp_id,
                         u"ersetzen": ersetzen}
        schreibe_einstellungen({u"name": name, u"massstab": massstab, u"abstand": abstand,
                                u"beschriftung": art, u"texttyp": texttyp_name,
                                u"ersetzen": ersetzen,
                                u"kategorien": sorted(self._kat_markiert)})
        self.fenster.Close()


def starte(uiapp, doc, aktive_ansicht):
    """Dialog zeigen. Rückgabe Auswahl (dict) oder None."""
    return LegendenFenster(uiapp, doc, aktive_ansicht).zeige()
