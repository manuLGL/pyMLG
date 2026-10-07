# -*- coding: utf-8 -*-
"""Dialog von DxfLegend (WPF): DXF wählen, Massstab und Präfix einstellen,
Vorschau auf weissem Papier und Liste der Stile, die angelegt werden.

    ┌ Datei: [C:\\…\\leyenda.dxf                     ] [Andere Datei…] ┐
    │ 42 Linien · 9 Texte · 3 Füllungen · Einheiten m · Text 0.25      │
    ├ Vorschau ─────────────────────────────┬ Stile ──────────────────┤
    │  ┌───────────────────────────────┐    │ Linienstile             │
    │  │ LEYENDA …                     │    │ ── LEY_FF7F00_DASHED_4  │
    │  │ - - -  RED SANEAMIENTO …      │    │ Texttypen               │
    │  └───────────────────────────────┘    │ Aa LEY_2.5mm_Arial_…    │
    ├ Einstellungen ────────────────────────┴─────────────────────────┤
    │ Name [LEYENDA DE SANEAMIENTO]  Ansicht [Legende ▾]  1:[50]      │
    │ Texthöhe [2.5] mm   Breite [180] mm   Präfix [LEY] □ abdunkeln  │
    └─────────────────────────────────────────────────────────────────┘
                                          [Legende erstellen] [Abbrechen]

Massstab: Texthöhe und Breite hängen zusammen (beide ändern den Faktor
Papier-mm je Zeichnungseinheit). Ohne Texte ist nur die Breite änderbar.
"""

import io
import json
import os
import traceback

from System.Windows import FontWeights, Point, Size, Thickness, VerticalAlignment
from System.Windows.Controls import Canvas, Orientation, StackPanel, TextBlock
from System.Windows.Media import (
    Color,
    DoubleCollection,
    FillRule,
    FontFamily,
    PathFigure,
    PathGeometry,
    PointCollection,
    PolyLineSegment,
    RotateTransform,
    ScaleTransform,
    SolidColorBrush,
    TransformGroup,
)
from System.Windows.Shapes import Line as WpfLinie
from System.Windows.Shapes import Path as WpfPfad
from System.Windows.Shapes import Polyline, Rectangle

from dxf_legende import dxf
from dxf_legende import geometrie as geo
from dxf_legende import logik as lg
from dxf_legende import revit as rv
from filter_manager import dialoge as dlg
from mlg_sprache import t, tt, uebersetze_xaml

TITEL = t(u"Legende aus DXF", u"Legend from DXF", u"Leyenda desde DXF")

FEHLERPROTOKOLL = os.path.join(dlg.protokollordner(), "DxfLegend_Fehler.log")
EINSTELLUNGEN = os.path.join(dlg.protokollordner(), "DxfLegend.json")

RAND_PX = 12.0
# Versalhöhe in Schriftgrössen (Arial) - für die Textgrösse der Vorschau
VERSAL = 0.72

ART_LEGENDE = 0
ART_ZEICHNUNG = 1

EINHEITEN = {0: (u"ohne", u"unitless", u"sin unidad"), 1: (u"Zoll", u"in", u"pulg"),
             2: (u"Fuss", u"ft", u"pies"), 4: (u"mm", u"mm", u"mm"),
             5: (u"cm", u"cm", u"cm"), 6: (u"m", u"m", u"m")}

XAML_TEXTE = {
    "titel": (u"Legende aus DXF (pyMLG)", u"Legend from DXF (pyMLG)",
              u"Leyenda desde DXF (pyMLG)"),
    "datei": (u"Datei:", u"File:", u"Archivo:"),
    "andere": (u"Andere Datei…", u"Other file…", u"Otro archivo…"),
    "vorschau": (u"Vorschau (weisses Papier)", u"Preview (white paper)",
                 u"Vista previa (papel blanco)"),
    "stile": (u"Stile und Typen", u"Styles and types", u"Estilos y tipos"),
    "einstellungen": (u"Einstellungen", u"Settings", u"Ajustes"),
    "name": (u"Name:", u"Name:", u"Nombre:"),
    "ansicht": (u"Ansicht:", u"View:", u"Vista:"),
    "massstab": (u"Massstab 1:", u"Scale 1:", u"Escala 1:"),
    "texthoehe": (u"Texthöhe auf dem Plan (mm):", u"Text height on sheet (mm):",
                  u"Altura de texto en plano (mm):"),
    "texthoehe_tip": (u"Die häufigste Texthöhe der DXF wird so gross. "
                      u"Alles andere skaliert mit.",
                      u"The most common text height of the DXF gets this "
                      u"size. Everything else scales along.",
                      u"La altura de texto más frecuente del DXF tendrá "
                      u"este tamaño. Todo lo demás escala igual."),
    "breite": (u"Breite (mm):", u"Width (mm):", u"Ancho (mm):"),
    "praefix": (u"Präfix der Stile:", u"Style prefix:", u"Prefijo de estilos:"),
    "praefix_tip": (u"Neue Linienstile, Texttypen und Füllbereichstypen "
                    u"beginnen damit, z.B. LEY_FF7F00_DASHED_4",
                    u"New line styles, text types and filled region types "
                    u"start with it, e.g. LEY_FF7F00_DASHED_4",
                    u"Los nuevos estilos de línea, tipos de texto y de "
                    u"región empiezan así, p.ej. LEY_FF7F00_DASHED_4"),
    "abdunkeln": (u"Helle Farben abdunkeln", u"Darken light colours",
                  u"Oscurecer colores claros"),
    "abdunkeln_tip": (u"Gelb, Cyan & Co. sind auf weissem Papier kaum lesbar. "
                      u"Weiss wird immer schwarz.",
                      u"Yellow, cyan & co. are hard to read on white paper. "
                      u"White always becomes black.",
                      u"Amarillo, cian, etc. apenas se leen en papel blanco. "
                      u"El blanco siempre pasa a negro."),
    "ersetzen": (u"Gleichnamige Legende ersetzen", u"Replace legend with the same name",
                 u"Reemplazar leyenda del mismo nombre"),
    "ersetzen_tip": (u"Gibt es schon eine Legende/Zeichnungsansicht mit dem Namen, wird "
                     u"ihr gesamter Inhalt gelöscht und neu gezeichnet. Auf Plänen "
                     u"bleibt sie an ihrem Platz. Ohne Haken entsteht \"Name (2)\".",
                     u"If a legend/drafting view with this name exists, all its content "
                     u"is deleted and redrawn. It stays in place on sheets. Unticked, "
                     u"\"Name (2)\" is created.",
                     u"Si ya existe una leyenda/vista de diseño con ese nombre, se borra "
                     u"todo su contenido y se vuelve a dibujar. En los planos se queda "
                     u"en su sitio. Sin marcar se crea \"Nombre (2)\"."),
    "ueberschreiben": (u"Vorhandene Stile überschreiben", u"Overwrite existing styles",
                       u"Sobrescribir estilos existentes"),
    "ueberschreiben_tip": (u"Gleichnamige Linienstile, Texttypen und Füllbereichstypen "
                           u"bekommen die Werte aus der DXF, statt _2, _3 … anzulegen. "
                           u"Wirkt überall, wo sie schon benutzt werden.",
                           u"Line styles, text types and filled region types with the "
                           u"same name get the values from the DXF instead of creating "
                           u"_2, _3 … Affects every place they are already used.",
                           u"Los estilos de línea, tipos de texto y de región con el mismo "
                           u"nombre toman los valores del DXF en vez de crear _2, _3 … "
                           u"Afecta a todos los lugares donde ya se usan."),
    "erstellen": (u"Legende erstellen", u"Create legend", u"Crear leyenda"),
    "abbrechen": (u"Abbrechen", u"Cancel", u"Cancelar"),
}

XAML = u"""
<Window %s Title="{{titel}}" Width="1100" Height="740"
        MinWidth="820" MinHeight="540" WindowStartupLocation="CenterOwner"
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
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <DockPanel>
      <TextBlock Text="{{datei}}" VerticalAlignment="Center" Margin="0,0,6,0"/>
      <Button x:Name="durchsuchen" DockPanel.Dock="Right" Content="{{andere}}"
              Padding="10,3" Margin="6,0,0,0"/>
      <TextBox x:Name="pfad" IsReadOnly="True"/>
    </DockPanel>
    <TextBlock x:Name="info" Grid.Row="1" Foreground="#555" Margin="0,6,0,6"
               TextWrapping="Wrap"/>

    <Grid Grid.Row="2">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="320"/>
      </Grid.ColumnDefinitions>
      <GroupBox Header="{{vorschau}}" Margin="0,0,8,0">
        <Border Background="White" BorderBrush="#C8CDD3" BorderThickness="1">
          <Canvas x:Name="leinwand" ClipToBounds="True" Background="White"/>
        </Border>
      </GroupBox>
      <GroupBox Header="{{stile}}" Grid.Column="1">
        <ListBox x:Name="stilliste" BorderThickness="0"
                 ScrollViewer.HorizontalScrollBarVisibility="Auto"/>
      </GroupBox>
    </Grid>

    <GroupBox Header="{{einstellungen}}" Grid.Row="3" Margin="0,8,0,0">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="160"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="90"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Text="{{name}}" VerticalAlignment="Center" Margin="0,0,6,0"/>
        <TextBox x:Name="name" Grid.Column="1"/>
        <TextBlock Text="{{ansicht}}" Grid.Column="2" VerticalAlignment="Center"
                   Margin="14,0,6,0"/>
        <ComboBox x:Name="art" Grid.Column="3"/>
        <TextBlock Text="{{massstab}}" Grid.Column="4" VerticalAlignment="Center"
                   Margin="14,0,6,0"/>
        <TextBox x:Name="massstab" Grid.Column="5"/>

        <TextBlock Grid.Row="1" Text="{{texthoehe}}" VerticalAlignment="Center"
                   Margin="0,6,6,0"/>
        <StackPanel Grid.Row="1" Grid.Column="1" Orientation="Horizontal"
                    Margin="0,6,0,0">
          <TextBox x:Name="texthoehe" Width="70" ToolTip="{{texthoehe_tip}}"/>
          <TextBlock Text="{{breite}}" VerticalAlignment="Center"
                     Margin="14,0,6,0"/>
          <TextBox x:Name="breite" Width="70"/>
        </StackPanel>
        <TextBlock Grid.Row="1" Grid.Column="2" Text="{{praefix}}"
                   VerticalAlignment="Center" Margin="14,6,6,0"/>
        <TextBox x:Name="praefix" Grid.Row="1" Grid.Column="3" Margin="0,6,0,0"
                 ToolTip="{{praefix_tip}}"/>
        <CheckBox x:Name="abdunkeln" Grid.Row="1" Grid.Column="4"
                  Grid.ColumnSpan="2" Margin="14,6,0,0"
                  VerticalAlignment="Center" Content="{{abdunkeln}}"
                  ToolTip="{{abdunkeln_tip}}"/>
        <StackPanel Grid.Row="2" Grid.ColumnSpan="6" Orientation="Horizontal"
                    Margin="0,8,0,0">
          <CheckBox x:Name="ersetzen" Content="{{ersetzen}}"
                    ToolTip="{{ersetzen_tip}}" VerticalAlignment="Center"/>
          <CheckBox x:Name="ueberschreiben" Content="{{ueberschreiben}}"
                    ToolTip="{{ueberschreiben_tip}}" Margin="24,0,0,0"
                    VerticalAlignment="Center"/>
        </StackPanel>
        <TextBlock x:Name="ansichthinweis" Grid.Row="3" Grid.ColumnSpan="6"
                   Foreground="#A05A00" TextWrapping="Wrap" Margin="0,6,0,0"/>
      </Grid>
    </GroupBox>

    <DockPanel Grid.Row="4" Margin="0,10,0,0">
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
        <Button x:Name="ok" Content="{{erstellen}}" Padding="14,3"
                Margin="0,0,6,0" IsDefault="True"/>
        <Button x:Name="abbrechen" Content="{{abbrechen}}" Width="90"
                IsCancel="True"/>
      </StackPanel>
      <TextBlock x:Name="hinweis" Foreground="#A05A00" VerticalAlignment="Center"
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


def frage_datei(startordner=None):
    """DXF-Datei per Dateidialog wählen. Rückgabe Pfad oder None."""
    from Microsoft.Win32 import OpenFileDialog
    dialog = OpenFileDialog()
    dialog.Title = t(u"DXF-Datei der Legende wählen", u"Select the DXF file of the legend",
                     u"Seleccione el archivo DXF de la leyenda")
    dialog.Filter = t(u"DXF-Dateien (*.dxf)|*.dxf", u"DXF files (*.dxf)|*.dxf",
                      u"Archivos DXF (*.dxf)|*.dxf")
    if startordner and os.path.isdir(startordner):
        dialog.InitialDirectory = startordner
    if dialog.ShowDialog() == True:  # noqa: E712 - Nullable<bool> aus .NET
        return dialog.FileName
    return None


def lade_zeichnung(pfad):
    """DXF lesen; fachliche Probleme als EingabeFehler."""
    try:
        zeichnung = dxf.lese_datei(pfad)
    except dxf.DxfFehler as fehler:
        if u"%s" % fehler == u"binary":
            raise EingabeFehler(t(
                u"Die Datei ist ein Binär-DXF. Bitte in AutoCAD als ASCII-DXF "
                u"speichern (DXFOUT, ohne Option \"Binär\").",
                u"The file is a binary DXF. Please save it as ASCII DXF in "
                u"AutoCAD (DXFOUT without the \"Binary\" option).",
                u"El archivo es un DXF binario. Guárdelo como DXF ASCII en "
                u"AutoCAD (DXFOUT sin la opción \"Binario\")."))
        raise EingabeFehler(t(u"Die Datei ist kein lesbares DXF (%s).",
                              u"The file is not a readable DXF (%s).",
                              u"El archivo no es un DXF legible (%s).") % fehler)
    except (IOError, OSError) as fehler:
        raise EingabeFehler(t(u"Die Datei lässt sich nicht öffnen:\n%s",
                              u"The file cannot be opened:\n%s",
                              u"No se puede abrir el archivo:\n%s") % fehler)
    if zeichnung.ist_leer():
        raise EingabeFehler(t(u"In der DXF-Datei wurden keine Linien, Texte oder "
                              u"Schraffuren gefunden.",
                              u"No lines, texts or hatches were found in the DXF file.",
                              u"No se encontraron líneas, textos ni sombreados en "
                              u"el archivo DXF."))
    return zeichnung


def _pinsel(rgb, alpha=255):
    pinsel = SolidColorBrush(Color.FromArgb(alpha, int(rgb[0]), int(rgb[1]), int(rgb[2])))
    pinsel.Freeze()
    return pinsel


def _zahl(text):
    try:
        return float((text or u"").strip().replace(u",", u"."))
    except ValueError:
        return None


class LegendenFenster(object):

    def __init__(self, uiapp, doc, pfad, zeichnung):
        self.doc = doc
        self.einstellungen = lies_einstellungen()
        self.legenden = rv.legenden(doc)
        self.fenster = dlg.lade_xaml(uebersetze_xaml(XAML, XAML_TEXTE))
        dlg.setze_besitzer(self.fenster, handle=uiapp.MainWindowHandle)
        self.c = self.fenster.FindName
        self.ergebnis = None
        self.plan = None
        self._still = False
        self._verdrahte()
        self._setze_zeichnung(pfad, zeichnung)

    def _h(self, funktion):
        return sicher(lambda: self.fenster, funktion)

    def _verdrahte(self):
        c = self.c
        art = c("art")
        art.Items.Add(t(u"Legende", u"Legend", u"Leyenda"))
        art.Items.Add(t(u"Zeichnungsansicht", u"Drafting view", u"Vista de diseño"))
        if self.legenden:
            art.SelectedIndex = ART_LEGENDE
            massstab = self.legenden[0].Scale
        else:
            art.SelectedIndex = ART_ZEICHNUNG
            massstab = 50
        c("massstab").Text = u"%d" % self.einstellungen.get(u"massstab", massstab)
        c("praefix").Text = self.einstellungen.get(u"praefix", u"LEY")
        c("abdunkeln").IsChecked = bool(self.einstellungen.get(u"abdunkeln", False))
        self._ansichthinweis()

        c("durchsuchen").Click += self._h(self._andere_datei)
        c("texthoehe").TextChanged += self._h(self._texthoehe_geaendert)
        c("breite").TextChanged += self._h(self._breite_geaendert)
        c("praefix").TextChanged += self._h(lambda s, a: self._aktualisiere())
        c("abdunkeln").Click += self._h(lambda s, a: self._aktualisiere())
        art.SelectionChanged += self._h(lambda s, a: self._ansichthinweis())
        c("name").TextChanged += self._h(lambda s, a: self._ansichthinweis())
        c("ersetzen").Click += self._h(lambda s, a: self._ansichthinweis())
        c("leinwand").SizeChanged += self._h(lambda s, a: self._zeichne())
        c("ok").Click += self._h(self._ok)

    def zeige(self):
        self.fenster.ShowDialog()
        return self.ergebnis

    # -- Datei -------------------------------------------------------------------

    def _andere_datei(self, sender, args):
        pfad = frage_datei(os.path.dirname(self.pfad))
        if not pfad:
            return
        self._setze_zeichnung(pfad, lade_zeichnung(pfad))

    def _setze_zeichnung(self, pfad, zeichnung):
        self.pfad = pfad
        self.zeichnung = zeichnung
        self.h_ref = lg.typische_texthoehe(zeichnung.texte)
        rahmen = lg.grenzen(zeichnung)
        self.breite_cad = (rahmen[2] - rahmen[0]) if rahmen else 0.0
        self.faktor, ziel = lg.vorschlag(zeichnung)
        c = self.c
        c("pfad").Text = pfad
        stamm = os.path.splitext(os.path.basename(pfad))[0]
        c("name").Text = lg.vorschlag_name(zeichnung, stamm)
        c("texthoehe").IsEnabled = self.h_ref is not None
        self._setze_felder()
        c("info").Text = self._infotext()
        self._aktualisiere()

    def _infotext(self):
        z = self.zeichnung
        teile = [t(u"%d Linien", u"%d lines", u"%d líneas") % len(z.kurven),
                 t(u"%d Texte", u"%d texts", u"%d textos") % len(z.texte),
                 t(u"%d Füllungen", u"%d fills", u"%d rellenos") % len(z.flaechen)]
        einheit = EINHEITEN.get(z.einheiten)
        if einheit:
            teile.append(t(u"Einheiten: %s", u"Units: %s", u"Unidades: %s") % tt(einheit))
        if self.h_ref:
            teile.append(t(u"häufigste Texthöhe %s", u"most common text height %s",
                           u"altura de texto más frecuente %s") % lg.zahl_text(self.h_ref, 4))
        text = u" · ".join(teile)
        if z.ignoriert:
            text += u"\n" + t(u"Nicht übernommen: ", u"Not imported: ", u"No importado: ") + \
                u", ".join(u"%s (%d)" % (typ, anzahl)
                           for typ, anzahl in sorted(z.ignoriert.items()))
        return text

    # -- Massstab ----------------------------------------------------------------

    def _setze_felder(self):
        self._still = True
        try:
            if self.h_ref:
                self.c("texthoehe").Text = lg.zahl_text(self.h_ref * self.faktor)
            else:
                self.c("texthoehe").Text = u""
            self.c("breite").Text = lg.zahl_text(self.breite_cad * self.faktor, 1)
        finally:
            self._still = False

    def _texthoehe_geaendert(self, sender, args):
        if self._still or not self.h_ref:
            return
        wert = _zahl(self.c("texthoehe").Text)
        if not wert or wert <= 0:
            return
        self.faktor = wert / self.h_ref
        self._still = True
        try:
            self.c("breite").Text = lg.zahl_text(self.breite_cad * self.faktor, 1)
        finally:
            self._still = False
        self._aktualisiere()

    def _breite_geaendert(self, sender, args):
        if self._still or self.breite_cad <= 0:
            return
        wert = _zahl(self.c("breite").Text)
        if not wert or wert <= 0:
            return
        self.faktor = wert / self.breite_cad
        if self.h_ref:
            self._still = True
            try:
                self.c("texthoehe").Text = lg.zahl_text(self.h_ref * self.faktor)
            finally:
                self._still = False
        self._aktualisiere()

    def _ansichthinweis(self):
        hinweis = u""
        name = (self.c("name").Text or u"").strip()
        vorhanden = rv.vorhandene_ansicht(self.doc, name) if name else None
        if vorhanden is not None:
            if self.c("ersetzen").IsChecked:
                self.c("ansichthinweis").Text = t(
                    u"\"%s\" wird geleert und neu gezeichnet.",
                    u"\"%s\" will be cleared and redrawn.",
                    u"\"%s\" se vaciará y se volverá a dibujar.") % name
            else:
                self.c("ansichthinweis").Text = t(
                    u"\"%s\" gibt es schon - es entsteht eine neue Ansicht \"%s (2)\". "
                    u"Zum Aktualisieren \"Gleichnamige Legende ersetzen\" anhaken.",
                    u"\"%s\" already exists - a new view \"%s (2)\" is created. "
                    u"To update it tick \"Replace legend with the same name\".",
                    u"\"%s\" ya existe - se crea una vista nueva \"%s (2)\". "
                    u"Para actualizarla marque \"Reemplazar leyenda del mismo nombre\".")                     % (name, name)
            return
        if not self.legenden:
            hinweis = t(u"Im Projekt gibt es noch keine Legende - per API lässt sich "
                        u"nur eine vorhandene duplizieren. Für eine echte Legende "
                        u"einmal Ansicht > Legenden > Legende anlegen, sonst wird eine "
                        u"Zeichnungsansicht erstellt.",
                        u"The project has no legend yet - the API can only duplicate "
                        u"an existing one. For a real legend create one once via View > "
                        u"Legends > Legend, otherwise a drafting view is created.",
                        u"El proyecto aún no tiene ninguna leyenda - la API solo puede "
                        u"duplicar una existente. Para una leyenda real cree una vez "
                        u"Vista > Leyendas > Leyenda; si no, se crea una vista de diseño.")
            if self.c("art").SelectedIndex == ART_LEGENDE:
                self.c("art").SelectedIndex = ART_ZEICHNUNG
        self.c("ansichthinweis").Text = hinweis

    # -- Plan und Vorschau -------------------------------------------------------

    def _aktualisiere(self):
        self.plan = lg.erstelle_plan(self.zeichnung, self.faktor,
                                     self.c("praefix").Text,
                                     bool(self.c("abdunkeln").IsChecked))
        self._fuelle_stile()
        self._zeichne()
        self.c("hinweis").Text = self._hinweis()

    def _hinweis(self):
        plan = self.plan
        klein = [typ for typ in plan.texttypen.values() if typ.hoehe_mm < 1.0]
        if klein:
            return t(u"Achtung: Texte kleiner als 1 mm - Massstab prüfen.",
                     u"Caution: texts smaller than 1 mm - check the scale.",
                     u"Atención: textos menores de 1 mm - revise la escala.")
        if plan.breite_mm > 1200:
            return t(u"Achtung: Die Legende wird über 1,2 m breit - Massstab prüfen.",
                     u"Caution: the legend gets wider than 1.2 m - check the scale.",
                     u"Atención: la leyenda supera 1,2 m de ancho - revise la escala.")
        return u""

    def _fuelle_stile(self):
        liste = self.c("stilliste")
        liste.Items.Clear()
        plan = self.plan

        def zaehle(eintraege):
            anzahl = {}
            for schluessel in eintraege:
                anzahl[schluessel] = anzahl.get(schluessel, 0) + 1
            return anzahl

        linien = zaehle(s for s, _p, _g in plan.linien)
        texte = zaehle(e.typ for e in plan.texte)
        flaechen = zaehle(f for f, _s in plan.flaechen)

        if plan.stile:
            liste.Items.Add(self._ueberschrift(t(u"Linienstile", u"Line styles",
                                                 u"Estilos de línea")))
            for name in sorted(plan.stile):
                liste.Items.Add(self._stilzeile(plan.stile[name], linien.get(name, 0)))
        if plan.texttypen:
            liste.Items.Add(self._ueberschrift(t(u"Texttypen", u"Text types",
                                                 u"Tipos de texto")))
            for name in sorted(plan.texttypen):
                liste.Items.Add(self._textzeile(plan.texttypen[name], texte.get(name, 0)))
        if plan.fuelltypen:
            liste.Items.Add(self._ueberschrift(t(u"Füllbereichstypen", u"Filled region types",
                                                 u"Tipos de región rellenada")))
            for name in sorted(plan.fuelltypen):
                liste.Items.Add(self._fuellzeile(plan.fuelltypen[name], flaechen.get(name, 0)))

    @staticmethod
    def _ueberschrift(text):
        block = TextBlock()
        block.Text = text
        block.FontWeight = FontWeights.SemiBold
        block.Margin = Thickness(0, 6, 0, 2)
        return block

    @staticmethod
    def _zeile(symbol, text):
        zeile = StackPanel()
        zeile.Orientation = Orientation.Horizontal
        symbol.Margin = Thickness(0, 0, 8, 0)
        symbol.VerticalAlignment = VerticalAlignment.Center
        zeile.Children.Add(symbol)
        block = TextBlock()
        block.Text = text
        block.VerticalAlignment = VerticalAlignment.Center
        zeile.Children.Add(block)
        return zeile

    def _stilzeile(self, stil, anzahl):
        feld = Canvas()
        feld.Width, feld.Height = 40.0, 12.0
        linie = WpfLinie()
        linie.X1, linie.Y1, linie.X2, linie.Y2 = 0.0, 6.0, 40.0, 6.0
        linie.Stroke = _pinsel(stil.farbe)
        dicke = max(1.0, min(5.0, lg.STIFTE_MM[stil.stift - 1] * 4.0))
        linie.StrokeThickness = dicke
        if stil.muster:
            striche = DoubleCollection()
            gesamt = sum(mm for _a, mm in stil.muster) or 1.0
            # Muster auf ~2 Wiederholungen im Feld strecken
            massstab = 20.0 / gesamt
            for _art, mm in stil.muster:
                striche.Add(max(1.0, mm * massstab) / dicke)
            linie.StrokeDashArray = striche
        feld.Children.Add(linie)
        text = u"%s  ·  %s %d  (%d)" % (stil.name, t(u"Stift", u"pen", u"pluma"),
                                        stil.stift, anzahl)
        return self._zeile(feld, text)

    def _textzeile(self, typ, anzahl):
        probe = TextBlock()
        probe.Text = u"Aa"
        probe.Width = 40.0
        probe.FontFamily = FontFamily(typ.schrift)
        probe.FontSize = 14.0
        probe.Foreground = _pinsel(typ.farbe)
        return self._zeile(probe, u"%s  (%d)" % (typ.name, anzahl))

    def _fuellzeile(self, typ, anzahl):
        feld = Rectangle()
        feld.Width, feld.Height = 40.0, 12.0
        feld.Fill = _pinsel(typ.farbe, 255 if typ.muster_name is None else 110)
        feld.Stroke = _pinsel((150, 150, 150))
        feld.StrokeThickness = 0.5
        return self._zeile(feld, u"%s  (%d)" % (typ.name, anzahl))

    def _zeichne(self):
        leinwand = self.c("leinwand")
        leinwand.Children.Clear()
        plan = self.plan
        breite, hoehe = leinwand.ActualWidth, leinwand.ActualHeight
        if plan is None or breite <= 2 * RAND_PX or hoehe <= 2 * RAND_PX:
            return
        if plan.breite_mm <= 0 and plan.hoehe_mm <= 0:
            return
        s = min((breite - 2 * RAND_PX) / max(plan.breite_mm, 1e-6),
                (hoehe - 2 * RAND_PX) / max(plan.hoehe_mm, 1e-6))

        def punkt(x, y):
            return Point(RAND_PX + x * s, RAND_PX + (plan.hoehe_mm - y) * s)

        for typ_name, schleifen in plan.flaechen:
            typ = plan.fuelltypen[typ_name]
            geometrie = PathGeometry()
            geometrie.FillRule = FillRule.EvenOdd
            for schleife in schleifen:
                zug = geo.abtasten(schleife, True)
                if len(zug) < 3:
                    continue
                figur = PathFigure()
                figur.StartPoint = punkt(*zug[0])
                figur.IsClosed = True
                abschnitt = PolyLineSegment()
                punkte = PointCollection()
                for x, y in zug[1:]:
                    punkte.Add(punkt(x, y))
                abschnitt.Points = punkte
                figur.Segments.Add(abschnitt)
                geometrie.Figures.Add(figur)
            pfad = WpfPfad()
            pfad.Data = geometrie
            pfad.Fill = _pinsel(typ.farbe, 255 if typ.muster_name is None else 110)
            leinwand.Children.Add(pfad)

        for stil_name, punkte_mm, geschlossen in plan.linien:
            stil = plan.stile[stil_name]
            linie = Polyline()
            punkte = PointCollection()
            for x, y in geo.abtasten(punkte_mm, geschlossen):
                punkte.Add(punkt(x, y))
            linie.Points = punkte
            linie.Stroke = _pinsel(stil.farbe)
            dicke = max(0.7, lg.STIFTE_MM[stil.stift - 1] * s)
            linie.StrokeThickness = dicke
            if stil.muster:
                striche = DoubleCollection()
                for _art, mm in stil.muster:
                    striche.Add(max(0.8, mm * s) / dicke)
                linie.StrokeDashArray = striche
            leinwand.Children.Add(linie)

        for eintrag in plan.texte:
            typ = plan.texttypen[eintrag.typ]
            block = TextBlock()
            block.Text = eintrag.inhalt
            block.FontFamily = FontFamily(typ.schrift)
            block.FontSize = max(1.0, typ.hoehe_mm * s / VERSAL)
            block.Foreground = _pinsel(typ.farbe)
            block.Measure(Size(float("inf"), float("inf")))
            w, h = block.DesiredSize.Width, block.DesiredSize.Height
            dx = {u"links": 0.0, u"mitte": w / 2.0, u"rechts": w}[eintrag.h_ausr]
            dy = {u"oben": 0.0, u"mitte": h / 2.0, u"unten": h}[eintrag.v_ausr]
            anker = punkt(eintrag.x, eintrag.y)
            Canvas.SetLeft(block, anker.X - dx)
            Canvas.SetTop(block, anker.Y - dy)
            gruppe = TransformGroup()
            if abs(typ.breitenfaktor - 1.0) > 0.005:
                gruppe.Children.Add(ScaleTransform(typ.breitenfaktor, 1.0, dx, dy))
            if abs(eintrag.drehung) > 1e-6:
                gruppe.Children.Add(RotateTransform(-eintrag.drehung * 57.29577951308232,
                                                    dx, dy))
            if gruppe.Children.Count:
                block.RenderTransform = gruppe
            leinwand.Children.Add(block)

    # -- Abschluss ---------------------------------------------------------------

    def _ok(self, sender, args):
        name = (self.c("name").Text or u"").strip()
        if not name:
            raise EingabeFehler(t(u"Bitte einen Namen für die Legende eingeben.",
                                  u"Please enter a name for the legend.",
                                  u"Introduzca un nombre para la leyenda."))
        massstab = _zahl(self.c("massstab").Text)
        if massstab is None or not 1 <= massstab <= 100000 or massstab != int(massstab):
            raise EingabeFehler(t(u"Der Massstab muss eine ganze Zahl sein, z.B. 50.",
                                  u"The scale must be a whole number, e.g. 50.",
                                  u"La escala debe ser un número entero, p.ej. 50."))
        if not (self.plan.linien or self.plan.texte or self.plan.flaechen):
            raise EingabeFehler(t(u"Es gibt nichts zu übernehmen.", u"There is nothing to import.",
                                  u"No hay nada que importar."))
        vorlage = None
        if self.c("art").SelectedIndex == ART_LEGENDE and self.legenden:
            vorlage = self.legenden[0]
        praefix = (self.c("praefix").Text or u"").strip() or u"LEY"
        abdunkeln = bool(self.c("abdunkeln").IsChecked)
        self.ergebnis = {u"plan": self.plan, u"name": name, u"massstab": int(massstab),
                         u"vorlage": vorlage, u"pfad": self.pfad,
                         u"ersetzen": bool(self.c("ersetzen").IsChecked),
                         u"ueberschreiben": bool(self.c("ueberschreiben").IsChecked)}
        schreibe_einstellungen({u"ordner": os.path.dirname(self.pfad), u"praefix": praefix,
                                u"abdunkeln": abdunkeln, u"massstab": int(massstab)})
        self.fenster.Close()


def starte(uiapp, doc):
    """Datei wählen, Dialog zeigen. Rückgabe Auswahl (dict) oder None."""
    einstellungen = lies_einstellungen()
    pfad = frage_datei(einstellungen.get(u"ordner"))
    if not pfad:
        return None
    try:
        zeichnung = lade_zeichnung(pfad)
    except EingabeFehler as fehler:
        meldung(None, u"%s" % fehler, warnung=True)
        return None
    return LegendenFenster(uiapp, doc, pfad, zeichnung).zeige()
