# -*- coding: utf-8 -*-
"""Fenster von LinkedIds (WPF).

    ┌ 2 Elemente aus 1 Verknüpfung ───────────────────────┐
    │ Rohre · Rundrohr: Standard                   398254 │
    │ TGA.rvt  (TGA.rvt : Position 1)                     │
    │ Wände · Basiswand: AW 36.5                    12034 │
    │ Aktuelles Modell                                    │
    └─────────────────────────────────────────────────────┘
      Strg+C kopiert die markierten IDs
      [IDs kopieren] [Tabelle kopieren] [Weitere wählen] [Schließen]

Das Fenster ist modal - solange es offen ist, nimmt Revit keine Auswahl an.
"Weitere wählen" schliesst es deshalb, lässt picken und öffnet es wieder mit
den gesammelten Einträgen.
"""

import io
import os
import traceback

import clr

clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Windows import Clipboard, FontWeights, GridLength, \
    GridUnitType, HorizontalAlignment, TextWrapping, Thickness, \
    VerticalAlignment  # noqa: E402
from System.Windows.Controls import ColumnDefinition, Grid, ListBoxItem, \
    StackPanel, TextBlock  # noqa: E402
from System.Windows.Input import Key, Keyboard, ModifierKeys  # noqa: E402
from System.Windows.Media import Color, SolidColorBrush  # noqa: E402

from filter_manager import dialoge as dlg  # noqa: E402
from linked_ids import logik as lg  # noqa: E402
from linked_ids import revit as rv  # noqa: E402
from mlg_sprache import t, uebersetze_xaml  # noqa: E402

TITEL = u"Linked IDs"

# Rückgabe von zeige(): der Benutzer möchte weitere Elemente wählen
WAEHLEN = "waehlen"

FEHLERPROTOKOLL = os.path.join(dlg.protokollordner(), "LinkedIds_Fehler.log")


def _pinsel(r, g, b):
    # Eingefroren, sonst gehört der Pinsel dem Thread, der das Modul geladen
    # hat (siehe filter_manager/fenster.py)
    pinsel = SolidColorBrush(Color.FromRgb(r, g, b))
    pinsel.Freeze()
    return pinsel


GRAU = _pinsel(110, 110, 110)
ROT = _pinsel(192, 57, 43)

XAML_TEXTE = {
    "titel": (u"Linked IDs (pyMLG)",
              u"Linked IDs (pyMLG)",
              u"Linked IDs (pyMLG)"),
    "hinweis": (u"Strg+C kopiert die markierten IDs, Doppelklick kopiert "
                u"eine einzelne. Ohne Markierung gilt die ganze Liste.",
                u"Ctrl+C copies the checked ids, double-click copies a "
                u"single one. Without a selection the whole list counts.",
                u"Ctrl+C copia los IDs marcados, doble clic copia uno solo. "
                u"Sin selección cuenta toda la lista."),
    "ids": (u"IDs kopieren",
            u"Copy IDs",
            u"Copiar IDs"),
    "ids_tip": (u"IDs mit Komma getrennt - so nimmt sie "
                u"\"Auswählen nach ID\" im verknüpften Modell an.",
                u"Ids separated by commas - ready for \"Select by ID\" in "
                u"the linked model.",
                u"IDs separados por comas, tal como los admite "
                u"\"Seleccionar por ID\" en el modelo vinculado."),
    "tabelle": (u"Tabelle kopieren",
                u"Copy table",
                u"Copiar tabla"),
    "tabelle_tip": (u"Dokument, Verknüpfung, Kategorie, Typ, Name, ID und "
                    u"eindeutige ID - für Excel.",
                    u"Document, link, category, type, name, id and unique "
                    u"id - for Excel.",
                    u"Documento, vínculo, categoría, tipo, nombre, ID e ID "
                    u"única, para Excel."),
    "waehlen": (u"Weitere wählen...",
                u"Pick more...",
                u"Elegir más..."),
    "schliessen": (u"Schließen",
                   u"Close",
                   u"Cerrar"),
}

XAML = u"""<Window %s
        Title="{{titel}}" Width="640" Height="480"
        MinWidth="420" MinHeight="300"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        FontFamily="Segoe UI" FontSize="12"
        ResizeMode="CanResizeWithGrip">
  <DockPanel Margin="10">
    <TextBlock x:Name="kopf" DockPanel.Dock="Top" FontSize="15"
               FontWeight="SemiBold" TextWrapping="Wrap"
               Margin="0,0,0,8"/>
    <TextBlock DockPanel.Dock="Bottom" Text="{{hinweis}}" Foreground="#6E6E6E"
               TextWrapping="Wrap" Margin="0,8,0,0"/>
    <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal"
                HorizontalAlignment="Right" Margin="0,10,0,0">
      <Button x:Name="ids" Content="{{ids}}" ToolTip="{{ids_tip}}"
              FontWeight="Bold" Padding="12,6" Margin="0,0,6,0"
              IsDefault="True"/>
      <Button x:Name="tabelle" Content="{{tabelle}}"
              ToolTip="{{tabelle_tip}}" Padding="12,6" Margin="0,0,6,0"/>
      <Button x:Name="waehlen" Content="{{waehlen}}" Padding="12,6"
              Margin="0,0,6,0"/>
      <Button x:Name="schliessen" Content="{{schliessen}}" Padding="12,6"
              IsCancel="True"/>
    </StackPanel>
    <TextBlock x:Name="status" DockPanel.Dock="Bottom" Foreground="#2E7D32"
               TextWrapping="Wrap" Margin="0,8,0,0"/>
    <ListBox x:Name="liste" SelectionMode="Extended"
             HorizontalContentAlignment="Stretch"
             ScrollViewer.VerticalScrollBarVisibility="Auto"/>
  </DockPanel>
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


def sicher(besitzer_liefern, funktion):
    """Ereignishandler mit Fehlerfang - eine Ausnahme im Handler würde
    ShowDialog() sonst wortlos beenden."""
    def handler(sender, args):
        try:
            funktion(sender, args)
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


class LinkedIdsFenster(object):

    def __init__(self, uiapp, eintraege):
        self.eintraege = eintraege
        self.aktion = None

        f = self.fenster = dlg.lade_xaml(uebersetze_xaml(XAML, XAML_TEXTE))
        try:
            dlg.setze_besitzer(f, handle=uiapp.MainWindowHandle)
        except Exception:
            pass
        for name in ("kopf", "liste", "status"):
            setattr(self, name, f.FindName(name))

        self.kopf.Text = lg.zusammenfassung(eintraege)
        self._baue_liste()
        self._verdrahte()

    # --- Aufbau -----------------------------------------------------------
    def _baue_liste(self):
        for eintrag in self.eintraege:
            self.liste.Items.Add(self._zeile(eintrag))

    def _zeile(self, eintrag):
        raster = Grid()
        for breite in (None, 110.0):
            spalte = ColumnDefinition()
            spalte.Width = (GridLength(breite) if breite
                            else GridLength(1.0, GridUnitType.Star))
            raster.ColumnDefinitions.Add(spalte)

        texte = StackPanel()
        beschreibung = TextBlock()
        beschreibung.Text = eintrag.beschreibung
        beschreibung.TextWrapping = TextWrapping.Wrap
        texte.Children.Add(beschreibung)
        herkunft = TextBlock()
        herkunft.Text = eintrag.herkunft
        herkunft.Foreground = ROT if not eintrag.geladen else GRAU
        herkunft.FontSize = 11.0
        herkunft.TextWrapping = TextWrapping.Wrap
        texte.Children.Add(herkunft)
        raster.Children.Add(texte)

        nummer = TextBlock()
        nummer.Text = u"%d" % eintrag.element_id
        nummer.FontSize = 14.0
        nummer.FontWeight = FontWeights.SemiBold
        nummer.HorizontalAlignment = HorizontalAlignment.Right
        nummer.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(nummer, 1)
        raster.Children.Add(nummer)

        item = ListBoxItem()
        item.Content = raster
        item.Padding = Thickness(4.0, 3.0, 4.0, 3.0)
        item.Tag = eintrag
        item.MouseDoubleClick += sicher(
            lambda: self.fenster,
            lambda sender, _a: self.kopiere_ids([sender.Tag]))
        return item

    def _verdrahte(self):
        f = self.fenster

        def s(funktion):
            return sicher(lambda: f, funktion)

        f.FindName("ids").Click += s(lambda _s, _a: self.kopiere_ids())
        f.FindName("tabelle").Click += s(lambda _s, _a: self.kopiere_tabelle())
        f.FindName("waehlen").Click += s(lambda _s, _a: self.waehle_weitere())
        f.FindName("schliessen").Click += s(lambda _s, _a: f.Close())
        f.PreviewKeyDown += s(self._bei_taste)

    def _bei_taste(self, _sender, args):
        if args.Key == Key.C and Keyboard.Modifiers == ModifierKeys.Control:
            self.kopiere_ids()
            args.Handled = True

    # --- Aktionen ---------------------------------------------------------
    def gewaehlt(self):
        """Markierte Einträge - ohne Markierung die ganze Liste."""
        markiert = [item.Tag for item in self.liste.SelectedItems]
        return markiert or self.eintraege

    def kopiere_ids(self, eintraege=None):
        eintraege = eintraege or self.gewaehlt()
        Clipboard.SetText(lg.ids_text(eintraege))
        self.status.Text = t(u"%d ID(s) in die Zwischenablage kopiert.",
                             u"%d id(s) copied to the clipboard.",
                             u"%d ID(s) copiados al portapapeles.") % len(
            eintraege)

    def kopiere_tabelle(self):
        eintraege = self.gewaehlt()
        Clipboard.SetText(lg.tabelle(eintraege))
        self.status.Text = t(u"%d Zeile(n) als Tabelle kopiert.",
                             u"%d row(s) copied as a table.",
                             u"%d fila(s) copiadas como tabla.") % len(
            eintraege)

    def waehle_weitere(self):
        self.aktion = WAEHLEN
        self.fenster.Close()

    def zeige(self):
        self.fenster.ShowDialog()
        return self.aktion


def starte(uiapp, uidoc):
    eintraege = rv.sammle_aus_auswahl(uidoc)
    if not eintraege:
        eintraege = rv.waehle(uidoc)
        if not eintraege:
            return
    while True:
        if LinkedIdsFenster(uiapp, eintraege).zeige() != WAEHLEN:
            return
        weitere = rv.waehle(uidoc)
        if weitere:
            lg.ergaenze(eintraege, weitere)
